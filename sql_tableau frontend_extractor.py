import os
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile

pd.options.mode.chained_assignment = None


class TableauDocument:
    """
    Parses a single Tableau workbook (.twb or .twbx) and exposes:
      - parameters   : pd.DataFrame of all Tableau Parameters
      - calculations : pd.DataFrame of all Calculated Fields with formula + impact analysis
    """

    def __init__(self, filePath: str):
        self.filePath = os.path.normpath(filePath)
        self.fileName = os.path.basename(self.filePath)
        self.workbookName = os.path.splitext(self.fileName)[0]

        self.xmlRoot = self._get_xml_root()
        self.parameters = self._get_parameters()
        self.calculations = self._get_calculations()

    # ---------- helpers ----------
    @staticmethod
    def _local_tag(tag: str) -> str:
        """Strip XML namespace prefix: '{ns}column' -> 'column'"""
        return tag.split("}")[-1] if tag else tag

    def _is_under_tag(self, element, parent_map, tag_name: str) -> bool:
        """Return True if any ancestor of element has local tag == tag_name."""
        p = parent_map.get(element)
        while p is not None:
            if self._local_tag(p.tag) == tag_name:
                return True
            p = parent_map.get(p)
        return False

    # ---------- IO ----------
    def _get_xml_root(self):
        """Parse .twb or .twbx into an XML root element (no disk write)."""
        filePath = self.filePath
        lower = filePath.lower()

        if lower.endswith(".twb"):
            return ET.parse(filePath).getroot()

        if lower.endswith(".twbx"):
            with zipfile.ZipFile(filePath, "r") as z:
                twb_names = [n for n in z.namelist() if n.lower().endswith(".twb")]
                if not twb_names:
                    raise ValueError(f"No .twb found inside twbx: {filePath}")
                with z.open(twb_names[0]) as twb_file:
                    return ET.parse(twb_file).getroot()

        raise ValueError("File must be .twb or .twbx")

    # ---------- Parameters ----------
    def _get_parameters(self):
        """Extract all Tableau Parameters into a DataFrame."""
        temp = []
        for ds in self.xmlRoot.iter():
            if self._local_tag(ds.tag) == "datasource" and ds.get("name") == "Parameters":
                for col in ds:
                    if self._local_tag(col.tag) == "column":
                        temp.append({
                            "caption": col.get("caption"),
                            "datatype": col.get("datatype"),
                            "name": col.get("name"),
                            "default_value": col.get("value"),
                            "param-domain-type": col.get("param-domain-type"),
                            "role": col.get("role"),
                            "type": col.get("type"),
                            "members": self.find_members(col)
                        })
                break
        return pd.DataFrame(temp)

    def find_members(self, parameter_col):
        """Return comma-separated list of allowed values for a list-type parameter."""
        members = []
        for child in parameter_col:
            if self._local_tag(child.tag) == "members":
                for m in child:
                    if self._local_tag(m.tag) == "member" and m.get("value") is not None:
                        members.append(m.get("value"))
        return ",".join(members)

    # ---------- Calculations ----------
    def _get_calculations(self):
        """
        Extract all Calculated Fields across all datasources (excluding Parameters).

        Strategy:
        - Iterate each datasource subtree
        - Find columns with a direct child <calculation class='tableau'>
        - Skip columns under <datasource-dependencies> to avoid duplicates
        - Replace internal identifiers with human-readable captions
        - Compute downstream impact for each field
        """
        temp = []

        for datasource in self.xmlRoot.iter():
            if self._local_tag(datasource.tag) != "datasource":
                continue
            if datasource.get("name") == "Parameters":
                continue

            ds_name = self.extract_alias_name(datasource)

            # Build parent map for ancestor lookups
            parent_map = {}
            for p in datasource.iter():
                for c in list(p):
                    parent_map[c] = p

            for col in datasource.iter():
                if self._local_tag(col.tag) != "column":
                    continue

                # Skip duplicate entries under datasource-dependencies
                if self._is_under_tag(col, parent_map, "datasource-dependencies"):
                    continue

                # Only process columns with a Tableau calculation
                calc_node = None
                for ch in list(col):
                    if self._local_tag(ch.tag) == "calculation" and ch.get("class") == "tableau":
                        calc_node = ch
                        break
                if calc_node is None:
                    continue

                temp.append({
                    "datasource": ds_name,
                    "caption": self.extract_alias_name(col),
                    "name": col.get("name"),
                    "role": col.get("role"),
                    "datatype": col.get("datatype"),
                    "type": col.get("type"),
                    "formula": calc_node.get("formula")
                })

        df = pd.DataFrame(temp)
        if len(df) == 0:
            return df

        # Per-datasource: humanize identifiers + compute impact
        out = pd.DataFrame()
        for ds in df["datasource"].dropna().unique():
            calcs_sub = df[df["datasource"] == ds].copy()
            calcs_sub = self.update_calculation_formula(calcs_sub)
            calcs_sub = self.add_impacted_fields(calcs_sub)
            out = pd.concat([out, calcs_sub], ignore_index=True)

        out = out.drop_duplicates(subset=["datasource", "name", "formula"]).reset_index(drop=True)
        return out

    def create_identifier_dict(self, df):
        """Build {internal_name: caption} lookup dict from a DataFrame."""
        if df is None or len(df) == 0:
            return {}
        df2 = df[["name", "caption"]].copy()
        df2.index = df2["name"]
        df2.drop(columns="name", inplace=True)
        return df2.to_dict()["caption"]

    def update_calculation_formula(self, calcs_sub):
        """Replace internal [Calculation_XXXXX] identifiers with human-readable captions."""
        search_dict = self.create_identifier_dict(calcs_sub)
        search_dict.update(self.create_identifier_dict(self.parameters))

        calcs_sub = calcs_sub.reset_index(drop=True)
        for i in range(len(calcs_sub)):
            cell = calcs_sub.at[i, "formula"]
            if not isinstance(cell, str):
                continue
            for key, val in search_dict.items():
                if key and key in cell and val:
                    cell = cell.replace(key, "[" + val + "]")
            calcs_sub.at[i, "formula"] = cell
        return calcs_sub

    def add_impacted_fields(self, calcs_sub):
        """
        For each calculated field, find all other fields whose formula references it.
        Result stored in the 'impacts' column as a comma-separated list.
        """
        calcs_sub = calcs_sub.reset_index(drop=True)
        calcs_sub["impacts"] = ""

        for i in range(len(calcs_sub)):
            caption_i = calcs_sub.at[i, "caption"]
            if not caption_i:
                continue
            token = "[" + str(caption_i) + "]"

            impacted = []
            for j in range(len(calcs_sub)):
                formula = calcs_sub.at[j, "formula"]
                if isinstance(formula, str) and token in formula:
                    impacted.append("[" + str(calcs_sub.at[j, "caption"]) + "]")
            calcs_sub.at[i, "impacts"] = ",".join(impacted)

        return calcs_sub

    def extract_alias_name(self, element):
        """Return caption if available, otherwise strip brackets from name."""
        cap = element.get("caption")
        if cap is None:
            val = element.get("name")
            if val is None:
                return None
            return val.replace("[", "").replace("]", "")
        return cap


def summarize_folder_to_excel(folder, out_xlsx, recursive=True, debug_print=True):
    """
    Scan a folder for Tableau workbooks and export a multi-sheet Excel summary:
      - Summary              : one row per workbook (counts of params & calc fields)
      - Parameters_Detail   : all parameters with metadata
      - Calculations_Detail : all calculated fields with formulas and impact analysis
      - Errors              : any files that failed to parse (if applicable)

    Args:
        folder      : path to folder containing .twb / .twbx files
        out_xlsx    : output Excel file path
        recursive   : if True, search subfolders recursively
        debug_print : if True, print progress per file
    """
    folder = os.path.normpath(folder)
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"Folder not found: {folder}")

    paths = []
    if recursive:
        for root, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith((".twb", ".twbx")):
                    paths.append(os.path.join(root, f))
    else:
        for f in os.listdir(folder):
            if f.lower().endswith((".twb", ".twbx")):
                paths.append(os.path.join(folder, f))

    rows_summary, rows_params, rows_calcs, rows_errors = [], [], [], []

    for p in sorted(paths):
        try:
            doc = TableauDocument(p)

            if debug_print:
                print(doc.fileName, "  params:", len(doc.parameters), "  calcs:", len(doc.calculations))

            # Collect parameter names
            param_names = []
            if len(doc.parameters) > 0:
                for _, r in doc.parameters.iterrows():
                    cap = r.get("caption")
                    nm = r.get("name")
                    param_names.append(cap if pd.notna(cap) and cap else nm)
            param_names = [str(x) for x in param_names if x and str(x) != "None"]

            # Collect calculated field names
            calc_names = []
            if len(doc.calculations) > 0 and "caption" in doc.calculations.columns:
                calc_names = [str(x) for x in doc.calculations["caption"].dropna().unique().tolist()]

            rows_summary.append({
                "Tableau Name": doc.workbookName,
                "File": doc.fileName,
                "Parameters": "; ".join(sorted(set(param_names))),
                "Calculated Fields": "; ".join(sorted(set(calc_names))),
                "Parameter Count": len(set(param_names)),
                "Calculated Field Count": len(set(calc_names)),
            })

            if len(doc.parameters) > 0:
                tmp = doc.parameters.copy()
                tmp.insert(0, "Tableau Name", doc.workbookName)
                rows_params.append(tmp)

            if len(doc.calculations) > 0:
                tmp = doc.calculations.copy()
                tmp.insert(0, "Tableau Name", doc.workbookName)
                rows_calcs.append(tmp)

        except Exception as e:
            rows_errors.append({"FilePath": p, "Error": repr(e)})

    df_summary = pd.DataFrame(rows_summary)
    df_params = pd.concat(rows_params, ignore_index=True) if rows_params else pd.DataFrame()
    df_calcs = pd.concat(rows_calcs, ignore_index=True) if rows_calcs else pd.DataFrame()
    df_errors = pd.DataFrame(rows_errors)

    with pd.ExcelWriter(out_xlsx) as writer:
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        df_params.to_excel(writer, sheet_name="Parameters_Detail", index=False)
        df_calcs.to_excel(writer, sheet_name="Calculations_Detail", index=False)
        if len(df_errors) > 0:
            df_errors.to_excel(writer, sheet_name="Errors", index=False)

    return out_xlsx


if __name__ == "__main__":
    # Update these paths before running
    out_file = summarize_folder_to_excel(
        folder=r"C:\Users\your_username\Desktop\tableau_extractor\workbooks",
        out_xlsx=r"C:\Users\your_username\Desktop\tableau_extractor\Tableau Parameters and Cal Fields.xlsx",
        recursive=True,
        debug_print=True
    )
    print("Done. Output:", out_file)
