import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import html
import pandas as pd
import re
from collections import Counter
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s:%(message)s'
)

# ====== PLEASE SET YOUR FOLDERS ======
# Example: BASE_FOLDER = Path(r'C:\Users\your_username\Desktop\tableau_extractor')
BASE_FOLDER = Path(r'C:\Users\your_username\Desktop\tableau_extractor')
OUTPUT_FOLDER = BASE_FOLDER

# Download directory: place all .twb / .twbx files here (subfolders supported)
DOWNLOAD_FOLDER = BASE_FOLDER / "Download"
DOWNLOAD_FOLDER.mkdir(parents=True, exist_ok=True)

# SQL output root: Tableau Folder/<WorkbookName>/*.sql
TABLEAU_FOLDER = OUTPUT_FOLDER / "Tableau Folder"
TABLEAU_FOLDER.mkdir(parents=True, exist_ok=True)

# Excel output directory
SUMMARY_OUTPUT_FOLDER = OUTPUT_FOLDER

# Temp folder for .twb files extracted from .twbx
UNZIP_FOLDER = BASE_FOLDER / "Unzip"
UNZIP_FOLDER.mkdir(parents=True, exist_ok=True)

# Final merged Excel output
SUMMARY_XLSX = SUMMARY_OUTPUT_FOLDER / "Tableau summary.xlsx"

# ====== PATTERNS ======
TABLES_PATTERN = r'(?:^|[\s;])(?:FROM|JOIN)\s+([^\n\r;]+)'
JOIN_PATTERN = r'\bjoin\b'

# ====== HELPERS ======
def safe_name(name):
    """Replace invalid filename characters with underscores."""
    return re.sub(r'[\\/*?:"<>|]', "_", str(name or "unknown"))

def extract_tables(sql, pattern):
    if not sql:
        return []
    return re.findall(pattern, sql, re.IGNORECASE | re.DOTALL)

def count_joins(sql):
    """Count JOIN occurrences in SQL (case-insensitive, word-boundary match)."""
    if not sql:
        return 0
    return len(re.findall(JOIN_PATTERN, sql, flags=re.IGNORECASE))

def count_sql_lines(sql):
    """Count total physical lines in a SQL string. Returns 0 for empty input."""
    if not sql or not sql.strip():
        return 0
    return len(sql.splitlines())

def make_summary(workbook_name, datasource_type, datasource_name, script_name, tables,
                 join_count=0, sql_line_count=0):
    return {
        "Workbook": workbook_name,
        "Datasource name": datasource_name,
        "Datasource type": datasource_type,
        "Script/File name": script_name,
        "Tables": tables,
        "Join count": join_count,
        "SQL line count": sql_line_count
    }

def unzip_twbx_only_twb(twbx_path: Path, extract_to: Path):
    """Extract only the .twb file(s) from a .twbx archive."""
    extract_to.mkdir(parents=True, exist_ok=True)
    extracted_paths = []

    with zipfile.ZipFile(twbx_path, "r") as z:
        members = [m for m in z.namelist() if m.lower().endswith(".twb")]
        if not members:
            logging.warning(f"No .twb found in {twbx_path}")
            return extracted_paths

        for member in members:
            z.extract(member, extract_to)
            extracted_paths.append(extract_to / member)

    logging.info(f"Extracted {len(extracted_paths)} .twb from {twbx_path.name} into {extract_to}")
    return extracted_paths

def write_extracted(sql, script_name, workbook_name, datasource_name):
    """
    Write extracted SQL to:
    OUTPUT_FOLDER/Tableau Folder/<WorkbookName>/<Datasource>__<Script>.sql
    """
    wb = safe_name(workbook_name)
    ds = safe_name(datasource_name)
    sc = safe_name(script_name)

    folder = TABLEAU_FOLDER / wb
    folder.mkdir(parents=True, exist_ok=True)

    file_path = folder / f"{ds}__{sc}.sql"
    file_path.write_text(sql, encoding="utf-8")

def clean_string(s):
    """Remove everything except letters, numbers, dots and underscores."""
    if s is None or pd.isna(s):
        return ''
    return re.sub(r'[^a-zA-Z0-9._]', '', str(s))

def tables_to_text(x):
    """Convert a list/tuple of table names to a comma-separated string for Excel output."""
    if isinstance(x, (list, tuple, set)):
        vals = [str(i).strip() for i in x if i is not None and not pd.isna(i) and str(i).strip()]
        return ", ".join(vals)
    if x is None or pd.isna(x):
        return ""
    return str(x)

# ====== 0) DISCOVER ALL TABLEAU FILES IN Download (RECURSIVE) ======
twb_in_download = list(DOWNLOAD_FOLDER.rglob("*.twb"))
twbx_in_download = list(DOWNLOAD_FOLDER.rglob("*.twbx"))

logging.info(f"Found {len(twb_in_download)} .twb in {DOWNLOAD_FOLDER}")
logging.info(f"Found {len(twbx_in_download)} .twbx in {DOWNLOAD_FOLDER}")

# ====== 1) BUILD PARSE JOBS ======
# Each job contains:
#   twb_path      - actual .twb file to parse
#   workbook_name - display name shown in Excel output
parse_jobs = []

for twb_path in twb_in_download:
    parse_jobs.append({
        "twb_path": twb_path,
        "workbook_name": twb_path.stem
    })

for twbx_path in twbx_in_download:
    rel_parent = twbx_path.parent.relative_to(DOWNLOAD_FOLDER)
    extract_to = UNZIP_FOLDER / rel_parent / twbx_path.stem
    extracted_twb_paths = unzip_twbx_only_twb(twbx_path, extract_to)

    for extracted_twb in extracted_twb_paths:
        parse_jobs.append({
            "twb_path": extracted_twb,
            "workbook_name": twbx_path.stem
        })

logging.info(f"Total .twb to parse: {len(parse_jobs)}")

# ====== 2) PARSE ======
summary = []

for job in parse_jobs:
    twb_path = job["twb_path"]
    workbook_name = job["workbook_name"]

    try:
        tree = ET.parse(twb_path)
    except Exception as e:
        logging.error(f"Failed to parse {twb_path}: {e}")
        continue

    root = tree.getroot()

    for datasource in root.findall('.//datasource'):
        datasource_name = safe_name(
            datasource.get("caption") or datasource.get("name") or "unknown_datasource"
        )

        # --- Flat file connections (Excel / CSV / cloud storage) ---
        for named_connection in datasource.findall('.//named-connection'):
            connection = named_connection.find('connection')
            if connection is None:
                continue

            if (
                connection.get('cloudFileExtension') in ('xlsx', 'csv') or
                connection.get('class') in ('excel-direct', 'textscan') or
                (connection.get('dbname') is not None and "CloudStorage" in connection.get('dbname', ''))
            ):
                if connection.get('cloudFileExtension') in ('xlsx', 'csv'):
                    script_name = safe_name(connection.get('cloudFileName'))
                elif connection.get('class') in ('excel-direct', 'textscan'):
                    script_name = safe_name(connection.get('filename'))
                else:
                    script_name = safe_name(connection.get('dbname', 'unknown_file'))

                summary.append(make_summary(
                    workbook_name=workbook_name,
                    datasource_name=datasource_name,
                    datasource_type="Excel/CSV",
                    script_name=script_name,
                    tables=[],
                    join_count=0,
                    sql_line_count=0
                ))

            # --- Database / DWH connection ---
            # Customize the keyword below to match your database name pattern
            if connection.get('dbname') is not None and "your_db" in connection.get('dbname', ''):
                script_name = safe_name(connection.get('dbname'))
                summary.append(make_summary(
                    workbook_name=workbook_name,
                    datasource_name=datasource_name,
                    datasource_type="Database table",
                    script_name=script_name,
                    tables=[],
                    join_count=0,
                    sql_line_count=0
                ))

        # --- Custom SQL relations ---
        for relation in datasource.findall('.//relation'):
            if relation.get('type') == 'text' and relation.text:
                sql = html.unescape(relation.text.strip())
                tables = extract_tables(sql, TABLES_PATTERN)
                script_name = safe_name(relation.get("name") or "custom_sql")

                join_count = count_joins(sql)
                sql_line_count = count_sql_lines(sql)

                write_extracted(sql, script_name, workbook_name, datasource_name)

                summary.append(make_summary(
                    workbook_name=workbook_name,
                    datasource_name=datasource_name,
                    datasource_type='Custom SQL',
                    script_name=script_name,
                    tables=tables,
                    join_count=join_count,
                    sql_line_count=sql_line_count
                ))

# ====== 3) BUILD DATAFRAMES ======
df = pd.DataFrame(summary)

if df.empty:
    raise SystemExit(
        "No data extracted. Please check: (1) Download folder contains .twb/.twbx files, "
        "(2) .twbx archives include a .twb, (3) workbooks have datasources/relations."
    )

# ---- Sheet 1: Table frequency (Custom SQL only) ----
df_res = df[df['Datasource type'] == 'Custom SQL'].copy()

df_res['Tables'] = df_res['Tables'].apply(
    lambda x: tuple(x) if isinstance(x, (list, tuple, set)) else tuple()
)
df_res = df_res.drop_duplicates()

df_tables = df_res['Tables'].explode().dropna()
df_tables = df_tables.astype(str).str.strip()
df_tables_clean = df_tables.apply(clean_string)
df_tables_clean = df_tables_clean[df_tables_clean != '']

table_counts = Counter(df_tables_clean.tolist())

df_data_table = pd.DataFrame(table_counts.items(), columns=['Table', 'Frequency'])
df_data_table = df_data_table.sort_values('Frequency', ascending=False).reset_index(drop=True)

# ---- Sheet 2: Workbook-level summary ----
def summarize_workbook(group):
    num_excel = (group['Datasource type'] == 'Excel/CSV').sum()
    num_db = (group['Datasource type'] == 'Database table').sum()

    num_custom_sql = (
        group[group['Datasource type'] == 'Custom SQL'][['Datasource name', 'Script/File name']]
        .drop_duplicates()
        .shape[0]
    )

    tables_series = group['Tables']
    exploded_tables = tables_series.explode().dropna()
    exploded_tables = exploded_tables.astype(str).str.strip()
    exploded_tables = exploded_tables[exploded_tables != '']

    # Customize these patterns to match your own table naming conventions
    num_source_a = len(
        exploded_tables[
            exploded_tables.str.contains('source_a', case=False, na=False)
        ].unique()
    )

    num_source_b = len(
        exploded_tables[
            exploded_tables.str.contains(r'source_b_|_ref', case=False, na=False, regex=True)
        ].unique()
    )

    return pd.Series({
        'Number of Excel': num_excel,
        'Number of Custom SQL': num_custom_sql,
        'Source A tables': num_source_a,
        'Source B tables': num_source_b + num_db
    })

df_data_source = (
    df.groupby('Workbook', dropna=False)
      .apply(lambda g: summarize_workbook(g.drop(columns=['Workbook'], errors='ignore')))
      .reset_index()
)

# ---- Sheet 3: Full detail (all datasource types) ----
df_ds_all = df.copy()
df_ds_all['Tables'] = df_ds_all['Tables'].apply(tables_to_text)

df_ds_all = df_ds_all[
    [
        'Workbook',
        'Datasource name',
        'Datasource type',
        'Script/File name',
        'Tables',
        'Join count',
        'SQL line count'
    ]
].drop_duplicates().reset_index(drop=True)

# ====== 4) WRITE EXCEL WITH THREE SHEETS ======
with pd.ExcelWriter(SUMMARY_XLSX) as writer:
    df_data_table.to_excel(writer, sheet_name='data table', index=False)
    df_data_source.to_excel(writer, sheet_name='data source', index=False)
    df_ds_all.to_excel(writer, sheet_name='ds all', index=False)

print("Done.")
print(f"Input  (Download):       {DOWNLOAD_FOLDER}")
print(f"SQL output root:         {TABLEAU_FOLDER}")
print(f"Combined summary Excel:  {SUMMARY_XLSX}")
print(f"Unzip folder:            {UNZIP_FOLDER}")
