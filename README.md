# Tableau_workbook_extractor
A Python toolkit to extract Tableau workbook metadata: data sources, custom SQL, calculated fields, and parameters, exported to structured Excel reports. Built from hands-on experience where manual documentation is impractical in corporate Tableau environment.

---
 
## Background
 
In large BI environments, Tableau workbooks accumulate complex logic fast:
- Custom SQL queries referencing dozens of source tables
- Calculated fields referencing each other in non-obvious chains
- Parameters controlling filter logic across multiple dashboards
These scripts automate the audit work instead of opening every workbook manually in Tableau Desktop.
 
---
 
## Scripts
 
### `tableau_backend_extractor.py` — Data source audit
 
Scans all `.twb` / `.twbx` files in a folder and extracts:
 
- **Custom SQL**: full query text saved as `.sql` files, with JOIN count and line count per script
- **Excel / CSV connections**: file names and paths
- **Database table connections**: server and database identifiers
- **Table frequency**: ranked list of which source tables appear most across all workbooks
- **Output — `Tableau summary.xlsx` (3 sheets):**
 
| Sheet | Contents |
|---|---|
| `data table` | All source tables ranked by frequency |
| `data source` | Per-workbook count of Excel sources, Custom SQL scripts, and table categories |
| `ds all` | Full detail: every datasource × workbook combination with JOIN count and SQL line count |
 
---
 
### `tableau_frontend_extractor.py` — Field audit
 
Parses all calculated fields and parameters from a folder of Tableau workbooks.
 
- **Parameters**: name, datatype, default value, domain type
- **Calculated fields**: formula with internal identifiers
- **Output — `Tableau Parameters and Cal Fields.xlsx`:**
 
| Sheet | Contents |
|---|---|
| `Summary` | One row per workbook: parameter count, calculated field count, full name lists |
| `Parameters_Detail` | Every parameter with full metadata |
| `Calculations_Detail` | Every calculated field: formula, role, datatype, downstream impacts |
 
---
 
## How it works
 
Tableau `.twb` files are XML. `.twbx` files are ZIP archives containing a `.twb`.
 
**Backend extractor:**
1. Discovers all `.twb` / `.twbx` recursively under the configured folder
2. Extracts `.twb` from `.twbx` archives into a temp folder
3. Walks `<datasource>` → `<named-connection>` and `<relation>` XML nodes
4. Classifies each connection: Excel/CSV, Database, or Custom SQL
5. For Custom SQL: extracts table references via regex (`FROM` / `JOIN`), counts JOINs and lines
6. Aggregates everything into a single Excel report
   
**Frontend extractor:**
1. Locates the `Parameters` datasource node and extracts all parameter metadata
2. Iterates all other datasource nodes for `<column>` elements with a `<calculation class='tableau'>` child
3. Skips columns under `<datasource-dependencies>` to avoid duplicates from cross-datasource references

