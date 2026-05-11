# Deploy to Snowflake Streamlit

This guide covers deploying the Weekly Inflight Report Generator as a **Snowflake Native Streamlit** app.

## Prerequisites

- Snowflake account with Streamlit enabled (Enterprise Edition or higher)
- `ACCOUNTADMIN` or privileges to create databases, schemas, stages, and streamlits
- A running Snowflake Virtual Warehouse

## Deployment Steps

### 1. Configure the SQL Script

Open `snowflake_deploy.sql` and replace the placeholders:
- `MY_DATABASE` → your target database (e.g., `REPORTING_DB`)
- `MY_SCHEMA` → your target schema (e.g., `APPS`)
- `MY_WAREHOUSE` → your warehouse (e.g., `COMPUTE_WH`)

### 2. Run the SQL in Snowflake

Execute the SQL commands from `snowflake_deploy.sql` in a Snowflake Worksheet to create:
- The stage (`inflight_report_stage`)
- The Streamlit app (`weekly_inflight_report`)

### 3. Upload Files to the Stage

Use **one** of the following methods:

#### Option A: Snowsight UI (Easiest)
1. Go to **Data > Databases > YOUR_DB > YOUR_SCHEMA > Stages > inflight_report_stage**
2. Click **+ Files**
3. Upload these files from this repo:
   - `app.py`
   - `environment.yml`

#### Option B: SnowSQL CLI
```bash
snowsql -a <YOUR_ACCOUNT> -u <YOUR_USER>

PUT file://app.py @YOUR_DB.YOUR_SCHEMA.inflight_report_stage AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
PUT file://environment.yml @YOUR_DB.YOUR_SCHEMA.inflight_report_stage AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
```

#### Option C: SQL PUT Command (from Snowflake)
```sql
PUT file:///full/path/to/app.py @YOUR_DB.YOUR_SCHEMA.inflight_report_stage AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
PUT file:///full/path/to/environment.yml @YOUR_DB.YOUR_SCHEMA.inflight_report_stage AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
```

### 4. Add the `openpyxl` Package

1. In Snowsight, go to **Projects > Streamlit > weekly_inflight_report**
2. Click the **Packages** tab (or dropdown depending on your Snowflake version)
3. Search for and add **`openpyxl`** from the Anaconda channel
4. Click **Save**

> **Note:** `streamlit` and `pandas` are pre-installed in Snowflake Streamlit and do not need to be added.

### 5. Launch the App

Click **Open** or refresh the app preview in Snowsight. It should load immediately.

## How It Works in Snowflake

- Users open the app inside Snowsight (no local Python installation needed)
- They upload their **Raw Data CSV** and optional **Media Plan Excel** via the file uploaders
- The app processes everything in-memory inside Snowflake's secure sandbox
- The generated Excel report is downloaded directly from the browser

## Data Security Note

Since this is now running inside your Snowflake tenant:
- No data leaves your Snowflake environment during processing
- All file uploads are session-scoped and ephemeral
- Consider converting this in the future to read directly from Snowflake tables/stages instead of CSV uploads for a fully governed pipeline

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `ModuleNotFoundError: openpyxl` | Add `openpyxl` in the Packages tab of the Streamlit editor |
| App won't load | Ensure your warehouse is running and the stage path matches the `ROOT_LOCATION` |
| Download button doesn't work | Clear browser cache or try a different browser; Snowsight sometimes caches previews |

## Future Enhancement: Native Snowflake Tables

Instead of CSV uploads, you can modify the app to read directly from Snowflake tables:

```python
from snowflake.snowpark.context import get_active_session
session = get_active_session()
df = session.table("YOUR_DB.YOUR_SCHEMA.RAW_CAMPAIGN_DATA").to_pandas()
```

This eliminates manual uploads and enables scheduled/automated reporting.
