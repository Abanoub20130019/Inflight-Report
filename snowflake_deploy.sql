-- ============================================================
-- Snowflake Streamlit Deployment Script
-- Weekly Inflight Report Generator
-- ============================================================
-- 
-- INSTRUCTIONS:
-- 1. Replace MY_DATABASE, MY_SCHEMA, MY_WAREHOUSE with your actual names.
-- 2. Run these commands in a Snowflake Worksheet (with ACCOUNTADMIN or appropriate privileges).
-- 3. Upload app.py, README.md, and environment.yml to the stage created below.
--
-- ============================================================

-- Step 1: Use your target database and schema
USE DATABASE MY_DATABASE;
USE SCHEMA MY_SCHEMA;

-- Step 2: Create a dedicated stage for the Streamlit app (if it doesn't exist)
CREATE STAGE IF NOT EXISTS inflight_report_stage
  DIRECTORY = (ENABLE = TRUE)
  COMMENT = 'Stage for Weekly Inflight Report Generator Streamlit app';

-- Step 3: Create the Streamlit app
-- Note: After running this, you must upload app.py to the stage via Snowsight or PUT command.
CREATE STREAMLIT IF NOT EXISTS weekly_inflight_report
  ROOT_LOCATION = '@MY_DATABASE.MY_SCHEMA.inflight_report_stage'
  MAIN_FILE = 'app.py'
  QUERY_WAREHOUSE = 'MY_WAREHOUSE'
  COMMENT = 'Weekly Inflight Report Generator';

-- Step 4: (Optional) Grant usage to other roles
-- GRANT USAGE ON STREAMLIT weekly_inflight_report TO ROLE ANALYST_ROLE;

-- ============================================================
-- UPLOADING FILES TO THE STAGE
-- ============================================================
-- Option A: Via Snowsight UI
--   1. Go to Data > Databases > MY_DATABASE > MY_SCHEMA > Stages > inflight_report_stage
--   2. Click "+ Files" and upload: app.py, environment.yml
--
-- Option B: Via SQL (if running from Snowflake)
--   PUT file:///path/to/app.py @MY_DATABASE.MY_SCHEMA.inflight_report_stage AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
--   PUT file:///path/to/environment.yml @MY_DATABASE.MY_SCHEMA.inflight_report_stage AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
--
-- Option C: Via SnowSQL CLI (from your local machine)
--   snowsql -a <account> -u <user>
--   PUT file://app.py @MY_DATABASE.MY_SCHEMA.inflight_report_stage AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
--   PUT file://environment.yml @MY_DATABASE.MY_SCHEMA.inflight_report_stage AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
--
-- ============================================================
-- AFTER UPLOAD
-- ============================================================
-- The app will appear in Snowsight under: Projects > Streamlit > weekly_inflight_report
-- Open it and add 'openpyxl' from the Packages dropdown if it doesn't auto-install from environment.yml.
