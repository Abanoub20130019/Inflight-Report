# Weekly Inflight Report Generator

A Streamlit tool to generate weekly inflight reports from raw campaign data.

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the app:
```bash
streamlit run app.py
```

## Usage

1. **Upload Raw Data**: Upload your `Raw Data.csv` file containing campaign performance data.

2. **Set Filters**:
   - **Campaign Name**: Enter a keyword or code that appears in campaign names (e.g., `signature`, `asc`). Leave empty to include all campaigns.
   - **Week Number**: Select the reporting week from available weeks in the data.
   - **Brand(s)**: Select one or more brands from the list.
   - **Market(s)**: Select one or multiple markets, or all.
   - **Channel(s)**: Select one or more channels (e.g., Facebook, TikTok, Snapchat).
   - **Funnel Type(s)**: Select Awareness, Consideration, Conversion, or all.

3. **Preview Data**: Review the filtered data and summary views by Market, Channel, and Funnel.

4. **Generate Report**: Click **Generate Excel Report** to download a multi-sheet Excel file containing:
   - **RAW**: Campaign-level data formatted like the template
   - **REPORT**: Structured report by Market and Funnel
   - **DETAIL**: Full filtered dataset
   - **SUMMARY**: Aggregated by Market, Channel, and Funnel
   - **AWARENESS**, **CONSIDERATION**, **CONVERSION**: Separate sheets per funnel type

## Data Requirements

The CSV file should contain the following columns:
- Year, Period, Week, Brand, Market, Channel, Funnel Type, Campaign
- Cost $, Clicks, CPC, Impressions, CPM, CTR, Reach, Frequency
- Full Video Views, VTR, Sessions, Click To Session Rate, App Installs
- Orders, Revenue $, Off. Orders, Off. Revenue $, CvR
