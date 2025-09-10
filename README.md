# LSEG-ETL
Data ETL project: pulling minute-level data from LSEG Refinitiv Workspace API and export in certain format to excel, and append it from time to time. 
Market Data ETL

A simple ETL pipeline to fetch minute-level futures and commodity data from LSEG Workspace, transform it into clean tabular format, and load it into Excel for historical tracking.

📌 Features

Extract

Fetches real-time & historical minute data from LSEG Workspace API

Supports multiple contracts (FCPO, DCP, DBY, BO, SM, Sc, LGO, etc.)

Configurable fields (e.g., OPEN_PRC, HIGH_1, LOW_1, TRDPRC_1, NUM_MOVES, ACVOL_UNS)

Transform

Normalizes timestamp field (date → Timestamp)

Appends only new rows (avoids duplicates)

Supports multiple contracts in a single run

Load

Saves cleaned data into Excel (.xlsx)

Daily rolling files with contract + date in filename (e.g., FCPOc1_minute_20250910.xlsx)

Appends new data to existing files automatically

⚙️ Requirements

Python 3.9+

lseg-data Python SDK

pandas

openpyxl

Install dependencies:

pip install lseg-data pandas openpyxl

🚀 Usage

Clone repo

git clone https://github.com/yourname/marketdata-etl.git
cd marketdata-etl


Configure

Set your LSEG App Key in the script

Adjust Contracts list for the products you want

Adjust FIELDS for the data fields to capture

Run

python fetch_data.py


Output

Data will be saved in the Workspace_data_backup/ directory

Example file:

FCPOc1_minute_20250910.xlsx

📊 Example
Timestamp            OPEN_PRC  HIGH_1  LOW_1  TRDPRC_1  NUM_MOVES  ACVOL_UNS
2025-09-10 09:00:00   4470     4475    4468   4472      12         55
2025-09-10 09:01:00   4472     4476    4470   4475      18         110

🛠️ Roadmap

 Add support for equities & FX

 Load to database (Postgres/SQL) instead of Excel

 Schedule automated ETL jobs (cron / Airflow)

 Add data quality checks
