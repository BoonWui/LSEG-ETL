import os
import json
import lseg.data as ld
import pandas as pd
from datetime import datetime, timedelta
import glob
# v2 Added append function
# added other contracts
# --- Config -----------------------------------------------------------------
STATE_FILE      = "fcpo_state.json"
LSEG_APP_KEY    = "API KEY"   # put your Workspace app key
WORKSPACE_NAME  = "workspace"
RIC             = [f"FCPOc{i}" for i in range(1, 3)]
Contracts       = [f"FCPOc{i}" for i in range(1, 25)] + [f"DCPc{i}" for i in range(1, 25)] + [f"DBYc{i}" for i in range(1, 25)]+ [f"BOc{i}" for i in range(1, 25)]+ [f"SMc{i}" for i in range(1, 25)] + [f"Sc{i}" for i in range(1, 25)]+ [f"LGOc{i}" for i in range(1, 25)]+ [f"DCPv{i}" for i in range(1, 2)]+ [f"DBYv{i}" for i in range(1, 2)]+ [f"BOv{i}" for i in range(1, 2)]+ [f"SMv{i}" for i in range(1, 2)] + [f"Sv{i}" for i in range(1, 2)]+ [f"LGOv{i}" for i in range(1, 2)]
SAVE_DIR        = r"C:\Users\reuters\Desktop\Workspace_data_backup"
FIELDS = ["OPEN_PRC", "HIGH_1", "LOW_1", "TRDPRC_1"] # minute

os.makedirs(SAVE_DIR, exist_ok=True)

# --- Start Desktop session ---
session = ld.session.desktop.Definition().get_session()
session.open()
ld.session.set_default(session)

today_str = datetime.today().strftime("%Y%m%d")

for RIC in Contracts:
    try:
        print(f"Fetching {RIC}...")

        # --- Fetch new data ---
        df_new = (
            ld.get_history(
                universe=RIC,
                interval="minute",
                count=1000000,
                fields=FIELDS
            )
            .fillna(0)
            .infer_objects(copy=False)
        )

        # --- Reset index to get datetime column ---
        df_new = df_new.reset_index()
        if 'date' in df_new.columns:
            df_new.rename(columns={'date': 'Timestamp'}, inplace=True)
        elif 'Date' in df_new.columns:
            df_new.rename(columns={'Date': 'Timestamp'}, inplace=True)
        elif 'Timestamp' not in df_new.columns:
            df_new['Timestamp'] = pd.to_datetime(df_new.index)

        # Keep only desired columns
        df_new = df_new[['Timestamp'] + FIELDS]

        # --- Look for previous files ---
        pattern = os.path.join(SAVE_DIR, f"{RIC}_minute_*.xlsx")
        files = glob.glob(pattern)
        files.sort()

        if files:
            latest_file = files[-1]
            df_old = pd.read_excel(latest_file)

            # Normalize old file's datetime column
            ts_col = [c for c in df_old.columns if 'time' in c.lower() or 'date' in c.lower()]
            if ts_col:
                df_old.rename(columns={ts_col[0]: 'Timestamp'}, inplace=True)
            else:
                df_old['Timestamp'] = pd.to_datetime(df_old.index)

            # Keep only desired columns
            df_old = df_old[['Timestamp'] + FIELDS]

            df_old['Timestamp'] = pd.to_datetime(df_old['Timestamp'])
            df_new['Timestamp'] = pd.to_datetime(df_new['Timestamp'])

            # Append only new timestamps
            df_to_append = df_new[~df_new['Timestamp'].isin(df_old['Timestamp'])]
            if not df_to_append.empty:
                df_combined = pd.concat([df_old, df_to_append], ignore_index=True)
            else:
                df_combined = df_old
                print("No new data to append.")
        else:
            df_combined = df_new

        # --- Save to Excel with today's date ---
        file_name = f"{RIC}_minute_{today_str}.xlsx"
        file_path = os.path.join(SAVE_DIR, file_name)
        df_combined.to_excel(file_path, index=False)
        print(f"✅ File saved: {file_path}\n")

    except Exception as e:
        print(f"❌ Error fetching {RIC}: {e}\n")

# --- Close session ---
session.close()