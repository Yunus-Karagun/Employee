import pandas as pd
from pandas.tseries.offsets import MonthEnd
from datetime import datetime, timedelta, time
import numpy as np

Mola=pd.read_excel(r"D:\EX\Parametreler\Config.xlsx", sheet_name="Mola")
Tatil=pd.read_excel(r"D:\EX\Parametreler\Config.xlsx", sheet_name="Tatil", index_col=0)
Param=pd.read_excel(r"D:\EX\Parametreler\Config.xlsx", sheet_name="Param", index_col=0)

tatil_gunu = Tatil[Tatil["Tatil"]==1].index

Clean = pd.read_excel(Param.iloc[5,0], header=1, usecols=[0, 2, 5, 7, 8])\
           .query("Tarih.notnull()")\
           .sort_values(by=['Departman', 'Adı Soyadı', 'Tarih', 'Saat'], ascending=[True, True, True, True])\
           .drop_duplicates()\
           .reset_index(drop=True)

Clean['Tarih'] = pd.to_datetime(Clean['Tarih'], format='%d.%m.%Y')
Clean['Saat'] = Clean.apply(lambda row: pd.to_datetime(row['Tarih'].strftime('%Y-%m-%d') + ' ' + str(row['Saat']), format='%Y-%m-%d %H:%M') if pd.notnull(row['Saat']) else row['Saat'], axis=1)

sequences = {}
sequence_numbers = []
for index, row in Clean.iterrows():
    key = (row['Adı Soyadı'], row['Tarih'], row['Durum'])
    if key in sequences:
        sequences[key] += 1
    else:
        sequences[key] = 1
    sequence_numbers.append(sequences[key])
Clean['GC_No'] = sequence_numbers

Clean["Durum_No"]= Clean["GC_No"].astype(str)+"-"+ Clean["Durum"]
Clean.drop(columns=['Durum', 'GC_No'], inplace=True)
pivot_df = Clean.pivot_table(index=['Adı Soyadı', 'Departman', 'Tarih'],
                                   columns='Durum_No', 
                                   values='Saat', 
                                   aggfunc='first').reset_index()
pivot_df = pivot_df.sort_values(by=['Departman', 'Adı Soyadı', 'Tarih'], ascending=True)

pivot_df['Brut1'] = (pivot_df['1-Çıkış'] - pivot_df['1-Giriş']).dt.total_seconds() / 3600
#pivot_df['Brut1'] = pivot_df['Brut1'].fillna(0)
pivot_df['Brut2'] = (pivot_df['2-Çıkış'] - pivot_df['2-Giriş']).dt.total_seconds() / 3600
#pivot_df['Brut2'] = pivot_df['Brut2'].fillna(0)
if '3-Giriş' in pivot_df.columns and '3-Çıkış' in pivot_df.columns:
    pivot_df['3-Giriş'] = pd.to_datetime(pivot_df['3-Giriş'], errors='coerce')
    pivot_df['3-Çıkış'] = pd.to_datetime(pivot_df['3-Çıkış'], errors='coerce')
    pivot_df['Brut3'] = (pivot_df['3-Çıkış'] - pivot_df['3-Giriş']).dt.total_seconds() / 3600
    pivot_df['Brut3'] = pivot_df['Brut3'].fillna(0)


date_range = pd.date_range(start=Param.iloc[2,0], end=Param.iloc[2,0]+ MonthEnd(1))
employee_departments = pivot_df.drop_duplicates(subset=['Adı Soyadı'], keep='last')[['Adı Soyadı', 'Departman']]
date_employee_grid = pd.MultiIndex.from_product([date_range, employee_departments['Adı Soyadı']], names=['Tarih', 'Adı Soyadı']).to_frame(index=False)
date_employee_grid = pd.merge(date_employee_grid, employee_departments, on='Adı Soyadı', how='left')
complete_df = pd.merge(date_employee_grid, pivot_df, on=['Adı Soyadı', 'Departman', 'Tarih'], how='left')
column_order = ['Adı Soyadı', 'Departman', 'Tarih'] + [col for col in complete_df.columns if col not in ['Adı Soyadı', 'Departman', 'Tarih']]
complete_df = complete_df[column_order]
complete_df = complete_df.sort_values(by=['Departman', 'Adı Soyadı', 'Tarih']).reset_index(drop=True)


df_merged = pd.merge(complete_df, Mola, on='Departman', how='left')

def categorize_time(row, time_column):
    time = row[time_column].time() if pd.notnull(row[time_column]) else None
    if time:
        if time < row['W_Start']:
            return "0"
        elif row['W_Start'] <= time < row['MB_Start']:
            return "1"
        elif row['MB_Start'] <= time <= row['MB_End']:
            return "2"
        elif row['MB_End'] < time < row['LB_Start']:
            return "3"
        elif row['LB_Start'] <= time <= row['LB_End']:
            return "4"
        elif row['LB_End'] < time < row['AB_Start']:
            return "5"
        elif row['AB_Start'] <= time <= row['AB_End']:
            return "6"
        elif row['AB_End'] < time <= row['W_End']:
            return "7"
        else:
            return "8"  # For times outside the defined intervals
    else:
        return "KY"  # For cases where time is NaN or NaT


# Apply the function for all four time columns
df_merged['Entry_Category_1'] = df_merged.apply(lambda row: categorize_time(row, '1-Giriş'), axis=1)
df_merged['Exit_Category_1'] = df_merged.apply(lambda row: categorize_time(row, '1-Çıkış'), axis=1)
df_merged['Entry_Category_2'] = df_merged.apply(lambda row: categorize_time(row, '2-Giriş'), axis=1)
df_merged['Exit_Category_2'] = df_merged.apply(lambda row: categorize_time(row, '2-Çıkış'), axis=1)

# Concatenate the categories for the first and second entries/exits into separate columns
df_merged['CC_1'] = df_merged.apply(
    lambda row: "KY" if "KY" in [row['Entry_Category_1'], row['Exit_Category_1']] else row['Entry_Category_1'] + row['Exit_Category_1'], axis=1
)

df_merged['CC_2'] = df_merged.apply(
    lambda row: "KY" if "KY" in [row['Entry_Category_2'], row['Exit_Category_2']] else row['Entry_Category_2'] + row['Exit_Category_2'], axis=1
)

df_merged.drop(columns=['Entry_Category_1', 'Exit_Category_1', 'Entry_Category_2', 'Exit_Category_2',], inplace=True)

# Extend the categorization function to include the third entry and exit if they exist
if '3-Giriş' in df_merged.columns and '3-Çıkış' in df_merged.columns:
    df_merged['Entry_Category_3'] = df_merged.apply(lambda row: categorize_time(row, '3-Giriş'), axis=1)
    df_merged['Exit_Category_3'] = df_merged.apply(lambda row: categorize_time(row, '3-Çıkış'), axis=1)

    # Concatenate the categories for the third entry/exit into a separate column
    df_merged['CC_3'] = df_merged.apply(
        lambda row: "KY" if "KY" in [row['Entry_Category_3'], row['Exit_Category_3']] else row['Entry_Category_3'] + row['Exit_Category_3'], axis=1
    )

    # Remove intermediate columns for the third entry/exit
    df_merged.drop(columns=['Entry_Category_3', 'Exit_Category_3'], inplace=True)

#M1
def calculate_duration(row):
    def time_diff(start_date, start_time, end_time):
        # Handles None values and combines date with time for start and end datetime objects
        if pd.isna(start_date) or pd.isna(start_time) or pd.isna(end_time):
            return pd.Timedelta(seconds=0)
        if isinstance(start_time, datetime):  # If start_time is already datetime, no need to combine
            start_datetime = start_time
        else:  # Combine date and time to create datetime
            start_datetime = datetime.combine(start_date, start_time)
        if isinstance(end_time, datetime):  # If end_time is already datetime, no need to combine
            end_datetime = end_time
        else:  # Combine date and time to create datetime
            end_datetime = datetime.combine(start_date, end_time)
        return end_datetime - start_datetime
    
    tarih = row['Tarih']
    cc_1 = row['CC_1']
    
    # Full durations dictionary to use time_diff with 'Tarih' for datetime.time columns
    durations = {
        "00": time_diff(tarih, row['1-Giriş'], row['1-Çıkış']),
        "01": time_diff(tarih, row['1-Giriş'], row.get('W_Start')),
        "02": time_diff(tarih, row['1-Giriş'], row.get('W_Start')) + time_diff(tarih, row['1-Çıkış'], row.get('MB_Start')),
        "03": time_diff(tarih, row['1-Giriş'], row.get('W_Start')) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')),
        "04": time_diff(tarih, row['1-Giriş'], row.get('W_Start')) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['1-Çıkış']),
        "05": time_diff(tarih, row['1-Giriş'], row.get('W_Start')) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "06": time_diff(tarih, row['1-Giriş'], row.get('W_Start')) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['1-Çıkış']),
        "07": time_diff(tarih, row['1-Giriş'], row.get('W_Start')) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "08": time_diff(tarih, row['1-Giriş'], row.get('W_Start')) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "11": pd.Timedelta(seconds=0),
        "12": time_diff(tarih, row.get('MB_Start'), row['1-Çıkış']),
        "13": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')),
        "14": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['1-Çıkış']),
        "15": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "16": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['1-Çıkış']),
        "17": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "18": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "22": time_diff(tarih, row['1-Giriş'], row['1-Çıkış']),
        "23": time_diff(tarih, row['1-Giriş'], row.get('MB_End')),
        "24": time_diff(tarih, row['1-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['1-Çıkış']),
        "25": time_diff(tarih, row['1-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "26": time_diff(tarih, row['1-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['1-Çıkış']),
        "27": time_diff(tarih, row['1-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "28": time_diff(tarih, row['1-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "33": pd.Timedelta(seconds=0),
        "34": time_diff(tarih, row.get('LB_Start'), row['1-Çıkış']),
        "35": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "36": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['1-Çıkış']),
        "37": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "38": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "44": time_diff(tarih, row['1-Giriş'], row['1-Çıkış']),
        "45": time_diff(tarih, row['1-Giriş'], row.get('LB_End')),
        "46": time_diff(tarih, row['1-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['1-Çıkış']),
        "47": time_diff(tarih, row['1-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "48": time_diff(tarih, row['1-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "55": pd.Timedelta(seconds=0),
        "56": time_diff(tarih, row.get('AB_Start'), row['1-Çıkış']),
        "57": time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "58": time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "66": time_diff(tarih, row['1-Giriş'], row['1-Çıkış']),
        "67": time_diff(tarih, row['1-Giriş'], row.get('AB_End')),
        "68": time_diff(tarih, row['1-Giriş'], row.get('AB_End')),
        "77": pd.Timedelta(seconds=0),
        "78": pd.Timedelta(seconds=0),
        "88": pd.Timedelta(seconds=0)
    }
    
    duration = durations.get(cc_1, pd.Timedelta(seconds=0))
    return duration.total_seconds() / 3600  # Convert duration to hours

# Assuming df_merged is prepared with the necessary columns
df_merged['M1'] = df_merged.apply(calculate_duration, axis=1)


#M2
def calculate_duration(row):
    def time_diff(start_date, start_time, end_time):
        # Handles None values and combines date with time for start and end datetime objects
        if pd.isna(start_date) or pd.isna(start_time) or pd.isna(end_time):
            return pd.Timedelta(seconds=0)
        if isinstance(start_time, datetime):  # If start_time is already datetime, no need to combine
            start_datetime = start_time
        else:  # Combine date and time to create datetime
            start_datetime = datetime.combine(start_date, start_time)
        if isinstance(end_time, datetime):  # If end_time is already datetime, no need to combine
            end_datetime = end_time
        else:  # Combine date and time to create datetime
            end_datetime = datetime.combine(start_date, end_time)
        return end_datetime - start_datetime
    
    tarih = row['Tarih']
    cc_2 = row['CC_2']
    
    # Full durations dictionary to use time_diff with 'Tarih' for datetime.time columns
    durations = {
        "00": time_diff(tarih, row['2-Giriş'], row['2-Çıkış']),
        "01": time_diff(tarih, row.get('W_Start'), row['2-Giriş']),
        "02": time_diff(tarih, row.get('W_Start'), row['2-Giriş']) + time_diff(tarih, row['2-Çıkış'], row.get('MB_Start')),
        "03": time_diff(tarih, row.get('W_Start'), row['2-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')),
        "04": time_diff(tarih, row.get('W_Start'), row['2-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['2-Çıkış']),
        "05": time_diff(tarih, row.get('W_Start'), row['2-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "06": time_diff(tarih, row.get('W_Start'), row['2-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['2-Çıkış']),
        "07": time_diff(tarih, row.get('W_Start'), row['2-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "08": time_diff(tarih, row.get('W_Start'), row['2-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "11": pd.Timedelta(seconds=0),
        "12": time_diff(tarih, row.get('MB_Start'), row['2-Çıkış']),
        "13": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')),
        "14": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['2-Çıkış']),
        "15": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "16": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['2-Çıkış']),
        "17": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "18": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "22": time_diff(tarih, row['2-Giriş'], row['2-Çıkış']),
        "23": time_diff(tarih, row['2-Giriş'], row.get('MB_End')),
        "24": time_diff(tarih, row['2-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['2-Çıkış']),
        "25": time_diff(tarih, row['2-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "26": time_diff(tarih, row['2-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['2-Çıkış']),
        "27": time_diff(tarih, row['2-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "28": time_diff(tarih, row['2-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "33": pd.Timedelta(seconds=0),
        "34": time_diff(tarih, row.get('LB_Start'), row['2-Çıkış']),
        "35": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "36": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['2-Çıkış']),
        "37": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "38": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "44": time_diff(tarih, row['2-Giriş'], row['2-Çıkış']),
        "45": time_diff(tarih, row['2-Giriş'], row.get('LB_End')),
        "46": time_diff(tarih, row['2-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['2-Çıkış']),
        "47": time_diff(tarih, row['2-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "48": time_diff(tarih, row['2-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "55": pd.Timedelta(seconds=0),
        "56": time_diff(tarih, row.get('AB_Start'), row['2-Çıkış']),
        "57": time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "58": time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "66": time_diff(tarih, row['2-Giriş'], row['2-Çıkış']),
        "67": time_diff(tarih, row['2-Giriş'], row.get('AB_End')),
        "68": time_diff(tarih, row['2-Giriş'], row.get('AB_End')),
        "77": pd.Timedelta(seconds=0),
        "78": pd.Timedelta(seconds=0),
        "88": pd.Timedelta(seconds=0)
    }
    
    duration = durations.get(cc_2, pd.Timedelta(seconds=0))
    return duration.total_seconds() / 3600  # Convert duration to hours

# Assuming df_merged is prepared with the necessary columns
df_merged['M2'] = df_merged.apply(calculate_duration, axis=1)


#M3
def calculate_duration(row):
    def time_diff(start_date, start_time, end_time):
        # Handles None values and combines date with time for start and end datetime objects
        if pd.isna(start_date) or pd.isna(start_time) or pd.isna(end_time):
            return pd.Timedelta(seconds=0)
        if isinstance(start_time, datetime):  # If start_time is already datetime, no need to combine
            start_datetime = start_time
        else:  # Combine date and time to create datetime
            start_datetime = datetime.combine(start_date, start_time)
        if isinstance(end_time, datetime):  # If end_time is already datetime, no need to combine
            end_datetime = end_time
        else:  # Combine date and time to create datetime
            end_datetime = datetime.combine(start_date, end_time)
        return end_datetime - start_datetime
    
    tarih = row['Tarih']
    cc_3 = row['CC_3']
    
    # Full durations dictionary to use time_diff with 'Tarih' for datetime.time columns
    durations = {
        "00": time_diff(tarih, row['3-Giriş'], row['3-Çıkış']),
        "01": time_diff(tarih, row.get('W_Start'), row['3-Giriş']),
        "02": time_diff(tarih, row.get('W_Start'), row['3-Giriş']) + time_diff(tarih, row['3-Çıkış'], row.get('MB_Start')),
        "03": time_diff(tarih, row.get('W_Start'), row['3-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')),
        "04": time_diff(tarih, row.get('W_Start'), row['3-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['3-Çıkış']),
        "05": time_diff(tarih, row.get('W_Start'), row['3-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "06": time_diff(tarih, row.get('W_Start'), row['3-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['3-Çıkış']),
        "07": time_diff(tarih, row.get('W_Start'), row['3-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "08": time_diff(tarih, row.get('W_Start'), row['3-Giriş']) + time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "11": pd.Timedelta(seconds=0),
        "12": time_diff(tarih, row.get('MB_Start'), row['3-Çıkış']),
        "13": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')),
        "14": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['3-Çıkış']),
        "15": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "16": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['3-Çıkış']),
        "17": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "18": time_diff(tarih, row.get('MB_Start'), row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "22": time_diff(tarih, row['3-Giriş'], row['3-Çıkış']),
        "23": time_diff(tarih, row['3-Giriş'], row.get('MB_End')),
        "24": time_diff(tarih, row['3-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row['3-Çıkış']),
        "25": time_diff(tarih, row['3-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "26": time_diff(tarih, row['3-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['3-Çıkış']),
        "27": time_diff(tarih, row['3-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "28": time_diff(tarih, row['3-Giriş'], row.get('MB_End')) + time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "33": pd.Timedelta(seconds=0),
        "34": time_diff(tarih, row.get('LB_Start'), row['3-Çıkış']),
        "35": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')),
        "36": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['3-Çıkış']),
        "37": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "38": time_diff(tarih, row.get('LB_Start'), row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "44": time_diff(tarih, row['3-Giriş'], row['3-Çıkış']),
        "45": time_diff(tarih, row['3-Giriş'], row.get('LB_End')),
        "46": time_diff(tarih, row['3-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row['3-Çıkış']),
        "47": time_diff(tarih, row['3-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "48": time_diff(tarih, row['3-Giriş'], row.get('LB_End')) + time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "55": pd.Timedelta(seconds=0),
        "56": time_diff(tarih, row.get('AB_Start'), row['3-Çıkış']),
        "57": time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "58": time_diff(tarih, row.get('AB_Start'), row.get('AB_End')),
        "66": time_diff(tarih, row['3-Giriş'], row['3-Çıkış']),
        "67": time_diff(tarih, row['3-Giriş'], row.get('AB_End')),
        "68": time_diff(tarih, row['3-Giriş'], row.get('AB_End')),
        "77": pd.Timedelta(seconds=0),
        "78": pd.Timedelta(seconds=0),
        "88": pd.Timedelta(seconds=0)
    }
    
    duration = durations.get(cc_3, pd.Timedelta(seconds=0))
    return duration.total_seconds() / 3600  # Convert duration to hours
    
if 'CC_3' in df_merged.columns:
    df_merged['M3'] = df_merged.apply(calculate_duration, axis=1)
else:
    print("Column 'CC_3' not found in df_merged. Skipping duration calculation.")


if 'Brut3' in df_merged.columns:
    df_merged['Brut_T'] = df_merged['Brut1'].fillna(0) + df_merged['Brut2'].fillna(0) + df_merged['Brut3'].fillna(0)
else:
    # If 'Brut3' does not exist, sum only 'Brut1' and 'Brut2' (NaN values already replaced by 0)
    df_merged['Brut_T'] = df_merged['Brut1'].fillna(0) + df_merged['Brut2'].fillna(0)
    
if 'M3' in df_merged.columns:
    df_merged['Mola_T'] = df_merged['M1'].fillna(0) + df_merged['M2'].fillna(0) + df_merged['M3'].fillna(0)
else:
    # If 'Brut3' does not exist, sum only 'Brut1' and 'Brut2' (NaN values already replaced by 0)
    df_merged['Mola_T'] = df_merged['M1'].fillna(0) + df_merged['M2'].fillna(0)
    
df_merged['Net_Calışma'] = df_merged['Brut_T'] - df_merged['Mola_T'] 




df_merged['Gun'] = df_merged['Tarih'].dt.dayofweek
df_merged['Carpan'] = df_merged['Gun'].apply(lambda x: 1 if x < 5 else (1.5 if x == 5 else (2 if x == 6 else np.nan)))

df_merged['Normal_calisma'] = np.where(
    (df_merged['Gun'] < 5) & (df_merged['Net_Calışma'] <= 9), df_merged['Net_Calışma'],
    np.where((df_merged['Gun'] < 5) & (df_merged['Net_Calışma'] > 9), 9,
             np.where(df_merged['Gun'] >= 5, 0, np.nan)
    )
)

df_merged['Fazla_Calisma'] = np.where(
    (df_merged['Gun'] < 5) & (df_merged['Net_Calışma'] <= 9.25), 0,
    np.where((df_merged['Gun'] < 5) & (df_merged['Net_Calışma'] > 9.25), df_merged['Net_Calışma'] - 9,
             np.where(df_merged['Gun'] >= 5, df_merged['Net_Calışma'], np.nan)
    )
)

df_merged['eksik_saat'] = np.where(
    (df_merged['Gun'] < 5) & (df_merged['Net_Calışma'] >= 9), 0,
    np.where(
        (df_merged['Gun'] < 5) & (df_merged['Net_Calışma'] < 9) & (df_merged['Net_Calışma'] > 0), 9 - df_merged['Net_Calışma'],
        np.where(
            (df_merged['Gun'] >= 5) | (df_merged['Tarih'].isin(tatil_gunu)), 0, np.nan
        )
    )
)

df_merged['eksik_gun'] = np.where(
    (df_merged['Gun'] < 5) & (~df_merged['Tarih'].isin(tatil_gunu)) & (df_merged['CC_1'] == "KY") & (df_merged['CC_2'] == "KY"), 1, 0
)

pivot_df = df_merged.pivot_table(index=["Adı Soyadı", 'Departman'], 
                                 values=['Normal_calisma', 'Fazla_Calisma', 'eksik_saat', 'eksik_gun'], 
                                 aggfunc='sum')
pivot_df_sorted = pivot_df.sort_values(by=['Departman', 'Adı Soyadı'], ascending=[True, True]).reset_index()

eksik_df = df_merged[
    ((df_merged["1-Giriş"].isna()) & (~df_merged["1-Çıkış"].isna())) | 
    ((~df_merged["1-Giriş"].isna()) & (df_merged["1-Çıkış"].isna()))
]
columns_to_match = ["Adı Soyadı", "Departman", "Tarih"]
final_eksik_df = pd.merge(df_merged, eksik_df[columns_to_match].drop_duplicates(), on=columns_to_match, how='inner')

with pd.ExcelWriter(Param.iloc[7,0]) as writer:
    df_merged.to_excel(writer, sheet_name='Detay')
    pivot_df_sorted.to_excel(writer, sheet_name='Ozet')
    eksik_df.to_excel(writer, sheet_name='ekik')