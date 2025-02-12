# @title Begin Automation

from google.colab import files, drive
import pandas as pd
import sys
import time
import shutil

# ====
# 0. Upload the POR and BOM Files; mount the shared Google Drive
# ====

uploaded = files.upload()

por_file = None
bom_file = None

try:
  drive.mount('/content/drive')
except Exception as ex:
  print('[ERROR] There was an issue connecting to the Google Drive. Please try again')
  sys.exit(1)

# ====
# 1. Validate the uploaded files
# ====

if len(uploaded) != 2:
  print('[ERROR] Please upload exactly 2 files')
  sys.exit(1)

for filename in uploaded.keys():
  if 'por' in filename.lower() and 'bom' in filename.lower():
    print(f'[ERROR] {filename} cannot contain both BOM and POR in the filename')
    shutil.copyfile(
        filename,
        f'/content/drive/MyDrive/DemandPlanning/error/{filename}')
    sys.exit(1)

  if 'por' in filename.lower():
    por_file = filename
  elif 'bom' in filename.lower():
    bom_file = filename

if por_file:
  print(f'Plan of Records file: {por_file}')
else:
  print('[ERROR] Plan of Records file is missing')
  sys.exit(1)

if bom_file:
  print(f'Bill of Materials file: {bom_file}')
else:
  print('[ERROR] Bill of Materials file is missing')
  sys.exit(1)

# ====
# 2. Validate file extensions
# ====

# read files into Pandas
if por_file.endswith('.csv'):
  por = pd.read_csv(por_file)
elif por_file.endswith(('.xls', '.xlsx')):
  por = pd.read_excel(por_file)
else:
  shutil.copyfile(
      por_file,
      f'/content/drive/MyDrive/DemandPlanning/error/{por_file}')
  print(f'[ERROR] {por_file} is an unsupported file -- please ensure that uploaded file is either an Excel or CSV file')
  sys.exit(1)

if bom_file.endswith('.csv'):
  bom = pd.read_csv(bom_file)
elif bom_file.endswith(('.xls', '.xlsx')):
  bom = pd.read_excel(bom_file)
else:
  shutil.copyfile(
      bom_file,
      f'/content/drive/MyDrive/DemandPlanning/error/{bom_file}')
  print(f'[ERROR] {bom_file} is an unsupported file -- please ensure that uploaded file is either an Excel or CSV file')
  sys.exit(1)

# ====
# 3. Validate columns and formats for uploaded files
# ====

for column in ['region', 'country', 'site_type', 'mf_part_number', 'part_type',
               'grouping', 'part_id', 'uin', 'vendor', 'manufacturer',
               'mf_item_description', 'cost', 'supply_source', 'quantity']:
  if column not in bom.columns:
    print(f'[ERROR] Please ensure that {column} is present as a column header in the file.')
    shutil.copyfile(
      bom_file,
      f'/content/drive/MyDrive/DemandPlanning/error/{bom_file}')
    sys.exit(1)

for column in ['region', 'country', 'investment_id', 'investment_name', 'site_id',
               'rad_site_infra', 'rad_site_eue', 'demand_type', 'business_line',
               'site_type', 'building_type', 'site_go_live',
               'need_by_date_infrastructure', 'need_by_date_end_user_equipment',
               'need_by_date_special_projects', 'mdf_need_by_date', 'idf_need_by_date',
               'mdf_type', 'mdf_quantity', 'ap_quantity', 'idf_quantity',
               'infra_value_added_reseller', 'car_status', 'por_confidence_level']:
  if column not in por.columns:
    print(f'[ERROR] Please ensure that {column} is present as a column header in the file.')
    shutil.copyfile(
      por_file,
      f'/content/drive/MyDrive/DemandPlanning/error/{por_file}')
    sys.exit(1)

# ====
# 4. Perform POR Cleansing; clean datetime fields and replace commas and periods
#    in the quantity fields. Make the resulting file available for download
# ====

date_columns = [
        'site_go_live',
        'need_by_date_infrastructure',
        'need_by_date_end_user_equipment',
        'need_by_date_special_projects',
        'mdf_need_by_date',
        'idf_need_by_date'
    ]

for col in date_columns:
  por[col] = pd.to_datetime(por[col], errors='coerce')

# clean and convert quantity columns (replace invalid entries with '0' and convert to numeric)
quantity_columns = ['mdf_quantity', 'ap_quantity', 'idf_quantity']
for col in quantity_columns:
  por[col] = por[col].astype(str).str.replace(',', '', regex=True)
  por[col] = pd.to_numeric(por[col], errors='coerce').fillna(0)

timestr = time.strftime("%Y%m%d%H%M%S")
por_filename = f'{timestr}_por_result.xlsx'
por.to_excel(por_filename, index=False)
shutil.copyfile(
    por_filename,
    f'/content/drive/MyDrive/DemandPlanning/processed/{por_filename}')
print(f'Uploaded {por_filename} to /processed folder.')
# files.download(por_filename) -- replace with download function if copy file to drive not working

# ====
# 5. Explode the BOM via the POR; merge, consolidate, determine rad site based
#    on part type; clean dates, costs, quantities. Make result file available
#    to download
# ====

bom = bom[bom['site_type'].notna() & (bom['site_type'] != '')]
# Merge POR and BOM on 'site_type'
df = por.merge(bom, on='site_type', how='left', suffixes=('', '_bom'))

# Process 'rad_site' column
df['rad_site'] = df['rad_site_infra']
df.loc[df['part_type'].str.lower() == 'end user equipment', 'rad_site'] = df['rad_site_eue']

# Process 'site_go_live' column
df['site_go_live'] = df['site_go_live'].replace(['', '00:00:00'], pd.NA)

# Process 'need_by_date' column
df.loc[df['part_type'].str.lower().isin(['infrastructure', '', None]), 'need_by_date'] = df['need_by_date_infrastructure']
df.loc[df['part_type'].str.lower() == 'end user equipment', 'need_by_date'] = df['need_by_date_end_user_equipment']
df['need_by_date'] = df['need_by_date'].replace(['', '00:00:00'], pd.NA)

# Process 'cost' column
df['cost'] = df['cost'].astype(str).replace(['', '.', ','], '0')
df['cost'] = df['cost'].str.replace('$', '').str.replace(',', '').astype(float)

# Process 'quantity' column
df['quantity'] = df['quantity'].astype(str).replace(['', '.', ','], '0')
df['quantity'] = df['quantity'].str.replace(',', '').astype(float)

# Handle missing 'need_by_date'
df.loc[(df['part_type'].str.lower() == 'infrastructure') & df['need_by_date_infrastructure'].isna(), 'quantity'] = 0
df.loc[(df['part_type'].str.lower() == 'end user equipment') & df[
    'need_by_date_end_user_equipment'].isna(), 'quantity'] = 0

result_df = df[['region', 'country', 'site_id', 'rad_site', 'demand_type', 'business_line', 'site_type',
                'building_type', 'car_status', 'por_confidence_level', 'mdf_type', 'mf_part_number',
                'mf_item_description', 'part_type', 'grouping', 'part_id', 'uin', 'vendor', 'manufacturer',
                'investment_id', 'investment_name', 'site_go_live', 'need_by_date', 'cost', 'quantity']]

# select distinct
result_df = result_df.drop_duplicates()

timestr = time.strftime("%Y%m%d%H%M%S")
unpivot_forecast = f'{timestr}_parts_forecast.xlsx'
result_df.to_excel(unpivot_forecast, index=False)
shutil.copyfile(
    unpivot_forecast,
    f'/content/drive/MyDrive/DemandPlanning/processed/{unpivot_forecast}')
print(f'Uploaded {unpivot_forecast} to /processed folder.')
# files.download(unpivot_forecast) -- replace if google drive upload not working

print('[FINISHED]')
