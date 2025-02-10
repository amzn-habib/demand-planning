# @title Begin Automation

from google.colab import files
from google.colab import drive
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
  print('There was an issue connecting to the Google Drive. Please try again')
  sys.exit(1)

# ====
# 1. Validate the uploaded files
# ====

if len(uploaded) != 2:
  print('Please upload exactly 2 files')
  sys.exit(1)

for filename in uploaded.keys():
  if 'por' in filename.lower() and 'bom' in filename.lower():
    print(f'Error: {filename} cannot contain both BOM and POR in the filename')
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
  print('Plan of Records file is missing')
  sys.exit(1)

if bom_file:
  print(f'Bill of Materials file: {bom_file}')
else:
  print('Bill of Materials file is missing')
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
  print(f'{por_file} is an unsupported file -- please ensure that uploaded file is either an Excel or CSV file')
  sys.exit(1)

if bom_file.endswith('.csv'):
  bom = pd.read_csv(bom_file)
elif bom_file.endswith(('.xls', '.xlsx')):
  bom = pd.read_excel(bom_file)
else:
  shutil.copyfile(
      bom_file,
      f'/content/drive/MyDrive/DemandPlanning/error/{bom_file}')
  print(f'{bom_file} is an unsupported file -- please ensure that uploaded file is either an Excel or CSV file')
  sys.exit(1)

# ====
# 3. Validate columns and formats for uploaded files
# ====

for column in ['region', 'country', 'site_type', 'mf_part_number', 'part_type',
               'grouping', 'part_id', 'uin', 'vendor', 'manufacturer',
               'mf_item_description', 'cost', 'supply_source', 'quantity']:
  if column not in bom.columns:
    print(f'Please ensure that {column} is present as a column header in the file.')
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
    print(f'Please ensure that {column} is present as a column header in the file.')
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
# files.download(por_filename) -- replace with download function if copy file to drive not working

# ====
# 5. Explode the BOM via the POR; merge, consolidate, determine rad site based
#    on part type; clean dates, costs, quantities. Make result file available
#    to download
# ====

# Standardize part_type column in bom
bom['part_type'] = bom['part_type'].str.lower().fillna('')

# Determine rad_site based on part_type
def get_rad_site(row):
    if row['part_type'] in ('infrastructure', ''):
        return row['rad_site_infra']
    elif row['part_type'] == 'end user equipment':
        return row['rad_site_eue']
    return None

# Merge bom and por using left join on site_type
merged = por.merge(bom, on='site_type', how='left')

# Apply transformations
merged['rad_site'] = merged.apply(get_rad_site, axis=1)
merged['site_go_live'] = pd.to_datetime(merged['site_go_live'], errors='coerce')
merged['need_by_date'] = pd.to_datetime(
    merged.apply(
        lambda row: row['need_by_date_infrastructure'] if row['part_type'] in ('infrastructure', '') else row[
            'need_by_date_end_user_equipment'], axis=1),
    errors='coerce'
)

# Clean cost and quantity columns
for col in ['cost', 'quantity']:
    merged[col] = pd.to_numeric(
        merged[col].astype(str).str.replace(r'[$,]', '', regex=True).replace({'': '0', '.': '0'}),
        errors='coerce'
    ).fillna(0)

# Remove rows where site_type is null or empty
merged = merged[merged['site_type'].notna() & (merged['site_type'] != '')]
merged.drop(columns=['rad_site_infra', 'rad_site_eue'], inplace=True, errors="ignore")
grps = {'country': ['country_x', 'country_y'], 'region': ['region_x', 'region_y']}
merged = pd.lreshape(merged, grps).drop_duplicates().convert_dtypes()

output_columns = ['region', 'country', 'site_id', 'rad_site', 'demand_type', 'business_line', 'site_type',
                  'building_type', 'car_status', 'por_confidence_level', 'mdf_type', 'mf_part_number',
                  'mf_item_description', 'part_type', 'grouping', 'part_id', 'uin', 'vendor', 'manufacturer',
                  'investment_id', 'investment_name', 'site_go_live', 'need_by_date', 'cost', 'quantity']

merged = merged[output_columns]

timestr = time.strftime("%Y%m%d%H%M%S")
unpivot_forecast = f'{timestr}_unpivot_forecast.xlsx'
merged.to_excel(unpivot_forecast, index=False)
shutil.copyfile(
    unpivot_forecast,
    f'/content/drive/MyDrive/DemandPlanning/processed/{unpivot_forecast}')
# files.download(unpivot_forecast) -- replace if google drive upload not working
