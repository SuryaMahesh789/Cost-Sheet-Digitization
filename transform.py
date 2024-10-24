import pandas as pd
import numpy as np

input_file='Plastic Part 5 Oct 24.xlsx'
output_file = 'Plastic Part Formatted.xlsx'

# Read the first sheet into a DataFrame
df = pd.read_excel(input_file, sheet_name="Sheet1")

# Example 1: Rename columns
df = df.rename(columns={
    'Sl No': 'SlNo',
    'Plant': 'PlantCode',
    'Part Number': 'PartNo',
    'Part Name': 'PartName',
    'RM Spec': 'RMCommodity',
    'RM Spec': 'RMGrade',
    'RM Spec': 'Specification',
    'Gross Wt': 'GrossWeight',
    'Scrap Recovery': 'ScrapRecoveryPercent',
    'Scrap / Runner Wt': 'JaliScrapWeight',
    'Supplier Name': 'SupplierName',
    'Supplier Code': 'SupplierCode',
    'RM Rate': 'RMRatePerKg',
    'Scrap / Regrind Rate': 'JaliScrapCostPerKg',
    'Part Number': 'NewColumn2',
    'Part Number': 'NewColumn2',
    'Part Number': 'NewColumn2',
    'Part Number': 'NewColumn2',
    'Part Number': 'NewColumn2'
})


df['Class']='VBC'
df['BOMLevel']=0
df['PartType']='Component'
df['BOMNo'] = ''
df['AssemblyPartNo']=''
df['RevisionNumber']=''
df['GroupCode']=''
df['Technology']='Plastic'
df['Quantity']=1
df['RMCode']=''
df['RMType']='STD'
df['Specification']='NA'
df['RMUOM']='Kilogram'
df['RMSourceVendorCode']=''
df['MultipleRM']=False
df['BlankSizeT']=''
df['BlankSizeL']=''
df['BlankSizeW']=''
df['BlankSizeP']=''
df['RunnerWeight']=0

df['FinishWeight']=df['Gross Wt']-df['Scrap / Runner Wt']

df['CircleScrapWeight']=0

df['SOB']=100

df['RMFreightCostPerKg']=0
df['LandedRMCostPerKg']=df['RMRatePerKg']+df['RMFreightCostPerKg']

df['ShearingCostPerKg']=0
df['RMCutOffPrice']=0

df['NetRMCostPerKg']=df['LandedRMCostPerKg']+df['ShearingCostPerKg']

df['JaliScrapCostPerKg']=0

df['CircleScrapCostPerKg']=''

df['ProcessingLossPercent']=''

df['ProcessingLossWeight']=''

df['BurningLossPercent']=''

df['BurningLossWeight']=''

df['RMCostPerPiece']=((df['NetRMCostPerKg']*((100-df['MasterBatchPercentage'])/100)+df['MasterBatchTotal']))*(df['GrossWeight']+df['ProcessingLossWeight']+df['BurningLossWeight']+df['RunnerWeight'])-(df['JaliScrapCostPerKg']*(df['GrossWeight']+df['ProcessingLossWeight']-df['FinishWeight']+df['RunnerWeight'])*df['ScrapRecoveryPercent']/100)

df['TotalRMCost'] = df['RMCostPerPiece']*df['Quantity']

df['BOPCategory']='Plastic'

df['BOPUOM']='Number'

df['BOPCostPerPiece']=df['BOPCostPerAssembly']+df['BOPHandlingCost']

df['BOPCostPerAssembly'] = df['BOPCostPerPiece'].shift(-1) * df['Quantity'].shift(-1)

df['BOPHandlingType']='Percentage'

df['BOPCostApplicableforHandling']=df['BOPCostPerAssembly']

df['BOPHandlingCost']=df['BOPCostPerAssembly']*(df['BOPHandlingCharges']/100)

df['ProgressiveStamping_NoOfCavity']=0

df['ProgressiveStamping_StrokeRate']=0

df['ProgressiveStamping_UnitofMeasurement']=0

df['ProgressiveStamping_Cost']=0

df['MachineTonnage']=0

df['Plastic Molding_NoOfCavity']

df['Plastic Molding_StrokeRate']

print(f'Transformation complete. Data saved to {output_file}.')

