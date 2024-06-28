import pandas as pd
import openpyxl

def count(x,value):
    return (x == value).sum()

UE = ['2UE Validated\nTest Platforms','2UE Validated\nTest Platforms','1UE Validated\nTest Platforms','1UE Validated Test Platforms with Exceptions']
NUM = ['292','168','207','251','300'] 
num_list = []
Case_count_list = []
Band_count_list = []
df_list = []

# Setting Input & Output Path
file_path = '3.92.0_20240429_r071_noTPCVInfo - Copy.xlsx'
target_path = 'Summary.xlsx'
Excel_file = pd.ExcelFile(file_path)

RedCap_Sheets = [sheet for sheet in Excel_file.sheet_names if 'RedCap' in sheet]
print(RedCap_Sheets)

for sheet in RedCap_Sheets:
    df = pd.read_excel(Excel_file,sheet_name=sheet,skiprows=1,engine='openpyxl',index_col=None)
    df_list.append(df)
df = pd.concat(df_list,ignore_index=True)
print(df.shape)

# Filter the Protocal
column_names = df.columns
df = df[df['Specification']=='38.523-1']


df['TC'] = df[UE].fillna('').apply(lambda x: ','.join(x.astype(str)),axis=1)
for num in NUM:
    num_list.append(num)
    # Tester
    df[f'TC_{num}']= ''
    df.loc[df['TC'].str.contains(num,na=False),f'TC_{num}'] = num
    # TC and TC_BAND
    Case_count = df[df[f'TC_{num}']==num]['Test Case'].nunique()
    Band_count = df[f'TC_{num}'].apply(lambda x: x == num).sum()
    Case_count_list.append(Case_count)
    Band_count_list.append(Band_count)
    
df_summary = pd.DataFrame({'NUM':num_list,'TC':Case_count_list,'TC_BAND':Band_count_list})
print(df_summary)
    
del df['TC']


df_copy = df.copy()



with pd.ExcelWriter(target_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='RedCap Combined Sheet',index=False)
    df_summary.to_excel(writer,sheet_name='Summary',index=False)
    
