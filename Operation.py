
import pandas as pd
from datetime import date
#%%



File_name = "Hubs"



#%%
ref_hot = pd.read_excel(r"N:\Purchasing Department\Ge\Terapeak_Match_NoMatch_NewSKU\Resources\Forecasting_2022_06_07_08_06_40.xlsx")[14:]
cross_ref = pd.read_excel(r"N:\Purchasing Department\Ge\Terapeak_Match_NoMatch_NewSKU\Resources\ItemCrossReferenceswithWMSCategoryResults.xlsx")
#%%
File_name = "Brake-Pads"
df = pd.read_excel(r"N:\Purchasing Department\Ge\Terapeak_Match_NoMatch_NewSKU\To_Do\{}.xlsx".format(File_name))
df_match = df[df['Match SKU'].notnull()]
df_nomatch = df[df['Match SKU'].isnull()]

match_sku = ','.join(df_match['Member Item SKUs'][df_match['Member Item SKUs'].notnull()])
man_num = ','.join(df_nomatch['Mfr Numbers'][df_nomatch['Mfr Numbers'].notnull()].astype(str))

correction = {' ':',', ', ':',', '(2)':'', '(2':'', '(1)':'', '(1':'', ',,':','}
for x,y in correction.items():
    match_sku = match_sku.replace(x,y)
    man_num = man_num.replace(x,y)

match = pd.DataFrame(match_sku.split(','), columns=(['AutoShack SKU']))
nomatch = pd.DataFrame(man_num.split(','), columns=(['Manufacturer Number']))

match = match[~match['AutoShack SKU'].isin(ref_hot['Parameter'].tolist())].drop_duplicates().reset_index(drop=True)
nomatch = pd.merge(nomatch, cross_ref[['Item', 'Cross Reference']], left_on='Manufacturer Number', right_on='Cross Reference', how='left').drop(['Cross Reference'], axis=1).rename(columns={'Item':'AutoShack SKU'})
nomatch = nomatch[~nomatch['AutoShack SKU'].isin(ref_hot['Parameter'].tolist())].drop_duplicates(subset=['Manufacturer Number']).reset_index(drop=True)

match[['Package Name', 'Manufacture Number', 'Total Sales', 'Lowest Sales Price', 'Min Year', 'Max Year', 'Years >= 2007', 'MMY', 'VIO']] = '?'
nomatch[['Total Sales', 'Lowest Sales Price', 'Min Year', 'Max Year', 'Years >= 2007', 'MMY', 'VIO']] = '?'

for i in range(len(match)):
    try:
        word = match['AutoShack SKU'].iloc[i]
        indices = [x for x, y in enumerate(df_match['Member Item SKUs'].astype(str).str.contains(word)) if y==True]
        index = [x for x in indices if df_match['Sale Price'].iloc[x] == min(df_match['Sale Price'].iloc[indices])][0]
        match['Package Name'].iloc[i] = df_match['Match SKU'].iloc[index]
        match['Manufacture Number'].iloc[i] = df_match['Member Item SKUs'].iloc[index]
        match['Lowest Sales Price'].iloc[i] = df_match['Sale Price'].iloc[index]
        match['Total Sales'].iloc[i] = sum(df_match['Sales'].iloc[indices])
    except Exception:
        pass

for i in range(len(nomatch)):
    try:
        word = nomatch['Manufacturer Number'].iloc[i]
        indices = [x for x, y in enumerate(df_nomatch['Mfr Numbers'].astype(str).str.contains(word)) if y==True]
        index = [x for x in indices if df_nomatch['Sale Price'].iloc[x] == min(df_nomatch['Sale Price'].iloc[indices])][0]
        nomatch['Lowest Sales Price'].iloc[i] = df_nomatch['Sale Price'].iloc[index]
        nomatch['Total Sales'].iloc[i] = sum(df_nomatch['Sales'].iloc[indices])
    except Exception:
        pass

to_Phu = nomatch['Manufacturer Number'][nomatch['AutoShack SKU'].isnull()].reset_index(drop=True)
nomatch=nomatch[~nomatch['AutoShack SKU'].isnull()]

#%%
file_name  = File_name + ' (Organized_on_' + str(date.today()) +').xlsx'
writer = pd.ExcelWriter("N:\Purchasing Department\Ge\Terapeak_Match_NoMatch_NewSKU\Completed\{}".format(file_name), engine='xlsxwriter')
df.to_excel(writer, sheet_name='Raw Sheet', index=False)
match.to_excel(writer, sheet_name='Matched', index=False)
nomatch.to_excel(writer, sheet_name='No Matched', index=False)
writer.save()
writer.close()

file_name  = File_name + ' (Phu_' + str(date.today()) +').xlsx'
to_Phu.to_excel("N:\Purchasing Department\Ge\Terapeak_Match_NoMatch_NewSKU\Completed\{}".format(file_name), index=False)

