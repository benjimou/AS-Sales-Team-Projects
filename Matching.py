# -*- coding: utf-8 -*-
"""
Created on Mon Mar 14 10:50:35 2022

@author: gou
"""


import pandas as pd
from datetime import date
import os
os.chdir(os.path.dirname(__file__))


cross = pd.read_excel(r"Resource\ItemCrossReferenceswithWMSCategoryResults.xlsx")
priority = pd.read_excel(r"Resource\Manu_Priority.xlsx")
kits = pd.read_excel(r"Resource\PartsinPackageResults.xlsx")
autoshack = pd.read_excel(r"Resource\Autoshack eBay IDs.xlsx")
previously_unknown = pd.read_excel(r"Result\Feedbacks\NoMatches.xlsx")


## Update the File name on the line below:
allmatchtodate = pd.read_excel(r"Result\AllMatchings\AllMatchesToDate.xlsx")
raw = pd.read_excel(r"Raw\MatchListings.xlsx")


#%%
# Remmove Duplicate and generate a sheet that conatins all single listing ID matches mulitple SKU to investigate
raw = pd.concat([raw, previously_unknown],ignore_index=True).drop_duplicates()
duplicateListing = raw[raw.duplicated(['ListingID'], keep=False)]
duplicateListing['ListingID'] = duplicateListing['ListingID'].astype(float)
duplicateListing = duplicateListing.sort_values(['ListingID']).reset_index(drop=True)
duplicateListing.to_excel(r"Result\Feedbacks\DeplicatedListing.xlsx", index=False)

print('Number of Rows of Duplicated Listing: ', len(duplicateListing))


#%%
# Remove those Listings found in the previous step, and also remove AutoShack Listings and Listings that have already found matched
cook = raw.drop_duplicates(['ListingID'], keep=False)
cook = cook[~cook['ListingID'].isin(autoshack['item_id'])] # Remove AutoShack Listing
cook = cook[~cook['ListingID'].isin(allmatchtodate['ListingID'])].reset_index(drop=True) # Remove Previous Match


# Kit Table Starts Building
kits = kits[['Package Name', 'Members', 'WMS Category']]
kits = kits[kits['Members'].str.len()>3]
kits['Formatted'] = kits['Members'].str.split(',')
def repeat_str(member):
    loc = member.find('(')
    loc2 = member.find(')')
    n = int(float(member[loc+1:loc2]))
    part = member[:loc]
    output = (part + ';') * n
    return output[:-1]
for i in range(len(kits)):
    for j in range(len(kits['Formatted'].iloc[i])):
        kits['Formatted'].iloc[i][j] = repeat_str(kits['Formatted'].iloc[i][j])
kits['Formatted'] = [';'.join(map(str, l)) for l in kits['Formatted']]
kits.insert(1, 'Type', 'Kit')
kits = kits.rename({'Package Name':'Item'}, axis=1)
# Kit Table is Done here


# Part Table Starts Building
part = cross[['Item', 'Manufacturer', 'Cross Reference', 'Order', 'Wms Category']].sort_values(by=['Item', 'Order'], ascending=[True, True]).reset_index(drop=True)
priority['key'] = priority['WMS Category'] + priority['Cross Reference Used for Purchasing']
priorityMapping = dict(zip(priority['key'], priority['Priority']))
part['Priority'] = (part['Wms Category']+part['Manufacturer']).map(priorityMapping)
part = part.sort_values(by=['Item','Priority'], ascending=[True, True]).reset_index(drop=True)
LabelMapping = dict(zip(priority['key'],priority['Bryce Label']))
part['Manufacturer'] = (part['Wms Category'] + part['Manufacturer']).map(LabelMapping).fillna(part['Manufacturer'])
part['Formatted'] = (part['Manufacturer'] + ";" + part['Cross Reference']).astype(str)
part = part.groupby(['Item', 'Manufacturer', 'Wms Category', 'Priority'], dropna=False).agg({'Formatted':'|'.join}).reset_index().sort_values(by=['Item', 'Priority'], ascending=[True, True], na_position='last').reset_index(drop=True)
part.insert(1, 'Type', 'Part')
# Part Table is Done here


#%%
# Putting the results together
result = pd.concat([part[['Item', 'Type', 'Formatted']], kits[['Item', 'Type', 'Formatted']]])
output = pd.merge(cook, result, how='left', left_on='Match SKU', right_on='Item')
output = output[['Match SKU', 'Item', 'Type', 'Formatted', 'ListingID']]


output_kit = output[output.Type == 'Kit'].drop(['Item', 'Type'], axis=1).rename({'Formatted':'KitContents'}, axis=1)
output_part = output[output.Type == 'Part'].drop(['Item', 'Type'], axis=1).rename({'Formatted':'CrossReference'}, axis=1)
output_part = output_part.drop_duplicates(subset='ListingID', keep='first').reset_index(drop=True)
output_unknown = output[output.Type.isnull()].drop(['Item', 'Type', 'Formatted'], axis=1)
allmatchtodate_updated = pd.concat([allmatchtodate, output]).drop(['Item', 'Type', 'Formatted'], axis=1).drop_duplicates().reset_index(drop=True)


#%%
# Save as Spreadsheets
allmatchtodate_updated.to_excel(r"Result\AllMatchings\AllMatchesToDate.xlsx", index=False)
#result.to_csv(r"M:\List Matching\ListingID\to_Bryce\Results\Dict_"+str(date.today())+".csv", index=False)
output_kit.to_csv(r"Result\Kits&Parts\KitMatches_"+str(date.today())+".csv", index=False)
output_part.to_csv(r"Result\Kits&Parts\PartMatches_"+str(date.today())+".csv", index=False)
output_unknown = pd.concat([previously_unknown, output_unknown], ignore_index=True).drop_duplicates().reset_index(drop=True)
output_unknown.to_excel(r"Result\Feedbacks\NoMatches.xlsx", index=False)
