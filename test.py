import pandas as pd
import datetime as dt
import numpy as np
import math as m

song_name = ['A','B','C','D','E','F','G']
song_data = []
data = pd.read_excel('Samples.xlsx',sheet_name=['A','B','C','D','E','F','G'])
for name in song_name:
        sdt = data.get(name)
        sdt.insert(0,'Name',name,True)
        song_data.append(sdt)

hr_per_interval = 0.5
min_per_interval = hr_per_interval*60
hr_interval = dt.timedelta(minutes=min_per_interval) #interval

hr_per_slot = 1
min_per_slot = hr_per_slot*60
hr_wanted = dt.timedelta(minutes=min_per_slot) #want 1 hr slot for each song

list_song = []
for i,sd in enumerate(song_data):
       available = sd[~(sd == 0).any(axis=1)]
       available= available.iloc[:,:4]
       print(available)
       list_song.append(available)
#all_dates = data['Date'].unique()
# available_A = song_data[0].loc[(song_data[0]['John'] == 1) & (song_data[0]['Mary'] == 1) & (song_data[0]['Tony'] == 1)]
# available_B = song_data[1].loc[(song_data[1]['Bell'] == 1) & (song_data[1]['Apple'] == 1) & (song_data[1]['Fah'] == 1)]
# available_A.insert(0,'Name',song_name[0],True)
# available_B.insert(0,'Name',song_name[1],True)
# available_A = available_A.iloc[:,:4] #cut off name columns
# available_B = available_B.iloc[:,:4] #cut off name columns
# list_song = [available_A,available_B]
#print(list_song[0])
print('\n')
#print(list_song[1])

str_time = '17:00:00'
end_time = '00:00:00'
#construct slot table
slot = pd.DataFrame(
        {'Start': pd.date_range('2022-12-10 '+str_time, '2022-12-11 '+end_time, freq=str(hr_per_slot)+'H', closed='left'),
        'End': pd.date_range('2022-12-10 '+str_time, '2022-12-11 '+end_time, freq=str(hr_per_slot)+'H', closed='right'),
        'Song': np.nan}
     )
print(slot)
slot_requirement = pd.DataFrame(columns=['Name', 'Time Start', 'Time End'])
print('\n')
#search for matched avaible time depends on slot table
for index,row in slot.iterrows():
        searching_start = row['Start']#timestamp to datetime
        searching_end = searching_start + hr_wanted - hr_interval #datetime
        searching_date = searching_start.date()
        #print('Searching with',searching_start,'to',searching_end,'on',searching_date)
        for song in list_song:
                dated_song = song.loc[song['Date'] == str(searching_date)]
                if searching_start.time() in dated_song['Time Start'].values and searching_end.time() in dated_song['Time Start'].values:
                        new_row = pd.DataFrame({'Name' : song.iloc[0,0], 'Time Start' : searching_start, 'Time End' : searching_end}, index=[0])
                        slot_requirement = pd.concat([new_row,slot_requirement.loc[:]]).reset_index(drop=True)
                        print(song.iloc[0,0],str(searching_start)," ",str(searching_end)," Appended")

slot_requirement['Count'] = slot_requirement.groupby('Name')['Name'].transform('count')
slot_requirement = slot_requirement.sort_values(by=['Count'])
slot_requirement = slot_requirement.reset_index(drop=True)

print('\nStarter slot requirement')
print(slot_requirement)

#max_iter = 5
max_iter = m.ceil(len(slot)/(60/min_per_interval))
#Priority method
for c in range(1,max_iter+11):
        k=c
        k=slot_requirement.Count.min()
        while(len(slot_requirement.loc[slot_requirement['Count']==k])!=0):
                print('\nWith k = ',k)
                require = slot_requirement.loc[slot_requirement['Count']==k]
                slot_start = require.loc[0,'Time Start']
                slot_end = require.loc[0,'Time End']
                slot_name = require.loc[0,'Name']
                print('\nInsert ',slot_start,'to',slot_end,'with',slot_name)
                slot.loc[(slot['Start'] >= str(slot_start)) & (slot['Start'] <= str(slot_end)),'Song'] = slot_name
                #delete duplicate time that duplicate the inserted time
                slot_requirement = slot_requirement.drop(slot_requirement[(slot_requirement['Time Start'] == str(slot_start)) & (slot_requirement['Time End'] == str(slot_end)) | (slot_requirement['Name'] == slot_name)].index)
                #recount
                slot_requirement['Count'] = slot_requirement.groupby('Name')['Name'].transform('count')
                slot_requirement = slot_requirement.sort_values(by=['Count'])
                slot_requirement = slot_requirement.reset_index(drop=True)
                #go back to min Count
                k=slot_requirement.Count.min()
                print('\nMin: ',k)
                print('\n New slot requirement')
                print(slot_requirement)

print(slot)
# for c in range(1,max_iter+1):
#         print(c)
#         require = slot_requirement.loc[slot_requirement['Count']==c]
#         print(require)
#         #if have the sime period do random
#         for index,row in require.iterrows():
#                 slot_start = row['Time Start']
#                 slot_end = row['Time End']
#                 slot_name = row['Name']
#                 slot.loc[(slot['Start'] >= str(slot_start)) & (slot['Start'] <= str(slot_end)),'Song'] = slot_name
#                 print(slot)