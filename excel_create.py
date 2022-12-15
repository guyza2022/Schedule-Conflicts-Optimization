import pandas as pd
import numpy as np
import datetime as dt

sum_data = pd.read_excel('Samples.xlsx',sheet_name=['Summary'])
member_data = sum_data.get('Summary')
song_names = member_data['Name'].unique()

hr_per_interval = 0.5
min_per_interval = hr_per_interval*60
hr_interval = dt.timedelta(minutes=min_per_interval) #interval

hr_per_slot = 1
min_per_slot = hr_per_slot*60
hr_wanted = dt.timedelta(minutes=min_per_slot) #want 1 hr slot for each song

str_time = '17:00:00'
end_time = '00:00:00'

dict_name_members = {}
for name in song_names:
    member_list = member_data.loc[member_data['Name']==name]['Members'].values
    dict_name_members[name] = member_list

writer = pd.ExcelWriter('Schedule.xlsx')

for song in dict_name_members:
    members = dict_name_members.get(song)
    sheet = pd.DataFrame(columns=['Date','Time Start','Time End'])
    print('\nSong',song)
    for i,member in enumerate(members):
        date_time_start = pd.date_range('2022-12-10 '+str_time, '2022-12-11 '+end_time, freq=str(hr_per_slot)+'H', closed='left')
        date_time_end = pd.date_range('2022-12-10 '+str_time, '2022-12-11 '+end_time, freq=str(hr_per_slot)+'H', closed='right')
        date = date_time_start.date
        time_start = date_time_start.time
        time_end = date_time_end.time
        sheet['Date'] = pd.to_datetime(date).strftime('%Y-%m-%d')
        sheet['Date'] = sheet['Date']
        sheet['Time Start'] = time_start
        sheet['Time End'] = time_end
        sheet.insert(len(sheet.columns),member,np.nan,True)
        sheet.to_excel(writer, sheet_name=song)
    print('Song',song,'Done\n')
writer.save()
# a = pd.DataFrame(columns=['Date','Time Start','Time End'])

# writer = pd.ExcelWriter('Test1.xlsx')
# a.to_excel(writer, sheet_name='A')
# a.to_excel(writer, sheet_name='B')
# writer.save()
