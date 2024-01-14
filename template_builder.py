import pandas as pd
import numpy as np
import datetime as dt

def build_template(template_file_name,hr_per_slot = 1):
    sum_data = pd.read_excel(template_file_name,sheet_name=['Summary','Rooms'])
    member_data = sum_data.get('Summary')
    room_data = sum_data.get('Rooms')
    #fill merged columns
    member_data['Song'] = member_data['Song'].fillna(method='ffill')
    member_data['Artist'] = member_data['Artist'].fillna(method='ffill')
    #get song names
    song_names = member_data['Song'].unique()
    part_names = member_data['Part'].unique()
    #get room names
    room_names = room_data['Room'].unique()

    #get maximum time of all rooms
    min_date = room_data['Date Start'].min()
    max_date = room_data['Date End'].max()
    str_time = room_data['Time Start'].min()
    if 'Yes' in room_data['Overnight'].unique():
        overnight_row = room_data.loc[room_data['Overnight'] == 'Yes']
        end_time = overnight_row['Time End'].max()
    else:
        end_time = room_data['Time End'].max()

    #convert time to string
    str_time = str_time.strftime('%H:%M:%S')
    end_time = end_time.strftime('%H:%M:%S')

    #song based version
    # dict_name_members = {}
    # for name in song_names:
    #     member_list = member_data.loc[member_data['Song']==name]['Members'].values
    #     dict_name_members[name] = member_list

    #part based version
    dict_name_members = {}
    for part in part_names:
        member_list = member_data.loc[member_data['Part']==part]['Members'].values
        dict_name_members[part] = member_list

    writer = pd.ExcelWriter('Schedule.xlsx')
    #request sheet
    request_sheet = pd.DataFrame({'Song': song_names,
                                'Requested Hours': 2,
                                'Priority Rank': 1})
    for room in room_names:
        request_sheet[room] = 1
        request_sheet['Hour Per Slot('+room+')'] = 1
    member_data.to_excel(writer, sheet_name='Summary', index=False)
    request_sheet.to_excel(writer, sheet_name='Request', index=False)
    #form sheets
    for song in dict_name_members:
        members = dict_name_members.get(song)
        sheet = pd.DataFrame(columns=['Date','Time Start','Time End'])
        #member_data.to_excel(writer, sheet_name='Summary', index=False)
        room_data['Date Start'] = room_data['Date Start'].apply(lambda x: pd.to_datetime(x).strftime('%Y-%m-%d'))
        room_data['Date End'] = room_data['Date End'].apply(lambda x: pd.to_datetime(x).strftime('%Y-%m-%d'))
        room_data.to_excel(writer, sheet_name='Rooms', index=False)
        all_date = pd.date_range(min_date.strftime('%Y-%m-%d'), max_date.strftime('%Y-%m-%d'), freq='1D')
        print('\nSong',song)
        for date in all_date:
            date_end = date
            if (end_time<str_time or end_time == '00:00:00'):
                date_end += dt.timedelta(days=1)
            date = pd.to_datetime(date).strftime('%Y-%m-%d')
            date_end = pd.to_datetime(date_end).strftime('%Y-%m-%d')
            #print('\n\n',date,date_end)
            start = pd.date_range(date+' '+str_time, date_end+' '+end_time, freq=str(hr_per_slot)+'H', closed='left')
            stop = pd.date_range(date+' '+str_time, date_end+' '+end_time, freq=str(hr_per_slot)+'H', closed='right')
            temp_sheet = pd.DataFrame(
                {'Date': start.date,
                'Time Start': start.time,
                'Time End': stop.time}
            )
            temp_sheet['Date'] = temp_sheet['Date'].apply(lambda x: x.strftime('%Y-%m-%d'))
            temp_sheet['Time Start'] = temp_sheet['Time Start'].apply(lambda x: x.strftime('%H:%M:%S'))
            temp_sheet['Time End'] = temp_sheet['Time End'].apply(lambda x: x.strftime('%H:%M:%S'))
            sheet = pd.concat([sheet,temp_sheet], ignore_index=True)
        for i,member in enumerate(members):
            sheet.insert(len(sheet.columns),member,0,True)
        sheet.to_excel(writer, sheet_name=song, index=False)
        print('Song',song,'Done\n')
    writer.close()
