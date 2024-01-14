#from chromosomes import chromosome as chms
from gene import gene, geneChunk, geneBranch
import pandas as pd
import numpy as np
import datetime as dt
import random as r
from copy import deepcopy
import math as m
import regex as re
import matplotlib.pyplot as plt
import matplotlib
import statistics
from sklearn import preprocessing
from collections import OrderedDict
import sys
import tkinter as tk
import threading
import time as t
from collections import Counter


class constraints:
    
    # Constant
    room_col = 'Room'
    task_col = 'Song'
    summary_col = 'Summary'
    request_col = 'Request'
    date_col = 'Date'
    time_start_col = 'Time Start'
    time_end_col = 'Time End'  
    date_start_col = 'Date Start'
    date_end_col = 'Date End'
    vocal_col = 'Vocal'
    guitarist_col = 'Guitarist'
    percussion_col = 'Percussion'
    members_col = 'Members'
    part_col = 'Part'
    hr_per_slot = 1
    max_score = 1 # Max possible score
    slot_max_score = 1 # Max possible score in 1 slot

    def __init__(self, constraint_sheet_name: str):
        # Ininitialize constraints
        self.sheet_dict = {}
        self.task_sheet_dict = {}
        self.task_member_dict = {} # Store member of each task
        self.score_dict = {} # Keys are tasks. Each key is consisted of list of length max slots possible showing score of that task based on index
        self.room_score_dict = {}
        self.memory_dict = {}
        self.blank_room_dict = {}
        self.room_days_dict = {} # Number of days by rooms
        self.room_chunks_dict = {} # Nuber of slots by rooms
        self.global_counter = pd.DataFrame(columns=[self.date_col,self.time_start_col,self.time_end_col])
        self.visit_counter_dict = {}
        self.request_dict = {} # Keys are rooms. Each room have tasks as keys and data structure inside is (Number of Slot, Hours per Slot)
        self.room_names = []
        self.task_names = []
        self.nan_index = [] # The indexes are global index
        self.max_day = 0
        self.max_chunk_size = 0
        self.n_rooms = 0
        # Read constraint sheet
        self.excelFile = pd.ExcelFile(constraint_sheet_name)
        # Assign sheets to dictionary
        for sheet in self.excelFile.sheet_names:
            self.sheet_dict[sheet] = self.excelFile.parse(sheet)
        # Assign room and summary sheet to variable
        self.room_sheet = self.sheet_dict.get(self.room_col+'s')
        self.summary_sheet = self.sheet_dict.get(self.summary_col)
        self.request_sheet = self.sheet_dict.get(self.request_col)
        # Assign names
        self.room_names = self.room_sheet[self.room_col].unique()
        self.task_names = self.summary_sheet[self.task_col].unique()
        # Assign numbers
        self.n_rooms = len(self.room_names)

        # Generate blank sheet and assign to dictionary
        self.assignBlankRoom()
        # Generate global and visit counters and assign to dictionary
        self.assignCounters()
        # Generate task sheet and assign to dictionary
        self.assignTaskSheet()
        # Generate request dictionries by rooms and assign to dictionary
        self.assignRequest()
        # Check nan index and assign to dictionary
        self.assignNanIndex()

        # Assign max values
        self._getMaxChunkSize()
        self._getMaxDay()

        # Generate score lists and assign to dictionary
        self.assignScoreDict()
        # Generate task member lists and assign to dictionary
        self.assignTaskMemberDict()

        # Generate chromosome template
        self.chmsTemplate = self.generateChromosomeTemplate()

    def assignBlankRoom(self):
        # Generate
        global_counter = pd.DataFrame(columns=[self.date_col,self.time_start_col,self.time_end_col])
        for room in self.room_sheet[self.room_col].unique():
            date_start = self.room_sheet.loc[self.room_sheet[self.room_col] == room, self.date_start_col].values[0]
            date_end = self.room_sheet.loc[self.room_sheet[self.room_col] == room, self.date_end_col].values[0]
            time_start = self.room_sheet.loc[self.room_sheet[self.room_col] == room, self.time_start_col].values[0]
            time_end = self.room_sheet.loc[self.room_sheet[self.room_col] == room, self.time_end_col].values[0]
            all_date = pd.date_range(pd.to_datetime(date_start).strftime('%Y-%m-%d'), pd.to_datetime(date_end).strftime('%Y-%m-%d'), freq='1D')
            blank_room = pd.DataFrame(columns=[self.date_col,self.time_start_col,self.time_end_col,self.task_col])
            for date in all_date:
                date_till = date
                if (time_end<time_start or time_end == '00:00:00'):
                    date_till += dt.timedelta(days=1)
                date = pd.to_datetime(date).strftime('%Y-%m-%d')
                date_till = pd.to_datetime(date_till).strftime('%Y-%m-%d')
                start = pd.date_range(date+' '+time_start, date_till+' '+time_end, freq=str(self.hr_per_slot)+'H', inclusive='left')
                stop = pd.date_range(date+' '+time_start, date_till+' '+time_end, freq=str(self.hr_per_slot)+'H', inclusive='right')
                temp_sheet = pd.DataFrame({
                    self.date_col: start.date,
                    self.time_start_col: start.time,
                    self.time_end_col: stop.time,
                    self.task_col : np.nan
                })
                temp_sheet[self.date_col] = temp_sheet[self.date_col].apply(lambda x: x.strftime('%Y-%m-%d'))
                temp_sheet[self.time_start_col] = temp_sheet[self.time_start_col].apply(lambda x: x.strftime('%H:%M:%S'))
                temp_sheet[self.time_end_col] = temp_sheet[self.time_end_col].apply(lambda x: x.strftime('%H:%M:%S'))
                blank_room = pd.concat([blank_room,temp_sheet], ignore_index=True)
            delta = (pd.to_datetime(date_end)-pd.to_datetime(date_start))
            # Update value in of days and chunks in dictionary
            self.room_days_dict[room] = delta.days + 1
            self.room_chunks_dict[room] = int(len(blank_room)/self.room_days_dict[room])
            # Assign blank room sheet to dictionary
            self.blank_room_dict[room] = blank_room

    def assignCounters(self):
        # Generate
        for room, blank_room in self.blank_room_dict.items():
            blank_counter = deepcopy(blank_room)
            blank_counter.drop(columns=[self.task_col],inplace=True)
            for task in self.task_names:
                blank_counter[task] = 0
            self.visit_counter_dict[room] = blank_counter
            self.global_counter =  pd.merge(self.global_counter,blank_room,how='outer')
            self.global_counter.drop(columns=[self.task_col],inplace=True)
            for task in self.task_names:
                self.global_counter[task] = 0
    
        return None

    def assignTaskSheet(self):
        # Assign tasks to dictionary
        for task in self.task_names:
            task_sheet = pd.DataFrame({
                self.date_col: self.sheet_dict.get(self.vocal_col)[self.date_col],
                self.time_start_col: self.sheet_dict.get(self.vocal_col)[self.time_start_col],
                self.time_end_col: self.sheet_dict.get(self.vocal_col)[self.time_end_col]
            })
            task_member_sheet = self.summary_sheet.loc[self.summary_sheet[self.task_col]==task,[self.members_col,self.part_col]]
            for index,row in task_member_sheet.iterrows():
                part_sheet = self.sheet_dict.get(row[self.part_col])
                task_sheet[row[self.members_col]] = part_sheet[row[self.members_col]]
            self.task_sheet_dict[task] = task_sheet
        
        return None

    def assignScoreDict(self):
        for task in self.task_names:
            task_score_list = []
            for index in range(self.max_chunk_size*self.max_day):
                task_score_list.append(self.calculateTaskScore(task,index))
            self.score_dict[task] = task_score_list
        
        for room in self.room_names:
            self.room_score_dict[room] = {}
            for task in self.score_dict.keys():
                room_masked_index = [x for x in list(range(self.max_chunk_size*self.max_day)) if x not in [n[0] for n in self.nan_index if n[1] == room]]
                room_task_list = [self.score_dict[task][i] for i in room_masked_index]
                self.room_score_dict[room][task] = room_task_list
        
        return None

    
    def assignTaskMemberDict(self):
        for task in self.task_names:
            task_member_list: list = list(self.summary_sheet.loc[self.summary_sheet[self.task_col]==task,self.members_col].values)
            self.task_member_dict[task] = task_member_list
        
        return None
    
    def assignRequest(self):
        for room in self.room_names:
            self.request_dict[room] = {}
        for task in self.task_names:
            for room in self.room_names:
                task_request_by_room = self.request_sheet.loc[self.request_sheet[self.task_col]==task,room].values[0]
                hr_per_slot = self._getHrPerSlot(room,task)
                self.request_dict[room][task] = (task_request_by_room/hr_per_slot,hr_per_slot)
    
    def assignNanIndex(self):
        # Generate
        nan_check_sheet = pd.DataFrame(columns=[self.date_col,self.time_start_col,self.time_end_col])
        # Check nan index
        for nr in self.blank_room_dict.keys():
            nan_check_sheet = pd.merge(nan_check_sheet, self.blank_room_dict.get(nr),how='outer')
        nan_check_sheet.sort_values(by=[self.date_col,self.time_start_col,self.time_end_col],inplace=True)
        dt_nan_check_data = nan_check_sheet.loc[:,[self.date_col,self.time_start_col,self.time_end_col]].values
        dt_nan_check_data = [list(x) for x in dt_nan_check_data]
        for nr in self.blank_room_dict.keys():
            nr_sheet = self.blank_room_dict.get(nr)
            dt_nr_data= nr_sheet.loc[:,[self.date_col,self.time_start_col,self.time_end_col]].values
            dt_nr_data = [list(x) for x in dt_nr_data]
            # Assign index to list
            for dt_index,dt_val in enumerate(dt_nan_check_data):
                if dt_val not in dt_nr_data:
                    self.nan_index.append((dt_index,nr))
        
        
        return None
    
    def _getMaxChunkSize(self):
        all_chunk_size = list(self.room_chunks_dict.values())
        self.max_chunk_size = max(all_chunk_size)
        return None

    def _getMaxDay(self):
        all_days = list(self.room_days_dict.values())
        self.max_day = max(all_days)
        return None

    def generateChromosomeTemplate(self):
        template = []
        for room in self.room_names:
            room_days = self.room_days_dict[room]
            chunk_size = self.room_chunks_dict[room]
            branch_nan_index = [x[0] for x in self.nan_index]
            template.append(geneBranch(room, room_days, chunk_size, self.max_chunk_size, branch_nan_index))
        return template
    
    def calculateTaskScore(self,task,index):
        if not pd.isna(task):
            task_sheet = self.task_sheet_dict.get(task)
            task_members = task_sheet.columns[3:]
            num_members = len(task_members)
            task_row = task_sheet.iloc[index,:]
            task_score = task_row.iloc[3:].sum()/num_members
            return task_score
        else:
            print('YYYYY')
            return 0
    
    def _getHrPerSlot(self,room,task):
        if not pd.isna(task):
            task_row = self.request_sheet.loc[self.request_sheet[self.task_col] == task]
            request_hr = task_row.filter(regex=room).iloc[0,1]
            return request_hr
        else:
            return 1 # Should be smallest slot of the schedule
    

            
        