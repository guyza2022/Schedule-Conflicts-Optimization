import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog, Menu
from tkcalendar import DateEntry
import pandas as pd
import numpy as np
import datetime as dt
import template_builder
import genetic
import threading

print(tk.TkVersion)
#variables
descale = 1.5
width = int(1920/descale)
height = int(1080/descale)
current_frame = 'generator'
file_session = False
menu_session = False
#time_change_session = False
backup_time_dict = {}
current_display_sheet = ''
room_data_dict = {}


#colors
primary = '#323232'
sub_primary = '#474747'
secondary = '#ff2c4a'
sub_secondary = '#ff667c'
font_color = '#f4f4f4'
info_font_color = 'e0e0e0'

#panel ratios
margin = 0.025
left_button_width = 0.9
left_button_height = 0.05
left_x_ratio = 0.0
center_x_ratio = 1-left_x_ratio

#generator
right_x_ratio = 0.15
generator_x_ratio = 1-right_x_ratio
top_header_height = margin
sheets_frame_width = 0.0
sheet_button_height = 0.05
display_height = 0.65
display_width = 1-sheets_frame_width
editor_height = 1-top_header_height-display_height
browse_button_height = 0.1
browse_button_width = 0.175
time_width = 0.4
add_width = 0.3
del_width = 0.3
tool_height = 1-browse_button_height
views_height = 0.075
views_button_width = 0.1
mode_option_height = 0.2
quality_frame_height = 0.2


root = tk.Tk()
root.title("Folkedule")
root.geometry(str(width)+"x"+str(height))
root.pack_propagate(False)
root.resizable(0,0)
gen_mode = tk.StringVar()
gen_mode.set('ma')
gen_quality = tk.StringVar()
gen_quality.set('a')
display_efficient_score = tk.StringVar()
#left frame
# left_frame = tk.Frame(root)
# left_frame.place(relheight=1,relwidth=left_x_ratio,relx=0,rely=0)

# main_button = tk.Button(left_frame, text="Main", command=lambda: indicate(main_button, 'main'))
# main_button.place(relwidth=left_button_width, relheight=left_button_height, relx=0, rely=top_header_height)

# generator_button = tk.Button(left_frame, text="Form Generator", command=lambda: indicate(generator_button, 'generator'))
# generator_button.place(relwidth=left_button_width, relheight=left_button_height, relx=0, rely=top_header_height+1*left_button_height)

# arr_button = tk.Button(left_frame, text="Arrange", command=lambda: indicate(arr_button, 'arrange'))
# arr_button.place(relwidth=left_button_width, relheight=left_button_height, relx=0, rely=top_header_height+2*left_button_height)

#left_buttons = [main_button, generator_button, arr_button]

def indicate(button, frame):
    global current_frame
    global menu_session
    if frame != current_frame:
        clear_indicate()
        clear_frames()
        button.config(bg='#ff88a7')
        if frame == 'main':
            print('\nClicked main')
            current_frame = 'main'
            print('New current frame: ',current_frame)
            main()
        elif frame == 'generator':
            print('\nClicked generator')
            current_frame = 'generator'
            print('New current frame: ',current_frame)
            generator()
        elif frame == 'arrange':
            print('\nClicked arrange')
            current_frame = 'arrange'
            print('New current frame: ',current_frame)
            arrange()
        else:
            print('Frame Not Found!')
    else:
        print('\nClicked the same frame')
    menu_session = False

def clear_indicate():
    pass
    # for button in left_buttons:
    #     button.configure(bg="#ffffff")

def get_name_in_sheet(*args):
    print('traced')
    global list_name_in_sheet
    sheet_name = from_stv.get()
    excluded_current_sheet_names(sheet_name)
    sheet = df.parse(sheet_name)
    all_names = sheet.columns[3:]
    move_combo['values'] = list(all_names)
    return None

def excluded_current_sheet_names(exclude):
    get_current_sheet_names()
    moveto_combo['values'] = [x for x in sheet_names if x != exclude]
    return None

#tree functions
def tree_pop_up_menu(event):
    global current_focus_column_index, current_focus_row_index
    column = tv1.identify_column(event.x)
    row = tv1.identify_row(event.y)
    current_focus_row_index = tv1.index(row)
    #print(tv1.item(row), tv1.index(row), column)
    col_index = int(column[-1])-1
    current_focus_column_index = col_index
    if tv1.item(row).get('values') != '':
        if col_index > 2:
            selected_val = tv1.item(row).get('values')
            selected_val = selected_val[col_index]
            if selected_val == 1:
                tree_menu_body_0.tk_popup(event.x_root, event.y_root)
            else:
                tree_menu_body_1.tk_popup(event.x_root, event.y_root)
    else:
        tree_menu_columns.tk_popup(event.x_root, event.y_root)

def set_table_value(val):
    focus_sheet = df.parse(current_display_sheet)
    focus_sheet.iloc[current_focus_row_index, current_focus_column_index] = val
    writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
    focus_sheet.to_excel(writer, sheet_name=current_display_sheet, index=False)
    writer.close()
    col = focus_sheet.columns[current_focus_column_index]
    row_date = focus_sheet.iloc[current_focus_row_index,0]
    row_time_start = focus_sheet.iloc[current_focus_row_index,1]
    row_time_end = focus_sheet.iloc[current_focus_row_index,2]
    print('Row_date :',row_date,row_time_start,row_time_end,'Col :',col,'Value',val)
    update_val = {'Column':col,'Row_date':row_date,'Row_time_start':row_time_start,'Row_time_end':row_time_end,'Sheet':current_display_sheet,'Value':val}
    update_backup(mode='table_val', dict_val=update_val)
    apply_changes(current_display_sheet)
    return None

def create_sub_menu(*args):
    global menu_session
    sub_menu = tk.Menu(tree_menu_columns, tearoff=0, bg = primary, fg = font_color)
    print('Sub menu created')
    for sheet in [x for x in df_sheet_names if x != dy_current_display_sheet.get()]:
        print(sheet,'DONE')
        sub_menu.add_command(label=sheet, command=lambda x=sheet: move(mode='pop_up', to=x))
    if menu_session:
        #delete old cascade
        tree_menu_columns.delete('Move')
    tree_menu_columns.add_cascade(label='Move',menu=sub_menu)
    menu_session = True
    return None

def delete(mode = 'column',from_ = current_display_sheet):
    if mode == 'column':
        #print('Focusing on',current_display_sheet,'column : ',current_focus_column_index)
        focus_sheet = df.parse(from_)
        del_col = focus_sheet.columns[current_focus_column_index]
        focus_sheet = focus_sheet.drop(columns=[del_col])
        writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
        focus_sheet.to_excel(writer, sheet_name=from_, index=False)
        writer.close()
        update_val = {'Column': del_col, 'Sheet': from_}
        print(update_val)
        update_backup(mode='delete_col',dict_val=update_val)
    apply_changes(current_display_sheet)

def update_progress():
    best_round_score = genetic.get_best_round_score()
    finish = genetic.get_run_status()
    current_process_round = genetic.get_current_round()
    efficeint_score_label.configure(text=str(best_round_score))
    current_round_time = genetic.get_round_time()
    
    progress_show_label.configure(text=str(current_process_round)+'/'+str(ng)+'(ETA :'+\
                                    '%.2f'%(current_round_time*(ng-current_process_round)/60)+'mins)')
    option_frame.update()
    if finish:
        efficeint_score_label.configure(text='done')
        return None
    root.after(100,update_progress)

def perform_genetic(mode,qt):
    global efficeint_score_label
    global progress_show_label
    global ng #num_gen
    #root.update()
    if file_session:
        is_done = False
        avb_threshold = int(available_percentage_threshold_combo.get().strip('%'))
        if mode == 'ma':
            cf_pnm_const = 1.0
            avb_pnm_const = 0.25
        elif mode == 'lc':
            cf_pnm_const = 1.2
            avb_pnm_const = 0.0
        elif mode == 'bl':
            cf_pnm_const = 1.0
            avb_pnm_const = 0.0
        #print(avb_threshold,cf_pnm_const,avb_pnm_const)
        if qt == 'a':
            np = 10
            ng = 50
        elif qt == 'h':
            np = 25
            ng = 350
        elif qt == 'b':
            np = 30
            ng = 750
        #print(np,ng)

        #set up value
        efficeint_label = tk.Label(option_frame,text='Efficiency Score :').place(relx=0.5,rely=0.8,anchor=tk.CENTER)
        progress_label = tk.Label(option_frame,text='Progress :').place(relx=0.5,rely=0.7,anchor=tk.CENTER)
        #option_frame.update()
        #root.update()
        efficeint_score_label = tk.Label(option_frame,text='0')
        efficeint_score_label.place(relx=0.5,rely=0.85,anchor=tk.CENTER)
        progress_show_label = tk.Label(option_frame,text='0/'+str(ng))
        progress_show_label.place(relx=0.5,rely=0.75,anchor=tk.CENTER)
        genetic.setup_value(df,cf_pnm_const,avb_threshold/100,avb_pnm_const)
        threading.Thread(target=genetic.run,args=(np,ng,False)).start()
        #efficeint_score_label.configure(text='1')
        threading.Thread(target=update_progress).start()
    else:
        print('Browse a file first')
        


#set default color to left panel' buttons
clear_indicate()

#----------Center frames----------

#Center
center_frame = tk.Frame(root)
center_frame.place(relheight=1-margin,relwidth=center_x_ratio-margin,relx=left_x_ratio+margin,rely=0)

#Main
def main():
    global main_frame
    main_frame = tk.Frame(center_frame)
    main_frame.place(relheight=1,relwidth=generator_x_ratio,relx=0,rely=0)
    test_lb = tk.Label(main_frame, text = 'Main').pack(pady=20)

#Generator
def generator():
    global generator_frame
    generator_frame = tk.Frame(center_frame)
    generator_frame.place(relheight=1,relwidth=generator_x_ratio,relx=0,rely=0)

    # global sheets_frame
    # sheets_frame = tk.LabelFrame(generator_frame, text="Songs")
    # sheets_frame.place(relheight=display_height,relwidth=sheets_frame_width,relx=0,rely=top_header_height)

    global views_frame
    views_frame = tk.LabelFrame(generator_frame, text="Views")
    views_frame.place(relheight=views_height,relwidth=display_width,relx=sheets_frame_width, rely=top_header_height)

    global data_frame
    data_frame = tk.LabelFrame(generator_frame, text="Excel Data")
    data_frame.place(relheight=display_height-views_height, relwidth=display_width,relx=sheets_frame_width,rely=top_header_height+views_height)

    global editor_frame
    editor_frame = tk.Frame(generator_frame)
    editor_frame.place(relheight=editor_height,relwidth=1,relx=0,rely=top_header_height+display_height)

    global time_frame
    time_frame = tk.LabelFrame(editor_frame, text="Time")
    time_frame.place(relheight=tool_height,relwidth=time_width,relx=0,rely=browse_button_height)
    
    global add_frame
    add_frame = tk.LabelFrame(editor_frame, text="Add")
    add_frame.place(relheight=tool_height,relwidth=add_width,relx=time_width,rely=browse_button_height)

    global move_frame
    move_frame = tk.LabelFrame(editor_frame, text="Move")
    move_frame.place(relheight=tool_height,relwidth=del_width,relx=time_width+add_width,rely=browse_button_height)

    global browse_button
    browse_button = tk.Button(editor_frame, text="Browse A Form", command=lambda: File_dialog(mode = 'browse'))
    browse_button.place(relheight=browse_button_height, relwidth=browse_button_width, relx=1-browse_button_width, rely=0)

    global build_button
    build_button = tk.Button(editor_frame, text="Build With Template", command=lambda: File_dialog(mode = 'build'))
    build_button.place(relheight=browse_button_height, relwidth=browse_button_width, relx=1-2*browse_button_width, rely=0)

    global tv1
    tv1 = ttk.Treeview(data_frame)
    tv1.place(relheight=1, relwidth=1)
    
    global treescroll_x
    global treescroll_y
    treescroll_y = tk.Scrollbar(data_frame, orient='vertical',command=tv1.yview)
    treescroll_x = tk.Scrollbar(data_frame, orient='horizontal', command=tv1.xview)
    tv1.configure(xscrollcommand=treescroll_x.set,yscrollcommand=treescroll_y.set)
    treescroll_x.pack(side='bottom', fill='x')
    treescroll_y.pack(side='right', fill='y')
    tv1.bind('<Button-3>', tree_pop_up_menu)

    #right frame
    global right_frame
    right_frame = tk.Frame(center_frame)
    right_frame.place(relheight=1,relwidth=right_x_ratio-margin,relx=generator_x_ratio,rely=0)

    global option_frame
    option_frame = tk.LabelFrame(right_frame,text='Generator')
    option_frame.place(relheight=1-margin,relwidth=1,relx=0,rely=top_header_height)

    global mode_option_frame
    mode_option_frame = tk.LabelFrame(option_frame, text='Mode')
    mode_option_frame.place(relheight=mode_option_height,relwidth=1-4*margin,relx=2*margin,rely=views_height+top_header_height)

    def enable_percentage():
        if gen_mode.get()=='ma':
            available_percentage_threshold_combo.configure(state='normal')
            available_percentage_threshold_combo.set('70%')
        else:
            available_percentage_threshold_combo.configure(state='disabled')
            available_percentage_threshold_combo.set('0%')
    tk.Radiobutton(mode_option_frame,text='Most Available',variable=gen_mode,value='ma',command=lambda:enable_percentage()).pack(anchor=tk.NW)
    tk.Radiobutton(mode_option_frame,text='Least Conflict',variable=gen_mode,value='lc',command=lambda:enable_percentage()).pack(anchor=tk.W)
    tk.Radiobutton(mode_option_frame,text='Balance',variable=gen_mode,value='bl',command=lambda:enable_percentage()).pack(anchor=tk.SW)

    global available_entry_label,available_percentage_threshold_combo
    available_entry_label = tk.Label(option_frame, text='Available Percentage')
    available_entry_label.place(relx=2*margin,rely=views_height+mode_option_height+margin)
    available_percentage_threshold_combo = ttk.Combobox(option_frame, values=[str(x)+'%' for x in range(50,100,5)])
    available_percentage_threshold_combo.place(relwidth=1-4*margin,relx=2*margin,rely=views_height+mode_option_height+2.5*margin)
    available_percentage_threshold_combo.configure(state='disable')
    enable_percentage()

    global quality_frame
    quality_frame = tk.LabelFrame(option_frame, text='Quality')
    quality_frame.place(relheight=quality_frame_height,relwidth=1-4*margin,relx=2*margin,rely=views_height+mode_option_height+5*margin)

    tk.Radiobutton(quality_frame,text='Acceptable',variable=gen_quality,value='a').pack(anchor=tk.NW)
    tk.Radiobutton(quality_frame,text='High',variable=gen_quality,value='h').pack(anchor=tk.W)
    tk.Radiobutton(quality_frame,text='Best',variable=gen_quality,value='b').pack(anchor=tk.SW)

    global generate_button
    generate_button = tk.Button(option_frame, text='Generate',command=lambda: perform_genetic(gen_mode.get(),gen_quality.get()))
    generate_button.place(rely=0.65,relx=0.5,anchor=tk.CENTER)

    generator_frame.update()

    global date_start_label, time_start_label, start_colon_label, stop_colon_label, date_to_label, time_to_label, per_slot_label, room_label
    global date_start_fill, date_stop_fill
    global hour_start_combo, min_start_combo, hour_stop_combo, min_stop_combo, per_slot_combo, room_combo
    global apply_button

    hours = ['00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23']
    mins = ['00','30']
    date_start_label = tk.Label(time_frame, text='Date')
    time_start_label = tk.Label(time_frame, text='Time')
    start_colon_label = tk.Label(time_frame, text=':')
    stop_colon_label = tk.Label(time_frame, text=':')
    date_to_label = tk.Label(time_frame, text='to')
    time_to_label = tk.Label(time_frame, text='to')
    per_slot_label = tk.Label(time_frame, text='Hours per slot')
    date_start_fill = DateEntry(time_frame, selectmode='day')
    date_stop_fill = DateEntry(time_frame, selectmode='day')
    hour_start_combo = ttk.Combobox(time_frame, values=hours)
    min_start_combo = ttk.Combobox(time_frame, values=mins)
    hour_stop_combo = ttk.Combobox(time_frame, values=hours)
    min_stop_combo = ttk.Combobox(time_frame, values=mins)
    per_slot_combo = ttk.Combobox(time_frame, values=[0.5,1,2])
    room_label = tk.Label(time_frame,text='Room')
    room_combo = ttk.Combobox(time_frame)
    apply_button = tk.Button(time_frame, text='Apply',command=lambda: apply_time())

    date_start_label.grid(row=0, column=0, pady=5)
    date_start_fill.grid(row=0, column=1, pady=5)
    date_to_label.grid(row=0,column=2, pady=5)
    date_stop_fill.grid(row=0,column=3, pady=5)
    time_start_label.grid(row=1, column=0, pady=5)
    hour_start_combo.grid(row=1,column=1, pady=5)
    start_colon_label.grid(row=1,column=2, pady=5)
    min_start_combo.grid(row=1,column=3, pady=5)
    hour_stop_combo.grid(row=2,column=1, pady=5)
    stop_colon_label.grid(row=2,column=2, pady=5)
    min_stop_combo.grid(row=2,column=3, pady=5)
    per_slot_label.grid(row=3,column=0, pady=5)
    per_slot_combo.grid(row=3,column=1, pady=5)
    room_label.grid(row=3,column=2, pady=5)
    room_combo.grid(row=3,column=3, pady=5)
    apply_button.grid(row=4,column=2, pady=5)

    global name_label, part_label, addto_label, add_button
    global name_fill, part_combo, addto_combo
    global sheet_names
    sheet_names = []
    name_label = tk.Label(add_frame, text='Name')
    name_fill = tk.Entry(add_frame)
    part_label = tk.Label(add_frame, text='Part')
    part_combo = ttk.Combobox(add_frame, values=['Vocal','Guitar','Piano','Bass','Percussion'])
    addto_label = tk.Label(add_frame, text='Add To')
    addto_combo = ttk.Combobox(add_frame, values=sheet_names)
    add_button = tk.Button(add_frame, text='Add', command=lambda: add())

    name_label.grid(row=0, column=0)
    part_label.grid(row=1, column=0)
    addto_label.grid(row=2, column=0)
    name_fill.grid(row=0, column=1, pady=5)
    part_combo.grid(row=1, column=1, pady=5)
    addto_combo.grid(row=2, column=1, pady=5)
    add_button.grid(row=3,column=1, pady=5)

    global from_label, move_label, moveto_label
    global from_combo, move_combo, moveto_combo
    global from_stv, list_name_in_sheet, list_other_sheet
    from_stv = tk.StringVar()
    list_name_in_sheet = []
    list_other_sheet = []
    from_label = tk.Label(move_frame, text='From')
    from_combo = ttk.Combobox(move_frame, values=sheet_names, textvariable=from_stv)
    move_label = tk.Label(move_frame, text='Move')
    move_combo = ttk.Combobox(move_frame, values=list_name_in_sheet)
    moveto_label = tk.Label(move_frame, text='Move To')
    moveto_combo = ttk.Combobox(move_frame, values=list_other_sheet)
    move_button = tk.Button(move_frame, text='Move', command=lambda: move())
    from_stv.trace('w', get_name_in_sheet)

    from_label.grid(row=0, column=0)
    move_label.grid(row=1, column=0)
    moveto_label.grid(row=2, column=0)
    from_combo.grid(row=0, column=1, pady=5)
    move_combo.grid(row=1, column=1, pady=5)
    moveto_combo.grid(row=2, column=1, pady=5)
    move_button.grid(row=3,column=1, pady=5)

    #Tree menu
    global tree_menu_body_0, tree_menu_body_1, tree_menu_columns
    tree_menu_body_0 = Menu(tv1, tearoff=0, bg='black', fg='white')
    tree_menu_body_0.add_command(label='Set 0',command = lambda: set_table_value(0))

    tree_menu_body_1 = Menu(tv1, tearoff=0, bg='black', fg='white')
    tree_menu_body_1.add_command(label='Set 1',command = lambda: set_table_value(1))

    tree_menu_columns = Menu(tv1, tearoff=0, bg='black', fg='white')
    tree_menu_columns.add_command(label='Rename')
    tree_menu_columns.add_command(label='Delete', command=lambda: delete())


    global file_session
    if file_session:
        restore_excel_data(goto='Summary')
        print('\nRestored file: ',filename)
    else:
        print('\nNo file opened yet')

#Arrange
def arrange():
    global arrange_frame
    arrange_frame = tk.Frame(center_frame)
    arrange_frame.place(relheight=1,relwidth=generator_x_ratio,relx=0,rely=0)

    sum_result_width = 0.4
    sum_result_height = display_height
    sum_result_width_offset = -0.1
    mid_frame_width_offset = 0.1
    mid_frame_width = 1-mid_frame_width_offset-2*sum_result_width
    
    global ar_sum_frame
    ar_sum_frame = tk.LabelFrame(arrange_frame, text='Summary')
    ar_sum_frame.place(relwidth=sum_result_width-sum_result_width_offset,relheight=sum_result_height,relx=0,rely=top_header_height)

    room_frame_width = 1
    room_frame_height = 0.1

    global ar_room_frame
    ar_room_frame = tk.LabelFrame(ar_sum_frame, text='Rooms',pady=5)
    ar_room_frame.place(relwidth=room_frame_width,relheight=room_frame_height,relx=0,rely=0)

    global sum_tv
    sum_tv = ttk.Treeview(ar_sum_frame)
    sum_tv.place(relwidth=room_frame_width,relheight=1-room_frame_height,relx=0,rely=room_frame_height)

    global ar_result_frame
    ar_result_frame = tk.LabelFrame(arrange_frame, text='Result')
    ar_result_frame.place(relwidth=sum_result_width+sum_result_width_offset+mid_frame_width_offset,relheight=sum_result_height,relx=mid_frame_width+sum_result_width
                                        -sum_result_width_offset,rely=top_header_height)

    #test_lb2 = tk.Label(arrange_frame, text = 'Arrange').pack(pady=20)

def clear_frames():
    for frame in center_frame.winfo_children():
        frame.destroy()

#----------Initialize----------
generator()


#------------Styling-----------
# for frame in root.winfo_children():
#     frame.configure(bg=primary )

# for frame in center_frame.winfo_children():
#     frame.configure(bg=primary)

# for frame in generator_frame.winfo_children():
#     frame.configure(bg=primary)


def File_dialog(mode):
    global filename, template_path
    if mode == 'browse':
        filename = filedialog.askopenfilename(initialdir="/",
                                                title='Select a File', filetypes=[('xlsx files','.xlsx'),('All Files','.*')] )   
        print(filename)
        #label_file['text'] = filename
        Load_excel_data(mode='browse')
        #display_sheet(df.sheet_names[0])
    elif mode == 'build':
        template_path = filedialog.askopenfilename(initialdir="/",
                                                title='Select a File', filetypes=[('xlsx files','.xlsx'),('All Files','.*')])
        template_filename = r"{}".format(template_path)
        template_builder.build_template(template_filename,hr_per_slot=1)
        Load_excel_data(mode='build')

def Load_excel_data(mode):
    global df
    global file_path
    global file_session
    global df_sheet_names
    if mode == 'browse':
        file_path = filename
    elif mode == 'build':
        file_path = 'Schedule.xlsx'
        pass 
    try:
        excel_filename = r"{}".format(file_path)
        df = pd.ExcelFile(excel_filename)
        file_session = True
    except ValueError:
        messagebox.showerror("Information", "Chosen file invalid.")
        return None
    except FileNotFoundError:
        messagebox.showerror("Information",f"No such File as {file_path}")
        return None
    df_sheet_names = [x for x in df.sheet_names if x not in ['Summary','Request','Rooms']]
    for sheet in df_sheet_names:
        backup_time_dict[sheet] = df.parse(sheet)
    #clear_data()
    #create_songs_panel(df)
    create_views_panel(df)
    create_room_data(df)
    #set first sheet to be displayed first
    print('\nDisplaying sheet: ',df_sheet_names[0])
    update_sheets_combobox()
    display_sheet(df_sheet_names[0])

def restore_excel_data(goto):
    #create_songs_panel(df)
    create_views_panel(df)
    create_room_data(df)
    print('\nDisplaying sheet: ',df_sheet_names[0])
    update_sheets_combobox()
    display_sheet(goto)

def create_room_data(data):
    global room_data_dict
    global room_sheet
    room_sheet = data.parse('Rooms')
    room_list = room_sheet['Room'].unique()
    room_data_dict = {}
    for room in room_list:
        room_data_dict[room] = (room_sheet.loc[room_sheet['Room']==room]['Date Start'].values[0],
                                room_sheet.loc[room_sheet['Room']==room]['Date End'].values[0],
                                room_sheet.loc[room_sheet['Room']==room]['Time Start'].values[0],
                                room_sheet.loc[room_sheet['Room']==room]['Time End'].values[0],
                                room_sheet.loc[room_sheet['Room']==room]['Overnight'].values[0])
    room_combo['values'] = list(room_data_dict.keys())
    print(room_data_dict)

def update_sheets_combobox():
    #update sheet names
    get_current_sheet_names()
    #update combo box values
    addto_combo['values'] = sheet_names
    from_combo['values'] = sheet_names

def clear_data():
    #delte previous table
    tv1.delete(*tv1.get_children())
    pass

# def create_songs_panel(excel_file):
#     rely = 0
#     #destroy existing buttons
#     for button in sheets_frame.winfo_children():
#         button.destroy()
#     for sheet in excel_file.sheet_names:
#         button = tk.Button(sheets_frame, text=sheet, command=lambda sheet=sheet: display_sheet(sheet))
#         button.place(relheight=sheet_button_height,relwidth=1,rely=rely)
#         #print(sheet,"packed")
#         rely += sheet_button_height
#     #clear_button = tk.Button(sheets_frame, text="+", command=lambda: add_sheet())
#     #clear_button.place(relheight=sheet_button_height,relwidth=1,rely=rely)

def create_views_panel(excel_file):
    relx = 0
    #destroy existing buttons
    for button in views_frame.winfo_children():
        button.destroy()
    for view in excel_file.sheet_names:
        button = tk.Button(views_frame, text=view, command=lambda view=view: display_sheet(view))
        button.place(relheight=1,relwidth=views_button_width,relx=relx)
        #print(sheet,"packed")
        relx += views_button_width
    #clear_button = tk.Button(sheets_frame, text="+", command=lambda: add_sheet())
    #clear_button.place(relheight=sheet_button_height,relwidth=1,rely=rely)

def add_sheet():
    sheet_name = simpledialog.askstring("Add Song","Song Name :")
    old_sheet_sample = df.parse(df_sheet_names[0]) #use first sheet as a sample to construct new sheet
    add_columns = old_sheet_sample.columns[0:3] #select main columns to add to new sheet
    new_sheet = pd.DataFrame(columns=add_columns)
    for column in add_columns:
        new_sheet[column] = old_sheet_sample[column]
    writer = pd.ExcelWriter(file_path, engine = 'openpyxl', mode= 'a', if_sheet_exists = 'replace')
    new_sheet.to_excel(writer, sheet_name = sheet_name, index=False)
    writer.close()
    #apply changes and update values
    apply_changes('Summary')
    return None

def apply_changes(goto):
    #load new df
    load_new_df()
    #update combo box values
    update_sheets_combobox()
    #refresh interface
    restore_excel_data(goto)
    return None

def load_new_df():
    global df
    global df_sheet_names
    excel_filename = r"{}".format(file_path)
    df = pd.ExcelFile(excel_filename)
    df_sheet_names = [x for x in df.sheet_names if x not in ['Summary','Request','Rooms']]

def display_sheet(sheet_name):
    global current_display_sheet
    current_display_sheet = sheet_name
    global dy_current_display_sheet
    dy_current_display_sheet = tk.StringVar()
    dy_current_display_sheet.set(sheet_name)
    dy_current_display_sheet.trace('w',create_sub_menu('Move',current_display_sheet))
    print("Got sheet:",sheet_name)
    clear_data()
    sheet = df.parse(sheet_name)
    num_cols = len(sheet.columns)
    if num_cols<12:
        col_width = tv1.winfo_width()//num_cols
    else:
        col_width = tv1.winfo_width()//12
    tv1['column'] = list(sheet.columns)
    tv1['show'] = 'headings'
    for column in tv1['columns']:
        tv1.column(column,stretch=False,width=col_width)
        tv1.heading(column, text=column)
    sheet_rows = sheet.to_numpy().tolist()
    for row in sheet_rows:
        tv1.insert("","end", values=row)
    return None

def get_current_sheet_names():
    global df
    global sheet_names
    if file_session:
        sheet_names = [x for x in df.sheet_names if x not in ['Summary','Request','Rooms']]
    return None

def add(mode = 'normal', dict_val={},cont_from='',cont_from_col='', to=''):
    if mode == 'normal':
        name = name_fill.get()
        part = part_combo.get()
        addto = addto_combo.get()
        content = 0
    elif mode == 'pop_up':
        focus_sheet = df.parse(to)
        name = cont_from_col
        #part = 
        addto = to
        backup_content = df.parse(cont_from).loc[:,['Date','Time Start','Time End',name]]
        content = df.parse(cont_from)[name].values
    adding_sheet = df.parse(addto)
    adding_sheet[name] = content
    writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
    adding_sheet.to_excel(writer, sheet_name=addto, index=False)
    writer.close()
    update_val = {'Name':name,'Content':backup_content,'Sheet':addto}
    update_backup(mode='add',dict_val=update_val)
    apply_changes(addto)

def move(mode = 'normal',from_='', to=''):
    if mode == 'normal':
        move_from = from_combo.get()
        name = move_combo.get()
        move_to = moveto_combo.get()
    elif mode == 'pop_up':
        move_from = current_display_sheet #move from displaying sheet
        focus_sheet = df.parse(move_from)
        name = focus_sheet.columns[current_focus_column_index]
        move_to = to
        add(mode='pop_up',cont_from=move_from,cont_from_col=name, to=move_to)
        delete(mode='column', from_=move_from)
        apply_changes(move_to)
        return None
    sheet_from = df.parse(move_from)
    column_data = sheet_from[name]
    sheet_to = df.parse(move_to)
    sheet_to[name] = column_data
    sheet_from = sheet_from.drop(columns=[name])
    writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
    sheet_from.to_excel(writer, sheet_name=move_from, index=False)
    sheet_to.to_excel(writer, sheet_name=move_to, index=False)
    writer.close()
    apply_changes(move_to)
    return None

def apply_time():
    global time_change_session
    date_start = date_start_fill.get_date()
    date_stop = date_stop_fill.get_date()
    hr_start = hour_start_combo.get()
    min_start = min_start_combo.get()
    hr_stop = hour_stop_combo.get()
    min_stop = min_stop_combo.get()
    hr_per_slot = per_slot_combo.get()
    room = room_combo.get()
    
    #clening data
    date_start = date_start.strftime('%Y-%m-%d')
    date_stop = date_stop.strftime('%Y-%m-%d')
    time_start = hr_start+':'+min_start+':00'
    time_stop = hr_stop+':'+min_stop+':00'
    print(date_start,'-',date_stop,'\n'+time_start,'-',time_stop)

    writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
    #update in table
    room_sheet.loc[room_sheet['Room'] == room,'Date Start'] = date_start
    room_sheet.loc[room_sheet['Room'] == room,'Date End'] = date_stop
    room_sheet.loc[room_sheet['Room'] == room,'Time Start'] = time_start
    room_sheet.loc[room_sheet['Room'] == room,'Time End'] = time_stop
    if time_stop<time_start:
        room_sheet.loc[room_sheet['Room'] == room,'Overnight'] = 'Yes'
    else:
        room_sheet.loc[room_sheet['Room'] == room,'Overnight'] = 'No'

    room_sheet.to_excel(writer, sheet_name = 'Rooms', index=False)

    min_date_start = room_sheet['Date Start'].min()
    max_date_end = room_sheet['Date End'].max()
    min_time_start = room_sheet['Time Start'].min()

    print(room_sheet)
    if len(room_sheet['Overnight']) != 0:
        is_overnight = True
        overnight_room_sheet = room_sheet.loc[room_sheet['Overnight'] == 'Yes']
        max_time_end = overnight_room_sheet['Time End'].max()
    else:
        is_overnight = False
        max_time_end = room_sheet['Time End'].max()
 

    if min_date_start < date_start:
        date_start = min_date_start
    if max_date_end > date_stop:
        date_stop = max_date_end
    if min_time_start < time_start:
        time_start = min_time_start
    if (time_start < time_stop): #not overnight
        if is_overnight:
            time_stop = max_time_end
        elif (time_stop < max_time_end):
            time_stop = max_time_end
    elif (time_start > time_stop): #overnight
        if is_overnight and (max_time_end > time_stop):
            time_stop = max_time_end
    
    print(date_start,'-',date_stop,'\n'+time_start,'-',time_stop)

    
    # if room_time_start<time_start:
    #     time_start = room_time_start
    # if (room_time_stop>time_stop and time_stop>time_start and room_overnight == 'No') or \
    #                 (room_time_stop>time_stop and time_start<time_stop and room_overnight == 'Yes'): #22.00 vs 23.00 or 2.00 vs 3.00
    #     time_stop = room_time_stop
    # elif (time_stop>time_start and room_overnight == 'Yes'): #23.00 vs 2.00
    #     time_stop = room_time_stop
    # if room_date_start < date_start:
    #     date_start = room_date_start
    # if room_date_stop > date_stop:
    #     date_stop = room_date_stop

    new_time = pd.DataFrame(columns=['Date','Time Start','Time End'])
    #print(date_start, date_stop)
    dates = pd.date_range(date_start, date_stop, freq='1D')
    for date in dates:
        date_stop = date
        if (time_stop<time_start or time_stop == '00:00:00'):
            date_stop += dt.timedelta(days=1)
        date = pd.to_datetime(date).strftime('%Y-%m-%d')
        date_stop = pd.to_datetime(date_stop).strftime('%Y-%m-%d')
        #print('\n\n',date)
        start = pd.date_range(date+' '+time_start, date_stop+' '+time_stop, freq=str(hr_per_slot)+'H', closed='left')
        stop = pd.date_range(date+' '+time_start, date_stop+' '+time_stop, freq=str(hr_per_slot)+'H', closed='right')
        sheet = pd.DataFrame(
            {'Date': start.date,
            'Time Start': start.time,
            'Time End': stop.time}
        )
        new_time = pd.concat([new_time,sheet], ignore_index=True)
    #print(new_time)
    new_time['Date'] = new_time['Date'].apply(lambda x: x.strftime('%Y-%m-%d'))
    new_time['Time Start'] = new_time['Time Start'].apply(lambda x: x.strftime('%H:%M:%S'))
    new_time['Time End'] = new_time['Time End'].apply(lambda x: x.strftime('%H:%M:%S'))
    for sheet in df_sheet_names:
        #old_time = df.parse(sheet)
        #if time_change_session:
        old_time = backup_time_dict.get(sheet)
        #print('\nUse back-up data')
        new_time_to_save = pd.merge(new_time, old_time, how='left', on = ['Date','Time Start','Time End'])
        new_time_to_save = new_time_to_save.fillna(0)
        update_val = {'New time':new_time_to_save,'Old time':old_time,'Sheet':sheet}
        update_backup(mode='time',dict_val=update_val)
        new_time_to_save.to_excel(writer, sheet_name = sheet, index=False)
    writer.close()
    time_change_session = True
    apply_changes('Rooms')
    # for sheet in df.sheet_names:
    #     print(backup_time_dict.get(sheet))
    return None

def update_backup(mode,dict_val):
    if mode == 'time':
        backup_time_data = pd.merge(dict_val.get('New time'), dict_val.get('Old time'), how='outer', on=['Date','Time Start','Time End'], suffixes=['_x',''])
        backup_time_data.drop(backup_time_data.filter(regex='_x$').columns, axis=1, inplace=True)
        backup_time_data = backup_time_data.sort_values(by=['Date','Time Start'],ignore_index=True)
        backup_time_data = backup_time_data.fillna(0)
        backup_time_dict[dict_val.get('Sheet')] = backup_time_data
    elif mode == 'table_val':
        focus_backup = backup_time_dict.get(dict_val.get('Sheet'))
        focus_backup.loc[(focus_backup['Date'] == dict_val.get('Row_date')) & (focus_backup['Time Start'] == dict_val.get('Row_time_start')) & (focus_backup['Time End'] == dict_val.get('Row_time_end')),dict_val.get('Column')] = dict_val.get('Value')
        #print(focus_backup)
        backup_time_dict[dict_val.get('Sheet')] = focus_backup
    elif mode == 'add':
        content =  backup_time_dict.get(current_display_sheet)
        content = content.loc[:,['Date','Time Start','Time End',dict_val.get('Name')]]
        backup_add_data = pd.merge(backup_time_dict.get(dict_val.get('Sheet')), content, how='left', on=['Date','Time Start','Time End'], suffixes=['_x',''])
        backup_add_data.drop(backup_add_data.filter(regex='_x$').columns, axis=1, inplace=True)
        backup_add_data = backup_add_data.fillna(0)
        print('Add backuped')
        print(backup_add_data)
        backup_time_dict[dict_val.get('Sheet')] = backup_add_data
    elif mode == 'delete_col':
        focus_backup = backup_time_dict.get(dict_val.get('Sheet'))
        focus_backup = focus_backup.drop(columns=[dict_val.get('Column')])
        print('Delete backuped')
        print(focus_backup)
        backup_time_dict[dict_val.get('Sheet')] = focus_backup
    elif mode == 'delete_row':
        pass
    else:
        print('No such a mode :',mode)
    return None

root.mainloop()