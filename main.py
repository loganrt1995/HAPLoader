import os
import sys
import pywinauto
from pywinauto import keyboard
import pandas as pd
import ctypes


# gets data from HAP Spreadsheet
def get_excel_data(excel_file):
    # gets two data frames from hap spreadsheet
    #table = pd.read_excel('C:/Users/ltaylor/Desktop/Logan/HAP Spreadsheet.xlsx', nrows=23, dtype=str)
    table = pd.read_excel(excel_file, nrows=23, dtype=str)
    table = table.fillna('')
    row_range = list(range(0,27)) + list(range(35,100))
    #small_table = pd.read_excel('C:/Users/ltaylor/Desktop/Logan/HAP Spreadsheet.xlsx', skiprows=row_range, usecols='A,B', dtype=str)
    small_table = pd.read_excel(excel_file, skiprows=row_range, usecols='A,B', dtype=str)

    # assigned project-wide variables from smaller dataframe
    window_height = small_table.iloc[1, 1]
    roof_load = small_table.iloc[2,1]
    floor_load = small_table.iloc[3,1]
    window_type = small_table.iloc[4, 1]
    file_name = small_table.iloc[5,1]

    # sets up pywinauto HAP
    app, dlg = hap_setup(window_height, window_type, file_name)

    # iterates over rows of bigger dataframe and calls auto_hap for each if it's name is not empty
    room_index = 1
    for index, row in table.iterrows():
        data_list = []
        if row.iloc[0] != '' and row.iloc[0] != 'NaN' and row.iloc[0] != 'nan':
            for index, item in row.iteritems():
                data_list.append(item)
            hap_spaces(data_list, roof_load, floor_load, app, dlg, room_index)
            room_index += 1

    # exits window after no activity for a few seconds
    ctypes.windll.user32.MessageBoxW(0, "Script complete.", "HAP Loader", 1)
    actionable_dlg = dlg.wait('visible')

# sets up hap (windows)
def hap_setup(window_height, window_type, file_name):
    from pywinauto.application import Application
    app = Application(backend='uia').start('C:/E20-II/HAP51/CODE/hap.exe') # this one for start up
    #app = Application(backend='uia').connect(path=r'hap.exe') # this one if HAP is already open

    # describe the window
    dlg = app['HAP51 - [Untitled]']  # specifies dialog

    #dlg.print_control_identifiers()    # prints all the controls

    # opens HAP template
    try:
        app['dlg']['OK'].click_input()
    except:
        ctypes.windll.user32.MessageBoxW(0, "Ensure HAP is not already running before using the program.",
                                         "HAP Loader", 1)
        sys.exit()
    app['dlg']["000HAP Loader Template"].click_input()
    app['dlg']['Open'].click_input()

    # checks if file_name already exists, deletes if it does
    projects_path = r'C:\E20-II\Projects'
    projects_list = os.listdir(projects_path)
    if file_name in projects_list:
        print(file_name + ' already exists, existing deleted')
        app['dlg']['Project'].click_input()
        pywinauto.keyboard.send_keys('{VK_DOWN 5}')
        pywinauto.keyboard.send_keys('{ENTER}')
        delete_test = False
        for i in range(10):
            try:
                pywinauto.keyboard.send_keys(file_name)
                app['dlg']['Delete'].click_input()
                app['dlg']['Yes'].click_input()
                delete_test = True
                break
            except:
                pywinauto.keyboard.send_keys('{VK_DOWN 20}')
        if delete_test == False:
            ctypes.windll.user32.MessageBoxW(0, "The existing HAP project could not be deleted, delete manually.", "HAP Loader", 1)
            sys.exit()
    else:
        print(file_name + ' does not already exist')

    # saves new HAP file
    app['dlg']['Project'].click_input()
    pywinauto.keyboard.send_keys('{VK_DOWN 4}')
    pywinauto.keyboard.send_keys('{ENTER}')
    pywinauto.keyboard.send_keys(file_name)
    pywinauto.keyboard.send_keys('{ENTER}')
    # builds window
    app['dlg']['Spaces'].click_input()
    pywinauto.keyboard.send_keys('{VK_DOWN 8}')
    pywinauto.keyboard.send_keys('{TAB 1}')
    pywinauto.keyboard.send_keys('{VK_DOWN 1}')
    pywinauto.keyboard.send_keys('{ENTER}')
    if window_type == 'Double Pane':
        pywinauto.keyboard.send_keys('Double Pane by Foot')
        pywinauto.keyboard.send_keys('{TAB 2}')
        pywinauto.keyboard.send_keys(window_height)
        pywinauto.keyboard.send_keys('{TAB 5}')
        pywinauto.keyboard.send_keys('{VK_DOWN 3}')
        pywinauto.keyboard.send_keys('{TAB 3}')
        pywinauto.keyboard.send_keys('{ENTER}')
    elif window_type == 'Single Pane':
        pywinauto.keyboard.send_keys('Single Pane by Foot')
        pywinauto.keyboard.send_keys('{TAB 2}')
        pywinauto.keyboard.send_keys(window_height)
        pywinauto.keyboard.send_keys('{TAB 8}')
        pywinauto.keyboard.send_keys('{ENTER}')
    else:
        print('window input did not work')

    return app, dlg


def hap_spaces(data_list, roof_load, floor_load, app, dlg, room_index):

    # space variables
    space_name = data_list[0]
    space_area = data_list[1]
    people = data_list[2]
    extra_btu = data_list[3]
    exp_1_dir = data_list[4]
    exp_1_area = data_list[6]
    exp_1_windows = data_list[7]
    exp_2_dir = data_list[8]
    exp_2_area = data_list[10]
    exp_2_windows = data_list[11]
    exp_3_dir = data_list[12]
    exp_3_area = data_list[14]
    exp_3_windows = data_list[15]

    directions = {
        'N': '1',
        'NE': '3',
        'E': '5',
        'SE': '7',
        'S': '9',
        'SW': '11',
        'W': '13',
        'NW': '15'
    }

    # builds space
    app['dlg']['Weather'].click_input()
    app['dlg']['Spaces'].click_input()
    pywinauto.keyboard.send_keys('{TAB 1}')
    pywinauto.keyboard.send_keys('{VK_DOWN ' + str(room_index) + '}')
    pywinauto.keyboard.send_keys('{ENTER}')

    # General Tab
    pywinauto.keyboard.send_keys(space_name)
    pywinauto.keyboard.send_keys('{TAB 1}')
    pywinauto.keyboard.send_keys(space_area)
    pywinauto.keyboard.send_keys('{TAB 8}')
    pywinauto.keyboard.send_keys('{VK_RIGHT}')

    # Internals Tab
    pywinauto.keyboard.send_keys('{TAB 15}')
    pywinauto.keyboard.send_keys(people)
    pywinauto.keyboard.send_keys('{TAB 5}')
    pywinauto.keyboard.send_keys(extra_btu)
    pywinauto.keyboard.send_keys('{TAB 9}')
    pywinauto.keyboard.send_keys('{VK_RIGHT}')

    # Walls, Windows, Doors Tab
    if exp_1_dir != 'None' and exp_1_dir != '':
        pywinauto.keyboard.send_keys('{TAB 1}')     # sets exposures
        pywinauto.keyboard.send_keys('{VK_DOWN ' + directions[exp_1_dir] + '}')
        pywinauto.keyboard.send_keys('{TAB 1}')
        pywinauto.keyboard.send_keys(exp_1_area)
        pywinauto.keyboard.send_keys('{TAB 1}')
        pywinauto.keyboard.send_keys(exp_1_windows)
        if exp_2_dir != 'None' and exp_2_dir != '':
            pywinauto.keyboard.send_keys('{TAB 3}')
            pywinauto.keyboard.send_keys('{VK_DOWN ' + directions[exp_2_dir] + '}')
            pywinauto.keyboard.send_keys('{TAB 1}')
            pywinauto.keyboard.send_keys(exp_2_area)
            pywinauto.keyboard.send_keys('{TAB 1}')
            pywinauto.keyboard.send_keys(exp_2_windows)
            if exp_3_dir != 'None' and exp_3_dir != '':
                pywinauto.keyboard.send_keys('{TAB 3}')
                pywinauto.keyboard.send_keys('{VK_DOWN ' + directions[exp_3_dir] + '}')
                pywinauto.keyboard.send_keys('{TAB 1}')
                pywinauto.keyboard.send_keys(exp_3_area)
                pywinauto.keyboard.send_keys('{TAB 1}')
                pywinauto.keyboard.send_keys(exp_3_windows)
                pywinauto.keyboard.send_keys('{TAB 23}')
                pywinauto.keyboard.send_keys('{VK_RIGHT}')
            else:
                pywinauto.keyboard.send_keys('{TAB 24}')
                pywinauto.keyboard.send_keys('{VK_RIGHT}')
        else:
            pywinauto.keyboard.send_keys('{TAB 25}')
            pywinauto.keyboard.send_keys('{VK_RIGHT}')
    else:
        pywinauto.keyboard.send_keys('{VK_RIGHT}')

    # Roof, Skylights Tab
    if roof_load == 'Yes':
        pywinauto.keyboard.send_keys('{TAB 1}')
        pywinauto.keyboard.send_keys('{VK_DOWN}')
        pywinauto.keyboard.send_keys('{TAB 7}')
        pywinauto.keyboard.send_keys('r')
        pywinauto.keyboard.send_keys('{TAB 6}')
        pywinauto.keyboard.send_keys('{VK_RIGHT 2}')
    else:
        pywinauto.keyboard.send_keys('{VK_RIGHT 2}')

    # Floors Tab then finish
    if floor_load == 'Above Conditioned':
        pywinauto.keyboard.send_keys('{TAB 2}')
        pywinauto.keyboard.send_keys('{ENTER}')
    elif floor_load == 'Above Unconditioned':
        pywinauto.keyboard.send_keys('{TAB 1}')
        pywinauto.keyboard.send_keys('{VK_DOWN}')
        pywinauto.keyboard.send_keys('{TAB 7}')
        pywinauto.keyboard.send_keys('{ENTER}')
    elif floor_load == 'Slab on Grade':
        pywinauto.keyboard.send_keys('{TAB 1}')
        pywinauto.keyboard.send_keys('{VK_DOWN 2}')
        pywinauto.keyboard.send_keys('{TAB 5}')
        pywinauto.keyboard.send_keys('{ENTER}')
    else:
        print('floor load had incorrect variable')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #excel_file = r"G:\3374000 - 15870 Midway Addison Tx\001 - Malin Company - 1\_M\1M-1P\HAP Spreadsheet - 1M-1P.xlsm"
    #ctypes.windll.user32.MessageBoxW(0, str(sys.argv[1]),
    #                                 "HAP Loader", 1)
    if len(sys.argv) == 1:
        ctypes.windll.user32.MessageBoxW(0, "No file path was passed to the program.",
                                         "HAP Loader", 1)
        sys.exit()
    else:
        excel_file = sys.argv[1]
    get_excel_data(excel_file)    # should run whole script - use this one

# INSTRUCTIONS: ENTER DATA INTO "HAP SPREADSHEET", SAVE SPREADSHEET, PRESS RUN.

# DO NOT USE IF NUMBER OF ROOMS EXCEED NUMBER OF ROWS IN "HAP SPREADSHEET"

# EVENTUALLY I WANT TO ADD A FUNCTION TO PUT HAP'S CFM INTO EACH ROOM ON CAD

# CURRENTLY THERE IS A BUG THAT IF THE FILE NAME ALREADY EXISTS IT WILL GO CRAZY.

# type the following into the teriminal to create exe:
# C:\Users\ltaylor\AppData\Local\Programs\Python\Python39\Scripts\pyinstaller main.py --onefile

