
#First we import all the Python bits and pieces needed to run the program
import sys                              # system functions
import os                               # Operating system functions
from email.header import UTF8           # needed to decode strings
import xlwings as xw                    # needed to communicate with Excel
from datetime import datetime, timedelta# date functions
from time import sleep                  # used for waiting
import numpy as np                      # Maths functions std dev and means

sys.path.append('U:/Flow_Cell_Leak_Test')  #add the path of the elve DLL library here

from array import array                 # We are going to put our data in an array
from ctypes import c_int32, c_double, byref     # The elve DLL uses C variable types

# Import the Elve functions we will need
from Elveflow64 import OB1_Initialization, OB1_Add_Sens, Elveflow_Calibration_Default, Elveflow_Calibration_Load, \
                        OB1_Calib,Elveflow_Calibration_Save,OB1_Set_Press,OB1_Get_Press,OB1_Get_Sens_Data, \
                        OB1_Get_Trig,OB1_Set_Trig, OB1_Destructor

wb = xw.books.active                    # link to the active work book


cnfSht = wb.sheets['Config']            #link to the config worksheet
tstSht = wb.sheets['Test']              #link to the test worksheet
dataSht = wb.sheets['Results']          #link to the results worksheet


#
# Initialization of OB1 (copied from their example)
# Note that the controller ID 01ED6986 is hard coded, use NIMAX to find
Instr_ID=c_int32()
error=OB1_Initialization('01ED6986'.encode('ascii'),5,5,0,0,byref(Instr_ID)) # connects to the controller
error=OB1_Add_Sens(Instr_ID, 1, 2, 1, 0, 7, 0)                               # connects to the flow sensor
Calib_path_def = str(tstSht.cells(1,6).value)                                # Load the calibration file
Calib=(c_double*1000)()                  #always define array this way, calibration should have 1000 elements
Calib_path=Calib_path_def
#Calib_path='C:\\temp\\Calibration\\Calib.txt'
error=Elveflow_Calibration_Load (Calib_path.encode('ascii'), byref(Calib), 1000)

set_channel=c_int32(1)                      # set the OB1 channel 1 by default - we could put this in the config worksheet
set_pressure=c_double()                     # initialise the pressure demand variable

data_sens=c_double()                        # Initialise the flow sensor variable
get_pressure=c_double()                     # Initialise the pressure sensor variable

the_row = 11                                # set the start row in the test worksheet
settle_time = tstSht.cells(5,4).value       # read in the settle time
read_delay = tstSht.cells(6,4).value        # read in the delay
while tstSht.cells(the_row,1).value > 0:    # start the test loop we set the last row to -1
    p_arr = []                              # empty the pressure array
    f_arr = []                              # empty the flow array
    print('testing {}'.format(tstSht.cells(the_row,2).value)) # print the test point (only visible in python)
    set_pressure=c_double(tstSht.cells(the_row,2).value)    #get the demand pressure
    error=OB1_Set_Press(Instr_ID.value, set_channel, set_pressure, byref(Calib),1000) # send the set pressure command to the OB1
    sleep(settle_time)                      #wait for stability
    for x in range(int(tstSht.cells(the_row,3).value)):  # start the data collecton loop (number of loops defined in the test)
        error=OB1_Get_Sens_Data(Instr_ID.value,set_channel, 1,byref(data_sens))                     #Read the flow
        error=OB1_Get_Press(Instr_ID.value, set_channel, 1, byref(Calib),byref(get_pressure), 1000) #Read the pressure
        p_arr.append(float(get_pressure.value)) # put the pressure into the pressure array
        f_arr.append(float(data_sens.value))    # put the flow into the flow array
        sleep(read_delay)                       # short wait to ensure new data
    #Save the values into the result sheet
    dataSht.cells(the_row,2).value = datetime.now()
    dataSht.cells(the_row,3).value = np.mean(f_arr)
    dataSht.cells(the_row,4).value = np.std(f_arr)
    dataSht.cells(the_row,5).value = np.mean(p_arr)
    dataSht.cells(the_row,6).value = np.std(p_arr)
    the_row += 1                                # increment row counter

# Save the data in a new sheet
newFilename = str(tstSht.cells(3,6).value) + '_' + str(datetime.now().strftime('%Y%m%d%H%M%S')) + '.xlsm' # Make the file name
the_path = str(cnfSht.cells(3,2).value) + str(tstSht.cells(2,6).value)                                    # Make the path
wb.save(os.path.join(the_path,newFilename))         #Save the file