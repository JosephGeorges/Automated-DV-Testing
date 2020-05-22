'''
This program reads the data from an excel file to generate ISO 16750-2 Load Dump Test B Pulse and sends the waveform to a Tektronix AFG1062.
Assigns the waveform to CH1 of the AFG1062 and sends 15 pulses.  

python 3.7.7 (https://www.python.org/downloads/)
pandas 1.0.3 (https://pandas.pydata.org/docs/index.html)
numpy 1.19   (https://numpy.org/devdocs/)
pyvisa 1.10.1 (https://pyvisa.readthedocs.io/en/latest/)
'''

import pandas as pd
import numpy as np
import pylab as plt
import pyvisa
import time

#Initialize list of data points 
filename = 'ISO_16750-2_2010 4_6_4b_load_dump.xlsx' ##NEEDS TO BE CHANGED TO FILE WITH 35V PEAK!!
df = pd.read_excel(filename)

waveform = []
waveform = np.array(df['Voltage'])

w_min, w_max = waveform.min(), waveform.max()
waveform = 2*(waveform - w_min)/(w_max - w_min) - 1

rm = pyvisa.ResourceManager()
afg = rm.open_resource('USB0::0x0699::0x0353::1813139::INSTR')
Inst_Name = afg.query("*IDN?")
print("Instrument Identified As: " + Inst_Name)
afg.timeout = 10000
afg.write('*rst')
afg.write('*cls')

num_points = len(waveform)
sample_period = 9.216e-5
sample_rate = 1/(sample_period) 
t = np.linspace(0,num_points/sample_rate, num_points, endpoint = True)
plt.plot(t,waveform)
#plt.show()

#Create an array with all values set to 0 in the above convention
to_transfer = np.ones(len(waveform), dtype=np.uint16)*(2**13)

# Convert your data to fit within the format
to_transfer += np.require(np.rint(8191*waveform), np.uint16)

#Check for errors
if to_transfer.max() > 16383 or to_transfer.min() < 0:
    raise ValueError('Analogical values out of range.')

afg.write('source1:function ememory') #Sets the AFG to ARB mode
afg.write('SOURce1:FREQuency:FIXed 1Hz') #Sets the period on AFG to 1 second. - Set timebase on SCOPE to 1s/divison.
afg.write('SOURce1:VOLTage:LEVel:IMMediate:AMPLitude 3.5Vpp')
afg.write('SOURce1:VOLTage:LEVel:IMMediate:OFFSet 1.75V')

#Write the data to the instrument
print("Writing waveform data to " + Inst_Name)
afg.write_binary_values("DATA EMEMory,", to_transfer, datatype='h', is_big_endian = True) #Copies the waveform into editable memory of AFG
print("Writing Data Complete...")

afg.query('*opc?') #Wait Command
afg.write('DATA:COPY USER5,EMEMory')#Copies waveform data from Edit Memory to USER5 memory location

#Loops the pulse for 15 minutes, sends 15 pulses, one per minute. 
for i in range(1,16):
    afg.query('*opc?')
    afg.write('source1:function user5')#Assigns waveform in USER5 location to CH1 of AFG
    afg.query('*opc?')
    afg.write('output1 on') #Turns on output 1 of the AFG
    print("Pulse Number " + str(i) + " Sent") 
    time.sleep(60) #time delay of 60 seconds before sending next pulse

print("Test Complete") 
