#!/usr/bin/env python3

from ctypes import * 
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
import numpy as np
from datetime import datetime
import time
import smtplib
import dwf
import os

# Biblioteka DWF
dwf = cdll.LoadLibrary("libdwf.so")

# Zmienne parametry badania
AcqTime_Min = 5
AcqTime_S = 0
Hz = 4000           # 40 harmonicznych po 50Hz = 2000Hz, tw Shannona x2 = 4000Hz
Preview = 20        # Ilość okresów sygnału do podglądu na wykresie
Iteration = 120     # Ilość wykonań akwizycji w pętli

# Definicje zmiennych
hdwf = c_int()
sts = c_byte()
hzAcq = c_double(Hz)
AcqTime = c_double()
AcqTime = AcqTime_S+60*AcqTime_Min
Start_Time = datetime.now()

# Długość próbek dla napięcia i natężenia
Preview_Samples = Preview*80
Acquisition_Samples = AcqTime*80*50
# ilość sampli dla 1 sekundy akwizycji sygnału (4096/50)Hz = 81.92 sampli dla okresu, * 50 -> dla 1 sekundy
# 2147483647   max ilość sampli -> 8729.6 minut -> 145.5 godzin -> ok. 6 dni

cValid = c_int(0)
cAvailable = c_int()
cLost = c_int()
cCorrupted = c_int()
fLost = 0
fCorrupted = 0

Final_Harmonic_Voltage = (c_double*40)()
Final_Harmonic_Current = (c_double*40)()
Final_X_oneside_Voltage = (c_double*Acquisition_Samples)()
Final_X_oneside_Current = (c_double*Acquisition_Samples)()

# VERSION
version = create_string_buffer(16)
dwf.FDwfGetVersion(version)
print("DWF Version: "+str(version.value))


def THD(Data):
    Sum = 0.0
    for x in range(len(Data)):
        Sum = Sum + (Data[x])**2
    Harmonics = Sum - (max(Data))**2.0 
    THD = 100 * Harmonics**0.5 / max(Data)
    return THD
        
for i in range(1, Iteration+1):
    
    # Dostęp do oscyloskopu
    print("Opening device")
    dwf.FDwfDeviceConfigOpen(c_int(-1), c_int(1), byref(hdwf)) 
    if hdwf.value == 0:                        # jeśli 2 arg DeviceOpen zwróci 0 to sie nie połączył
        type_error = create_string_buffer(512)
        dwf.FDwfGetLastErrorMsg(type_error)
        print(str(type_error.value))
        print("Failed to open device")
        quit()
    print("Device has been open")
    
    # Ustawienia akwizycji
    dwf.FDwfAnalogInChannelEnableSet(hdwf, c_int(0), c_bool(True))
    dwf.FDwfAnalogInChannelEnableSet(hdwf, c_int(1), c_bool(True))
    dwf.FDwfAnalogInChannelRangeSet(hdwf, c_int(0), c_double(5))  # zakres w woltach
    dwf.FDwfAnalogInAcquisitionModeSet(hdwf, c_int(3)) # Tryb akwizycji danych oscyloskopu - acqmodeRecord 
    dwf.FDwfAnalogInFrequencySet(hdwf, hzAcq)  # Częstotliwosć próbkowania
    dwf.FDwfAnalogInRecordLengthSet(hdwf, c_double(Acquisition_Samples/Hz)) # Długość akwizycji sygnału

    # 1s stabilizacji i uruchomienie oscyloskopu
    time.sleep(1)
    dwf.FDwfAnalogInConfigure(hdwf, c_int(0), c_int(1))# 0 - nie konfiguruje nic / 1 - start akwizycji

    # Pętla akwizycji danych
    cSamples = 0
    
    # Bufory na próbki Napięcia i Natężenia
    rgdSamples_Voltage_Preview = (c_double*Preview_Samples)()
    rgdSamples_Voltage_FFT = (c_double*Acquisition_Samples)()
    rgdSamples_Current_Preview = (c_double*Preview_Samples)()
    rgdSamples_Current_FFT = (c_double*Acquisition_Samples)()
    
    Final_rgdSamples_Voltage_FFT = (c_double*Acquisition_Samples)()
    Final_rgdSamples_Current_FFT = (c_double*Acquisition_Samples)()
    
    print('Recording samples... | Iteration number: ' + str(i))
    
    while cSamples < Acquisition_Samples:
        dwf.FDwfAnalogInStatus(hdwf, c_int(1), byref(sts))
        
        # Jeśli akwizycja się jeszcze nie rozpoczęła - wraca do początku pętli
        if cSamples == 0 and (sts == c_ubyte(4) or sts == c_ubyte(5) or sts == c_ubyte(1)) : # Stan urządzenia badany przez poszczególne bity: Config / Prefill / Armed
            continue
        
        dwf.FDwfAnalogInStatusRecord(hdwf, byref(cAvailable), byref(cLost), byref(cCorrupted))
    
        cSamples += cLost.value
    
        if cLost.value :
            fLost = 1
        if cCorrupted.value :
            fCorrupted = 1
        if cAvailable.value==0 :
            continue
    
        if cSamples+cAvailable.value > Acquisition_Samples:
            cAvailable = c_int(Acquisition_Samples-cSamples)
        
        dwf.FDwfAnalogInStatusData(hdwf, c_int(1), byref(rgdSamples_Voltage_FFT, sizeof(c_double)*cSamples), cAvailable) # Kanał Napięca
        dwf.FDwfAnalogInStatusData(hdwf, c_int(0), byref(rgdSamples_Current_FFT, sizeof(c_double)*cSamples), cAvailable) # Kanał Natężenia
        
        cSamples += cAvailable.value
    
    dwf.FDwfAnalogOutReset(hdwf, c_int(1))
    dwf.FDwfAnalogOutReset(hdwf, c_int(0))
    dwf.FDwfDeviceCloseAll()
    
    print("Recording done")
    if fLost:
        print("Samples were lost!")
    if fCorrupted:
        print("Samples could be corrupted!")
    
    print ("Collected: " + str(len(rgdSamples_Current_FFT)))
    
        
    rgdSamples_Voltage_Preview = rgdSamples_Voltage_FFT[:Preview_Samples]
    rgdSamples_Current_Preview = rgdSamples_Current_FFT[:Preview_Samples]
    
    window = np.blackman(Acquisition_Samples)
    rgdSamples_Voltage_FFT = rgdSamples_Voltage_FFT*window
    rgdSamples_Current_FFT = rgdSamples_Current_FFT*window
    
    X_oneside_Voltage = np.abs(np.fft.rfft(rgdSamples_Voltage_FFT))/(len(rgdSamples_Voltage_FFT)/2)
    X_oneside_Current = np.abs(np.fft.rfft(rgdSamples_Current_FFT))/(len(rgdSamples_Current_FFT)/2)
    f_oneside = np.fft.rfftfreq(len(rgdSamples_Voltage_FFT), 1/Hz)
    
    Harmonic_Voltage = (c_double*40)()
    Harmonic_Voltage = X_oneside_Voltage[50*AcqTime::50*AcqTime]
    Harmonic_Current = (c_double*40)()
    Harmonic_Current = X_oneside_Current[50*AcqTime::50*AcqTime]
    
    print("Calculation done")
    
    
    if i == 1:
        Final_X_oneside_Voltage = X_oneside_Voltage
        Final_X_oneside_Current = X_oneside_Current
        Final_Harmonic_Voltage = Harmonic_Voltage
        Final_Harmonic_Current = Harmonic_Current
    else:
        Final_X_oneside_Voltage = (Final_X_oneside_Voltage + X_oneside_Voltage)/2
        Final_X_oneside_Current = (Final_X_oneside_Current + X_oneside_Current)/2
        Final_Harmonic_Voltage = (Final_Harmonic_Voltage + Harmonic_Voltage)/2
        Final_Harmonic_Current = (Final_Harmonic_Current + Harmonic_Current)/2
   
    print("Total Harmonic Voltage Distortion : " + str(THD(Final_Harmonic_Voltage)) + " %")
    print("Total Harmonic Current Distortion : " + str(THD(Final_Harmonic_Current)) + " %")
    
    
    # Tworzenie wykresu harmonicznych napięcia
    plt.figure(figsize=(15,5))
    y_pos = np.arange(len(Final_Harmonic_Voltage))
    plt.bar(y_pos+1, Final_Harmonic_Voltage, color = '#969696')
    plt.xticks(np.arange(1, 41, 1))
    plt.xlabel('Numer harmonicznej', fontsize=12, color='#323232', fontweight='bold')
    plt.ylabel('Napięcie dla harmonicznej', fontsize=12, color='#323232', fontweight='bold')
    plt.title('Harmoniczne napięcia - ' + 'Sampling ' + str(Hz) + ' Hz', fontsize=16, color='#323232')
    if i == Iteration:
        plt.savefig('Final_Plots/Harmonic_Voltage_' + str(i) +'.png', bbox_inches='tight') 
    #else:
    if (i%5) == 0:
        plt.savefig('Harmonic_Voltage/Harmonic_Voltage_' + str(i) +'.png', bbox_inches='tight')  
    plt.cla() 
    plt.clf() 
    plt.close()  
    
    # Tworzenie wykresu harmonicznych natężenia
    plt.figure(figsize=(15,5))
    y_pos = np.arange(len(Final_Harmonic_Current))
    plt.bar(y_pos+1, Final_Harmonic_Current, color = '#969696')
    plt.xticks(np.arange(1, 41, 1))
    plt.xlabel('Numer harmonicznej', fontsize=12, color='#323232', fontweight='bold')
    plt.ylabel('Natężenie dla harmonicznej', fontsize=12, color='#323232', fontweight='bold')
    plt.title('Harmoniczne natężenia - ' + 'Sampling ' + str(Hz) + ' Hz', fontsize=16, color='#323232')
    if i == Iteration:
        plt.savefig('Final_Plots/Harmonic_Current_' + str(i) +'.png', bbox_inches='tight') 
    #else:
    if (i%5) == 0:
        plt.savefig('Harmonic_Current/Harmonic_Current_' + str(i) +'.png', bbox_inches='tight')  
    plt.cla() 
    plt.clf() 
    plt.close() 
    
    # Tworzenie wykresu Napięcia
    if i == Iteration:
        plt.figure(figsize=(15, 5))
        plt.title('Próbkowanie ' + str(Hz) + ' Hz | Numer iteracji: ' + str(i))
        plt.xlabel('Częstotliwość [Hz] (Czas = '+ str(AcqTime*i) + ' sekund)', fontweight='bold')
        plt.ylabel('Amplituda znormalizowanego napięcia FFT |X(freq)|', fontweight='bold')
        plt.xticks(np.arange(0, Hz/2+1, 100))
        plt.plot(f_oneside, Final_X_oneside_Voltage, 'r-', lw=0.5)
        plt.savefig('Final_Plots/Voltage_FFT_' + str(i) +'.png', bbox_inches='tight') 
        plt.cla() 
        plt.clf() 
        plt.close()   
    
    plt.figure(figsize=(15, 5))
    plt.title('Próbkowanie ' + str(Hz) + ' Hz | Numer iteracji: ' + str(i))
    plt.xlabel('Częstotliwość [Hz] (Czas = '+ str(AcqTime*i) + ' sekund)', fontweight='bold')
    plt.ylabel('Amplituda znormalizowanego napięcia FFT |X(freq)|', fontweight='bold')
    plt.xticks(np.arange(0, Hz/2+1, 100))
    plt.plot(f_oneside, 20 * np.log10(Final_X_oneside_Voltage), 'r-', lw=0.5)
    if i == Iteration:
        plt.savefig('Final_Plots/Voltage_FFT_dB_' + str(i) +'.png', bbox_inches='tight') 
    #else:
    if (i%5) == 0:
        plt.savefig('Voltage_FFT_dB/Voltage_FFT_dB_' + str(i) +'.png', bbox_inches='tight')
    plt.cla() 
    plt.clf() 
    plt.close()  
    
    if i == Iteration:
        plt.figure(figsize=(15, 4))
        plt.xlabel('Czas [s]', fontweight='bold')
        plt.ylabel('Napięcie [V]', fontweight='bold')
        Time_line = np.linspace(0.0, Preview*0.02, Preview_Samples)
        plt.plot(Time_line, rgdSamples_Voltage_Preview, 'b-', lw=1)
        plt.savefig('Preview/Voltage_Preview.png', bbox_inches='tight')
        plt.cla() 
        plt.clf() 
        plt.close()  
    
    # Tworzenie wykresu Natężenia
    if i == Iteration:
        plt.figure(figsize=(15, 5))
        plt.title('Próbkowanie ' + str(Hz) + ' Hz | Numer iteracji: ' + str(i))
        plt.xlabel('Częstotliwość [Hz] (Czas = '+ str(AcqTime*i) + ' sekund)', fontweight='bold')
        plt.ylabel('Amplituda znormalizowanego natężenia FFT |X(freq)|', fontweight='bold')
        plt.xticks(np.arange(0, Hz/2+1, 100))
        plt.plot(f_oneside, Final_X_oneside_Current, 'r-', lw=0.5)
        plt.savefig('Final_Plots/Current_FFT_' + str(i) +'.png', bbox_inches='tight') 
        plt.cla() 
        plt.clf() 
        plt.close() 
    
    plt.figure(figsize=(15, 5))
    plt.title('Próbkowanie ' + str(Hz) + ' Hz | Numer iteracji: ' + str(i))
    plt.xlabel('Częstotliwość [Hz] (Czas = '+ str(AcqTime*i) + ' sekund)', fontweight='bold')
    plt.ylabel('Amplituda znormalizowanego natężenia FFT |X(freq)|', fontweight='bold')
    plt.xticks(np.arange(0, Hz/2+1, 100))
    plt.plot(f_oneside, 20 * np.log10(Final_X_oneside_Current), 'r-', lw=0.5)
    if i == Iteration:
        plt.savefig('Final_Plots/Current_FFT_dB_' + str(i) +'.png', bbox_inches='tight') 
    #else:
    if (i%5) == 0:
        plt.savefig('Current_FFT_dB/Current_FFT_dB_' + str(i) +'.png', bbox_inches='tight')
    plt.cla() 
    plt.clf() 
    plt.close() 
    
    if i == Iteration:
        plt.figure(figsize=(15, 4))
        plt.xlabel('Czas [s]', fontweight='bold')
        plt.ylabel('Natężenie [mA]', fontweight='bold')
        Time_line = np.linspace(0.0, Preview*0.02, Preview_Samples)
        plt.plot(Time_line, rgdSamples_Current_Preview, 'g-', lw=1)
        plt.savefig('Preview/Current_Preview.png', bbox_inches='tight')
        plt.cla() 
        plt.clf() 
        plt.close()
    
    print("Plots has been saved")
    time.sleep(1)
    
    # Clearing the Screen
    os.system('clear')

print('After time: ' + str(AcqTime*Iteration) +'s')
print("Final Total Harmonic Voltage Distortion : " + str(THD(Final_Harmonic_Voltage)) + " %")
print("Final Total Harmonic Current Distortion : " + str(THD(Final_Harmonic_Current)) + " %")

Harmonic_Voltage_Rat = (c_double*40)()
Harmonic_Current_Rat = (c_double*40)()

for v in range(len(Final_Harmonic_Voltage)):
    Harmonic_Voltage_Rat[v] = (Final_Harmonic_Voltage[v]/Final_Harmonic_Voltage[0]) * 100 # Stosunek % do 1 harmonicznej

for v in range(len(Final_Harmonic_Current)):
    Harmonic_Current_Rat[v] = (Final_Harmonic_Current[v]/Final_Harmonic_Current[0]) * 100 # Stosunek % do 1 harmonicznej
    
Harmonic_Voltage_Rat = Harmonic_Voltage_Rat[1::]
Harmonic_Current_Rat = Harmonic_Current_Rat[1::]

plt.figure(figsize=(15,5))
y_pos = np.arange(0, 39, 1)
plt.bar(y_pos+2, Harmonic_Voltage_Rat, color = '#808080')
plt.xticks(np.arange(2, 41, 1))
for i in range(len(Harmonic_Voltage_Rat)):
    Harmonic_Voltage_Rat[i] = round(Harmonic_Voltage_Rat[i], 2)
    if Harmonic_Voltage_Rat[i] == 0 or Harmonic_Voltage_Rat[i] < 0.01:
        Harmonic_Voltage_Rat[i] = int(0)
        plt.annotate(Harmonic_Voltage_Rat[i],(y_pos[i]+1.8,Harmonic_Voltage_Rat[i]+0.01))
    else:
        plt.annotate(Harmonic_Voltage_Rat[i],(y_pos[i]+1.4,Harmonic_Voltage_Rat[i]+0.05*Harmonic_Voltage_Rat[i]))
plt.ylim([0, 1.15*max(Harmonic_Voltage_Rat)])
plt.xlabel('Numer harmonicznej', fontsize = 12, fontweight='bold')
plt.ylabel('Stosunek do 1 harmonicznej [%]', fontsize=12, fontweight='bold')
plt.title('Stosunek harmonicznych napięcia', fontsize=16)
plt.savefig('Ratio/Harmonic_Voltage_Rat_' + str(datetime.now()) + '.png', bbox_inches='tight')  
plt.cla() 
plt.clf() 
plt.close()  

plt.figure(figsize=(15,5))
y_pos = np.arange(0, 39, 1)
plt.bar(y_pos+2, Harmonic_Current_Rat, color = '#808080')
plt.xticks(np.arange(2, 41, 1))
for i in range(len(Harmonic_Current_Rat)):
    Harmonic_Current_Rat[i] = round(Harmonic_Current_Rat[i], 2)
    if Harmonic_Current_Rat[i] == 0 or Harmonic_Current_Rat[i] < 0.2:
        Harmonic_Current_Rat[i] = int(0)
        plt.annotate(Harmonic_Current_Rat[i], (y_pos[i] + 1.8, Harmonic_Current_Rat[i] + 0.3))
    else:
        plt.annotate(Harmonic_Current_Rat[i], (y_pos[i] + 1.4, Harmonic_Current_Rat[i] + 0.05*Harmonic_Current_Rat[i]))
plt.ylim([0, 1.15*max(Harmonic_Current_Rat)])
plt.xlabel('Numer harmonicznej', fontsize=12, fontweight='bold')
plt.ylabel('Stosunek do 1 harmonicznej [%]', fontsize=12, fontweight='bold')
plt.title('Stosunek harmonicznych natężenia', fontsize=16)
plt.savefig('Ratio/Harmonic_Current_Rat_' + str(datetime.now()) + '.png', bbox_inches='tight')  
plt.cla() 
plt.clf() 
plt.close()  

Finish_Time = datetime.now()
print ("Total operation time: " + str(Finish_Time-Start_Time))


email = "wozniak_test@gmail.com"
password = "htcgoprfyzrpuhhc"
message = """
After time: """ + str(AcqTime*Iteration) +"""s\n
Final Total Harmonic Voltage Distortion : """ + str(THD(Final_Harmonic_Voltage)) + """ %\n
Final Total Harmonic Current Distortion : """ + str(THD(Final_Harmonic_Current)) + """ %\n
Total operation time: """ + str(Finish_Time-Start_Time)

with smtplib.SMTP(host="smtp.gmail.com", port=587) as connection:
    connection.starttls()
    connection.login(user=email, password=password)
    connection.sendmail(
        from_addr=email,
        to_addrs=email,
        msg=message.encode("utf-8")
    )
