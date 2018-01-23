import serial

ser =serial.Serial('COM8')
file = open("log.txt","r")
for i in range(5000):
    byte=bytes(file.read(1),'utf-8')
    ser.write(byte)
