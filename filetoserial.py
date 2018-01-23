import serial

ser =serial.Serial('COM5')
file=open("log.txt","r")
for i in range(6000):
    char=bytes(file.read(1),'ascii')
    ser.write(char)
    print(char)
print('Ciaociao')
