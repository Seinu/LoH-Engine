from PIL import Image 
import numpy 
img=Image.open('colorcode.png') 
width, height = img.size
rgb_img = img.convert('RGB')
f = open("output.csv", "w")
l = []
for x in range(height):
    for y in range(width):
        r, g, b = rgb_img.getpixel((y,x))
        f.write('{0},{1},{2}\n'.format(r,g,b))
f.close()
f = open("output.csv", "r")
mar = open("output.mar", "w+b")
Lines = f.readlines()
for line in Lines:
    #print(line)
    if(line == "0,38,248\n"):
        mar.write(b'\xE0\x07')
    elif(line == "248,240,192\n"):
        mar.write(b'\x40\x08')
    elif(line == "0,0,156\n"):
        mar.write(b'\x00\x08')
    elif(line == "230,230,230\n"):
        mar.write(b'\x20\x00')
    elif(line == "190,230,190\n"):
        mar.write(b'\xA0\x02')
    elif(line == "248,248,248\n"):
        mar.write(b'\x40\x00')
    elif(line == "188,188,224\n"):
        mar.write(b'\x60\x00')
    elif(line == "64,112,48\n"):
        mar.write(b'\x80\x00')
    elif(line == "112,192,112\n"):
        mar.write(b'\xA0\x00')
    elif(line == "152,112,72\n"):
        mar.write(b'\xC0\x00')
    elif(line == "224,184,120\n"):
        mar.write(b'\x20\x08')
    elif(line == "192,56,48\n"):
        mar.write(b'\xC0\x05')
    elif(line == "0,72,0\n"):
        mar.write(b'\x00\x03')
    elif(line == "64,32,56\n"):
        mar.write(b'\xA0\x07')
    else:
        mar.write(b'\xA0\x07')
f.close()
mar.close()
print('Done!')