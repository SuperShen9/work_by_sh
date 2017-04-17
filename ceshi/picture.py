from pygame import *
from math import *
screen=display.set_mode((994,994))
m=image.load('super.jpg')
screen.blit(m,(0,0))
mon=open('super.py','w')
mon.write('''from pygame import *
from math import *

screen=display.set_mode((994,994))
''')
running=True
display.flip()
for x in range(0,994,5):
    for y in range(0,994,5):
        c=str(screen.get_at((x,y)))
        mon.write('draw.rect(screen,'+c+',\
        ('+str(x)+','+str(y)+',4,4))\n')
    mon.write('display.flip()\n')
mon.write('''running=True
while running:
    for evnt in event.get():
        if evnt.type==QUIT:
            running=False
quit()''')
mon.close()

