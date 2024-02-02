from time import sleep
x=1000
ti=0.95
while x>7:
    print(f"{x} - 7 = {x-7}")
    x -= 7
    if ti>=0.05:
        ti-=0.05                
    sleep(ti)
print('Ты свой...')