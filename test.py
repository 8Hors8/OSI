# import random
#
from time import time
from time import sleep


k = '1-1'
n = 1

if len(k)<=2:
    k = f'{k}-{n}'
else:
    k = k.split('-')
    k[-1] = str(n+1)
    k = '-'.join(k)

print(k,type(k))



