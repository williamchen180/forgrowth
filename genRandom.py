


import random

with open('random.txt', 'w') as f:
    for i in range(62, 2792, 10):
        r = random.sample(range(10), 10)
        for j in range(0, 10):
            #f.write(str(r[j]) + '\n')
            f.write(str(j) + '\n')

