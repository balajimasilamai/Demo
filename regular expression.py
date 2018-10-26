import re

s = 'My name is Conrad, and blahblah@gmail.com is my email.'

domain = re.search('@', s)
if domain:
    print ('abc')
else:
    print ('afaf')

print (domain)
