import json
d=open('data.json').read()
t=open('template.html').read().replace('/*DATA*/null',d)
open('index.html','w').write(t)
print(len(t))
