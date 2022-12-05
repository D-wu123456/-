import re
s='网神洞鉴[2022]数鉴字第150号'
cost=re.findall(r'[1-9]+\.?[0-9]*',s)
print(cost)