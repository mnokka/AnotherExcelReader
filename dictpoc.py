d={}
a="kissa"
b="koira"
c="hirvi"

d[a]=1
d[a]=2
d[b]=1

# use dictionaruy as a counter
if (c in d):
    value=d.get(a,"10000") # 1000 is default value
    value=value+1 
    #d[a]=value
    print "yes a key"
    print "a:{0}".format(value)
    d[a]=value
else:
    print "not a key, setting"
    d[c]=0    

for key,value in d.items():
    print "**************************************"
    if (key in d):
        #value=d.get(a,"10000") # 1000 is default value
        value=value+1 
        d[key]=value
    else:
        print "not key"  
    print "KEY:{0}  => VALUE: {1}".format(key,value)