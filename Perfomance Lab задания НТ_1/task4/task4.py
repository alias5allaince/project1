import sys
 
a=str(sys.argv[1])
b=str(sys.argv[2])

c1= list(a)
c2= list(b)

if len(c1)==len(c2):
    for i in range(0,len(c2)):
        if c1[i]==c2[i] or c2[i]=="*":
            result='OK'
        else: 
            result='KO' 
  
elif len(c1)>len(c2):
    for i in range(0,len(c2)):
        if c2[i]!="*":
            result='KO'
        elif c1[i]==c2[i] or c2[i]=="*":
            result='OK'
            
elif len(c1)<len(c2):
     for i in range(0,len(c1)):
        if c2[i]!="*":
            result='KO'
        elif c1[i]==c2[i] or c2[i]=="*":
            result='OK'

print(result)
    
    


     
