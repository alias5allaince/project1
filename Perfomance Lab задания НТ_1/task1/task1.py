import sys


a=sys.argv[1]
b=sys.argv[2]


def convert_to_01 (a,b):
    try:
        a=int(sys.argv[1])
        b=str(sys.argv[2])
    except (TypeError, ValueError, NameError, SyntaxError, IndexError): pass
    if b=="01":
        x=bin(a)
        return x
    else:
        pass
    
    
try:
    print(convert_to_01 (a,b)[2:])
except (TypeError, ValueError, NameError,SyntaxError, IndexError): 
       print("usage")
