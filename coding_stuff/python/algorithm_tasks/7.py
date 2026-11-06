def function(x):
    if x==0:
        return 0
    elif x>0:
        str_x=str(x)
        reversed_str=str_x[::-1]
        reversed_x=int(reversed_str)
    else:
        str_x=str(-x)
        reversed_str = str_x[::-1]
        reversed_x = int(reversed_str)*(-1)

    if reversed_x>2**31-1 or reversed_x<-2**31:
        return 0
    print(reversed_x)

x_1=123
x_2=-123
x_3=120

function(x_1)
function(x_2)
function(x_3)
