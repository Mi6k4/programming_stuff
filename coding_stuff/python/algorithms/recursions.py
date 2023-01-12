def privet(x):
    if x == 0:
        return
    else:
        print("Hello")
        privet(x-1)


def sum(x):
    if x == 0:
        return 0
    elif x == 1:
        return 1
    else:
        return x+sum(x-1)

def factorial(x):
    if x==0:
        return 1
    else:
        return x*factorial(x-1)


def fibbonachi(x):
    if x == 0:
        return 0
    elif x==1:
        return 1
    else:
        return fibbonachi(x-1)+fibbonachi(x-2)

privet(5)

s = sum(5)
print(s)
f = factorial(5)
print(f)

fib=fibbonachi(4)
print(fib)
