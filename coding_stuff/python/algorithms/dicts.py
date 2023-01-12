dct={}
dct2=dict()


age={"john":3,"jwosh":5,"won":7}

print(age)

age["seg"]=25

print(age)

print(age["won"])

age2=age.copy()
age.clear()

print(age2.get("john"))
age2.pop("john")
print(age2)
print(age2.values())
age3={1:333,"Alex":"Kiev",444:True}
age2.update(age3)
print(age2)

print(age3)

for key in age3:
    print(key, "-", age3[key])

for key,value in age3.items():
    print(key,"-",value)

for key in age3.keys():
    print(key)

for values in age3.values():
    print(values)


#del age[0]

