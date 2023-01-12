def linearsearch(list,element):
    for i in range(len(list)):
        if list[i]==element:
            return i
    return -1


list=[1,1,2,4,5,7,-1]

print(linearsearch(list,-6))

