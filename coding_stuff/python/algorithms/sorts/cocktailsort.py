def cocktailsort(list):
    for i in range(len(list)-1,0,-1):
        swapped = False
        for j in range(i,0,-1):
            if list[j]<list[j-1]:
                list[j], list[j-1]=list[j-1],list[j]
                swapped = True
        for j in range(i):
            if list[j]>list[j+1]:
                list[j],list[j+1]=list[j+1],list[j]
                swapped = True
        if  not swapped:
            return list



list_1=[1,2,345,4567,3,687,3,8,-1,453,-6]

print(list_1)
print(cocktailsort(list_1))


