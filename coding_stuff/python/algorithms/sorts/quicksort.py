def quick_sort(list):
    length=len(list)//2

    if len(list)<=1:
        return list
    else:
        element = list[length]
         #print(element)
    items_greater=[]
    items_lower=[]

    for i in list:
        if i > element:
            items_greater.append(i)
        elif i<element:
            items_lower.append(i)

    return quick_sort(items_lower) + [element] + quick_sort(items_greater)

#function doesnt work well actualy

list=[0,-1,24,54,456,3,0,3,3,45,653]
list1=list
list1.sort()
print(quick_sort(list))
print(list1)
