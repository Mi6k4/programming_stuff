def bubble_sort(mylist):
    last_item=len(mylist)-1
    for z in range(0,last_item):
        for x in range(0,last_item-z):
            if mylist[x]>mylist[x+1]:
                mylist[x],mylist[x+1]=mylist[x+1],mylist[x]
    return mylist



list = [12,4534,345,6,3,9,234,84,8,97,4]

print(list)
new_list=bubble_sort(list)

print(new_list)

