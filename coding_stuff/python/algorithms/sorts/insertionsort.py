def insertion_sort(list):
    for i in range(1,len(list)):
        current_value=list[i]
        position=i
        while position>0 and list[position-1]>current_value:
            list[position]=list[position-1]
            position-=1
        list[position]=current_value
    return list

list=[1,2,3,6,-1,345,547,32,5,56,-3]
list1=[2,2,2,2,2,2,1]
print(insertion_sort(list))
print(insertion_sort(list1))