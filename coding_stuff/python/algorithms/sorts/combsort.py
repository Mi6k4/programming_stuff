def combsort(list):
    step=int(len(list)/1.247)
    swap = 1
    while step > 1 or swap > 0:
        swap = 0
        i = 0
        while i + step < len(list):
            if list[i]>list[i+step]:
                list[i],list[i+step]=list[i+step],list[i]
                swap+=1
            i+=1
        if step>1:
            step=int(step/1.247)
    return list

list=[1,2,56,-45,5,87,890,45,-435,0]

print(list)
print(combsort(list))