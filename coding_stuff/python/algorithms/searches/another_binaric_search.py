def binaricsearch(list,number):
    first=0
    last=len(list)-1
    index=-1
    while (first<=last) and (index==-1):
        mid=(first+last)//2
        if list[mid]==number:
            index=mid
        elif number<list[mid]:
            last=mid-1
        elif number>list[mid]:
            last=mid+1
    return index


list=[1,2,3,4,5,6,7,8]

print(binaricsearch(list,2))