
#ne rabotaet
def search_recursion(list,target):
    start=0
    end=len(list)
    if start > end:
        return -1
    else:
        mid = (start+end//2)
        if target == list[mid]:
            return mid
        elif target < list[mid]:

            end = mid - 1
            return search_recursion(list,target)
        else:
            start=mid+1
            return  search_recursion(list,target)



#voobshe doljen rabotat
def search_while(list,target):
    start=0
    end=len(list)-1
    while start<=end:
        mid=(start+end)//2
        if list[mid] == target:
            return mid
        elif target<list[mid]:
            end = mid -1
        else:
            star = mid + 1
    return -1



list=[-1,0,3,5,9,12]
target=9

print(search_while(list,target))

