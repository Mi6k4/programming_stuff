

def binarysearch(mylist,iskat,start,stop):
    if start>stop:
        return False
    else:
        mid=(start+stop)//2
        if iskat == mylist[mid]:
            return mid
        elif iskat < mylist[mid]:
            return binarysearch(mylist,iskat,start,mid-1)
        else:
            return binarysearch(mylist,iskat,mid+1,stop)



list = [1,2,3,5,6,7,36,68,78]

iskat=78

start=0
stop=len(list)
x=binarysearch(list,iskat,start,stop)

if x==False:
    print("Ne naiden")
else:
    print(iskat,"at index",x)