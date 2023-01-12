#find Pivot Index
#The pivot index is the index where the sum of all the numbers strictly to the left of the index is equal
# to the sum of all the numbers strictly to the index's right.


def PivotIndex(list):
    if list is None:
        return -1

    left=0
    mid=0
    right=sum(list)

    for i in range(len(list)):
        left = left + mid
        print("left: ", left)
        mid = list[i]
        print("mid: ", mid)
        right = right - list[i]
        print("right: ", right)
        if left == right:
            return i
    return -1

       # right_sum = right_sum + list[-i-1]
       # print(right_sum)
        #left_sum = left_sum+list[i]
       #print(left_sum)



#по сути разбиваем на отрезки, левый отрезок равен нулю,
#второй сумме всего, прибавляем к левому отрезку, вычитаем из правого,когда они сравняются это наш искомый поинтер


list=[1,7,3,6,5,6]

list2=[60,0,20,0,30,0,10]

list3=[1,2,3]

list4=[1,0]
print(PivotIndex(list4))