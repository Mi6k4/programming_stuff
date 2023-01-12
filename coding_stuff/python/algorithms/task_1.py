# дан list целых чисел arr и целое число k. Рассчитать list средних арифметических всех
# непрерывных подмассивов длины k. Если k больше длины входного листа, вернуть None.
# Например, [1, 2, 3, -2, -1, 6], k=3 -> [2.0, 1.0, 0.0, 1.0].
def summ(arr,k):
    new_arr=[]
    sum_1=0
    if k>len(arr):
        return print("None")
    else:
        for i in range(len(arr)-k+1):
            for j in range(k):
                sum_1+=arr[i+j]
            sum=sum_1/k
            sum_1=0
            new_arr.append(sum)
        return new_arr



arr = [1, 2, 3, -2, -1, 6]; k=3
print(arr, k)




print(summ(arr,k))
