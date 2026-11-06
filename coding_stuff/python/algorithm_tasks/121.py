def func(prices):
    min_price=None
    result=0
    for i in range(len(prices)):
        if min_price is None or prices[i]<min_price:
            min_price=prices[i]
        elif min_price is not None:
            result = max(result, prices[i] - min_price)



    if result>0:
        return result
    else:
        return 0


prices = [7,1,5,3,6,4]
#prices = [7,6,4,3,1]
#prices =[2,4,1]
print(func(prices))