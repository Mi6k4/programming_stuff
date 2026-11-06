def func(nums):
    dict={}
    for num in nums:
        if num not in dict:
            dict[num]=1
        else:
            return True
    return False


nums = [1,2,3,1]
nums = [1,2,3,4]
nums = [1,1,1,3,3,4,3,2,4,2]
print(func(nums))