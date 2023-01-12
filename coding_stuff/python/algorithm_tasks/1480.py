#list of sums

def sum(nums):
    value=0
    result = []
    for i in range(len(nums)):
        value= value + nums[i]
        result.append(value)
    return result


list=[1,2,3,4]

print(sum(list))