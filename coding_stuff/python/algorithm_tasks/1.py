def func(nums,target):
    dict={}
    output_list=[]
    for i in range(len(nums)):
        if nums[i] not in dict:
            sub_num=target-nums[i]
            dict[sub_num]=i
        else:
            output_list.append(dict[nums[i]])
            output_list.append(i)
            return output_list




nums = [2,7,11,15]
target = 9
nums = [3,2,4]
target = 6
nums = [3,3]
target = 6
print(func(nums,target))