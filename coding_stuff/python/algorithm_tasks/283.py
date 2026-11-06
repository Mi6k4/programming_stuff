"""

Given an integer array nums, move all 0's to the end of it while maintaining the relative order of the non-zero elements.

Note that you must do this in-place without making a copy of the array.



Example 1:

Input: nums = [0,1,0,3,12]
Output: [1,3,12,0,0]
Example 2:

Input: nums = [0]
Output: [0]

"""
def move_zeros(nums):
    nums[:] = [x for x in nums if x != 0] + [0] * nums.count(0)
    return nums


def func(nums):
    zero_counts=0
    for num in nums:
        if num == 0:
            nums.remove(num)
            zero_counts+=1
        else:
            pass
    for i in range(zero_counts):
        nums.append(0)
    return nums

nums = [0,1,0,3,12]
nums = [0,0,1]
print(move_zeros(nums))