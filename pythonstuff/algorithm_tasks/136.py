


def single_number(list):
    answer = 0
    new_list=[]

    while answer == 0 :
        pointer=list.pop()
        new_list.append(pointer)
        if pointer not in list and pointer not in new_list:
            answer = 1
            return pointer


nums = [4,1,2,1,2]
nums2 = [1,1,1,1,2]
nums3 = [1,0,1]

print(single_number(nums))


