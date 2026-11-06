def function(nums,k):
    counter_dict={}
    list_of_numbers=[]
    final_answer=[]
    for number in nums:
        if number in counter_dict:
            counter_dict[number]+=1
        else:
            counter_dict[number]=1
    print(counter_dict)
    for key,value in counter_dict.items():
        list_of_numbers.append(value)
    print(list_of_numbers)
    list_of_numbers.sort(reverse=True)
    print(list_of_numbers)
    for i in range(k):
        for key, value in counter_dict.items():
            if value==list_of_numbers[i] and key not in final_answer:
                final_answer.append(key)
    print(final_answer)




nums_1 =[1,1,1,2,2,3]
k_1=2

nums_2=[1]
k_2=1

nums_3=[]
k_3=0

function(nums_1,k_1)
function(nums_2,k_2)
function(nums_3,k_3)