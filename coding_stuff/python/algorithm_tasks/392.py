#mine not working exactly correct
def solution(first_string,second_string):
    counter1=0
    counter2=0
    for counter2 in range(len(second_string)):
        if first_string[counter1] == second_string[counter2]:
            counter1+=1
            counter2+=1
        else:
            counter2+=1

    if counter1 == len(first_string):
        return True
    return False




# not mine
def solution2(first_string,second_string):
    pointer1=0
    pointer2=0
    while pointer1<len(first_string) and pointer2<len(second_string):
        if first_string[pointer1] == second_string[pointer2]:
            pointer1+=1
        pointer2+=1
    if pointer1==len(first_string):
        return True
    return False




string1="qwe"
string2="qwerty"
string3="qewrty"

s="abc"
t="ahbgdc"

print(solution(string1,string2  ))