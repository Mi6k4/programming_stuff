def function(s):
    my_stack=[]
    for letter in s:
        print(letter)
        if letter =='(' or letter=='[' or letter=='{':
            my_stack.append(letter)
            print(my_stack)
        else:
            print(my_stack)
            if len(my_stack)==0:
                print('I am here')
                return False
            elif letter ==')' and my_stack[-1]!='(' or letter ==']' and my_stack[-1]!='[' or letter =='}' and my_stack[-1]!='{' :
                return False
            else:
                print(my_stack)
                my_stack.pop()
    #print(my_stack)
    #print(len(my_stack))
    if len(my_stack)==0:
        return True
    else:
        return False


str_1="()"
str_2="()[]{}"
str_3="(]"

function(str_1)
#function(str_2)
#function(str_3)