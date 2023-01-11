def PalindromeNumber(number):
    if number<0:
        return False
    number_to_string=str(number)
    straight=number_to_string
    backwards=number_to_string[::-1]
    if straight == backwards:
        return True
    else: return False


def better_solution(number):
    return str(number)==str(number)[::-1]


number=101
number1=102
number2=-100
number3=11

print(PalindromeNumber(number))
print(PalindromeNumber(number1))
print(PalindromeNumber(number2))
print(PalindromeNumber(number3))

print(better_solution(number))