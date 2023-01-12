mystring="Hello,that's my string"
print(mystring.istitle())

print("you" in mystring)

print(mystring.find(' '))
print(mystring.index(' '))

print(len(mystring))

print(mystring.count("l"))

print(mystring.capitalize())

print(mystring.isnumeric()) #contains only number

string_list=mystring.split()
print(string_list)
string_list=mystring.split(",")
print(string_list)

print(mystring.islower())

print(mystring[0].islower())

reversed_string=mystring[::-1]
print(reversed_string)

new_string=''.join(string_list)
print(new_string)

print(mystring.upper())

print(mystring.lower())

print(mystring[1:10:2])

#converting int into str
number=123
str_number=str(number)
print(type(str_number))

#checking if string contains only alphabet characters
print(mystring.isalpha())

new_string_two= mystring.replace("Hello","Hi")
print(new_string_two)

print(mystring.startswith("H"))

print(mystring*2)

#capitilize first character in each word
print(mystring.title())

print(mystring+new_string_two)




