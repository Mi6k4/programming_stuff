def string_comp(input_str):
    compressed_string=""
    count=1
    for i in range(len(input_str)-1):
        if input_str[i]==input_str[i+1]:
            count+=1
        else:
            compressed_string+=input_str[i]+str(count)
            count=1
    compressed_string+=input_str[i+1]+str(count)

    if len(compressed_string)>=len(input_str):
        return input_str
    else: return compressed_string


first_string="aaaeeerertghfffgdff"
second_string="abcdefg"
third_string="aaaaaaeeeeeerrrr"
forth_string="aaaaeeeerrrrtttfggggr"

print(string_comp(first_string))
print(string_comp(second_string))
print(string_comp(third_string))
print(string_comp(forth_string))