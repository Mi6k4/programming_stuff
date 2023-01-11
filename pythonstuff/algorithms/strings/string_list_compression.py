def list_of_str_compression(list):
    index = 0
    i = 0
    while i < len(list):
        j = i
        while j < len(list) and list[j]==list[i]:
            j+=1
        index+=1
        list[index]=list[i]
        if (j-i)>1:
            count=str(j-i)
            for c in range(len(list)):
                #index+=1
                list[index]=count
        i = j

    return list

list_of_str=["a","a","a","b","b","b"]

print(list_of_str_compression(list_of_str))