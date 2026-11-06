def find_min(lst):
    min_value = lst[0]
    print(min(lst))
    for i in range(1, len(lst)):
        if lst[i] < min_value:
            min_value = lst[i]
    return min_value

print(find_min([3, 1, 4, 1, 5, 9]))