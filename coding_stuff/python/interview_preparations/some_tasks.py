"""
Условие:
Напишите функцию, которая принимает строку и возвращает словарь с количеством вхождений каждого символа.

Пример:

python
Копировать код
input_string = "hello"
output = {'h': 1, 'e': 1, 'l': 2, 'o': 1}
"""

def dict_func(input_string):
    output_dict={}
    for symb in input_string:
        if symb not in output_dict:
            output_dict[symb]=1
        else:
            output_dict[symb]+=1
    return output_dict

input_string = "hello"
input_string_2='rrrrrttttttooooop'
input_string_3='oppposdfahggppoasdf'
#print(dict_func(input_string))
#print(dict_func(input_string_2))
#print(dict_func(input_string_3))

"""
2. Задача на работу с массивами и списками:
Условие:
Напишите функцию, которая находит пару чисел в списке, сумма которых равна заданному числу target.

Пример:

python
Копировать код
nums = [2, 7, 11, 15]
target = 9
output = [2, 7]
"""

def func_2(list_of_nums,target):
    dict={}
    for num in list_of_nums:
        output_number = target - num

        if num in dict:
            return [dict[num], num]  # Возвращаем найденную пару
        else:
            dict[output_number] = num

nums = [1, 2, 3, 4, 5]
target = 10
#print(func_2(nums,target))

"""
Условие:
Дан список целых чисел. Найдите первое пропущенное положительное число.

Пример:

python
Копировать код
nums = [3, 4, -1, 1]
output = 2
"""


def func_3(list_of_numbers):
    list_of_numbers.sort()  # Сортируем список
    current = 1  # Начинаем с 1

    for num in list_of_numbers:
        if num == current:  # Если число равно текущему, увеличиваем current
            current += 1
        elif num > current:  # Если число больше current, значит current пропущено
            break

    return current  # Возвращаем первое пропущенное положительное число


def func_3_with_set(list_of_numbers):
    nums = set(list_of_numbers)  # Преобразуем список в множество, чтобы исключить дубли
    i = 1  # Начинаем с 1
    while i in nums:  # Пока число i присутствует в множестве
        i += 1  # Переходим к следующему числу
    return i  # Когда число не найдено, возвращаем его

nums = [1, 2, 2, 3, 4, 4, 5]
#print(func_3(nums))




"""
1. Задача: Нахождение недостающих чисел в последовательности
Условие: Дан список, содержащий числа от 1 до 
𝑛
n, но одно или несколько чисел пропущены. Найдите недостающие числа.

Пример:

python
Копировать код
nums = [1, 2, 4, 6, 7, 8]
# Ожидаемый результат: [3, 5]
"""

def func_4(list_of_nums):
    missed_numbers=[]
    current = 1
    for num in list_of_nums:
        if num == current:
            current+=1
        else:
            missed_count=num-current
            for i in range(missed_count):
                missed_numbers.append(current)
                current+=1
            current += 1
    return missed_numbers

nums = [1, 2, 4, 6, 7, 8]
# Ожидаемый результат: [3, 5]

#print(func_4(nums))



"""
2. Задача: Найти повторяющиеся элементы в списке
Условие: Дан список целых чисел, необходимо найти все числа, которые встречаются более одного раза.

Пример:

python
Копировать код
nums = [4, 3, 2, 7, 8, 2, 3, 1]
# Ожидаемый результат: [3, 2]
"""


def func_5(list_of_nums):
    dict={}
    output_set = set()
    for num in list_of_nums:
        if num not in dict:
            dict[num]=1
        else:
            output_set.add(num)
    return list(output_set)

nums = [4, 3, 2, 7, 8, 2, 3, 1]
#print(func_5(nums))



"""
 Задача: Найти максимальную длину подпоследовательности, которая состоит только из уникальных элементов
Условие: Дан список чисел. Необходимо найти максимальную длину подпоследовательности, в которой все элементы уникальны.

Пример:

python
Копировать код
nums = [1, 2, 3, 1, 2, 3, 4, 5]
# Ожидаемый результат: 5 (Подпоследовательность: [1, 2, 3, 4, 5])
"""

def func_6(list_of_nums):
    set_of_nums=set(list_of_nums)
    return len(set_of_nums)

nums = [1, 2, 3, 1, 2, 3, 4, 5]
print(func_6(nums))


