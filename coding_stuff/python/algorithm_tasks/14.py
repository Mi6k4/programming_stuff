"""

Write a function to find the longest common prefix string amongst an array of strings.

If there is no common prefix, return an empty string "".



Example 1:

Input: strs = ["flower","flow","flight"]
Output: "fl"
Example 2:

Input: strs = ["dog","racecar","car"]
Output: ""
Explanation: There is no common prefix among the input strings.
"""

"это придумал я но написал чат гпт"
def longestCommonPrefix(strs):
    if not strs:
        return ""

    # Начинаем с первой буквы первой строки
    for i in range(len(strs[0])):
        char = strs[0][i]

        # Проверяем, начинается ли каждая строка с этой буквы
        for string in strs:
            if i >= len(string) or string[i] != char:
                return strs[0][:i]  # Возвращаем строку до текущей позиции

    return strs[0]  # Если префикс совпал полностью


"это придумал чат гпт"
def func(strs):
    if not strs:
        return ""

    prefix = strs[0]

    for string in strs[1:]:
        while not string.startswith(prefix):
            prefix = prefix[:-1]
            if not prefix:
                return ""

    return prefix

strs = ["flower","flow","flight"]

print(func(strs))