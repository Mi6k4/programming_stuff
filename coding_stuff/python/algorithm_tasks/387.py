"""

Given a string s, find the first non-repeating character in it and return its index. If it does not exist, return -1.



Example 1:

Input: s = "leetcode"

Output: 0

Explanation:

The character 'l' at index 0 is the first character that does not occur at any other index.

Example 2:

Input: s = "loveleetcode"

Output: 2

Example 3:

Input: s = "aabb"

Output: -1
"""


def func(s):
    dict={}
    index_pointer=None
    for i in range(len(s)):
        if s[i] not in dict:
            dict[s[i]]=[1,i]
        else:
            dict[s[i]][0]+=1
    for key in list(dict.keys()):
        if dict[key][0]!=1:
            dict.pop(key)
        else:
            pass
    if len(dict) !=0:
        for key in dict:
            if index_pointer is None or dict[key][1]<index_pointer:
                index_pointer=dict[key][1]
            else:
                pass
        return index_pointer
    else:
        return -1


s = "leetcode"
s = "loveleetcode"
s = "aabb"
s ="aadadaad"
print(func(s))