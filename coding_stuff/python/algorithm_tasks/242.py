"""

Given two strings s and t, return true if t is an
anagram
 of s, and false otherwise.



Example 1:

Input: s = "anagram", t = "nagaram"

Output: true

Example 2:

Input: s = "rat", t = "car"

Output: false
"""

def func(s,t):
    if len(s)!=len(t):
        return False
    s_dict={}
    t_dict={}
    for letter in s:
        if letter not in s_dict:
            s_dict[letter]=1
        else:
            s_dict[letter]+=1
    for letter in t:
        if letter not in t_dict:
            t_dict[letter] = 1
        else:
            t_dict[letter]+= 1
    if s_dict==t_dict:
        return True
    else:
        return False

s = "anagram"
t = "nagaram"


s = "rat"
t = "car"
print(func(s,t))
