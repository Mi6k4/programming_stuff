"""
Задача: Даны две матрицы A (M x N) и B (N x K). Задача — написать функцию для перемножения этих матриц C = A * B (M x K)
M, N, K != 0
Пример входных данных:
A = [
        [1,2],
        [3,4]
    ]
B = [
        [5,6],
        [7,8]
    ]
Результат:
C = A * B = [

        [1 * 5 + 2 * 7, 1 * 6 + 2 * 8],

        [3 * 5 + 4 * 7, 3 * 6 + 4 * 8]

    ]

def matrix_product(A: List[List[int]], B: List[List[int]]) -> List[List[int]]:

"""



def matrix_product(A, B):
    C=[]
    for i in range(len(A)):
        sub_sum= 0
        vektor=[]
        for j in range(len(A[i])):
            for k in range(len(B[i])):
                sub_sum+=A[i][j]*B[j][k]
            vektor.append(sub_sum)
        C.append(vektor)
    return C


A = [
        [1,2],
        [3,4]
    ]
B = [
        [5,6],
        [7,8]
    ]
print(matrix_product(A,B))


def append(value: int, list: list[int] = []) -> list[int]:
    list.append(value)
    return list


alist1 = append(10)
alist2 = append([123], [])
alist3 = append('a')

print(f"{alist1=}, ", end="")
print(f"{alist2=}, ", end="")
print(f"{alist3=}")

