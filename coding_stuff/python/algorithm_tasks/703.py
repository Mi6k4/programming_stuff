import heapq
class KthLargest(object):

    def __init__(self, k, nums):
        self.k=k
        self.heap=[]
        for num in nums:
            self.add(num)


    def add(self, val):
        heapq.heappush(self.heap,val)
        if len(self.heap)>self.k:
            heapq.heappop(self.heap)
        return self.heap[0]

k = 3
nums = [4, 5, 8, 2]
kthLargest = KthLargest(k, nums)

print(kthLargest.add(3))  # Вывод: 4
print(kthLargest.add(5))  # Вывод: 5
print(kthLargest.add(10)) # Вывод: 5
print(kthLargest.add(9))  # Вывод: 8
print(kthLargest.add(4))  # Вывод: 8