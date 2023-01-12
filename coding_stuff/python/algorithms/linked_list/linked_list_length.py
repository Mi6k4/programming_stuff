class Node:
    def __init__(self,data):
        self.data=data
        self.next=None

    def __str__(self):
        return f"[{self.data}]->{self.next}"

class LinkedList:
    def __init__(self):
        self.head=None
    def __str__(self):
        return str(self.head)
    def length(self):
        count=0
        tmp=self.head
        while tmp:
            count+=1
            tmp=tmp.next
        return count


linked_list=LinkedList()
tmp=Node(1)
linked_list.head=tmp
for i in range(2,10):
    tmp.next=Node(i)
    tmp=tmp.next
print(linked_list)




