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


linked_list=LinkedList()
temp=Node(1)
linked_list.head=temp
for i in range(2,10):
    temp.next=Node(i)
    temp=temp.next
print(linked_list)



node1=Node(1)
node2=Node(2)
node1.next=node2
print(node1)
print(node2.data)

