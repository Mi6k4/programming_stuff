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


class Solution:
    def reverselist(self,head: Node)->Node:
        if not head:
            return head
        tmp=head
        tail=Node(data=head.data)
        while tmp.next:
            tmp=tmp.next
            tail=Node(data=tmp.data,next=tail.next)




