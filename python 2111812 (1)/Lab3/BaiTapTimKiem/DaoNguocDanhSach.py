class Node:
    # Hàm tạo để khởi tạo đối tượng nút
    def __init__(self, data):
        self.data = data
        self.next = None
class LinkedList:
    # Hàm khởi tạo phần đầu
    def __init__(self):
        self.head = None
    # Chức năng đảo ngược danh sách liên kết
    def reverse(self):
        prev = None
        current = self.head
        while(current is not None):
            next = current.next
            current.next = prev
            prev = current
            current = next
        self.head = prev
    # Chức năng chèn một nút mới vào đầu
    def push(self, new_data):
        new_node = Node(new_data)
        new_node.next = self.head
        self.head = new_node
    #Chức năng tiện ích trong LinkedList
    def printList(self):
        temp = self.head
        while(temp):
            print (temp.data,end=" ")
            temp = temp.next
# Chương trình driver để kiểm tra các chức năng trên
llist = LinkedList()
llist.push(20)
llist.push(4)
llist.push(15)
llist.push(85)
print ("Given Linked List")
llist.printList()
llist.reverse()
print ("\nReversed Linked List")
llist.printList()
