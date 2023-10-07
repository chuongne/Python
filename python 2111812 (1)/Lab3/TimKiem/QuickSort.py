# Việc triển khai này sử dụng trục xoay làm phần tử cuối cùng trong danh sách số
# Nó có một con trỏ để theo dõi các phần tử nhỏ hơn trục xoay
# Ở cuối hàm partition(), con trỏ được hoán đổi với trục xoay
# để đưa ra một số được "sắp xếp" tương ứng với trục xoay
# Chức năng tìm vị trí phân vùng
def partition(array, low, high):
	# choose the rightmost element as pivot
	pivot = array[high]

	# pointer for greater element
	i = low - 1

	# traverse through all elements
	# compare each element with pivot
	for j in range(low, high):
		if array[j] <= pivot:

			# Nếu tìm thấy phần tử nhỏ hơn trục
# hoán đổi nó với phần tử lớn hơn được chỉ bởi i
			i = i + 1

			# Hoán đổi phần tử tại i với phần tử tại j
			(array[i], array[j]) = (array[j], array[i])

	# Hoán đổi phần tử trụ với phần tử lớn hơn được chỉ định bởi i
	(array[i + 1], array[high]) = (array[high], array[i + 1])

	# Trả về vị trí nơi thực hiện phân vùng trả lại tôi + 1
	return i + 1

# function to perform quicksort

def quickSort(array, low, high):
	if low < high:
		# Tìm phần tử trục sao cho
        # phần tử nhỏ hơn trục nằm ở bên trái
        # phần tử lớn hơn trục nằm ở bên phải
		pi = partition(array, low, high)
		# Lệnh gọi đệ quy ở bên trái của trục
		quickSort(array, low, pi - 1)
		# Lệnh gọi đệ quy bên phải trục xoay
		quickSort(array, pi + 1, high)

data = [1, 7, 4, 1, 10, 9, -2]
print("Unsorted Array")
print(data)

size = len(data)

quickSort(data, 0, size - 1)

print('Sorted Array in Ascending Order:')
print(data)
