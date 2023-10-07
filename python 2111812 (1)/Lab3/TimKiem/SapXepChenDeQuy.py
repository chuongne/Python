def insertionSortRecursive(arr, n):
    if n <= 1:
        return
    # Sort first n-1 elements
    insertionSortRecursive(arr, n - 1)
    # Chèn phần tử cuối cùng vào đúng vị trí của nó trong mảng đã sắp xếp.
    last = arr[n - 1]
    j = n - 2
    # Di chuyển các phần tử của mảng[0..i-1], tức là
    # lớn hơn khóa, tới một vị trí phía trước
    # vị trí hiện tại của họ
    while (j >= 0 and arr[j] > last):
        arr[j + 1] = arr[j]
        j = j - 1
    arr[j + 1] = last
if __name__ == '__main__':
    A = [-7, 11, 6, 0, -3, 5, 10, 2]
    n = len(A)
    insertionSortRecursive(A, n)
    print(A)



def insertion_sort_recursive(arr):
    # base case: return when array has only one element
    if len(arr) <= 1:
        return arr
    # sắp xếp đệ quy nửa đầu của mảng
    mid = len(arr) // 2
    left_half = insertion_sort_recursive(arr[:mid])
    # sắp xếp đệ quy nửa sau của mảng
    right_half = insertion_sort_recursive(arr[mid:])
    # hợp nhất các nửa đã sắp xếp thành một mảng đã sắp xếp
    # print(left_half)
    # print(right_half)
    i, j = 0, 0
    sorted_arr = []
    while i < len(left_half) and j < len(right_half):
        if left_half[i] < right_half[j]:
            sorted_arr.append(left_half[i])
            i += 1
        else:
            sorted_arr.append(right_half[j])
            j += 1
    sorted_arr += left_half[i:]
    sorted_arr += right_half[j:]
    return sorted_arr
arr = [5, 2, 4, 6, 1, 3]
sorted_arr = insertion_sort_recursive(arr)
print(sorted_arr)  # Output: [1, 2, 3, 4, 5, 6]
