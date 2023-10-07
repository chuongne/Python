def insertionSort(arr):
    n = len(arr)  # Get the length of the array 
    if n <= 1:
        return  # If the array has 0 or 1 element, it is already sorted, so return
    for i in range(1, n):  # lặp lại mảng từ phần tử thứ 2
        key = arr[i]  # Store the current element as the key to be inserted in the right position
        j = i-1
        while j >= 0 and key < arr[j]:  # Move elements greater than key one position ahead
            arr[j+1] = arr[j]  # Dịch chuyển phần tử sang trái
            j -= 1
        arr[j+1] = key  # Chèn key vào vị trí
# Sorting the array [12, 11, 13, 5, 6] using insertionSort
arr = [12, 11, 13, 5, 6]
insertionSort(arr)
print(arr)