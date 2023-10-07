def count(S, m, n):
    # Chúng ta cần n+1 hàng để xây dựng bảng
     # theo cách từ dưới lên sử dụng giá trị trường hợp cơ sở 0
     # trường hợp (n = 0)
    table = [[0 for x in range(m)] for x in range(n+1)]
   # Điền các mục cho trường hợp giá trị 0 (n = 0)
    for i in range(m):
        table[0][i] = 1
    # Điền các mục còn lại trong bảng theo cách từ dưới lên
    for i in range(1, n+1):
        for j in range(m):
            # Số lượng giải pháp bao gồm S[j]
            x = table[i - S[j]][j] if i-S[j] >= 0 else 0
            # Số lượng giải pháp không bao gồm S[j]
            y = table[i][j-1] if j >= 1 else 0
            # tổng số
            table[i][j] = x + y
    return table[n][m-1]
# Driver program to test above function
arr = [1, 2, 3]
m = len(arr)
n = 4
print(count(arr, m, n))


def count(S, m, n):
    # table[i] sẽ lưu trữ số lượng lời giải cho
     # giá trị i. Chúng ta cần n+1 hàng khi xây dựng bảng
     # theo cách từ dưới lên sử dụng trường hợp cơ sở (n = 0)
     # Khởi tạo tất cả các giá trị trong bảng là 0
    table = [0 for k in range(n+1)]
    # Trường hợp cơ sở (Nếu giá trị cho trước là 0)
    table[0] = 1
    # Chọn từng đồng xu một và cập nhật giá trị bảng[]
     # sau chỉ số lớn hơn hoặc bằng giá trị của
     # nhặt được đồng xu
    for i in range(0,m):
        for j in range(S[i],n+1):
            table[j] += table[j-S[i]]
    return table[n]
# Driver program to test above function
arr = [1, 2, 3]
m = len(arr)
n = 4
x = count(arr, m, n)
print (x)



def count_coins(coins, target):
    memo = {}
    def helper(amount, idx):
        # Check if the solution for this subproblem already exists
        if (amount, idx) in memo:
            return memo[(amount, idx)] 
        # Base case: If the target sum is reached
        if amount == 0:
            return 1
        # Base case: If the target sum cannot be reached using remaining coins
        if amount < 0 or idx >= len(coins):
            return 0
        # Recursively calculate the number of possible ways using the current coin or skipping it
        memo[(amount, idx)] = helper(amount - coins[idx], idx) + helper(amount, idx + 1)
        return memo[(amount, idx)]
    # Call the recursive function with the initial parameters
    return helper(target, 0)
# Test the function
arr = [1, 2, 3]
n = 4
x = count_coins(arr, n)
print(x)