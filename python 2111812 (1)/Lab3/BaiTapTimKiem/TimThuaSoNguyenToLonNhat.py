import math
def maxPrimeFactors (n):
	# Khởi tạo số nguyên tố lớn nhất
    # biến có giá trị thấp nhất
	maxPrime = -1
	# In ra số 2 chia hết cho n
	while n % 2 == 0:
		maxPrime = 2
		n >>= 1 # tương đương với n /= 2
        # n chắc chắn là số lẻ ở điểm này,
        # do đó bỏ qua các số chẵn và
        # chỉ lặp với số nguyên lẻ
	for i in range(3, int(math.sqrt(n)) + 1, 2):
		while n % i == 0:
			maxPrime = i
			n = n / i
	# Điều kiện này là để xử lý
    # trường hợp n là số nguyên tố
    # lớn hơn 2
	if n > 2:
		maxPrime = n
	return int(maxPrime)
# Mã driver để kiểm tra chức năng trên
n = 15
print(maxPrimeFactors(n))
n = 25698751364526
print(maxPrimeFactors(n))

