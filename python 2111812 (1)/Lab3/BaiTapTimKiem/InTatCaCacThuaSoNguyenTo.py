import math
def primeFactors(n):
	# In ra số hai chia hết cho n
	while n % 2 == 0:
		print(2)
		n = n // 2
    #n chắc chắn là số lẻ ở điểm này
    # nên có thể bỏ qua 2 ( i = i + 2)
    #Tạo một dãy số từ 3 đến int(math.sqrt(n))+1 nhưng tăng thêm 2 thay vì 1:
	for i in range(3,int(math.sqrt(n))+1,2): 
		# trong khi i chia n , print i ad chia n
		while n % i== 0:
			print(i)
			n = n // i #chia lam tron	
	# Điều kiện nếu n là số nguyên tố
    # số lớn hơn 2
	if n > 2:
		print(n)
# Chương trình Driver để kiểm tra chức năng trên
n = 315
primeFactors(n)


