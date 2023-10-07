import math
def sumofoddFactors( n ):
   res = 1
   # bỏ qua các thừa số chẵn bằng
    # trên 2
   while n % 2 == 0:
       n = n // 2
     
   for i in range(3, int(math.sqrt(n) + 1)):
    #Trong khi i chia n, in
        # i và chia n
       count = 0
       curr_sum = 1
       curr_term = 1
       while n % i == 0:
           count+=1
           n = n // i
           curr_term *= i
           curr_sum += curr_term
       res *= curr_sum
   # Điều kiện này là
    # xử lý vụ việc khi
    # n là số nguyên tố.
   if n >= 2:
       res *= (1 + n)
   return res
# Driver code
n = 30
print(sumofoddFactors(n))
