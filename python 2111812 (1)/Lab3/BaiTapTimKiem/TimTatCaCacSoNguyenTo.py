def SieveOfEratosthenes(num):
    prime = [True for i in range(num+1)]
    # boolean array
    p = 2
    while (p * p <= num):
        # Nếu prime[p] không phải
        # đã thay đổi thì đó là số nguyên tố
        if (prime[p] == True):
            # Cập nhật tất cả bội số của p
            for i in range(p * p, num+1, p):
                prime[i] = False
        p += 1
        #print(p)
    # In tất cả các số nguyên tố
    for p in range(2, num+1):
        if prime[p]:
            print(p)
    #nó sẽ xuất các giá trị đúng
    #print(prime)
# Driver code
if __name__ == '__main__':
    num = 30
    print("Following are the prime numbers smaller"),
    print("than or equal to", num)
    SieveOfEratosthenes(num)