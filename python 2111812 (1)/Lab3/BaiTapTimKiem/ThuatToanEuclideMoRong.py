def gcdExtended(a, b):
    # Base Case
    if a == 0 :
        return b,0,1        
    gcd,x1,y1 = gcdExtended(b%a, a)
    # Cập nhật x và y sử dụng kết quả đệ quy
     # gọi
    x = y1 - (b//a) * x1
    y = x1
    return gcd,x,y
     
 
# Driver code
a, b = 35,15
g, x, y = gcdExtended(a, b)
print("gcd(", a , "," , b, ") = ", g)