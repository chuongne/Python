def binarypalindrome(num):
    #chuyển số sang nhị phân
    binary = bin(num)
    # bỏ qua hai ký tự đầu tiên của chuỗi
     # vì hàm bin thêm '0b' vào
     # tiền tố trong biểu diễn nhị phân của
     # một số
    binary = binary[2:]
    # bây giờ đảo ngược chuỗi nhị phân và so sánh
     # nó có bản gốc
    return binary == binary[-1::-1]
#binary[-1::-1] đảo ngược chuỗi
# Driver program
if __name__ == "__main__":
    num = 9
    print (binarypalindrome(num))