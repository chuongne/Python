def triangle(input1):
    if input1 < 1:
        return 'Zero'
    else:
          # Bộ đếm để kiểm soát tiền liên tiếp
        a = 1       
        height = 0
        while input1 > 0:
            if input1 - a >= 0:
                input1 = input1 - a
                a = a + 1
                height = height + 1
                #print(input1)
                print(a)
            else:
                break
        return height
    
if __name__ == "__main__":
    testinput1 = 22
    print(triangle(testinput1))