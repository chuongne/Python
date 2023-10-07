def find_min_sum(num):
    min_sum = num
    for i in range(2, int(num**0.5) + 1):
        if num % i == 0:
            factor = num // i
            min_sum = min(min_sum, i + factor)
    print(i)
    return min_sum
 
# driver code
number = 12
 
# Call the function and print the result
result = find_min_sum(number)
print("The minimum sum of factors for", number, "is", result)




'''Đầu vào : 12 
Đầu ra : 7 
Giải thích:
 Sau đây là các cách khác nhau để phân tích 12 và 
tổng các thừa số theo những cách khác nhau. 
12 = 12 * 1 = 12 + 1 = 13 
12 = 2 * 6 = 2 + 6 = 8 
12 = 3 * 4 = 3 + 4 = 7 
12 = 2 * 2 * 3 = 2 + 2 + 3 = 7 
Do đó tối thiểu tổng là 7 
Đầu vào : 105 
Đầu ra : 15'''