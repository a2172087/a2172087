import itertools

# 生成所有由0到9中且不重複數字所組成的五位數
numbers = list(range(10))
combinations = list(itertools.permutations(numbers, 5))
five_digit_numbers = [int(''.join(map(str, combo))) for combo in combinations if combo[0] != 0]

# 最大值和最小值
max_number = max(five_digit_numbers)
min_number = min(five_digit_numbers)

# 判斷質數的函數
def is_prime(num):
    if num <= 1:
        return False
    if num <= 3:
        return True
    if num % 2 == 0 or num % 3 == 0:
        return False
    i = 5
    while i * i <= num:
        if num % i == 0 or num % (i + 2) == 0:
            return False
        i += 6
    return True

# 找到質數
prime_numbers = [num for num in five_digit_numbers if is_prime(num)]
# prime_numbers中包含五位數字的數列
five_digit_prime_numbers = [num for num in prime_numbers if set(str(num)) <= set(['1', '2', '3', '6', '8'])]

# 輸出結果
print("最大值:", max_number)
print("最小值:", min_number)
print("質數是:", prime_numbers)
print("包含1、2、3、6、8的質數是:", five_digit_prime_numbers)
