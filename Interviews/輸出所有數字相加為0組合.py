def find_zero_sum_combinations(nums):
    result = []
    nums.sort()  
    n = len(nums)

    for i in range(n - 2):
        if i > 0 and nums[i] == nums[i - 1]:
            continue

        left, right = i + 1, n - 1

        while left < right:
            total = nums[i] + nums[left] + nums[right]
            if total == 0:
                result.append([nums[i], nums[left], nums[right]])
                while left < right and nums[left] == nums[left + 1]:
                    left += 1
                while left < right and nums[right] == nums[right - 1]:
                    right -= 1
                left += 1
                right -= 1
            elif total < 0:
                left += 1
            else:
                right -= 1

    return result

# 可從input_list中自由輸入值，會自動輸出所有數字相加為0組合
input_list = [-9, 0, 1, 3, -1, -7, 8, -1 , -2, -3 ]
result = find_zero_sum_combinations(input_list)
print(result)
