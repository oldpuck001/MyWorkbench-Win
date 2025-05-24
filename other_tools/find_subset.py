#find_subset.py

import pandas as pd
import os

def find_subset():

    target = 1062051.89
    tolerance=0.01

    file_path = os.path.join('C:\\Users\\lilei\\Desktop', 'book1.xlsx')
    nums_df = pd.read_excel(file_path)

    # 将第一列转换为列表
    numbers = nums_df.iloc[:, 0].tolist()
    print(numbers)

    # 首先尝试将浮点数转换为整数进行处理
    scaled_numbers = [round(num * 100) for num in numbers]
    scaled_target = round(target * 100)
    
    # 整数版本的动态规划
    dp = {0: []}
    for num in scaled_numbers:
        print(num)
        new_dp = {}
        for s in dp:
            new_sum = s + num
            if new_sum not in dp and new_sum not in new_dp:
                new_dp[new_sum] = dp[s] + [num]
        dp.update(new_dp)
    
    # 在整数版本中寻找最佳匹配
    best_diff = float('inf')
    best_subset = None
    for s in dp:
        current_diff = abs(s - scaled_target)
        if current_diff <= tolerance * 100 and current_diff < best_diff:
            best_diff = current_diff
            best_subset = [num / 100 for num in dp[s]]
    
    if best_diff <= tolerance * 100:
        return best_subset
    
    # 如果整数版本没有找到足够接近的解，尝试浮点数版本（带容差）
    dp_float = {0: []}
    for num in numbers:
        print(num)
        new_dp = {}
        for s in dp_float:
            new_sum = s + num
            # 检查是否足够接近目标
            if abs(new_sum - target) <= tolerance:
                return dp_float[s] + [num]
            if new_sum not in dp_float and new_sum not in new_dp:
                new_dp[new_sum] = dp_float[s] + [num]
        dp_float.update(new_dp)
    
    # 在浮点数版本中寻找最接近的解
    best_diff = float('inf')
    best_subset = None
    for s in dp_float:
        current_diff = abs(s - target)
        if current_diff < best_diff:
            best_diff = current_diff
            best_subset = dp_float[s]
    
    if best_diff <= tolerance:
        return best_subset
    else:
        return None


result = find_subset()
print(result)