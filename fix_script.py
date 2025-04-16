#!/usr/bin/env python
# -*- coding: utf-8 -*-

# 脚本用于修复电费节省使用最佳容量和自定义容量对比.py文件中的重复标志定义问题

with open('电费节省使用最佳容量和自定义容量对比.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 替换重复的变量定义
fixed_content = content.replace(
    "                        # 添加一个标志，表示第一次套利是否成功放电\n                        first_arbitrage_completed = True",
    "                        # 使用前面已定义的first_arbitrage_completed标志"
)

# 写入修复后的文件
with open('电费节省使用最佳容量和自定义容量对比_fixed.py', 'w', encoding='utf-8') as f:
    f.write(fixed_content)

print("修复完成，新文件已保存为：电费节省使用最佳容量和自定义容量对比_fixed.py") 