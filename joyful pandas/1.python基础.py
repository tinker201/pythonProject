# 列表推导式
print([i * 2 for i in range(5)])

# 列表推导式两层嵌套
print([i * j for i in [2, 3] for j in [4, 5]])

# 多项赋值
a, b = 3, 4
print(a, b)

# if条件赋值
print('cat' if 2 > 1 else 'dog')
print([i if i < 5 else 5 for i in range(10)])

# 匿名函数
my_func = lambda x: x * 3
print(my_func(3))
my_func2 = lambda a, b: a * b
print(my_func2(3, 4))

# map方法
print(list(map(lambda a: a * 2, range(4))))

print(r'.skghs'[0] == r'.')