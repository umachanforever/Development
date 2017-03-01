# -*- coding: utf-8 -*-

#递归函数,容易溢出
def fact(n):
    if n==1:
        return 1
    return n * fact(n - 1)
	
#尾递归优化，超出时，显示数值本身
	
def fact(n):
    return fact_iter(n,1)

def fact_iter(num,product):
    if  num ==1:
        return product
    return fact_iter(num-1,num*product)

#汉诺塔
def move(n,x,y,z):
    if n==1:
        print(x,'-->',z)
    else:
        move(n-1,x,z,y)#将前n-1个盘子从x移动到y上
        move(1,x,y,z)#将最底下的最后一个盘子从x移动到z上
        move(n-1,y,x,z)#将y上的n-1个盘子移动到z上
n=int(input('请输入汉诺塔的层数：'))
move(n,'x','y','z')



