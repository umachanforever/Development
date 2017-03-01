#!/usr/bin/env python3
# -*- coding: utf-8 -*-
 
#函数参数

def print_scores(**kw):
    print('      Name  Score')
    print('------------------')
    for name, score in kw.items():
        print('%10s  %d' % (name, score))
    print()

print_scores(Adam=99, Lisa=88, Bart=77)

data = {
    'Adam Lee': 99,
    'Lisa S': 88,
    'F.Bart': 77
}

print_scores(**data)

def print_info(name, *, gender, city='Beijing', age):
    print('Personal Info')
    print('---------------')
    print('   Name: %s' % name)
    print(' Gender: %s' % gender)
    print('   City: %s' % city)
    print('    Age: %s' % age)
    print()

print_info('Bob', gender='male', age=20)
print_info('Lisa', gender='female', city='Shanghai', age=18)

r'''
#结果
      Name  Score
------------------
      Lisa  88
      Bart  77
      Adam  99

      Name  Score
------------------
    Lisa S  88
    F.Bart  77
  Adam Lee  99

Personal Info
---------------
   Name: Bob
 Gender: male
   City: Beijing
    Age: 20

Personal Info
---------------
   Name: Lisa
 Gender: female
   City: Shanghai
    Age: 18
    '''