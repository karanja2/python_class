from ast import And


A=float(input("enter the first number"))
B=float(input("enter the second number"))
C=float(input("enter the third number"))

if A>B and A>C:
    print("the largest number is",A) 
if B>A and B>C:
    print("the largest number is",B)
else:
    print("the largest number is ",C)

    