from re import A


grades1=float(input("enter your grades in maths"))
grades2=float(input("enter your grades in english"))
grades3=float(input("enter your grades in kiswahili grades"))
grades4=float(input("enter your grades in science grades"))

grades=(grades1+grades2+grades3+grades4)
average=grades/4
print (average)

if average>=70 and average<=100:
    grades=A
    print(A)
if average>=60 and average<=70:
    B=average
    grades=B
    print(B)
if average>=40 and average<=60:
    C=average
    grades=C
    print(C)
if average>=0 and average<=40:
    D=average
    grades=D
    print(D)

    
    
    

    
    
                     