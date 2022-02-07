salary=int(input("enter your salary"))
year_of_service=int(input("enter your year of service"))
if year_of_service<=6 and year_of_service<=10:
    bonus=year_of_service*0.08
    total=bonus+salary
    print(total)
if year_of_service>10:
    bonus=year_of_service*0.1
    total=bonus+salary
    print(total)
if year_of_service<6:
    bonus=year_of_service*0.02
    total=bonus+salary
    print(total)
    