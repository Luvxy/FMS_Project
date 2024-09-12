add_month = 200000
benefit = 0.5
month = 50
benefit_number = 15
total = 0

for i in range(month):
    total += add_month
    for j in range(benefit_number):
        total += (total*benefit)/100
    print(f"price : {total}")