# Farnell CSV Reader
import csv

def readFarnellCSVfile():
    names = []                                      # Names of all who should pay
    orders = []
    duplicateRegister = []                          # To find any duplicate entries that Farnell have put in the csv file
    with open('Farnell.csv', mode='r') as file:
        csvFile = csv.reader(file)
        next(csvFile)                               # Skips header
        for line in csvFile:
            POnumber = line[0]
            order_date = line[6]
            order_code = line[15]
            customer_surname = line[16]
            customer_first_name = line[17].upper()
            order_quantity = line[21]
            unit_price = line[22]
            order_price = line[23]
            customer_full_name = [customer_first_name,customer_surname]
            customer_order = [order_date,order_code,order_quantity,unit_price,order_price]
            duplicateCheck = [customer_full_name,customer_order]
            if duplicateCheck in duplicateRegister:
                print('Duplicate found and removed:',duplicateCheck)
            else:
                duplicateRegister.append(duplicateCheck)
                if customer_full_name in names:
                    customer_index = names.index(customer_full_name)
                    orders[customer_index].append(customer_order)
                else:
                    names.append(customer_full_name)
                    orders.append([])
                    customer_index = names.index(customer_full_name)
                    orders[customer_index].append(customer_order)
    totals = []                                 # The total amounts of money each person should pay
    for person in orders:
        total = 0
        for order in person:
            price = float(order[-1][3:])
            total += price
        totals.append(total)
    allOrdersInCSVfile = []                          # All the orders for all the people
    for number in enumerate(names):
        index = number[0]
        entry = [names[index],orders[index],totals[index],index]
        allOrdersInCSVfile.append(entry)
    return allOrdersInCSVfile, POnumber


def main():
    print('This is the CSV reader file.')
    allOrdersinCSVfile = readFarnellCSVfile()
    print(allOrdersinCSVfile)


if __name__ == '__main__':
    main()