
#---CORE REQUIREMENTS---

#Requirement No.1 Read the data from the spreadsheet
import csv
def read_data():
    data = []
    with open('sales.csv', 'r') as sales_csv:
        spreadsheet = csv.DictReader(sales_csv)
        for row in spreadsheet:  # Ensure this is inside the 'with' block
            data.append(row)
    return data
# read_data()


#Requirement No.2 Collect all of the sales from each month into a single list
def monthly_sales():
    data = read_data()                  # using read_data function to read file
    for row in data:
        print(row['month'], row['sales'])
monthly_sales()


#Requirement No.3 Output the total sales across all months
#updated version of the example code
def total_sales():
    data = read_data()
    sales = [int(row['sales']) for row in data]     #list comprehension: storing each iteration in a list saved to the sales variable
    total = sum(sales)
    print(f'Total sales: {total}')
    return total
total_sales()


# ---EXTENSIONS---

#Monthly changes as a percentage list
def change_percent():
    data = read_data()
    sales = [int(row['sales']) for row in data]
    for i in range(len(sales)-1):                   #len-1 because you will use i + 1 below
        old = sales[i]                              #eg. Jan sales
        new = sales[i+1]                            #eg. Feb sales
        percent_change = round(((new - old) / old) * 100,2)
        print(f"% change: {percent_change} %")
change_percent()

#Monthly changes as a percentage based on input
def percentage_difference():
  data = read_data()
  sales_by_month = {row['month'].lower(): int(row['sales']) for row in data}
  month1 = input("Enter a month (e.g., jan): ").lower()
  month2 = input("Enter another month you want to compare the first months sales against (e.g., feb): ").lower()
  sales1 = sales_by_month[month1]
  sales2 = sales_by_month[month2]
  percentage_difference = round(((sales2 - sales1) / sales1) * 100,2)
  print(f"Sales in {month1}: {sales1}")
  print(f"Sales in {month2}: {sales2}")
  print(f"Percentage difference in sales between {month1} and {month2}: {percentage_difference}%")
# percentage_difference()

#Average option 1:
def avg_sales1():
    data = read_data()
    sales = [int(row['sales']) for row in data]
    AVG = round(sum(sales)/len(sales),2)                # use sum()/len()
    print(f'Average sales: {AVG}')
    return AVG
# avg_sales1()

#alternative option 2:
def avg_sales2():
    from statistics import mean                         # import mean from statistics module
    data = read_data()
    sales = [int(row['sales']) for row in data]
    avg = round(mean(sales),2)                          #call mean function for list saved in sales variable, then round to 2 decimal places
    print(f"Average: {avg}")
avg_sales2()


#Highest sale
def highest_sale():
    data = read_data()
    sales = [int(row['sales']) for row in data]
    highest_sales = max(sales)
    print(f"Highest sales figure: {highest_sales}")
    return highest_sales        # added for exporting to spreadsheet below
highest_sale()

#Lowest sale
def lowest_sale():
    data = read_data()
    sales =[int(row['sales']) for row in data]
    lowest = min(sales)
    print(f"Lowest sales figure: {lowest}")
    return lowest # added for exporting to spreadsheet below
lowest_sale()



# Data visualisation using Seaborn

import seaborn as sns
import matplotlib.pyplot as plt                 #not required in googleColab
import pandas as pd

data = pd.read_csv("sales.csv")                 #access sales.csv using pandas + read_Csv function
sns.barplot(data = data, x="month", y="sales")  #this line can be amended for different types of graphs or additional features

plt.title("Sales per Month")                    #naming the Title and axis
plt.xlabel("month")
plt.ylabel("Sales Amount")

# plt.show()                                      #call function and graph pops up





#Output summary of the results to a spreadsheet


#update monthly_sales function so it returns a list containing dictionaries.
def monthly_sales1():
    data = read_data()
    monthly_output = [] # variable ready to have a list of dictionaries added
    for row in data:
        output = {'Month':row['month'],'Sales':row['sales']} # each output line is a dictionary being stored in a temporary variable so we can use .append() below
        monthly_output.append(output)   # adding to the monthly_output list
    # print(monthly_output)             # use this to test the output is looking correct in the terminal
    return monthly_output
# monthly_sales()


#Create a dictionary with a summary of our data
summary = {
    'Summary' :['Average Sales', 'Highest Sale', 'Lowest Sale', 'Total Sales'],
    'Result' : [avg_sales1(), highest_sale(), lowest_sale(), total_sales()]         #calling functions with print statements so this re-prints the outputs
}

#--Exporting to Excel--

#OPTION 1: export to Excel workbook with one sheet only

# import pandas as pd           #pandas required to export to excel - already imported above for seaborn
dataset1 = monthly_sales1()     #specifying dataset - calling the updated monthly_sale1 function above
df1 = pd.DataFrame(dataset1)    #turning dataset into dataframe(df) - pd calls panda module, + DataFrame() for df creation
# print(df1)                      #test to check df is being created
df1.to_excel('Python_Project111.xlsx', sheet_name ='Monthly_sales') #create & populate Excel file (only one page)
#variable with dataframe df1. followed by to_excel()function, in brackets('create name of Excel file.xlsx', sheet_name = name of Excel sheet) (Excel files are .xlsx - MUST INCLUDE THIS)
#cannot repeat this line to create a different sheet as it will just overwrite what is already there


#OPTION 2: export to Excel workbook with multiple sheets

#specifying datasets
dataset2 = monthly_sales1()  #calling the updated monthly_sale1 function above
dataset3 = summary

#create dataframe(df)
df1 = pd.DataFrame(dataset2) #pd calls panda module, + DataFrame() for df creation
df2 = pd.DataFrame(dataset3)

with pd.ExcelWriter('Python_Project_x3.xlsx') as writer: # need to use ExcelWriter for multiple sheets and name the Excel file .xlsx

    df1.to_excel(writer, sheet_name='Monthly_Sales') # to_excel function for each sheet, instead of file name, use 'writer', followed by sheet name
    df2.to_excel(writer, sheet_name='Summary')

#Excel file will appear on the left hand side, right click --> Open in --> Open in Associated Application. Excel file will open with populated summary info.
