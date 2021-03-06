import os, sys, distutils, psutil
import csv
import agate
import agateexcel
import re
import locale

locale.setlocale(locale.LC_ALL, '')


# CUST_BRYAN_withdates_12-29-16.xlsx

# Make sure using Python 3
if sys.version_info[0] < 3:
    sys.exit('Requires Python 3.0 or higher')


## SETUP
# file = 'customers_trans_dates.csv'
seperator = '\n' + '========================================' + '\n'
yes = set(['yes','ye','y',''])
no = set(['no','n'])


##################
### processMissingBillTos
##################

def processMissingBillTos():
    # openfile as an agate Table
    # can open csv, xls or xlsx
    if openFile.endswith('.csv'):
        data = agate.Table.from_csv(openFile)
    elif openFile.endswith('.xls'):
        data = agate.Table.from_xls(openFile)
    elif openFile.endswith('.xlsx'):
        data = agate.Table.from_xlsx(openFile)
    else:
        sys.exit('Aborting... csv, xls, or xlsx files only')


    # Rows in original file
    # rc = row count
    rc = data.aggregate(agate.Count())
    print(seperator + 'Original File - Number of rows: ' + str(rc) + seperator)

    # Print original table
    data.print_table(max_rows=10,max_columns=10)

    # def to finish processMissingBillTos
    def continueProcessing():

        # Save customers that have bill to to a new csv
        haveBillTo = data.where(lambda row: row['Bill_To'] != 0 )
        rc = haveBillTo.aggregate(agate.Count())
        print(seperator + 'Have Bill To - Number of rows: ' + str(rc) + seperator)
        haveBillTo.print_table(max_rows=10,max_columns=10)
        haveBillTo.to_csv('customers_with_BillTo.csv')

        # customers missing Bill to
        missingBillTo = data.where(lambda row: row['Bill_To'] == 0 )
        rc = missingBillTo.aggregate(agate.Count())
        print(seperator + 'Missing Bill To - Number of rows: ' + str(rc) + seperator)
        missingBillTo.print_table(max_rows=10,max_columns=10)

        # copy ship to as bill to
        ## THIS IS A HACK FOR NOW ##
        ## Assumes that the bill to is 0 and then adds the Customer Number (Ship To) to it
        ## Maybe set columns as .Text() or str() and then copy directly?
        replacedBillTo = missingBillTo.compute([
            ('Bill_To', agate.Change('Bill_To', 'Customer'))
        ], replace=True)

        # save the customers with their bill tos applied from ship to
        replacedBillTo.to_csv('customers_with_replaced_BillTo.csv')

        # Count rows
        rc = replacedBillTo.aggregate(agate.Count())

        # print the table
        print(seperator + 'Replaced Bill_To - Number of rows: ' + str(rc) + seperator)
        replacedBillTo.print_table(max_rows=10,max_columns=10)

    # Do we need to filter?
    filterDate = input('Do you want to filter to date of last purchase? (y/n) ')

    if filterDate in yes:
        as400Date = input('What is the date in the AS400 date format?:  ')
        data = data.where(lambda row: int(as400Date) < row['Date_Last_Pmt'])
        data.print_table(max_rows=10, max_columns=10)
        # Did that filter look good?
        goAhead = input('Looks good? (y/n):  ')
        if goAhead in yes:
            # Then keep on going
            continueProcessing()
        else:
            # Start processing back over
            processMissingBillTos()
        # continueProcessing()
    elif filterDate in no:
        # No date filter needed. Continue processing
        continueProcessing()


def createNewCustomerFile():
    # join both csv files

    # create tables
    haveBillTo = agate.Table.from_csv('customers_with_BillTo.csv')
    replacedBillTo = agate.Table.from_csv('customers_with_replaced_BillTo.csv')
    # merge tables
    newCustomerTable = agate.Table.merge([haveBillTo,replacedBillTo])
    sortednewCustomerTable = newCustomerTable.order_by('Customer')

    newFileName = input('New csv file name... (include \'.csv\'): ')
    print('Creating file...')
    sortednewCustomerTable.to_csv(newFileName)

    # sort by ship to number

openFile = input('What file do you want process? (csv, xls, xlsx) ')

if openFile:
    processMissingBillTos()

joinFiles = input("Create new customers file with Bill To's? (y/n): ")

if joinFiles in yes:
    createNewCustomerFile()
elif joinFiles in no:
    sys.exit('Closing')
else:
    print("Please respond with 'yes' or 'no'.")


print('\n')
#
# clean_salesPerson_data.print_table();

#### Filter is working
# Filter value
# greaterThan = 1000
# data2 = data.where(lambda row: greaterThan < row['Sales'])
#
# # Count filtered rows
# rc2 = data2.aggregate(agate.Count())
#
# # print Header
# print(seperator)
# print('Sales above ' + str(greaterThan))
# # print row count
# print('Row Count: ' + str(rc2))
# print(seperator)
#
# data2.print_table()
# data2.to_csv('test2.csv')


# n2 = data2.Count('Sales')
# print = n2

# customers.print_table()
# Get all transactions after Jan 01, 2014
# customersAfter = data.where(lambda row: 1140101 < row['Date_Last_Pmt'])
#
# customersAfter.print_table()

#
# saveFile = input("Save file? (y/n)")
#
# yes = set(['yes','ye','y',''])
# no = set(['no','n'])
#
# if saveFile in yes:
#     print('Ok, I will save it')
#     data.to_csv('test2.csv')
#
#
# elif saveFile in no:
#     sys.exit('Ok, I will end script')
# else:
#     print("Please respond with 'yes' or 'no'.")
#     saveFile = input("Save file? (y/n)")
