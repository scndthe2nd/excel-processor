
#### Process ####
# Pull information from an excel sheet
## Get correct cell collumn x

import openpyxl
workbook = openpyxl.load_workbook(r'Demo/Demo Sheet.xlsx', data_only=True, read_only=True)
#print(workbook.sheetnames)
sheet = workbook['MASTER WORKSHEET']
devices = [] # Create list "devices"
for i, column in enumerate(sheet):
        if i <= 1: # Skip number of rows
                continue
        device = column[1].value # get value from column number, starting at 0
        devices.append(device) # add to list "devices"
print (devices)


config_book = openpyxl.load_workbook(r'Config/Config.xlsx', data_only=True, read_only=True)
#print(config_book.sheetnames)
device_sheet = config_book['Device Types']
device_types = [] # Create list "devices"
for i, column in enumerate(device_sheet):
        if i <= 0: # Skip number of rows
                continue
        device_type = column[1].value # get value from column number, starting at 0
        device_types.append(device_type) # add to list "devices"
print (device_types)

test_sheet = config_book['Test Types']
test_types = [] # Create list "devices"
for i, column in enumerate(test_sheet):
        if i <= 0: # Skip number of rows
                continue
        test_type = column[0].value # get value from column number, starting at 0
        test_types.append(test_type) # add to list "devices"
print (test_types)


#Sample_Value_1 = sheet['B3:B11']
#Sample_Value_2 = sheet['B4'].value
#Sample_Value_3 = sheet['C3'].value
#Sample_Value_4 = sheet['B5'].value
#print(Sample_Value_1)
#print(Sample_Value_2)
#print(Sample_Value_3)
#print(Sample_Value_4)

# Label information by filter
## Create a list of POSSIBLE_VALUES that SAMPLE_VALUE can have
## Create a list of OUTPUT_VALUES
## Create an association of POSSIBLE_VALUES with OUTPUT VALUES
## Relate OUTPUT_VALUE to SAMPLE_VALUE

## IF SAMPLE_VALUE = A; THEN OUTPUT_VALUE = 1,2,3
## IF SAMPLE_VALUE = B; THEN OUTPUT_VALUE = 2,4,6

# SAMPLE # 
# Create an array
#import array as arr
#b1= "stuff"
#Possible_Value = arr.array('i', [1,2,3,4,5,6,7,8,9,0])
#Sample_Value = ("1","2","3","4","5")
#Association = arr.array( "a" ,["1","2","3"])
#print(Possible_Value[4:8])

# Create a template for each label
## Define template

# Add values to a list



# Set up HTML page for print
## these are items that create the PDF Queue
## Add associated listed items into 


# Sample #
html_open = "<html><body>"
header_block = "<h1>Header</h1><br>"
html_close = "</body></html>"

QUEUE = ""
file = open('view.html', 'a')
file.write(html_open)
file.write(header_block)
#for each TEMPLATE_ITEM in QUEUE
#    echo TEMPLATE_ITEM >> view.html
file.write(html_close)
file.close()


# Create a pdf from that information
## send to pdf printer

#### Interface ####
# Create user interface
## Create interface window
## Create filepath selection field
## Create browse button
## Create sort dropdown
## Creat "Run" button