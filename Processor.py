
#### Process ####
# Pull information from an excel sheet
## Get correct cell collumn x

import openpyxl

# Pull information from Demo Resource Sheet
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


# Pull information from Config Workbook, this has the device types and device tests
config_book = openpyxl.load_workbook(r'Config/Config.xlsx', data_only=True, read_only=True)
#print(config_book.sheetnames)
device_sheet = config_book['Device Types']
device_types = [] 
for i, column in enumerate(device_sheet):
        if i <= 0: 
                continue
        device_type = column[1].value 
        device_types.append(device_type) 
print (device_types)

test_sheet = config_book['Test Types']
test_types = [] 
for i, column in enumerate(test_sheet):
        if i <= 0:
                continue
        test_type = column[0].value 
        test_types.append(test_type) 
print (test_types)



# Label information by filter
## Create a list of POSSIBLE_VALUES that SAMPLE_VALUE can have
## Create a list of OUTPUT_VALUES
## Create an association of POSSIBLE_VALUES with OUTPUT VALUES
## Relate OUTPUT_VALUE to SAMPLE_VALUE

## IF SAMPLE_VALUE = A; THEN OUTPUT_VALUE = 1,2,3
## IF SAMPLE_VALUE = B; THEN OUTPUT_VALUE = 2,4,6


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