
#### Process ####
# Pull information from an excel sheet
## Get correct cell collumn x

test_workbook = 'Demo/Demo Sheet.xlsx'
config_workbook = 'Config/Config.xlsx'
print_preview = 'view.html'

import openpyxl

# Pull information from Demo Resource Sheet
workbook = openpyxl.load_workbook(test_workbook, data_only=True, read_only=True)
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
config_book = openpyxl.load_workbook(config_workbook, data_only=True, read_only=True)
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
## these are items that create the print preview
## Add associated listed items into print preview

Test_Header = "<table><tr><th>TESTNAME</th> <th>Yes</th> <th>No</th> <th>N/A</th> <th>Comment</th></tr>"
Test_Line = '<td></td><td><input type="checkbox"></td> <td><input type="checkbox"></td> <td><input type="checkbox"></td> <td>_______________________________________________</td></tr>'
Test_Close = "</table>"







html_open = "<html><body>"
header_block = "<h1>Header</h1><br>"
html_close = "</body></html>"

# Define Print Preview File and Clear Print Preview

file = open(print_preview, 'w')
file.write("")
file.close()
# Open Print Preview to be appended
file = open(print_preview, 'a')
file.write(html_open)
file.write(header_block)
file.write(Test_Header)
file.write(Test_Line)
file.write(Test_Line)
file.write(Test_Line)
file.write(Test_Line)
file.write(Test_Line)
file.write(Test_Close)

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