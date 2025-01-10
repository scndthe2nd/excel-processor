
#### Process ####
# Pull information from an excel sheet
## Get correct cell collumn

# SAMPLE #
import openpyxl

wb = openpyxl.load_workbook(r'c:*specific destination of file*.xlsx')
sheet = wb.active
x1 = sheet['B3'].value
x2 = sheet['B4'].value
y1 = sheet['C3'].value
y2 = sheet['C4'].value



print(x1,x2,y1,y2)


# Label information by filter

# Create a template for each label
## Define template
# Add templates to a pdf queue
## Create a pdf queue
## add items in to pdf queue 
# Create a pdf from that information
## send to pdf printer

#### Interface ####
# Create user interface
## Create interface window
## Create filepath selection field
## Create browse button
## Create sort dropdown
## Creat "Run" button