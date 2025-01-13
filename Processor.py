
#### Process ####
# Pull information from an excel sheet
## Get correct cell collumn x

#import openpyxl
#workbook = openpyxl.load_workbook(r'Demo/Demo Sheet.xlsx', data_only=True, read_only=True)
#print(workbook.sheetnames)
#sheet = workbook['MASTER WORKSHEET']
#Sample_Value_1 = sheet['B'].value
#Sample_Value_2 = sheet['B4'].value
#Sample_Value_3 = sheet['C3'].value
#Sample_Value_4 = sheet['B5'].value
#print(Sample_Value_1,Sample_Value_2,Sample_Value_3,Sample_Value_4)

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