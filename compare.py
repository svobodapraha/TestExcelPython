import xlrd
import xlwt
import sys

# "..\\testfiles\\LA_TSAnS_baseline_1.8.xlsx" "..\\testfiles\\LA_TSAnS_baseline_2.1.xlsx"

def intersection(lst1, lst2): 
    lst3 = [value for value in lst1 if value in lst2] 
    return lst3

if len(sys.argv) != 3:
   print ("usage: compare old_excel_file new_excel_file")
   sys.exit()  


#workbook = xlrd.open_workbook("..\\testfiles\\LA_TSAnS_baseline_1.8.xlsx")
#workbook_new = xlrd.open_workbook("..\\testfiles\\LA_TSAnS_baseline_2.1.xlsx")
workbook = xlrd.open_workbook(sys.argv[1])
workbook_new = xlrd.open_workbook(sys.argv[2])

sheet = workbook.sheet_by_index(0)
sheet_new = workbook_new.sheet_by_index(0)

print("Files to compare:\nold: {} \nnew: {}\n".format(sys.argv[1], sys.argv[2]))



col_ObjectType = -1
for i in range(0, sheet.ncols):
  value = str(sheet.cell(0,i).value).strip().replace("\r"," ").replace("\n"," ")
  if(value == "Object Type"):
    col_ObjectType = i
    break
#print(col_ObjectType)

col_ObjectTypeNew = -1
for i in range(0, sheet_new.ncols):
  value = str(sheet_new.cell(0,i).value).strip().replace("\r"," ").replace("\n"," ")
  if(value == "Object Type"):
    col_ObjectTypeNew = i
    break
#print(col_ObjectTypeNew)


#find common column for both files for comparing and their position
#find names
col_list=sheet.row_values(0)
col_list.pop(0)
col_list_new=sheet_new.row_values(0)
col_list_new.pop(0)
col_list_common = intersection(col_list, col_list_new)

#find column position in both files
col_list_num = []
col_list_new_num = []
for col_name in col_list_common:
    for i in range (1, sheet.ncols):
      if sheet.cell(0,i).value == col_name:
        #print(i)
        col_list_num.append(i)
    for i in range (1, sheet_new.ncols):
      if sheet_new.cell(0,i).value == col_name:
        #print(i)
        col_list_new_num.append(i)




for i in range(1, sheet.nrows):
#for i in range(1, 10):
  #new row, find req.ID
  reqID = str(sheet.cell(i,0).value).strip().replace("\r"," ").replace("\n"," ")
  reqID_short = reqID[reqID.rindex("_")+1:]
  #find matching req in new file
  idFound = False
  for i_new in range(1, sheet_new.nrows):
    reqID_new = str(sheet_new.cell(i_new,0).value).strip().replace("\r"," ").replace("\n"," ")
    if reqID == reqID_new:
      idFound = True
      #print(reqID, i, i_new)
      #check columns
      reported_id = ""
      report_str = []
      for j, col_name in enumerate(col_list_common):
        k = col_list_new_num[j]
        j = col_list_num[j]
        oldValue = str(sheet.cell(i,j).value).strip().replace("\r"," ").replace("\n"," ")
        newValue = str(sheet_new.cell(i_new,k).value).strip().replace("\r"," ").replace("\n"," ")
        if((oldValue != newValue) and \
           (\
           (str(sheet.cell(i,col_ObjectType).value).strip() == "Functional Requirement")
           or \
           (str(sheet_new.cell(i_new,col_ObjectTypeNew).value).strip() == "Functional Requirement") \
           )):

          reported_id = "ID {}".format(reqID_short)
          report_str.append("{} OLD:".format(col_name))  
          report_str.append("  {}".format(oldValue))
          report_str.append("{} NEW:".format(col_name))
          report_str.append("  {}".format(newValue))

  if not idFound:
    print("ID %s" % reqID_short)
    print("********************")
    print("not found: %s"% reqID)
    print()
    
  if idFound & (report_str != []):
    print(reported_id)
    print("********************")
    for text_to_print in report_str:
      print(text_to_print)
    print()

#report new req:
for i_new in range(1, sheet_new.nrows):

  #new row, find req.ID
  reqID_new = str(sheet_new.cell(i_new,0).value).strip().replace("\r"," ").replace("\n"," ")
  reqID_short = reqID_new[reqID.rindex("_")+1:]
  #find matching req in old file (if not found it means new req)
  idFound = False
  for i in range(1, sheet.nrows):
    reqID_old = str(sheet.cell(i,0).value).strip().replace("\r"," ").replace("\n"," ")
    if reqID_old == reqID_new:
      idFound = True
  

  if not idFound:
    print("ID %s" % reqID_short)
    print("********************")
    print("new req: %s"% reqID_new)
    print()
    
  
    










