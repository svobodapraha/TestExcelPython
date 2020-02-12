import xlrd
import xlsxwriter
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
#files to compare
workbook = xlrd.open_workbook(sys.argv[1])
workbook_new = xlrd.open_workbook(sys.argv[2])

sheet = workbook.sheet_by_index(0)
sheet_new = workbook_new.sheet_by_index(0)

print("Files to compare:\nold: {} \nnew: {}\n".format(sys.argv[1], sys.argv[2]))

#report file
report_sheet_name = sys.argv[2]
report_sheet_name = report_sheet_name[report_sheet_name.rfind("\\")+1:]
if report_sheet_name.rfind(".") >= 0:
    report_sheet_name = report_sheet_name[:report_sheet_name.rfind(".")]
report_file_name = report_sheet_name+"_compare.xlsx"

#print (report_file_name)
#print (report_sheet_name)

workbook_report = xlsxwriter.Workbook(report_file_name)
sheet_report = workbook_report.add_worksheet(report_sheet_name)

sheet_report.write("A1", "Compared files")
sheet_report.write("B1", sys.argv[1])
sheet_report.write("C1", sys.argv[2])

rowReport = 1
colRep_requirement = 0
bold     = workbook_report.add_format({'bold': True})
bg_green = workbook_report.add_format({'bg_color': 'green'})
bg_red   = workbook_report.add_format({'bg_color': 'red'})

sheet_report.write(rowReport, colRep_requirement, "requirement", bold)
colRep_shortReq = 1
sheet_report.write(rowReport, colRep_shortReq, "req short", bold)
colRep_status = 2
sheet_report.write(rowReport, colRep_status, "Status", bold)
colRep_paramName = 3
sheet_report.write(rowReport, colRep_paramName, "Parameter Name", bold)
colRep_paramOld  = 4
sheet_report.write(rowReport, colRep_paramOld, "Old Value", bold)
colRep_paramNew  = 5
sheet_report.write(rowReport, colRep_paramNew, "New Value", bold)
                   
#workbook_report.close()
#sys.exit()

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
      report_ex_param = []
      report_ex_old = []
      report_ex_new = []
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

          report_ex_param.append(col_name)
          report_ex_old.append(oldValue)
          report_ex_new.append(newValue)

#these req are not in new file
  if not idFound:
    print("ID %s" % reqID_short)
    print("********************")
    print("not found: %s"% reqID)
    print()

#excel
    rowReport = rowReport +1
    sheet_report.write(rowReport, colRep_requirement, reqID)
    sheet_report.write(rowReport, colRep_shortReq, reqID_short)
    sheet_report.write(rowReport, colRep_status, "Missing", bg_red)
    
#changed req
  if idFound & (report_str != []):
    print(reported_id)
    print("********************")
    for text_to_print in report_str:
      print(text_to_print)
    print()

    #excel:
    for i_rep, param in enumerate(report_ex_param):
      rowReport = rowReport +1
      if i_rep==0:
        sheet_report.write(rowReport, colRep_requirement, reqID)  
        sheet_report.write(rowReport, colRep_shortReq, reqID_short)
        sheet_report.write(rowReport, colRep_status, "Changed")
      sheet_report.write(rowReport, colRep_paramName, report_ex_param[i_rep])
      sheet_report.write(rowReport, colRep_paramOld,  report_ex_old[i_rep])
      sheet_report.write(rowReport, colRep_paramNew,  report_ex_new[i_rep])
      

#report new req (only in new file):
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
    #excel
    rowReport = rowReport +1
    sheet_report.write(rowReport, colRep_requirement, reqID_new)
    sheet_report.write(rowReport, colRep_shortReq, reqID_short)
    sheet_report.write(rowReport, colRep_status, "New", bg_green)

sheet_report.set_column(0,3,40)
workbook_report.close()
    
    
  
    










