import xlrd, xlwt
import sys
import datetime
now = datetime.datetime.now()

"""
This is a test case generator for State and Action Test Cases.
By: Sagar Suri
"""


path = "C:\\Users\\....\\test.xlsx"
def prepare(path):
    """
    Opens the business document for reading and workbook for writing.

    @param path: the destination of the business document
    @return: A tuple with the business document and the workbook
    """
    #Open the business document from which cases will be generated 
    BusinessDocument = xlrd.open_workbook(path)
 
    # print the number of sheets, sheet names
    # print(BusinessDocument.nsheets)
    # print(BusinessDocument.sheet_names())

    #Open the excel workbook for writing stuff.
    
    workbook = xlwt.Workbook(encoding = 'ascii')
  
    return (BusinessDocument, workbook)
 
def case_writer(BusinessDocument, workbook):
    """
    Reads the all sheets of the business document and writes to the worksheet.
    External and Internal test cases will written in seperate sheets.

    @param BusinessDocument: Document to read from.
    @param workbook: Document to write to
    @return: None
    """
    # Styling stuff
    font = xlwt.Font()
    font.bold = True
    title_text = xlwt.easyxf("align: vert top, horiz left")
    title_text.font = font
    style_text = xlwt.easyxf("align: wrap on, vert top, horiz left")

    #Create the worksheets for the test cases.
    worksheet_ext = workbook.add_sheet("External")
    worksheet_int = workbook.add_sheet("Internal")

    for worksheet in [worksheet_ext, worksheet_int]:
        # Write the Header row 
        worksheet.write(0, 0, "Residence Type", title_text)
        worksheet.write(0, 1, "HPQC Folder", title_text)
        worksheet.write(0, 2, "User ID", title_text)
        worksheet.write(0, 3, "Testcase ID", title_text) 
        worksheet.write(0, 4, "Test Description", title_text)
        worksheet.write(0, 5, "Pre-condition", title_text)
        worksheet.write(0, 6, "Execution", title_text)
        worksheet.write(0, 7, "Expected Result", title_text)
        worksheet.write(0, 8, "Business Rule/Comment", title_text)
        worksheet.write(0, 9, "Release", title_text)
        worksheet.write(0, 10, "Creator", title_text)
        worksheet.write(0, 11, "Creation date", title_text)
        # Size adjustments 
        worksheet.col(0).width = 30*100
        worksheet.col(1).width = 30*200
        worksheet.col(2).width = 30*100
        worksheet.col(3).width = 30*256
        worksheet.col(4).width = 30*256
        worksheet.col(5).width = 30*100
        worksheet.col(6).width = 30*256
        worksheet.col(7).width = 30*256
        worksheet.col(8).width = 30*250
        worksheet.col(9).width = 30*60
        worksheet.col(10).width = 30*60
        worksheet.col(11).width = 30*60
  
    # The first sheet is revision history so we don't need it.
    business_sheets = BusinessDocument.sheets()[1:]

    # We're going to write to both the external and internal sheets simultaneously
    # a will serve as the counter for external, and b for internal
    a = 1
    b = 1
    # Iterate through the business documents, first iteration would be CRA, next YJ etc
    for sheet in business_sheets:
        for i in range (1, sheet.nrows):
            # easy references
            ActRole = sheet.row_values(i)[0]
            CurrentInternalStatus = sheet.row_values(i)[1]
            CurrentExternalDisplay = sheet.row_values(i)[2]
            Action = sheet.row_values(i)[3]
            MinistryActionDisplay = sheet.row_values(i)[4]
            ResultingInternalStatus = sheet.row_values(i)[5]
            ResultingExternalDisplay = sheet.row_values(i)[6]
            NotifyRoleInternal = sheet.row_values(i)[7]
            NotificationIDInt = sheet.row_values(i)[8]
            NotifyRoleExternal = sheet.row_values(i)[9]
            NotificationIDExt = sheet.row_values(i)[10]
            BusinessRules = sheet.row_values(i)[11]

                    
            # Empty strings don't look nice in test cases so change to N/A
            if CurrentExternalDisplay == "":
                CurrentExternalDisplay = "N/A"
            if CurrentInternalStatus == "":
                CurrentInternalStatus = "N/A"
            if ResultingExternalDisplay == "":
                ResultingExternalDisplay = "N/A"
            if ResultingInternalStatus == "":
                ResultingInternalStatus = "N/A"

             # Setting up the content for the cells
            ExpectedResult = ("Verify that the resulting statuses are: \n" +
                        "External: " + ResultingExternalDisplay + "\n" +
                        "Internal: " + ResultingInternalStatus + "\n")

            Description = ("Verify that when the " + ActRole + " performs the action: " + Action +
            " at the following workflow statuses: \n" +
                "External: " + CurrentExternalDisplay + "\n" +
                "Internal: " + CurrentInternalStatus + "\n" + ExpectedResult)

            Precondition = ""

            Execution = ("1. Log in as " + ActRole + ".\n" + 
                        "2. Progress a director's approval request to the following workflow statuses: \n" +
                        "External: " + CurrentExternalDisplay + "\n" +
                        "Internal: " + CurrentInternalStatus + "\n" + 
                        "3. Perform the following action on the director's approval request:\n" + Action)

            # Special Case
            if ActRole == "System":
              
                Execution = ("1. Progress a director's approval request to the following workflow statuses: \n" +
                            "External: " + CurrentExternalDisplay + "\n" + "Internal: " + CurrentInternalStatus + "\n" +
                            "2. To perform the following action: " + Action + ",\n"
                            "Perform the action listed in the Business Rule/Comment column OR in the associated business rule.")
            
            # Internal and External seperation
            # Assign worksheet and counter increments accordingly
            # a is for external, b is for internal
            if "RU" in ActRole or "SPA" in ActRole or "Site Designate" in ActRole:
                worksheet = worksheet_ext
                c = a
                case_id = "State&Action_" + sheet.name + "_" + str(a)
                a+=1
            
            else:
                worksheet = worksheet_int
                c = b
                case_id = "State&Action_" + sheet.name + "_" + str(b)
                b+=1
            
            # Now that we have all the information, we can just write it to the correct sheet
            worksheet.write(c, 0, sheet.name, style_text)
            worksheet.write(c, 1, "SORRL_DA_State&Action", style_text)
            worksheet.write(c, 2, "surisu", style_text)
            worksheet.write(c, 3, case_id, style_text)
            worksheet.write(c, 4, Description, style_text)
            worksheet.write(c, 5, Precondition, style_text)
            worksheet.write(c, 6, Execution, style_text)
            worksheet.write(c, 7, ExpectedResult, style_text)
            worksheet.write(c, 8, BusinessRules, style_text)
            worksheet.write(c, 9, "2.0", style_text)
            worksheet.write(c, 10, "Sagar", style_text)
            worksheet.write(c, 11, now.strftime("%m-%d-%Y"), style_text)
    
    workbook.save("state_action_cases.xls")



def main():
    #if the user supplied a path, should use that instead.
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else: 
        # my default path
        path = "C:\\Users\\surisu\\OneDrive - Government of Ontario\\Documents\\test.xlsx"
        pass
    Document = prepare(path)
    case_writer(Document[0], Document[1])
    ret_string = "Test cases saved succesfully!"
    print(ret_string)
  

if __name__ == "__main__": 
    main()



