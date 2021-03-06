import xlrd, xlwt
import sys
import datetime
now = datetime.datetime.now()

"""
This is a test case generator for State and Action Test Cases.
By: Sagar Suri
"""


# path = "C:\\Users\\....\\name.xlsx"
def prepare(path):
    """
    Opens the business document for reading and workbook for writing.

    @param path: the destination of the business document
    @return: A tuple with the business document and the workbook
    """
    #Open the business document from which cases will be generated 
    BusinessDocument = xlrd.open_workbook(path)
 
    # print the number of sheets, sheet names
    print(BusinessDocument.nsheets)
    print(BusinessDocument.sheet_names())

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
    worksheet = workbook.add_sheet("Dashboard")

    # for worksheet in [worksheet_ext, worksheet_int]:
    # Write the Header row 
    worksheet.write(0, 0, "Test Type", title_text)
    worksheet.write(0, 1, "HPQC Test Plan Folder name", title_text)
    worksheet.write(0, 2, "Your User ID", title_text)
    worksheet.write(0, 3, "Residence Type", title_text)
    worksheet.write(0, 4, "Test Case ID(Business ReqID+ BR/DFS/SC+screen Name)", title_text) 
    worksheet.write(0, 5, "Test Case  Description", title_text)
    worksheet.write(0, 6, "Pre-Condition", title_text)
    worksheet.write(0, 7, "Execution", title_text)
    worksheet.write(0, 8, "Expected Result", title_text)
    worksheet.write(0, 9, "Release", title_text)
    worksheet.write(0, 10, "Created by(optional)", title_text)
    worksheet.write(0, 11, "Creation Date(optional)", title_text)

    # Size adjustments 
    worksheet.col(0).width = 30*70
    worksheet.col(1).width = 30*200
    worksheet.col(2).width = 30*70
    worksheet.col(3).width = 30*70
    worksheet.col(4).width = 30*180
    worksheet.col(5).width = 30*270
    worksheet.col(6).width = 30*70
    worksheet.col(7).width = 30*256
    worksheet.col(8).width = 30*220
    worksheet.col(9).width = 30*70
    worksheet.col(10).width = 30*70
    worksheet.col(11).width = 30*100

  
    # # The first sheet is revision history so we don't need it.
    # business_sheets = BusinessDocument.sheets()[1:]

    # We're going to write to both the external and internal sheets simultaneously
    # a will serve as the counter for external, and b for internal
     
    c = 1
    # Iterate through the business documents, first iteration would be CRA, next YJ etc

    
    RuleSheet = BusinessDocument.sheets()[0]

    RoleNameList = []
    for num in range(5, 20):
        RoleName = RuleSheet.row_values(0)[num]
        
        for i in range (1, RuleSheet.nrows):
            # easy references
            ExternalStatus = RuleSheet.row_values(i)[0]
            InternalStatus = RuleSheet.row_values(i)[1]
            Section = RuleSheet.row_values(i)[num]

            # Empty strings don't look nice in test cases so change to N/A
            if ExternalStatus == "":
                ExternalStatus = "N/A"
            
            if InternalStatus == "":
                InternalStatus = "N/A"

            SOR_Status = ("Verify for the " + RoleName + ", when an SOR is under the status:\n\n" +
                            "External: " + ExternalStatus + "\n" +
                            "Internal: " + InternalStatus + "\n\n")
            
            Execution = ("1. Log in as " + RoleName + ".\n" + 
                            "2. Progress an SOR to the following workflow statuses: \n" +
                            "External: " + ExternalStatus + "\n" +
                            "Internal: " + InternalStatus + "\n" + 
                            "3. Verify the location of the SOR in the Dashboard ")

            if Section != "N/A":
                Description = (SOR_Status + "It is in the following section of the dashboard:\n\n" + Section)
                Expected_Result = "The SOR is in the following section on the dashboard:\n\n" + Section
                            
            else: 
                Description = SOR_Status + "It is not present in the dashboard." 
                        
                Expected_Result = "The SOR is not present in the dashboard."


        
            case_id = "SO_Dashboard_" + str(c)

            
            # Now that we have all the information, we can just write it to the correct sheet
            worksheet.write(c, 0, "Manual", style_text)
            worksheet.write(c, 1, "SORRL_SO_Dashboard", style_text)
            worksheet.write(c, 2, "surisu", style_text)
            worksheet.write(c, 3, "All", style_text)
            worksheet.write(c, 4, case_id, style_text)
            worksheet.write(c, 5, Description, style_text)
            worksheet.write(c, 6, "N/A", style_text)
            worksheet.write(c, 7, Execution, style_text)
            worksheet.write(c, 8, Expected_Result, style_text)
            worksheet.write(c, 9, "2.0", style_text)
            worksheet.write(c, 10, "Sagar", style_text)
            worksheet.write(c, 11, now.strftime("%m/%d/%Y"), style_text)
            c+=1

    
    workbook.save("Dashboard_test_Cases.xls")



def main():
    #if the user supplied a path, should use that instead.
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else: 
        # my default path
        path = "C:\\Users\\....\\dashboard.xlsx"
            

    Document = prepare(path)
    case_writer(Document[0], Document[1])
  

if __name__ == "__main__": 
    main()



