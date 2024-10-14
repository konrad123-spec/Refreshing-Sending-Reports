Attribute VB_Name = "MONTHLY_reports"
' Subroutine to send monthly reports via Outlook
Sub sending_MONTHLY_reports()

    ' Initialize Outlook application and create a new mail item
    Dim OutApp As Object, OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    ' Retrieve the file path from cell N25
    Dim path As String
    path = Range("N25").Value

    ' Define the range of recipients from column A (starting from A2 downwards)
    Dim RList As Range
    Set RList = Range("A2", Range("A2").End(xlDown))

    ' Loop through each recipient in the list
    Dim R As Range
    For Each R In RList
        
        ' Create a new mail item for each recipient
        Set OutMail = OutApp.CreateItem(0)
        
        ' Define the email details (recipient, subject, attachment, and body)
        With OutMail
            .To = R.Offset(0, 0)  ' Email recipient (column A)
            .Subject = R.Offset(0, 2)  ' Email subject (column C)
            .Attachments.Add (path & R.Offset(0, 1))  ' Attach file (column B)
            .Body = R.Offset(0, 3)  ' Email body (column D)
            .Display  ' Display the email for review before sending
        End With

    ' Move to the next recipient in the list
    Next R

End Sub

' Subroutine to refresh monthly reports from SAP and update Excel files
Sub RefreshMonthlyReports()

    ' Variables to store file path and workbook references
    Dim path As String
    Dim wb As Workbook
    Dim sheetName As String
    Dim period As String
    
    ' Prompt the user for the reporting period (e.g., mm.yyyy)
    period = InputBox("Please enter the period:(mm.yyyy)", "Period Input")
    
    ' Retrieve the file path from cell N25
    path = Range("N25").Value

    ' Define the range of workbook names from column G (starting from G2 downwards)
    Dim workbookRange As Range
    Set workbookRange = Range("G2", Range("G2").End(xlDown))

    ' Loop through each workbook listed in column G
    Dim wbName As Range
    For Each wbName In workbookRange
        ' Open the workbook
        Set wb = Workbooks.Open(path & wbName.Value)
        
        ' Refresh the data from SAP and set the period variable
        Refresh = Application.Run("SAPExecuteCommand", "RefreshData")
        a = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_1")

        ' Close the workbook after refreshing, saving the changes
        wb.Close SaveChanges:=True
    Next wbName
    
    ' Additional workbooks that require different treatment are refreshed outside the loop

    ' Open and refresh the "IAMP WIP WD Mensile.xls" workbook
    Workbooks.Open (path & "IAMP WIP WD Mensile.xls")
    Refresh = Application.Run("SAPExecuteCommand", "RefreshData")
    a = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_1")
    b = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_2")
    c = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_3")
    d = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_4")
    e = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_5")
    ActiveWorkbook.Close SaveChanges:=True

    ' Open and refresh the "IAMP ORB SL1C Mensile.xls" workbook
    Workbooks.Open (path & "IAMP ORB SL1C Mensile.xls")
    Refresh = Application.Run("SAPExecuteCommand", "RefreshData")
    a = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_1")
    b = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_2")
    ActiveWorkbook.Close SaveChanges:=True

    ' Open and refresh the "SOX controll 3.1 IAMA.xls" workbook
    Workbooks.Open (path & "SOX controll 3.1 IAMA .xls")
    Refresh = Application.Run("SAPExecuteCommand", "RefreshData")
    a = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_1")
    b = Application.Run("SAPSetVariable", "ZPERCOMP", period, "", "DS_2")
    ActiveWorkbook.Close SaveChanges:=True

    ' Open and refresh the "SD Fatture da emettere IAMP (weekly).xls" workbook
    Workbooks.Open (path & "SD Fatture da emettere IAMP (weekly).xls")
    Refresh = Application.Run("SAPExecuteCommand", "RefreshData")
    ActiveWorkbook.Close SaveChanges:=True

End Sub

