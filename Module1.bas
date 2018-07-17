Attribute VB_Name = "Module1"
Sub Merger()

    Call date2check
    
    ActiveDocument.MailMerge.OpenDataSource Name:= _
        ThisDocument.Path & "\time.xls", ConfirmConversions:=False, _
        ReadOnly:=False, LinkToSource:=True, AddToRecentFiles:=False, _
        PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
        WritePasswordTemplate:="", Revert:=False, Format:=wdOpenFormatAuto, _
        Connection:= _
        "DSN=Excel Files;DBQ=F:\Desktop\bt morning\time.xls;DriverId=1046;MaxBufferSize=2048;PageTimeout=5;" _
        , SQLStatement:="SELECT * FROM `Sheet1$`", SQLStatement1:="", SubType:= _
        wdMergeSubTypeOther
    ActiveDocument.MailMerge.ViewMailMergeFieldCodes = wdToggle
    
End Sub

Public Sub date2check()

Dim objExcel
Dim objDoc
Dim objSelection



Set objExcel = CreateObject("Excel.Application")
Set objDoc = objExcel.Workbooks.Add

objExcel.Visible = True
objExcel.Application.DisplayAlerts = False

Set objSelection = objExcel.Selection

objExcel.sheets("Sheet1").Activate


    objExcel.sheets("Sheet1").Cells(1, 1).Value = "CurentDate"
    objExcel.sheets("Sheet1").Cells(1, 2).Value = "Minus2BussinessDays"
    objExcel.sheets("Sheet1").Cells(2, 1).Value = "=Today()"
    objExcel.sheets("Sheet1").Cells(2, 1).NumberFormat = "dd/MM/yyyy"
    
    objExcel.sheets("Sheet1").Cells(2, 2).Value = "=Workday(Today(),-2)"
    objExcel.sheets("Sheet1").Cells(2, 2).NumberFormat = "dd/MM/yyyy"


objExcel.Activeworkbook.SaveAs FileName:=ThisDocument.Path & "\time.xls"
objExcel.Activeworkbook.Close True

objExcel.Application.DisplayAlerts = True

End Sub

Sub CreateNewMDBFile()

   Dim ws As Workspace
   Dim db As Database
   Dim LFilename As String

   'Get default Workspace
   Set ws = DBEngine.Workspaces(0)

   'Path and file name for new mdb file
   LFilename = ThisDocument.Path & "\time.mdb"

   'Make sure there isn't already a file with the name of the new database
   If Dir(LFilename) <> "" Then Kill LFilename

   db.Close
   Set db = Nothing

End Sub
