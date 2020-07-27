VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TitleBarPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'        Name: titlebar-path.cls
'      Author: Brian Ferguson
'     Website: https://github.com/brianferguson/excel-addins/titlebar-path/
'        Date: 2020.07.27
'     Version: 1.0
'     License: CC BY-NC-SA 4.0  https://creativecommons.org/licenses/by-nc-sa/4.0/
' Description: This replaces the name in the titlebar with the full path to the file when opened or saved.
'    Platform: Microsoft Office Professional Plus Excel 2013

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub Workbook_Open()
    Set App = Application
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ActiveWindow.Caption = Wb.FullName
End Sub

Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    ActiveWindow.Caption = Wb.FullName
End Sub
