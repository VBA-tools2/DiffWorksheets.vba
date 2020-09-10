VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffWorksheetsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("DiffWorksheets")

Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Public Event UpdateModel(ByVal sht As Object)

Private Sub Class_Initialize()
    Set App = Application
End Sub

Private Sub App_WorkbookActivate(ByVal wkb As Workbook)
    RaiseEvent UpdateModel(wkb.ActiveSheet)
End Sub

Private Sub App_SheetActivate(ByVal sht As Object)
    RaiseEvent UpdateModel(sht)
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub
