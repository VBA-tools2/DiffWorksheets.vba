VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffWorksheetsPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("DiffWorksheets")

'Use this class with e.g.
'------------------------------------------------------------------------------
'Private Presenter As DiffWorksheetsPresenter
'
'Public Sub DiffWorksheets()
'    Set Presenter = New DiffWorksheetsPresenter
'    Presenter.Show
'End Sub
'------------------------------------------------------------------------------
'in a module

Option Explicit

Private WithEvents AppEvents As DiffWorksheetsAppEvents
Attribute AppEvents.VB_VarHelpID = -1
Private WithEvents View As DiffWorksheetsForm
Attribute View.VB_VarHelpID = -1

Private Model As DiffWorksheetsModel
Private Diff As DiffWorksheetsDiff

Private Sub Class_Initialize()
    Set AppEvents = New DiffWorksheetsAppEvents
    Set Model = New DiffWorksheetsModel
    AppEvents_UpdateModel ActiveSheet
End Sub

Private Sub AppEvents_UpdateModel(ByVal sht As Object)
    Set Model.sht = sht
    If Not View Is Nothing Then View.ValidateWksButtons
End Sub

'---
'used from <https://stackoverflow.com/a/47358692/5776000>
Private Sub View_FormCancelled(ByRef Cancel As Boolean)
'BUG: commented these lines because 'View' is closed regardless of button pressed
'    'setting Cancel to True will leave the form open
'    Cancel = MsgBox("Cancel this operation?", vbYesNo + vbExclamation) = vbNo
'    If Not Cancel Then
'        ' modeless form was cancelled and is now hidden.
'        ' ...
        Set View = Nothing
'    End If
End Sub

Public Sub Show()
'BUG: uncommenting results in an runtime error. (Did it work for @MathieuGuindon?)
    'If Not View Is Nothing Then
    '   View.Visible = True 'just to ensure visibility & honor the .Show call
    '   Exit Sub
    'End If
    
    With Model
        If .IsSheetValid Then
            View_SetWks1
            
            Dim wks As Worksheet
            Set wks = .WorksheetWithSameNameInOtherWorkbook(.wks1)
'            Set wks = .NextWorksheet(.wks1)
            
            If Not wks Is Nothing Then
                Set .sht = wks
                View_SetWks2
                Set .sht = .wks1
            End If
        End If
    End With
    
    Set View = New DiffWorksheetsForm
    With View
        Set .Model = Model
        .Show vbModeless
    End With
End Sub
'---

Private Sub View_SetWks1()
    With Model
        If .IsSheetValid Then Set .wks1 = .sht
    End With
End Sub

Private Sub View_SetWks2()
    With Model
        If .IsSheetValid Then Set .wks2 = .sht
    End With
End Sub

Private Sub View_FormStartDiffing()
    Set Diff = New DiffWorksheetsDiff
    With Model
        Diff.DiffWorksheets .wks1, .wks2
    End With
    Set View = Nothing
End Sub
