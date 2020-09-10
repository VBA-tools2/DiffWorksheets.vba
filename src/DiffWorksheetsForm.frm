VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DiffWorksheetsForm
   Caption         =   "Diff Worksheets"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5130
   OleObjectBlob   =   "DiffWorksheetsForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "DiffWorksheetsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("DiffWorksheets")

Option Explicit

Public Event SetWks1()
Public Event SetWks2()
Public Event FormStartDiffing()
Public Event FormCancelled(ByRef Cancel As Boolean)

Private Type TView
    Model As DiffWorksheetsModel
    FormOnTop As cFormOnTop
End Type
Private This As TView

Public Property Set Model(ByVal Value As DiffWorksheetsModel)
    
    Set This.Model = Value
    
    With This.Model
        If Not .wks1 Is Nothing Then Set1Labels
        If Not .wks2 Is Nothing Then Set2Labels
    End With
    
    ValidateWksButtons
    ValidateDiffButton
    
End Property

'returns True if cancellation was cancelled by handler
Private Function OnCancel() As Boolean
    Dim cancelCancellation As Boolean
    RaiseEvent FormCancelled(cancelCancellation)
    
    If Not cancelCancellation Then
        Set This.FormOnTop = Nothing
        Me.Hide
    End If
    
    OnCancel = cancelCancellation
End Function

Private Sub UserForm_Initialize()
    With This
        Set .FormOnTop = New cFormOnTop
        Set .FormOnTop.TheUserform = Me
        .FormOnTop.InitializeMe
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then Cancel = Not OnCancel
End Sub

Private Sub cmdDiff_Click()
    Set This.FormOnTop = Nothing
    Me.Hide
    RaiseEvent FormStartDiffing
End Sub

Private Sub cmdSet1Worksheet_Click()
    RaiseEvent SetWks1
    Set1Labels
    ValidateDiffButton
End Sub

Private Sub cmdSet2Worksheet_Click()
    RaiseEvent SetWks2
    Set2Labels
    ValidateDiffButton
End Sub

Public Sub ValidateWksButtons()
    With This.Model
        cmdSet1Worksheet.Enabled = .IsSheetValid
        cmdSet2Worksheet.Enabled = .IsSheetValid
    End With
End Sub

Private Sub ValidateDiffButton()
    cmdDiff.Enabled = This.Model.CanWorksheetsBeDiffed
End Sub

Private Sub Set1Labels()
    With This.Model
        lblWorkbook1.Caption = .wkb1Name
        lblWorkbook1.Enabled = True
        lblWorksheet1.Caption = .wks1Name
        lblWorksheet1.Enabled = True
    End With
End Sub

Private Sub Set2Labels()
    With This.Model
        lblWorkbook2.Caption = .wkb2Name
        lblWorkbook2.Enabled = True
        lblWorksheet2.Caption = .wks2Name
        lblWorksheet2.Enabled = True
    End With
End Sub
