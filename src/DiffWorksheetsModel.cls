VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffWorksheetsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("DiffWorksheets")

Option Explicit

Private Type TModel
    sht As Object
    
    wks1 As Worksheet
    wks2 As Worksheet
    
    wkb1Name As String
    wks1Name As String
    wkb2Name As String
    wks2Name As String
End Type
Private This As TModel

Public Property Get sht() As Object
    Set sht = This.sht
End Property

Public Property Set sht(ByVal Value As Object)
    Set This.sht = Value
End Property

Public Property Get wks1() As Worksheet
    Set wks1 = This.wks1
End Property

Public Property Set wks1(ByVal Value As Worksheet)
    Set This.wks1 = Value
    AdaptNamesRelatedToWks1
End Property

Public Property Get wks2() As Worksheet
    Set wks2 = This.wks2
End Property

Public Property Set wks2(ByVal Value As Worksheet)
    Set This.wks2 = Value
    AdaptNamesRelatedToWks2
End Property

Public Property Get wkb1Name() As String
    wkb1Name = This.wkb1Name
End Property

Public Property Get wks1Name() As String
    wks1Name = This.wks1Name
End Property

Public Property Get wkb2Name() As String
    wkb2Name = This.wkb2Name
End Property

Public Property Get wks2Name() As String
    wks2Name = This.wks2Name
End Property

Public Function IsSheetValid() As Boolean
    IsSheetValid = (sht.Type = xlWorksheet)
End Function

Public Function WorksheetWithSameNameInOtherWorkbook( _
    ByVal wks As Worksheet _
        ) As Worksheet
    
    Set WorksheetWithSameNameInOtherWorkbook = Nothing
    
    Dim wkb As Workbook
    Set wkb = wks.Parent
    
    Dim wkbName As String
    wkbName = wkb.Name
    
    Dim wksName As String
    wksName = wks.Name
    
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name <> wkbName Then
            Dim ws As Worksheet
            For Each ws In wb.Worksheets
                If ws.Name = wksName Then
                    Set WorksheetWithSameNameInOtherWorkbook = ws
                    Exit Function
                End If
            Next
        End If
    Next
    
End Function

Public Function NextWorksheet( _
    ByVal wks As Worksheet _
        ) As Worksheet
    
    Set NextWorksheet = Nothing
    
    Dim wkb As Workbook
    Set wkb = wks.Parent
    
    On Error Resume Next
    Set NextWorksheet = wkb.Worksheets(wks.Index + 1)
    On Error GoTo 0
    
End Function

Public Function CanWorksheetsBeDiffed() As Boolean
    With This
        If .wks2 Is Nothing Then
            CanWorksheetsBeDiffed = False
        ElseIf .wks1 Is Nothing Then
            CanWorksheetsBeDiffed = False
        ElseIf .wks1Name <> .wks2Name Then
            CanWorksheetsBeDiffed = True
        ElseIf .wkb1Name <> .wkb2Name Then
            CanWorksheetsBeDiffed = True
        Else
            CanWorksheetsBeDiffed = False
        End If
    End With
End Function

Private Sub AdaptNamesRelatedToWks1()
    With This
        .wkb1Name = .wks1.Parent.Name
        .wks1Name = .wks1.Name
    End With
End Sub

Private Sub AdaptNamesRelatedToWks2()
    With This
        .wkb2Name = .wks2.Parent.Name
        .wks2Name = .wks2.Name
    End With
End Sub
