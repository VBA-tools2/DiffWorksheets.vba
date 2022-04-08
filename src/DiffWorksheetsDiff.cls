VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffWorksheetsDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("DiffWorksheets")

Option Explicit

Private Const DiffWorksheetName As String = "diff"

Private Enum eDiffWorksheetsDiffError
    [_First] = vbObjectError + 1
    ErrIsNothing = [_First]
    ErrWorksheetHasDiffName
    [_Last] = ErrWorksheetHasDiffName
End Enum

Public Sub DiffWorksheets( _
    ByVal wks1 As Worksheet, _
    ByVal wks2 As Worksheet _
)
    
    If wks1 Is Nothing Then RaiseErrorIsNothing
    If wks2 Is Nothing Then RaiseErrorIsNothing
    
    If HasWorksheetDiffName(wks1) Then RaiseErrorWorksheetHasDiffName
    If HasWorksheetDiffName(wks2) Then RaiseErrorWorksheetHasDiffName
    
    Dim DiffWorksheet As Worksheet
    Set DiffWorksheet = CreateDiffWorksheet(wks1)
    
    Dim NoOfRows As Long
    NoOfRows = GetUsedRangeMinRows(wks1, wks2)
    
    Dim NoOfColumns As Long
    NoOfColumns = GetUsedRangeMinColumns(wks1, wks2)
    
    SetDiffFormula _
            DiffWorksheet, _
            wks1, _
            wks2, _
            NoOfRows, _
            NoOfColumns
    
    DiffWorksheet.Activate
    
End Sub

Private Function HasWorksheetDiffName( _
    ByVal wks As Worksheet _
        ) As Boolean
    HasWorksheetDiffName = (LCase$(wks.Name) = LCase$(DiffWorksheetName))
End Function

Private Function CreateDiffWorksheet( _
    ByVal wks1 As Worksheet _
        ) As Worksheet
    
    Dim wkb As Workbook
    Set wkb = wks1.Parent
    
    Dim wks As Worksheet
    Set wks = GetAlreadyExistingDiffWorksheet(wkb)
    
    If wks Is Nothing Then
        Set wks = wkb.Worksheets.Add(Before:=wks1)
        
        With wks
            .Name = DiffWorksheetName
'            .Name = wks1.Name & "_" & DiffWorksheetName
            .Tab.ThemeColor = xlThemeColorDark1
            .Tab.TintAndShade = -0.499984740745262
            .Columns.ColumnWidth = 6.57
        End With
    Else
        With wks
            .UsedRange.Clear
        End With
    End If
    
    Set CreateDiffWorksheet = wks
    
End Function

Private Function GetAlreadyExistingDiffWorksheet( _
    ByVal wkb As Workbook _
        ) As Worksheet
    On Error Resume Next
    Dim wks As Worksheet
    Set wks = wkb.Worksheets(DiffWorksheetName)
    On Error GoTo 0
    
    Set GetAlreadyExistingDiffWorksheet = wks
End Function

Private Function GetUsedRangeMinRows( _
    ByVal wks1 As Worksheet, _
    ByVal wks2 As Worksheet _
        ) As Long
    
    GetUsedRangeMinRows = Application.WorksheetFunction.Max( _
            wks1.UsedRange.Rows(wks1.UsedRange.Rows.Count).Row, _
            wks2.UsedRange.Rows(wks2.UsedRange.Rows.Count).Row _
    )
    
End Function

Private Function GetUsedRangeMinColumns( _
    ByVal wks1 As Worksheet, _
    ByVal wks2 As Worksheet _
        ) As Long
    
    GetUsedRangeMinColumns = Application.WorksheetFunction.Max( _
            wks1.UsedRange.Columns(wks1.UsedRange.Columns.Count).Column, _
            wks2.UsedRange.Columns(wks2.UsedRange.Columns.Count).Column _
    )
    
End Function

Private Sub SetDiffFormula( _
    ByVal DiffWorksheet As Worksheet, _
    ByVal wks1 As Worksheet, _
    ByVal wks2 As Worksheet, _
    ByVal NoOfRows As Long, _
    ByVal NoOfColumns As Long _
)
    
    Dim wkb1Name As String
    wkb1Name = wks1.Parent.Name
    
    Dim wks1Name As String
    wks1Name = wks1.Name
    
    Dim wkb2Name As String
    wkb2Name = wks2.Parent.Name
    
    Dim wks2Name As String
    wks2Name = wks2.Name
    
    Dim rng1Address As String
    rng1Address = GetRng1Address(wks1Name)
    
    Dim rng2Address As String
    rng2Address = GetRng2Address( _
            wkb1Name, _
            wkb2Name, _
            wks2Name _
    )
    
    Dim rng As Range
    Set rng = DiffWorksheet.Cells(1, 1).Resize(NoOfRows, NoOfColumns)
    
    Dim FormulaString As String
    FormulaString = GetFormulaStringWithoutPrecisionChoice(rng1Address, rng2Address)
'    FormulaString = GetFormulaStringWithPrecisionChoice(rng1Address, rng2Address)
    
    With rng
        .FormulaR1C1 = FormulaString
        .NumberFormat = "0;;"
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -4.99893185216834E-02
            .PatternTintAndShade = 0
        End With
    End With
    
End Sub

Private Function GetRng1Address( _
    ByVal wks1Name As String _
        ) As String
    
    GetRng1Address = "'" & wks1Name & "'!RC"
    
End Function

Private Function GetRng2Address( _
    ByVal wkb1Name As String, _
    ByVal wkb2Name As String, _
    ByVal wks2Name As String _
        ) As String
    
    If wkb1Name <> wkb2Name Then
        Dim PrefixString As String
        PrefixString = "[" & wkb2Name & "]"
    End If
    
    GetRng2Address = "'" & PrefixString & wks2Name & "'!RC"
    
End Function

Private Function GetFormulaStringWithoutPrecisionChoice( _
    ByVal rng1Address As String, _
    ByVal rng2Address As String _
        ) As String
    
    'function description:
    '- if both cells contain errors, compare them
    '- elseif (at least) one cell contains a formula compare them
    '- else compare the cells
    GetFormulaStringWithoutPrecisionChoice = _
        "=IF(AND(ISERROR(" & rng1Address & "),ISERROR(" & rng2Address & "))," & _
        Chr$(10) & _
        "N(ERROR.TYPE(" & rng1Address & ")<>ERROR.TYPE(" & rng2Address & "))," & _
        Chr$(10) & _
        "IF(OR(ISFORMULA(" & rng1Address & "),ISFORMULA(" & rng2Address & "))," & _
            Chr$(10) & Space(4) & _
            "IFERROR(N(NOT(EXACT(FORMULATEXT(" & rng1Address & "),FORMULATEXT(" & rng2Address & ")))),1)," & _
            Chr$(10) & Space(4) & _
            "IFERROR(N(NOT(EXACT(" & rng1Address & "," & rng2Address & "))),1)))"
    
End Function

Private Function GetFormulaStringWithPrecisionChoice( _
    ByVal rng1Address As String, _
    ByVal rng2Address As String _
        ) As String
    
    '==========================================================================
    'add a defined name with this name and give it a value <=14
    '(14 means no precision loss, thus, larger values don't make sense)
    Const DefinedNameOfRoundPrecision As String = "DiffRoundPrecision"
    '==========================================================================
    
    'function description:
    '- '#REF!' errors should remain '#REF!' errors
    '- elseif both cells contain errors compare them
    '- elseif both cells contain numbers compare them
    '    comparison is based on given round precision (defined name: 'DefinedNameOfRoundPrecision') and
    '    order of magnitude of numbers
    '- else compare the cells
    GetFormulaStringWithPrecisionChoice = _
        "=IF(OR(AND(ISERROR(" & rng1Address & "),IFERROR(ERROR.TYPE(" & rng1Address & "),0)=4),AND(ISERROR(" & rng2Address & "),IFERROR(ERROR.TYPE(" & rng2Address & ")=4,0))),#REF!," & _
        Chr$(10) & _
        "N(IF(AND(ISERROR(" & rng1Address & "),ISERROR(" & rng2Address & ")),ERROR.TYPE(" & rng1Address & ")<>ERROR.TYPE(" & rng2Address & ")," & _
        Chr$(10) & _
        "IF(AND(ISNUMBER(" & rng1Address & "),ISNUMBER(" & rng2Address & ")),IF(" & rng1Address & "<=" & rng2Address & "," & _
            Chr$(10) & Space(4) & _
            "ROUND(" & rng1Address & "," & DefinedNameOfRoundPrecision & "-INT(IFERROR(LOG10(ABS(" & rng1Address & ")),0)))<>ROUND(" & rng2Address & "," & DefinedNameOfRoundPrecision & "-INT(IFERROR(LOG10(ABS(" & rng1Address & ")),0)))," & _
            Chr$(10) & Space(4) & _
            "ROUND(" & rng1Address & "," & DefinedNameOfRoundPrecision & "-INT(IFERROR(LOG10(ABS(" & rng2Address & ")),0)))<>ROUND(" & rng2Address & "," & DefinedNameOfRoundPrecision & "-INT(IFERROR(LOG10(ABS(" & rng2Address & ")),0))))," & _
        Chr$(10) & _
        "IFERROR(NOT(EXACT(" & rng1Address & "," & rng2Address & ")),1)))))"
    
End Function

'==============================================================================
Private Sub RaiseErrorIsNothing()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eDiffWorksheetsDiffError.ErrIsNothing, _
            Description:= _
                    "'wks1' or 'wks2' is 'Nothing'." & _
                    vbCrLf & _
                    "Please check beforehand for that."
End Sub

Private Sub RaiseErrorWorksheetHasDiffName()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eDiffWorksheetsDiffError.ErrWorksheetHasDiffName, _
            Description:= _
                    "One of the worksheets has the name: '" & DiffWorksheetName & "'." & _
                    vbCrLf & _
                    "This is not allowed/supported."
End Sub
