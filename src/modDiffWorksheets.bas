Attribute VB_Name = "modDiffWorksheets"

'@Folder("DiffWorksheets")

Option Explicit

Private Presenter As DiffWorksheetsPresenter

Public Sub DiffWorksheets()
    Set Presenter = New DiffWorksheetsPresenter
    Presenter.Show
End Sub
