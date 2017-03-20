Option Explicit
'The following Code extract the data from mail
'https://social.msdn.microsoft.com/Forums/en-US/22631e7e-53df-47c4-b625-22c9e935f02b/copy-a-table-from-body-of-an-email-to-excel-spreadsheet?forum=outlookdev
'Dazhong Li 11/11/2016
Sub dd()
Dim item As MailItem, x%
Dim r As Object  'As Word.Range
Dim doc As Object 'As Word.Document
Dim xlApp As Object, wkb As Object
Set xlApp = CreateObject("Excel.Application")
Set wkb = xlApp.Workbooks.Add
xlApp.Visible = True

Dim wks As Object
Set wks = wkb.Sheets(1)

For Each item In Application.ActiveExplorer.Selection
Set doc = item.GetInspector.WordEditor
    For x = 1 To doc.tables.Count
     Set r = doc.tables(x)
        r.Range.Copy
       wks.Paste
       wks.Cells(wks.Rows.Count, 1).End(3).Offset(1).Select
    Next
Next
End Sub
