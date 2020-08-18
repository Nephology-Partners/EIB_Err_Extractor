Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles cmd_EIB_ERROR.Click
        Dim xl As Excel.Application
        Dim sh As Excel.Worksheet
        Dim iRow As Integer 'TODO: Should this be long?
        Dim eestring As String
        'Reference to excel
        xl = Globals.ThisAddIn.Application
        xl.Cursor = Excel.XlMousePointer.xlWait
        'Reference first data row
        iRow = 6 'TODO: make option?
        'Reference to worksheet
        sh = xl.Sheets(1)
        If sh.Name = "Overview" Then 'TODO: make option? 
            sh = xl.Sheets(2)
        End If
        'Get spreadsheet key
        eestring = sh.Cells(iRow, 2).value2
        Do Until eestring Is Nothing
            sh.Cells(iRow, 1) = GetNote(sh.Range(sh.Cells(iRow, 1), sh.Cells(iRow, 2)))
            iRow += 1
            eestring = sh.Cells(iRow, 2).value2
        Loop
        'Reset cursor
        xl.Cursor = Excel.XlMousePointer.xlDefault
    End Sub
End Class
