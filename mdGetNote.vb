Module mdGetNote
    Public Function GetNote(a As Excel.Range) As String
        Dim c As Excel.Range
        GetNote = ""
        For Each c In a
            If c.Comment IsNot Nothing Then
                GetNote = c.Comment.Text
            Else
                GetNote &= ""
            End If
        Next c
    End Function
End Module
