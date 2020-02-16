Sub OpenSelection()
    
    Dim matches, match, selection As Object
    Dim selTxt As String
    Dim regEx As New RegExp
    
    With regEx
        .Global = True
        .Pattern = "[A-Z]{3}[0-9]{4}[A-Z0-9]{3}"
    End With
    
    Set selection = ActiveExplorer.selection.Item(1).GetInspector
    selTxt = selection.WordEditor.Application.selection.Range.Text
    
    If regEx.test(selTxt) Then
        Set matches = regEx.Execute(selTxt)
        For Each match In matches
            CreateObject("Wscript.Shell").Run "https://mesa.mtweb.mouzenidis-travel.ru/ru/ru/book?c=" & match & "#Main"
        Next
    End If
    
    selection.Close olSave
    Set matches = Nothing
    
End Sub