Sub regExpr()
    
    Dim currItem, i As Long
    Dim currMail, matches, match As Object
    Dim regEx As New RegExp
    
    With regEx
        .Global = True
        .Pattern = "[A-Z]{3}[0-9]{4}[A-Z0-9]{3}"
    End With
    
    currItem = ActiveExplorer.selection.Count
    
    For i = 1 To currItem
        Set currMail = ActiveExplorer.selection(i)
        If regEx.test(currMail.Subject) Then
            Set matches = regEx.Execute(currMail.Subject)
            For Each match In matches
                CreateObject("Wscript.Shell").Run "https://mesa.mtweb.mouzenidis-travel.ru/ru/ru/book?c=" & match & "#Main"
            Next
        End If
    Next
    
End Sub