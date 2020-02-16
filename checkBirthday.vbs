Option Explicit
Sub ReadBirthday()
    
    Dim lStart, lStop, i As Long
    Dim currItem As Object
    Const sSearchStart As String = "ï¿½ï¿½ï¿½ï¿½ ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½: "
    Const sSearchStop As String = " <br>"
    Dim birthday, todaysDay, daysCount As Date
    Dim sItemBody, currItemC, d, m, y As String
    todaysDay = Now
    todaysDay = FormatDateTime(todaysDay, vbShortDate)
    daysCount = Date
    
    For i = 1 To ActiveExplorer.selection.Count
        
        Set currItem = ActiveExplorer.selection(i)
         
        If InStr(currItem.Body, "Mouzenidis Travel: ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ ï¿½ï¿½ï¿½ï¿½ï¿½ ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ ") > 0 Then
            sItemBody = currItem.HTMLBody
            
            lStart = InStr(sItemBody, sSearchStart) + Len(sSearchStart)
            If lStart > 0 Then
                lStop = InStr(lStart, sItemBody, sSearchStop)
                If lStop > 0 And lStart > 0 Then
                    currItemC = Mid(sItemBody, lStart, lStop - lStart)
                    d = Left(currItemC, 2)
                    m = Mid(currItemC, 4, 2)
                    y = Right(currItemC, 4)
                    birthday = CDate(d & "/" & m & "/" & y)
                    daysCount = DateDiff("d", birthday, todaysDay)
                    
                    If daysCount < 6570 Then
                        currItem.UnRead = True
                        
                    End If
                End If
            End If
        End If
    Next
    
End Sub