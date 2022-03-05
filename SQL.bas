'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc          Util Class SQL
'@lastUpdate    05.12.2014
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Private pQuery As String
Private pRes As ADODB.Recordset
Private pConnStr As String
Private pSrcWkbPath As String
Private pSrcSht As String
Private pReg As Object
Private pCondition As String
Private pAggregate As String

Public Sub condition(s)
    
    If Trim(pCondition) = "" Then
        s = " WHERE ( " & s & " )"
    Else
        s = pCondition & " AND ( " & s & " )"
    End If
    
    pCondition = s

End Sub


Public Sub aggregate(s)
    
    pAggregate = " " & s

End Sub

Public Function datumRng(field As String, dStart As Date, dEnd As Date) As String
    datumRng = "[" & field & "]" & " BETWEEN " & CDbl(dStart) & " AND " & CDbl(dEnd)
End Function



Public Sub connect(Optional wkbPath As String, Optional targSht As String)
    On Error GoTo ConnectionError
    
    ' only the relative path allowed
    If IsMissing(wkbPath) Or wkbPath = "" Then
        pSrcWkbPath = ActiveWorkbook.Path & "\data\" & "src.xlsx"
    Else
        pSrcWkbPath = ActiveWorkbook.Path & wkbPath
    End If
    
    ' test if the Path correct
    Application.Workbooks.Open pSrcWkbPath
    
    If IsMissing(targSht) Or targSht = "" Then
        targSht = ActiveWorkbook.Worksheets(1).Name
    Else
    ' test if the shtName correct
        ActiveWorkbook.Worksheets(targSht).Activate
    End If
    
    pSrcSht = "[" & targSht & "$]"
    
    ActiveWorkbook.Close False

ConnectionError:
    If Err.Number <> 0 Then
         Debug.Print "Please check the existence of the Target Workbook or Sht : " & pSrcWkbPath & " -> " & pSrcSht
         Debug.Assert Err.Number = 0
    End If
    

End Sub

'''''''''''''''''''''
' @param  s : SQL to execute
' @param  "-" in the query after from will be replaced by pSrcSht
'''''''''''''''''''''

Public Sub query(s As String)
    On Error GoTo SQLError
    
    
    If pReg Is Nothing Then
        Set pReg = CreateObject("Vbscript.regexp")
        
        With pReg
            .pattern = "\bfrom\s*(\-)\s*"
            .Global = True
            .IgnoreCase = True
        End With
    End If
    
    If pReg.test(s) Then
        pQuery = pReg.Replace(s, "FROM " & pSrcSht)
    Else
        pQuery = s
    End If
    
  
    pQuery = pQuery & pCondition & pAggregate

    
    If pRes Is Nothing Then
    
        Set pRes = CreateObject("ADODB.Recordset")

        pConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
               "Data Source=" & pSrcWkbPath & ";" & _
               "Extended Properties=Excel 12.0"
        pRes.Open pQuery, pConnStr
    Else
        pRes.Open pQuery, pConnStr
    End If
    
SQLError:
    If Err.Number <> 0 Then
         Debug.Print "Please Check the Query Instruction :" & pQuery
         Debug.Assert Err.Number = 0
    End If
    
    
   
End Sub

Public Property Get res()
    
    Set res = pRes

End Property

Public Function clear()
    
    pQuery = ""
    pCondition = ""
    pAggregate = ""
    Set pRes = Nothing

End Function


Public Sub write2Sht(targSht As String, topLeftCell As Range)
    Dim tmpname As String
    Dim i As Integer
    
    tmpname = ActiveSheet.Name
    If Trim(targSht) = "" Then
        targSht = tmpname
    End If
    
    If Not pRes.EOF Then
        topLeftCell.Offset(1, 0).CopyFromRecordset res
        
        For i = 0 To res.Fields.Count - 1
         Worksheets(targSht).Cells(topLeftCell.Row, i + topLeftCell.Column).Value = res.Fields(i).Name
        Next i
    
    Else
        Debug.Print "Not Match Results found!"
    End If

End Sub

Public Sub closeCon()
    pRes.Close
End Sub


Public Function thisMonday(ByVal d As Date) As Date

     thisMonday = DateAdd("d", 1 - Weekday(d, vbMonday), d)
    
End Function
