Sub Macro1()
    Dim SQL$, MyFile$, s$, str$
    Dim sh As Worksheet
    Dim wb As Workbook
    s = ThisWorkbook.Worksheets(1).[c2]
    str = ThisWorkbook.Worksheets(1).[e2]
    MyFile = Dir(ThisWorkbook.Path & "\*.xls")
    While MyFile <> ""
       
        If MyFile <> ThisWorkbook.Name Then
            Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & MyFile)
            wb.Application.Visible = False
            For Each sh In wb.Worksheets
                sh.UsedRange.Replace What:=s, Replacement:=str, LookAt:=xlWhole, SearchOrder:=xlByRows, ReplaceFormat:=False, _
                    MatchCase:=True, SearchFormat:=False
            Next
            wb.Close True
        End If
        
        MyFile = Dir()
    Wend
    
    Set rs = Nothing
    Set cnn = Nothing
End Sub
