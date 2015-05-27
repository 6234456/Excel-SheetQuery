Sub main()
    
   Dim q As New sql
   
   ' connect to the excel sheet with location "\data\src.xlsx" by default
   ' different path could also be specified
   q.connect
   
  ' to query all the payments to DHL due in a week
  ' the usage of query, condition and aggregate is similar to the Hibernate Framework
   q.aggregate "GROUP BY [Alpha-Matchcode] ORDER BY SUM([Buchungsbetrag])"
   q.condition q.datumRng("Fälligkeit Verkaufserlös", Date, DateAdd("d", 7, Date))
   q.condition "[Alpha-Matchcode] LIKE '%DHL%'"
  
   q.query "SELECT [Alpha-Matchcode] as Liferant, SUM([Buchungsbetrag]) as Summe FROM -"
  
   
   ' write the result set to the first sheet
   ' if the set is empty, warning will be printed at the console
   Worksheets(1).Cells.clear
   
   q.write2Sht Worksheets(1).Name, Cells(1, 1)

   ' close connection
   q.closeCon
    

End Sub
