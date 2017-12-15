Imports System.IO
Imports ClosedXML.Excel
Imports PocoPropertyData

Public Class ExcelImporter
    
    public Shared Function ImportToDataTable(of tt As class)(fileInfo As FileInfo, optional columnMapping As Dictionary(Of String, String) = nothing) As DataTable
        Dim file = new XLWorkbook(fileInfo.FullName)
        Dim sheet = file.Worksheets(0)
        Dim tp = GetType(TT)
        Dim item = CType(tp.Assembly.CreateInstance(tp.FullName, True), TT)
        Dim headerLine As IXLRow = Nothing
        Dim headerLineNumber = 0
        'Dim propList = item.GetPropertyNames()
        Dim cnt = 1
        Dim fnd = False
        While (headerLineNumber = 0 AndAlso cnt <= sheet.Rows.Count) AndAlso Not fnd
            Dim line = sheet.Row(cnt) ' file.Lines(cnt)
            For Each col In line.Cells()
                If (item.DoesPropertyExist(col.Value.ToString().Trim()) OrElse (columnMapping IsNot nothing AndAlso columnMapping.ContainsKey(col.Value.ToString().Trim()))) Then
                    headerLine = line
                    headerLineNumber = cnt
                    fnd = True
                End If
            Next
            cnt += 1
        End While
        Debug.WriteLine("Header Line Number = " & headerLineNumber)
        
        Dim table = ToDataTable(sheet, headerLine:=headerLineNumber)


        Return table
    End Function


    Public Shared Function ImportFromFile(Of tt As Class)(fileInfo As FileInfo, columnMapping As Dictionary(Of String, String), getNewObject As Func(Of tt), optional afterLoad As Action(Of tt, DataTable, DataRow, List(Of string)) = nothing) As List(Of tt)
        Dim table = ImportToDataTable(Of tt)(fileInfo, columnMapping)

        Return table.ToList(getNewObject, columnMapping, afterLoad)
    End Function


    Public shared Function ToDataTable(sheet As IXLWorksheet, Optional hasHeader As Boolean = True, Optional headerLine As Integer = 0) As DataTable
        Dim maxcolumns = (From l In sheet.Rows
                              From c in l.Cells()
                              Where Not string.IsNullOrWhiteSpace(c.Value.ToString())
                          Select c.Address.ColumnNumber).Max()
        Dim maxrows = (From l In sheet.Rows
                From c in l.Cells()
                Where Not string.IsNullOrWhiteSpace(c.Value.ToString())
                Select c.Address.RowNumber).Max()
        Dim table = New DataTable()

        For cnt = 1 To maxcolumns
            Dim colName = ""
            If hasHeader AndAlso sheet.Row(headerLine).CellCount()>cnt Then
                colName = sheet.Row(headerLine).Cell(cnt).Value.ToString()
            End If
            If String.IsNullOrEmpty(colName) Then
                colName = "Column_" & cnt
            End If
            table.Columns.Add(colName)
        Next
        Dim srow = 1
        If hasHeader Then
            srow = headerLine + 1
        End If
        For cnt = srow To maxrows
            Dim row = table.NewRow
            For cnum = 1 To maxcolumns
                Dim v = ""
                If (sheet.Row(cnt).CellCount()>cnum) Then
                    v=sheet.Row(cnt).Cell(cnum).Value.ToString()
                End If
                row(cnum-1) = v
            Next
            table.Rows.Add(row)
        Next

        Return table
    End Function
End Class
