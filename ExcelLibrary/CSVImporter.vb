Imports System.IO
Imports System.Reflection
Imports PocoPropertyData.Extensions

Public Class CSVImporter

    public Shared Function ImportToDataTable(of tt As class)(fileInfo As FileInfo, optional columnMapping As Dictionary(Of String, String) = nothing) As DataTable
        Dim file = CSVFile.LoadFromFile(fileInfo.FullName)
        Dim tp = GetType(TT)
        Dim item = CType(tp.Assembly.CreateInstance(tp.FullName, True), TT)
        Dim headerLine As CSVFile.CSVLine = Nothing
        Dim headerLineNumber = 0
        'Dim propList = item.GetPropertyNames()
        Dim cnt = 0
        Dim fnd = False
        While (headerLineNumber = 0 AndAlso cnt < file.Lines.Count) AndAlso Not fnd
            Dim line = file.Lines(cnt)
            For Each col In line.Columns
                If (item.DoesPropertyExist(col.Trim()) OrElse (columnMapping IsNot nothing AndAlso columnMapping.ContainsKey(col.Trim()))) Then
                    headerLine = line
                    headerLineNumber = cnt
                    fnd = True
                End If
            Next
            cnt += 1
        End While
        Debug.WriteLine("Header Line Number = " & headerLineNumber)
        Dim table = file.ToDataTable(headerLine:=headerLineNumber)


        Return table
    End Function


    Public Shared Function ImportFromFile(Of tt As Class)(fileInfo As FileInfo, columnMapping As Dictionary(Of String, String), getNewObject As Func(Of tt), optional afterLoad As Action(Of tt, DataTable, DataRow, List(Of string)) = nothing) As List(Of tt)
        Dim table = ImportToDataTable(Of tt)(fileInfo, columnMapping)

        Return table.ToList(getNewObject, columnMapping, afterLoad)
    End Function

End Class
