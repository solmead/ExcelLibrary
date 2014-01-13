Imports System.IO
Imports System.Reflection
Imports PocoPropertyData.Extensions

Public Class CSVImporter

    Public Shared Function ImportFromFile(Of tt As Class)(fileInfo As FileInfo, columnMapping As Dictionary(Of String, String), getNewObject As Func(Of tt)) As List(Of tt)
        Dim file = CSVFile.LoadFromFile(fileInfo.FullName)
        Dim item = getNewObject()
        Dim headerLine As CSVFile.CSVLine = Nothing
        Dim headerLineNumber = 0
        Dim propList = item.GetPropertyNames()
        Dim cnt = 0
        Dim fnd = False
        While (headerLineNumber = 0 AndAlso cnt < file.Lines.Count) AndAlso Not fnd
            Dim line = file.Lines(cnt)
            For Each col In line.Columns

                If ((From p In propList Where p.Trim.ToUpper = col.Trim.ToUpper Select p).Any) OrElse columnMapping.ContainsKey(col.Trim()) Then
                    headerLine = line
                    headerLineNumber = cnt
                    fnd = True
                End If
            Next
            cnt += 1
        End While
        Debug.WriteLine("Header Line Number = " & headerLineNumber)
        Dim table = file.ToDataTable(headerLine:=headerLineNumber)

        Return table.ToList(getNewObject, columnMapping)
    End Function

End Class
