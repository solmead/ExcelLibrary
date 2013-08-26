Imports System.IO
Imports System.Reflection

Public Class CSVImporter

    Public Shared Function ImportFromFile(Of tt)(fileInfo As FileInfo, columnMapping As Dictionary(Of String, String), getNewObject As Func(Of tt)) As List(Of tt)
        Dim file = CSVFile.LoadFromFile(fileInfo.FullName)
        Dim item = getNewObject()
        Dim headerLine As CSVFile.CSVLine = Nothing
        Dim headerLineNumber = 0
        Dim propList = GetPropertyNames(item)
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
        Dim itemList = New List(Of tt)()
        For row = headerLineNumber + 1 To file.Lines.Count - 1
            item = getNewObject()
            Dim line = file.Lines(row)
            For col = 0 To line.Columns.Count
                Dim headCol = headerLine.Column(col).Trim()
                Dim column = line.Column(col)
                If (columnMapping.ContainsKey(headCol)) Then
                    headCol = columnMapping(headCol)
                End If
                If (DoesPropertyExist(item, headCol)) Then
                    Dim tpe = GetPropertyType(item, headCol)
                    Dim tName = tpe.FullName.ToUpper()
                    If (tName.Contains("DATETIME")) Then
                        Dim v As Date
                        DateTime.TryParse(column, v)
                        SetValue(item, headCol, v)
                    ElseIf (tName.Contains("BOOL")) Then
                        column = column.ToUpper().Replace("YES", "TRUE").Replace("NO", "FALSE").Replace("0", "FALSE").
                                Replace("1", "TRUE")
                        Dim v As Boolean
                        Boolean.TryParse(column, v)
                        SetValue(item, headCol, v)
                    ElseIf (tName.Contains("INT") OrElse tName.Contains("FLOAT") OrElse tName.Contains("DOUBLE") OrElse tName.Contains("LONG")) Then
                        column = column.Replace("$", "").Replace(",", "")
                        Dim v As Double
                        Double.TryParse(column, v)
                        SetValue(item, headCol, v)
                    ElseIf (tName.Contains("DECIMAL")) Then
                        column = column.Replace("$", "").Replace(",", "")
                        Dim v As Decimal
                        Decimal.TryParse(column, v)
                        SetValue(item, headCol, v)
                    Else
                        Dim v = Convert.ChangeType(column, tpe)
                        SetValue(item, headCol, v)
                    End If

                ElseIf (headCol <> "") Then
                    'Throw New Exception("Column Not Handled: [" + headCol + "]")
                End If
            Next

            itemList.Add(item)
        Next
        Return itemList


    End Function
    Private Shared Function GetPropertyType(item As Object, propertyName As String) As Type
        Dim tp As Type = item.GetType
        Dim prop = tp.GetProperty(propertyName)

        If (prop IsNot Nothing) Then
            Return prop.PropertyType
        End If
        Return GetType(String)
    End Function
    Private Shared Sub SetValue(item As Object, propertyName As String, value As Object)

        Dim tp As Type = item.GetType
        Dim prop = tp.GetProperty(propertyName)

        If (prop IsNot Nothing) Then
            prop.SetValue(item, value, Nothing)
        End If
    End Sub
    Private Shared Function GetValue(item As Object, propertyName As String) As Object
        Dim retVal As Object = Nothing
        Dim tp As Type = item.GetType
        Dim prop = tp.GetProperty(propertyName)
        If (prop IsNot Nothing) Then
            retVal = prop.GetValue(item, Nothing)
        End If

        Return retVal
    End Function
    Private Shared Function DoesPropertyExist(item As Object, propertyName As String) As Boolean
        ' Dim retVal As Object = Nothing
        Dim tp As Type = item.GetType
        Dim prop = tp.GetProperty(propertyName)
        Return (prop IsNot Nothing)
    End Function
    Private Shared Function GetPropertyNames(item As Object, Optional onlyWritable As Boolean = True, Optional onlyBaseTypes As Boolean = False) As List(Of String)
        Dim tp As Type = item.GetType
        Dim props = tp.GetProperties((BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.FlattenHierarchy)).ToList

        If onlyWritable Then
            props = (From p In props Where p.CanWrite Select p).ToList
        End If
        If onlyBaseTypes Then
            Try
                props = (From p In props
                         Where Not (p.PropertyType.FullName.Contains("Record") OrElse
                         p.PropertyType.FullName.Contains("Set") OrElse
                         p.PropertyType.FullName.Contains("EntitySet") OrElse
                         (p.PropertyType.BaseType IsNot Nothing AndAlso
                          (p.PropertyType.BaseType.FullName.Contains("Record") OrElse
                         p.PropertyType.BaseType.FullName.Contains("Set") OrElse
                         p.PropertyType.BaseType.FullName.Contains("EntitySet"))))).ToList
            Catch ex As Exception

            End Try
        End If

        Return (From p In props Select p.Name).ToList
    End Function
End Class
