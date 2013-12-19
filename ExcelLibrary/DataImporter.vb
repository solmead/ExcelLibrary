Imports System.IO
Imports System.Reflection
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.ComponentModel

Public Class DataImporter
    Public Shared Function FromDataTable(Of tt)(table As DataTable, columnMapping As Dictionary(Of String, String), getNewObject As Func(Of tt)) As List(Of tt)
        Dim mappings = GetDefinedMappings(getNewObject())

        For Each key In columnMapping.Keys
            If (Not mappings.ContainsKey(key)) Then
                mappings.Add(key, columnMapping(key))
            Else
                mappings(key) = columnMapping(key)
            End If
        Next

        Dim itemList = New List(Of tt)()
        For Each row In table.Rows
            Dim item = getNewObject()
            For col = 0 To table.Columns.Count
                Try
                    Dim headCol = table.Columns(col).ColumnName.Trim()
                    Dim column = row(col)
                    If (mappings.ContainsKey(headCol)) Then
                        headCol = mappings(headCol)
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
                        ElseIf (tName.Contains("INT")) Then
                            column = column.Replace("$", "").Replace(",", "")
                            Dim v As Double
                            Double.TryParse(column, v)
                            SetValue(item, headCol, CInt(v))
                        ElseIf (tName.Contains("FLOAT")) Then
                            column = column.Replace("$", "").Replace(",", "")
                            Dim v As Single
                            Single.TryParse(column, v)
                            SetValue(item, headCol, v)
                        ElseIf (tName.Contains("DOUBLE")) Then
                            column = column.Replace("$", "").Replace(",", "")
                            Dim v As Double
                            Double.TryParse(column, v)
                            SetValue(item, headCol, v)
                        ElseIf (tName.Contains("LONG")) Then
                            column = column.Replace("$", "").Replace(",", "")
                            Dim v As Double
                            Double.TryParse(column, CLng(v))
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
                Catch ex As Exception

                End Try
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
    Friend Shared Function GetPropertyNames(item As Object, Optional onlyWritable As Boolean = True, Optional onlyBaseTypes As Boolean = False) As List(Of String)
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
    Private Shared Function GetDefinedMappings(item As Object, Optional onlyWritable As Boolean = True, Optional onlyBaseTypes As Boolean = False) As Dictionary(Of String, String)
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


        Dim lst = (From p In props
                   Where (From ca In p.GetCustomAttributes(True)
                          Where ca.TypeId.FullName.Contains("ColumnAttribute")
                          Select ca).Any() OrElse
                    p.GetCustomAttributes(GetType(DisplayAttribute), True).Any() OrElse
                    p.GetCustomAttributes(GetType(DisplayNameAttribute), True).Any()
                   Select p).ToList()

        'Dim lst = (From p In props
        '           Where p.GetCustomAttributes(GetType(ColumnAttribute), True).Any()
        '           Select p)


        Dim dic As New Dictionary(Of String, String)
        For Each itm In lst
            Dim attrs As IEnumerable(Of Object) = (From ca In itm.GetCustomAttributes(True) Where ca.TypeId.FullName.Contains("ColumnAttribute") Select ca).ToList()
            'Dim attr As ColumnAttribute = itm.GetCustomAttributes(GetType(ColumnAttribute), True).SingleOrDefault()
            For Each attr In attrs
                Try
                    If (attr IsNot Nothing) AndAlso Not String.IsNullOrWhiteSpace(attr.Name) Then
                        dic.Add(attr.Name, itm.Name)
                    End If
                Catch ex As Exception

                End Try
            Next

            Dim attr2 As DisplayAttribute = itm.GetCustomAttributes(GetType(DisplayAttribute), True).FirstOrDefault()
            If (attr2 IsNot Nothing) AndAlso Not dic.ContainsKey(attr2.Name) Then
                dic.Add(attr2.Name, itm.Name)
            End If

            Dim attr3 As DisplayNameAttribute = itm.GetCustomAttributes(GetType(DisplayNameAttribute), True).FirstOrDefault()
            If (attr3 IsNot Nothing) AndAlso Not dic.ContainsKey(attr3.DisplayName) Then
                dic.Add(attr3.DisplayName, itm.Name)
            End If

        Next



        Return dic
    End Function
End Class
