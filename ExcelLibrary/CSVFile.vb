'/*
' * Copyright (C) 2009-2012 Solmead Productions
' *
' * == BEGIN LICENSE ==
' *
' * Licensed under the terms of any of the following licenses at your
' * choice:
' *
' *  - GNU General Public License Version 2 or later (the "GPL")
' *    http://www.gnu.org/licenses/gpl.html
' *
' *  - GNU Lesser General Public License Version 2.1 or later (the "LGPL")
' *    http://www.gnu.org/licenses/lgpl.html
' *
' *  - Mozilla Public License Version 1.1 or later (the "MPL")
' *    http://www.mozilla.org/MPL/MPL-1.1.html
' *
' * == END LICENSE ==
' */

Imports Microsoft.VisualBasic.CompilerServices
Imports CsvHelper.Configuration
Imports CsvHelper
Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Linq.Dynamic
Imports Microsoft.VisualBasic.FileIO.TextFieldParser

Public Class CSVFile
    Public Class CSVLine
        Private m_Columns As New List(Of String)

        Public Function ColumnNumber(value As String) As Integer
            Return Columns.IndexOf((From c In Columns Where c.ToUpper = value.ToUpper Select c).FirstOrDefault)
        End Function


        Public ReadOnly Property Columns() As List(Of String)
            Get
                Return m_Columns
            End Get
        End Property

        Public Sub AddColumn(value As Object)
            Dim s = ""

            If value IsNot Nothing Then
                s = CStr(value)
            End If
            AddColumn(s)
        End Sub
        Public Sub AddColumn(value As String)
            Dim curPos = m_Columns.Count - 1
            Column(curPos + 1) = value
        End Sub


        Public Property Column(ByVal Pos As Integer) As String
            Get
                If Columns.Count <= Pos Then
                    Return ""
                End If
                Return Columns(Pos)
            End Get
            Set(ByVal value As String)
                While Columns.Count <= Pos
                    Columns.Add("")
                End While
                Columns(Pos) = value
            End Set
        End Property
        Public Sub New()

        End Sub
        Public Sub New(ByVal Cols As List(Of String))
            m_Columns = Cols
        End Sub
        Public Function GetCSVLine(ByVal Delimiter As String) As String
            Dim sb As New System.Text.StringBuilder
            Dim s As String
            Dim First As Boolean = True
            For Each s In Columns
                If Not First Then
                    sb.Append(Delimiter)
                Else
                    First = False
                End If
                sb.Append("""" & s & """")
            Next

            Return sb.ToString
        End Function
    End Class
    Public Lines As New List(Of CSVLine)
    Public ColumnDelimiter As String = ","

    Public Sub RemoveColumn(pos As Integer)
        For Each l In Lines
            If l.Columns.Count > pos Then
                l.Columns.RemoveAt(pos)
            End If
        Next
    End Sub


    Public Function GetAsCSV() As String
        Dim sb As New System.Text.StringBuilder
        Dim Line As CSVLine

        For Each Line In Lines
            sb.AppendLine(Line.GetCSVLine(ColumnDelimiter))
        Next
        Return sb.ToString
    End Function


    Public Shared Function LoadFromFileData(ByVal Data As String, Optional ByVal ColDelimiter As String = ",") As CSVFile
        Dim mem As New MemoryStream
        mem.Write(System.Text.Encoding.UTF8.GetBytes(Data), 0, Data.Length)
        mem.Seek(0, SeekOrigin.Begin)
        Dim SR As New StreamReader(mem)
        Return LoadFromFileData(SR, ColDelimiter)
    End Function
    Public Shared Function LoadFromFileData(ByVal Data As MemoryStream, Optional ByVal ColDelimiter As String = ",") As CSVFile
        Dim SR As New StreamReader(Data)
        Return LoadFromFileData(SR, ColDelimiter)
    End Function
    Public Shared Function LoadFromFileData(ByVal ReadFile As Stream, Optional ByVal ColDelimiter As String = ",") As CSVFile
        Dim SR As New StreamReader(ReadFile)
        Return LoadFromFileData(SR, ColDelimiter)
    End Function
    Public Shared Function LoadFromFileData(ByVal ReadFile As StreamReader, Optional ByVal ColDelimiter As String = ",") As CSVFile
        Dim CSVF As New CSVFile
        CSVF.ColumnDelimiter = ColDelimiter

        Dim parser = New CsvParser(ReadFile, New CsvConfiguration With {.Delimiter = ColDelimiter})
        While (True)
            Dim line = parser.Read()

            If (line Is Nothing) Then
                Exit While
            Else
                CSVF.Lines.Add(New CSVLine(line.ToList))
            End If

        End While
        Return CSVF

        'Dim afile As FileIO.TextFieldParser = New FileIO.TextFieldParser(ReadFile)
        'Dim CurrentRecord As String() ' this array will hold each line of data
        'afile.TextFieldType = FileIO.FieldType.Delimited
        'afile.Delimiters = New String() {ColDelimiter}
        'afile.HasFieldsEnclosedInQuotes = True

        '' parse the actual file
        'Do While Not afile.EndOfData
        '    Try
        '        CurrentRecord = afile.ReadFields
        '        CSVF.Lines.Add(New CSVLine(CurrentRecord.ToList))

        '    Catch ex As FileIO.MalformedLineException
        '        Stop
        '    End Try
        'Loop
    End Function

    Public Shared Function LoadFromFileDataOld(ByVal ReadFile As StreamReader, Optional ByVal ColDelimiter As String = ",") As CSVFile
        'Dim FileHolder As FileInfo = New FileInfo(strPath)
        'Dim ReadFile As StreamReader = FileHolder.OpenText()
        Dim strLine As String = "start"
        Dim strLineNext As String = "next"
        Dim CSVF As New CSVFile
        CSVF.ColumnDelimiter = ColDelimiter
        Dim First As Boolean = True
        Dim HasQuote As Boolean = False
        strLine = ReadFile.ReadLine
        If strLine.Substring(0, 1) = """" Then
            HasQuote = True
        End If
        While Not ReadFile.EndOfStream
            strLineNext = ReadFile.ReadLine
            If strLineNext = "" Then strLineNext = " "
            If HasQuote Then
                While Not ReadFile.EndOfStream AndAlso (strLineNext.Substring(0, 1) <> """" OrElse strLineNext.Substring(0, 3) = """,""")
                    strLine = strLine & strLineNext
                    strLineNext = ReadFile.ReadLine
                    If strLineNext = "" Then strLineNext = " "
                End While
                If strLineNext.Substring(0, 1) <> """" Then
                    strLine = strLine & strLineNext
                    strLineNext = ""
                ElseIf strLineNext.Substring(0, 3) = """,""" Then
                    strLine = strLine & strLineNext
                    strLineNext = ""
                End If
            End If
            If Not strLine = "" Then
                Debug.WriteLine(strLine)
                strLine = strLine.Replace(vbCrLf, "<BR/>")
                strLine = strLine.Replace(vbLf, "<BR/>")
                strLine = strLine.Replace(vbCr, "<BR/>")
                CSVF.Lines.Add(New CSVLine(ParseCSVLine(strLine & CSVF.ColumnDelimiter, CSVF.ColumnDelimiter)))
            End If
            strLine = strLineNext
        End While
        If (Not String.IsNullOrEmpty(strLine) AndAlso strLine.Trim <> "") Then
            Debug.WriteLine(strLine)
            strLine = strLine.Replace(vbCrLf, "<BR/>")
            strLine = strLine.Replace(vbLf, "<BR/>")
            strLine = strLine.Replace(vbCr, "<BR/>")
            CSVF.Lines.Add(New CSVLine(ParseCSVLine(strLine & CSVF.ColumnDelimiter, CSVF.ColumnDelimiter)))
        End If
        Return CSVF
    End Function
    Public Shared Function LoadFromFile(ByVal strPath As String, Optional ByVal ColDelimiter As String = ",") As CSVFile
        Dim FileHolder As FileInfo = New FileInfo(strPath)
        Dim ReadFile As StreamReader = FileHolder.OpenText()
        Dim CSVF = LoadFromFileData(ReadFile, ColDelimiter)
        ReadFile.Close()
        ReadFile = Nothing
        Return CSVF
    End Function
    Public Shared Function LoadFromIEnumerable(ByVal list As IEnumerable, Optional ByVal HasHeader As Boolean = True, Optional useDisplayName As Boolean = False) As CSVFile
        Dim olist As List(Of Object) = (From i In list Select i).ToList
        Return LoadFromDataTable(olist.ToDataTable(useDisplayName), HasHeader)
    End Function
    Public Shared Function LoadFromDataTable(ByVal DT As DataTable, Optional ByVal HasHeader As Boolean = True) As CSVFile

        Dim CSVF As New CSVFile
        If HasHeader Then
            Dim L As New CSVLine

            For Each C As DataColumn In DT.Columns
                L.Columns.Add(C.ColumnName)
            Next
            CSVF.Lines.Add(L)
        End If
        For Each R In DT.Rows
            Dim L As New CSVLine

            For Each C As DataColumn In DT.Columns
                L.Columns.Add(R(C.ColumnName).ToString)
            Next
            CSVF.Lines.Add(L)

        Next
        Return CSVF
    End Function



    Private Shared Function ParseCSVLine(ByVal CSVstr As String, Optional ByVal ColDelimiter As String = ",") As List(Of String)

        Dim startPos As Integer
        Dim endPos As Integer
        Dim currPos As Integer
        Dim tempstr As String
        Dim commaPos As Integer
        Dim quotePos As Integer
        Dim strLen As Integer
        Dim charLen As Integer

        Dim a As New List(Of String)

        startPos = 1
        currPos = 1

        strLen = Len(CSVstr)


        Do While strLen <> 0
            'CSVstr = Replace(CSVstr, "," & Chr(34) & ",", ", ,")
            'CSVstr = Replace(CSVstr, ", " & Chr(34) & ",", ", ,")
            'CSVstr = Replace(CSVstr, "," & Chr(34) & " ,", ", ,")
            CSVstr = Replace(CSVstr, "," & Chr(34) & Chr(34), ", ")
            CSVstr = Replace(CSVstr, Chr(34) & Chr(34) & ",", " ,")
            CSVstr = Replace(CSVstr, Chr(34) & Chr(34), "&quot;")
            commaPos = InStr(currPos, CSVstr, ColDelimiter)
            quotePos = InStr(currPos, CSVstr, Chr(34))
            'last data
            If commaPos = 0 Then
                If quotePos = 0 Then
                    If Not currPos > endPos Then
                        endPos = strLen + 1
                        charLen = endPos - currPos
                        tempstr = Mid(CSVstr, currPos, charLen)
                        'If Not tempstr = "" Then
                        a.Add(ReadChars(tempstr, 1, charLen, charLen).ToString().Replace("&quot;", """"))
                        'End If
                    End If
                Else
                    currPos = quotePos
                    endPos = InStr(currPos + 1, CSVstr, Chr(34))
                    charLen = endPos - currPos
                    tempstr = Mid(CSVstr, currPos + 1, charLen - 1)

                    'If Not tempstr = "" Then
                    a.Add(ReadChars(tempstr, 1, charLen, charLen).ToString().Replace("&quot;", """"))
                    'End If
                End If
                Exit Do
            End If
            'no " in line
            If quotePos = 0 Then

                endPos = commaPos
                charLen = endPos - currPos
                tempstr = Mid(CSVstr, currPos, charLen)
                'If Not tempstr = "" Then
                a.Add(ReadChars(tempstr, 1, charLen, charLen).ToString().Replace("&quot;", """"))
                'End If


            ElseIf (quotePos <> 0) Then
                '" in line
                If commaPos < quotePos Then
                    endPos = commaPos
                    charLen = endPos - currPos
                    tempstr = Mid(CSVstr, currPos, charLen)
                    'If Not tempstr = "" Then
                    a.Add(ReadChars(tempstr, 1, charLen, charLen).ToString().Replace("&quot;", """"))
                    'End If
                Else
                    currPos = quotePos
                    endPos = InStr(currPos + 1, CSVstr, Chr(34))
                    If endPos <= 0 Then endPos = CSVstr.Length
                    charLen = endPos - currPos
                    tempstr = Mid(CSVstr, currPos + 1, charLen - 1)

                    'If Not tempstr = "" Then
                    a.Add(ReadChars(tempstr, 1, charLen, charLen).ToString().Replace("&quot;", """"))
                    'End If
                    endPos = endPos + 1
                End If
            End If
            currPos = endPos + 1
        Loop

        Return a

    End Function

    Private Shared Function ReadChars(ByVal str As String, ByVal StartPos As Integer, ByVal EndPos As Integer, ByVal strLen As Integer) As String

        Dim c As Integer
        Dim s As String = ""
        For c = StartPos - 1 To EndPos
            Try
                If str IsNot Nothing AndAlso (c + 1) <= str.Length Then
                    s &= str.Substring(c, 1)
                End If
            Catch exp As Exception
            End Try
        Next
        Return s

        'Dim strArray As String = str
        'Dim b(strLen) As Char
        'Dim sr As New StringReader(strArray)

        'sr.Read(b, 0, EndPos)

        'sr.Close()

        'Dim s As New String(b)
        'Return s.Replace(ControlChars.Cr, "").Replace(ControlChars.Lf, "").Replace(ControlChars.CrLf, "").Replace(ControlChars.NewLine, "")

    End Function
End Class


