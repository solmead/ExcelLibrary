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

Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Data
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls



Public Class ExcelExport
    Private Enum CellType
        None
        [String]

    End Enum
    Private TheWorkBook As New XmlDocument

    Private CurrentTable As XmlElement
    Public WorkBookname As String = ""

    Public Sub New(ByVal WorkBookname As String)
        Me.WorkBookname = Replace(WorkBookname, " ", "_")
        SetupExcel(Me.WorkBookname)
    End Sub
    Public Sub New(ByVal WorkBookName As String, ByVal DT As DataTable)
        Me.WorkBookname = Replace(WorkBookName, " ", "_")
        SetupExcel(Me.WorkBookname)
        AddSheet(Me.WorkBookname, DT)
    End Sub
    Public Sub New(ByVal WorkBookName As String, ByVal Grid As GridView)
        Me.WorkBookname = Replace(WorkBookName, " ", "_")
        SetupExcel(Me.WorkBookname)
        AddSheet(Me.WorkBookname, Grid)
    End Sub
    Public Sub AddSheet(ByVal Name As String)
        Dim Melem As XmlElement
        Melem = TheWorkBook.CreateElement("Worksheet")
        TheWorkBook.DocumentElement.AppendChild(Melem)
        Melem.SetAttribute("Name", "urn:schemas-microsoft-com:office:spreadsheet", Name)
        CurrentTable = TheWorkBook.CreateElement("Table")
        Melem.AppendChild(CurrentTable)
    End Sub
    Public Sub AddSheet(ByVal Name As String, ByVal Tb As DataTable)
        AddSheet(Name)
        AddSheetHeader(Tb)
        Dim R As DataRow
        For Each R In Tb.Rows
            AddRow(R)
        Next
    End Sub
    Public Sub AddSheet(ByVal Name As String, ByVal Grid As GridView)
        Grid.DataBind()
        Dim DT As New Data.DataTable
        Dim GC As DataControlField
        For Each GC In Grid.Columns
            DT.Columns.Add(New Data.DataColumn(GC.HeaderText.Replace(" ", "_")))
        Next
        Dim GR As GridViewRow
        For Each GR In Grid.Rows
            Dim DR As Data.DataRow = DT.NewRow
            Dim i As Integer
            For i = 0 To Grid.Columns.Count - 1
                'Dim TW As New System.IO.StringWriter()
                'Dim HTW As New HtmlTextWriter(TW)

                Dim GC2 As DataControlFieldCell
                GC2 = GR.Cells(i)
                Dim tstr As String = ""
                tstr = GC2.Text
                tstr = tstr & TreeControl(GC2)
                tstr = RemoveBetween(tstr, "<", ">", False)
                DR(i) = tstr.Trim

            Next
            DT.Rows.Add(DR)
        Next
        AddSheet(Name, DT)
    End Sub
    Public Sub ThrowResponse()
        Dim Response As HttpResponse = HttpContext.Current.Response
        Response.Clear()
        Response.ContentType = "application/vnd.ms-excel"

        Response.AppendHeader("Content-Disposition", "attachment; filename=" & WorkBookname & ".xls")
        Response.BinaryWrite(System.Text.Encoding.UTF8.GetBytes(GetExcelData))
        Response.End()
    End Sub
    Public Function GetExcelData() As String
        Dim tstr As String = ""
        tstr = "<?xml version=""1.0""?>"
        tstr = tstr & "<?mso-application progid=""Excel.Sheet""?>"
        Return tstr & TheWorkBook.DocumentElement.OuterXml
    End Function
    Private Function DataElem(ByVal Name As String, ByVal Value As String) As XmlElement
        Dim Melem As XmlElement = TheWorkBook.CreateElement(Name)
        Melem.InnerText = Value
        Return Melem
    End Function
    Private Sub SetupExcel(ByVal Name As String)
        '        <?xml version="1.0"?>
        '<?mso-application progid="Excel.Sheet"?>
        Dim Melem As XmlElement
        Melem = TheWorkBook.CreateElement("Workbook")
        Melem.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:spreadsheet")
        Melem.SetAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office")
        Melem.SetAttribute("xmlns:x", "urn:schemas-microsoft-com:office:excel")
        Melem.SetAttribute("xmlns:ss", "urn:schemas-microsoft-com:office:spreadsheet")
        Melem.SetAttribute("xmlns:html", "http://www.w3.org/TR/REC-html40")
        TheWorkBook.AppendChild(Melem)

        Dim Melem2 As XmlElement
        Melem2 = TheWorkBook.CreateElement("DocumentProperties")
        Melem.AppendChild(Melem2)

        Melem2.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:office")
        Melem2.AppendChild(DataElem("Author", ""))
        Melem2.AppendChild(DataElem("LastAuthor", ""))
        Melem2.AppendChild(DataElem("Created", Now))
        Melem2.AppendChild(DataElem("LastSaved", Now))
        Melem2.AppendChild(DataElem("Company", ""))
        Melem2.AppendChild(DataElem("Version", "12.00"))


        Melem2 = TheWorkBook.CreateElement("Styles")
        Melem.AppendChild(Melem2)


        Dim Melem3 As XmlElement
        Dim Melem4 As XmlElement
        Dim Melem5 As XmlElement
        Melem3 = TheWorkBook.CreateElement("Style")
        Melem2.AppendChild(Melem3)
        Melem3.SetAttribute("ID", "urn:schemas-microsoft-com:office:spreadsheet", "Default")
        Melem3.SetAttribute("Name", "urn:schemas-microsoft-com:office:spreadsheet", "Normal")

        Melem4 = TheWorkBook.CreateElement("Alignment")
        Melem3.AppendChild(Melem4)
        Melem4.SetAttribute("Vertical", "urn:schemas-microsoft-com:office:spreadsheet", "Bottom")
        Melem3.AppendChild(DataElem("Borders", ""))
        Melem3.AppendChild(DataElem("Font", ""))
        Melem3.AppendChild(DataElem("Interior", ""))
        Melem3.AppendChild(DataElem("NumberFormat", ""))
        Melem3.AppendChild(DataElem("Protection", ""))

        Melem3 = TheWorkBook.CreateElement("Style")
        Melem2.AppendChild(Melem3)
        Melem3.SetAttribute("ID", "urn:schemas-microsoft-com:office:spreadsheet", "header")

        Melem4 = TheWorkBook.CreateElement("Alignment")
        Melem3.AppendChild(Melem4)
        Melem4.SetAttribute("Horizontal", "urn:schemas-microsoft-com:office:spreadsheet", "Center")
        Melem4.SetAttribute("Vertical", "urn:schemas-microsoft-com:office:spreadsheet", "Bottom")

        Melem4 = TheWorkBook.CreateElement("Borders")
        Melem3.AppendChild(Melem4)

        Melem5 = TheWorkBook.CreateElement("Border")
        Melem4.AppendChild(Melem5)
        Melem5.SetAttribute("Position", "urn:schemas-microsoft-com:office:spreadsheet", "Bottom")
        Melem5.SetAttribute("LineStyle", "urn:schemas-microsoft-com:office:spreadsheet", "Continuous")
        Melem5.SetAttribute("Weight", "urn:schemas-microsoft-com:office:spreadsheet", "2")

        Melem4 = TheWorkBook.CreateElement("Font")
        Melem3.AppendChild(Melem4)
        Melem4.SetAttribute("Family", "urn:schemas-microsoft-com:office:excel", "Swiss")
        Melem4.SetAttribute("Bold", "urn:schemas-microsoft-com:office:spreadsheet", "1")
        Melem4 = TheWorkBook.CreateElement("Interior")
        Melem3.AppendChild(Melem4)
        Melem4.SetAttribute("Color", "urn:schemas-microsoft-com:office:spreadsheet", "#99CCFF")
        Melem4.SetAttribute("Pattern", "urn:schemas-microsoft-com:office:spreadsheet", "Solid")

    End Sub
    Private Function GetHeaderCell(ByVal Value As String) As XmlElement
        Dim Cell As XmlElement
        Dim Data As XmlElement

        Cell = TheWorkBook.CreateElement("Cell")
        Cell.SetAttribute("StyleID", "urn:schemas-microsoft-com:office:spreadsheet", "header")
        Data = DataElem("Data", Value)
        Data.SetAttribute("Type", "urn:schemas-microsoft-com:office:spreadsheet", "String")
        Cell.AppendChild(Data)
        Return Cell
    End Function
    Private Function GetCell(ByVal Value As String, ByVal Type As CellType) As XmlElement
        Dim Cell As XmlElement
        Dim Data As XmlElement

        Cell = TheWorkBook.CreateElement("Cell")
        Data = DataElem("Data", Value)
        If Type <> CellType.None Then
            Data.SetAttribute("Type", "urn:schemas-microsoft-com:office:spreadsheet", Type.ToString)
        Else
            Data.SetAttribute("Type", "urn:schemas-microsoft-com:office:spreadsheet", "")
        End If
        Cell.AppendChild(Data)
        Return Cell
    End Function

    Public Sub AddSheetHeader(ByVal DT As DataTable)
        Dim Row As XmlElement = TheWorkBook.CreateElement("Row")
        CurrentTable.AppendChild(Row)

        Dim DF As DataColumn
        For Each DF In DT.Columns
            Row.AppendChild(GetHeaderCell(DF.ColumnName))
        Next
    End Sub
    Public Sub AddRow(ByVal DR As DataRow)
        Dim Row As XmlElement = TheWorkBook.CreateElement("Row")
        CurrentTable.AppendChild(Row)
        Dim DC As DataColumn
        For Each DC In DR.Table.Columns
            Dim Type As CellType
            Type = GetCellType(DC.DataType)
            If DR.IsNull(DC.ColumnName) Then
                Row.AppendChild(GetCell("", Type))
            Else
                Dim s As String = ""
                Try
                    s = DR(DC.ColumnName).ToString
                Catch ex As Exception

                End Try
                Row.AppendChild(GetCell(s, Type))
            End If
        Next
    End Sub
    Private Function GetCellType(ByVal Type As System.Type) As CellType



        Return CellType.String
    End Function

    Private Shared Function TreeControl(ByVal Con As Control) As String
        Dim Tstr As String = ""
        Dim c As Object
        For Each c In Con.Controls
            Try
                If c.visible Then
                    Tstr = Tstr & c.text
                End If
            Catch ex As Exception

            End Try
            'Try
            '    Tstr = Tstr & c.value
            'Catch ex As Exception

            'End Try
            Try
                Tstr = Tstr & TreeControl(c)
            Catch ex As Exception

            End Try
        Next
        Return Tstr
    End Function
    Private Shared Function RemoveBetween(ByVal TheContent As String, ByVal BeginTagName As String, ByVal EndTagName As String, ByVal TruncateContent As Boolean) As String
        TheContent = Replace(Replace(TheContent, "<p>", vbCrLf), "</p>", vbCrLf)
        Dim st : st = 1
        Dim cnt : cnt = 1
        Do While InStr(st, TheContent, BeginTagName) > 0
            Dim BeginComment : BeginComment = InStr(st, TheContent, BeginTagName)
            Dim EndComment : EndComment = InStr(BeginComment + Len(BeginTagName) + 1, TheContent, EndTagName) + Len(EndTagName)
            Dim CheckOtherBegin : CheckOtherBegin = InStr(BeginComment + Len(BeginTagName) + 1, TheContent, "<")
            If CheckOtherBegin = 0 Or (CheckOtherBegin = EndComment - Len(EndTagName)) Or CheckOtherBegin >= EndComment Then
                TheContent = Mid(TheContent, 1, BeginComment - 1) & Mid(TheContent, EndComment)
            Else
                st = CheckOtherBegin
            End If
            cnt = cnt + 1
            If cnt > 20000 Then Exit Do
        Loop
        If TheContent IsNot Nothing AndAlso TruncateContent AndAlso TheContent.Length > 45 Then
            TheContent = TheContent.Substring(0, 45) & "..."
        End If
        'TheContent = Replace(TheContent, vbCrLf, "<br/>")
        Return TheContent
    End Function

    Public Shared Sub ThrowResponse(ByVal WorkBookName As String, ByVal DT As DataTable)
        Dim EE As New ExcelExport(WorkBookName, DT)
        Dim Response As HttpResponse = HttpContext.Current.Response
        Response.Clear()
        Response.ContentType = "application/vnd.ms-excel"
        Response.AppendHeader("Content-Disposition", "attachment; filename=" & EE.WorkBookname & ".xls")
        Response.BinaryWrite(System.Text.Encoding.UTF8.GetBytes(EE.GetExcelData))
        Response.End()
    End Sub
    Public Shared Sub ThrowResponse(ByVal WorkBookName As String, ByVal Grid As GridView)
        Dim EE As New ExcelExport(WorkBookName, Grid)
        Dim Response As HttpResponse = HttpContext.Current.Response
        Response.Clear()
        Response.ContentType = "application/vnd.ms-excel"
        Response.AppendHeader("Content-Disposition", "attachment; filename=" & EE.WorkBookname & ".xls")
        Response.BinaryWrite(System.Text.Encoding.UTF8.GetBytes(EE.GetExcelData))
        Response.End()
    End Sub
End Class


