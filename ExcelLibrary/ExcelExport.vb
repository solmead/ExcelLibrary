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
Imports System.IO
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports ClosedXML.Excel


Public Class ExcelExport

    private workbook As XLWorkbook
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
        workbook.Worksheets.Add(name)
    End Sub
    Public Sub AddSheet(ByVal Name As String, ByVal Tb As DataTable)
        workbook.Worksheets.Add(Tb, name)
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
    public sub SaveWorkbook(file As FileInfo)
        workbook.SaveAs(file.FullName)
    End sub
    Public Sub SaveWorkbook(stream As Stream)
        
        'Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        'Response.AddHeader("content-disposition", "attachment;filename=""" & WorkBookname & ".xlsx" & """")
        'Using memoryStream = New MemoryStream()
            workbook.SaveAs(stream)
            'memoryStream.WriteTo(Response.OutputStream)
            'memoryStream.Close()
        'end using
        
        'Response.End()
    End Sub
    Public Sub SaveWorkbook(Response As HttpResponse)
        
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        Response.AddHeader("content-disposition", "attachment;filename=""" & WorkBookname & ".xlsx" & """")
        Using memoryStream = New MemoryStream()
            workbook.SaveAs(memoryStream)
            memoryStream.WriteTo(Response.OutputStream)
            memoryStream.Close()
        end using
        
        Response.End()
    End Sub
    
    Private Sub SetupExcel(ByVal Name As String)
        workbook = new XLWorkbook()
        workbook.Properties.Title = Name
    End Sub
    

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

    Public Shared Sub RespondWith(ByVal WorkBookName As String, ByVal DT As DataTable)
        Dim EE As New ExcelExport(WorkBookName, DT)
        Dim Response As HttpResponse = HttpContext.Current.Response
        EE.SaveWorkbook(Response)
    End Sub
    Public Shared Sub RespondWith(ByVal WorkBookName As String, ByVal Grid As GridView)
        Dim EE As New ExcelExport(WorkBookName, Grid)
        Dim Response As HttpResponse = HttpContext.Current.Response
        EE.SaveWorkbook(Response)
    End Sub
End Class


