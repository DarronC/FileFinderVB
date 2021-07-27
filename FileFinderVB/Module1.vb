Imports System.Data.OleDb
Imports Microsoft.Office.Interop
'This module uses oleDB to query Windows Index Folder location.
'Network folders cannot by indexed by Windows
'Current build reads and writes to an excel sheet [optional]

Module Module1
    Dim XLApp As Microsoft.Office.Interop.Excel.Application
    Dim XLBook As Excel._Workbook
    Dim XLRow As Integer
    Dim query1 As String
    Dim searchname As String
    Dim query2 As String
    Dim query3 As String
    Dim query4 As String
    Dim filepath As String = "file:C:\Users\DCharles\Documents\2023"
    Dim XLpath As String = "G:\Users\SHARED\ARC_Attachments\2022-23 ARC\AROW_No_Images_7-26-21.xlsx"

    Sub Main()
        XLApp = CreateObject("Excel.Application") 'Create the excel Application object
        If XLApp Is Nothing Then
            MsgBox("Incomplete MS Excel installation on the client computer. Cannot continue.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        XLBook = XLApp.Workbooks.Open(XLpath)
        XLApp.Workbooks.Application.Visible = True

        With XLBook.Worksheets(1)
            For index = 1 To 903
                searchname = .Range("F" & Format(index)).Value
                Dim connection As New OleDbConnection("Provider=Search.CollatorDSO;Extended Properties=""Application=Windows""")

                ' File name search (case insensitive), also searches sub directories
                query1 = "SELECT System.ItemName FROM SystemIndex " +
                                "WHERE scope ='" & filepath & "' AND System.ItemName LIKE '%" & searchname & "%' AND System.ItemType = '.pdf' AND System.ItemName NOT LIKE '%AR20%'"

                ' File name search (case insensitive), does Not search sub directories
                query2 = "SELECT System.ItemName FROM SystemIndex " +
                                "WHERE Directory ='" & filepath & "' AND System.ItemName LIKE '%E730_001%'"

                'Folder name search (case insensitive)
                query3 = "SELECT System.ItemName FROM SystemIndex " +
                                "WHERE scope = '" & filepath & "' AND System.ItemType = 'Directory' AND System.Itemname LIKE '%E23_001%' "

                ' Folder name search (case insensitive), does Not search sub directories
                query4 = "SELECT System.ItemName FROM SystemIndex " +
                                "WHERE directory = '" & filepath & "' AND System.ItemType = 'Directory' AND System.Itemname LIKE '%Summary%' "

                connection.Open()
                Dim command As New OleDbCommand(query1, connection)
                Dim r As OleDbDataReader
                Dim rowcount As Integer = 0
                r = command.ExecuteReader
                Using r
                    While (r.Read)
                        Console.WriteLine(r(0))
                        rowcount = rowcount + 1
                    End While

                    If rowcount >= 1 Then
                        .Range("G" & Format(index)).Value = "FAIL"
                    Else
                        .Range("G" & Format(index)).Value = "PASS"
                    End If
                End Using
                connection.Close()

            Next

        End With
        Console.WriteLine("Complete. Enter to Close")
        Console.ReadKey()
    End Sub

End Module
