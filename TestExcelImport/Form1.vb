Imports Microsoft.Office.Interop.Excel
Imports System.Data
Public Class Form1
    Dim returnSet As New DataSet

    Public Shared Function IsNullOrEmpty(
    value As String
) As Boolean
    End Function
    Function TestExcelImport()

        Dim openFile = New OpenFileDialog
        openFile.Title = "Select an Excel File"
        openFile.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"

        'If use does not hit ok when selecting file, function does nothing
        If openFile.ShowDialog() <> DialogResult.OK Then
            Return Nothing
        End If

        'Selecting the excel application, using the workbooks, selecting an excel file
        Dim xl As New Microsoft.Office.Interop.Excel.Application
        Dim xlBooks As Workbooks = xl.Workbooks
        Dim thisFile As Workbook = xlBooks.Open(openFile.FileName)


        'For every sheet, create a new data table that is called Table1,2,3 etc
        For s As Integer = 1 To thisFile.Sheets.Count
            Dim returnTable As New System.Data.DataTable
            returnTable.TableName = String.Format("Table{0}", s)
            Dim firstSheet As Range = thisFile.Sheets(s).UsedRange
            'Create a new column in the data table for every column in excel sheet
            For c As Integer = 1 To firstSheet.Columns.Count
                Dim newCol As New DataColumn
                newCol.ColumnName = String.Format("Column{0}", c)
                returnTable.Columns.Add(newCol)
            Next
            'Create a new row
            For r As Integer = 1 To firstSheet.Rows.Count
                Dim newRow As DataRow = returnTable.NewRow()
                For c As Integer = 1 To firstSheet.Columns.Count

                    If firstSheet.Cells(r, c).Value2 = Nothing Then
                        newRow(c - 1) = ""
                    Else
                        newRow(c - 1) = firstSheet.Cells(r, c).Value.ToString()
                    End If

                Next
                returnTable.Rows.Add(newRow)
                Console.WriteLine(String.Format("Read {0} row(s) from sheet {1}.", r - 1, s))
            Next
            returnSet.Tables.Add(returnTable)

        Next

        DataGridView1.DataSource = returnSet.Tables(0).DefaultView 'Or whatever

        thisFile.Close()
        xlBooks.Close()
        xl.Quit()

    End Function

    Function TestDataRetrieval()


        Dim s = returnSet.Tables(0)

        If s.Rows(1).Item(1) IsNot "" Then         'And s.Rows(1).Item(1) > 0 Then
            TextBox2.Text = String.Format("Update tblmenuitems Set price1 = ' {0}' where itemnum in (100-111)", s.Rows(1).Item(1))
        End If

        Dim modSet As New DataSet
        Dim modTable As New System.Data.DataTable
        Dim r = returnSet.Tables(0)

        modTable.Columns.Add("Regular")
        modTable.Columns.Add("Sub Total")
        modTable.Columns.Add("Delivery Fee")

        modTable.Rows.Add(r.Rows(1).Item(0), r.Rows(1).Item(1), r.Rows(1).Item(2))
        modSet.Tables.Add(modTable)

        Form2.Show()
        Form2.DataGridView1.DataSource = modSet.Tables(0).DefaultView



        'Else
        '   TextBox2.Text = ""
        ' End If

        'Parse Excel File for values, add them if they're not null or 0
        'associate them with variables
        'compile variables into a data table that is user friendly
        '   ask user if data is correct, if not, allow for user to override pre-existing values
        '       Take new acquired data from data table, generate necessary scripts for end user


    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim testForm As Form = New Form
        Try
            TestExcelImport()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub


    Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            TestDataRetrieval()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
