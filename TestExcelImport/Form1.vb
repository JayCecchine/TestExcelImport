Imports Microsoft.Office.Interop.Excel
Imports System.Data
Public Class Form1

    Sub TestExcelImport()


        Dim table As System.Data.DataTable = GetTable()
        ' Create new Application.
        Dim excel As Application = New Application
        Dim table1 As New System.Data.DataTable
        ' Open Excel spreadsheet.
        Dim FilePath = OpenFileDialog1.FileName
        Dim w As Workbook = excel.Workbooks.Open(FilePath)

        For i As Integer = 0 To w.Worksheets.Count


            Dim sheet As Worksheet = w.Sheets(1)

            Dim r As Range = sheet.UsedRange

            'Get used range of selected sheet
            'Make a for loop array that gets the data type and value of 
            '   the selected cell, i.e 1,1 , would minus by -1 
            '       and convert into a data table via add row.
            'Figure out how to create the columns, as the columns decide

            'table.Columns.AddRange()

            ' Dim array(,) As Object = r.Value(XlRangeValueDataType.xlRangeValueDefault)


            'Dim c As String = w.Sheets(1).range(r).Value.ToString()
            'TextBox2.Text = c

            'Have to figure out how to construct a data tabel from the excel file.
            'this might help http://stackoverflow.com/questions/14261655/best-fastest-way-to-read-an-excel-sheet-into-a-datatable
            'Dim b As String = w.Sheets(1).Cells(1, 1).Value.ToString()
        Next




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = GetTable()
        Try
            TestExcelImport()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Function TestTable() As System.Data.DataTable
        Dim tt As New System.Data.DataTable

        tt.Columns.Add("An", GetType(String))
        tt.Columns.Add("ba", GetType(String))
    End Function
    Function GetTable() As System.Data.DataTable
        ' Create new DataTable instance.
        Dim table As New System.Data.DataTable

        '' Create four typed columns in the DataTable.
        table.Columns.Add("Dosage", GetType(Integer))
        table.Columns.Add("Drug", GetType(String))
        table.Columns.Add("Patient", GetType(String))
        table.Columns.Add("Date", GetType(DateTime))


        ' Add five rows with those columns filled in the DataTable.
        table.Rows.Add(25, "Indocin", "David", DateTime.Now)
        table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now)
        table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now)
        table.Rows.Add(21, "Combivent", "Janet", DateTime.Now)
        table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now)
        Return table
    End Function

    Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
