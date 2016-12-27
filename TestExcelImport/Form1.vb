Imports Microsoft.Office.Interop.Excel
Imports System.Data
Public Class Form1

    Sub TestExcelImport()

        Dim openFile = New OpenFileDialog
        openFile.Title = "Select an Excel File"
        openFile.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"
        If openFile.ShowDialog() <> True Then
            Return
        End If

        Dim xl As New Microsoft.Office.Interop.Excel.Application
        Dim xlBooks As Workbooks = xl.Workbooks
        Dim thisFile As Workbook = xlBooks.Open(openFile.FileName)
        Dim returnSet As New DataSet

        For s As Integer = 1 To thisFile.Sheets.Count
            Dim returnTable As New System.Data.DataTable
            returnTable.TableName = String.Format("Table{0}", s)
            Dim firstSheet As Range = thisFile.Sheets(s).UsedRange
            For c As Integer = 1 To firstSheet.Columns.Count
                Dim newCol As New DataColumn
                newCol.ColumnName = String.Format("Column{0}", c)
                returnTable.Columns.Add(newCol)
            Next

            For r As Integer = 1 To firstSheet.Rows.Count
                Dim newRow As DataRow = returnTable.NewRow()
                For c As Integer = 1 To firstSheet.Columns.Count
                    newRow(c - 1) = firstSheet.Cells(r, c).Value.ToString()
                Next
                returnTable.Rows.Add(newRow)
                Console.WriteLine(String.Format("Read {0} row(s) from sheet {1}.", r - 1, s))
            Next
            returnSet.Tables.Add(returnTable)

        Next

        DataGridView1.DataSource = returnSet.Tables("Sheet1").DefaultView 'Or whatever

        thisFile.Close()
        xlBooks.Close()
        xl.Quit()




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
        Dim pharmacy As New System.Data.DataTable

        '' Create four typed columns in the DataTable.
        pharmacy.Columns.Add("Dosage", GetType(Integer))
        pharmacy.Columns.Add("Drug", GetType(String))
        pharmacy.Columns.Add("Patient", GetType(String))
        pharmacy.Columns.Add("Date", GetType(DateTime))


        ' Add five rows with those columns filled in the DataTable.
        pharmacy.Rows.Add(25, "Indocin", "David", DateTime.Now)
        pharmacy.Rows.Add(50, "Enebrel", "Sam", DateTime.Now)
        pharmacy.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now)
        pharmacy.Rows.Add(21, "Combivent", "Janet", DateTime.Now)
        pharmacy.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now)
        Return pharmacy
    End Function

    Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
