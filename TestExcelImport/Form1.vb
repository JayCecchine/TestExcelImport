Imports Microsoft.Office.Interop.Excel
Imports System.Data
Public Class Form1
    Sub TestExcelImport()

        Dim table As System.Data.DataTable = GetTable()
        ' Create new Application.
        Dim excel As Application = New Application

        ' Open Excel spreadsheet.
        Dim w As Workbook = excel.Workbooks.Open("C:\Excel\test.xlsx")
        ' Dim Finals, Dimen1, Dimen2 As String

        Dim b As String = w.Sheets(1).Range("A3").Value.ToString()

        TextBox2.Text = b
        ' Loop over all sheets.
        For i As Integer = 1 To w.Sheets.Count

            ' Get sheet.
            Dim sheet As Worksheet = w.Sheets(i)

            ' Get range.
            Dim r As Range = sheet.UsedRange


            ' Load all cells into 2d array.
            Dim array(,) As Object = r.Value(XlRangeValueDataType.xlRangeValueDefault)
            'Dim butts As String = r.Value(XlRangeValueDataType.xlRangeValueDefault)
            'TextBox1.Text = butts
            ' Scan the cells.
            If array IsNot Nothing Then
                Debug.WriteLine("Length: {0}", array.Length)

                ' Get bounds of the array.
                Dim bound0 As Integer = array.GetUpperBound(0)
                Dim bound1 As Integer = array.GetUpperBound(1)





                ' Loop over all elements.
                For j As Integer = 1 To bound0
                    For x As Integer = 1 To bound1
                        Dim s1 As String = array(j, x)
                        Debug.Write(s1)
                        Debug.Write(" "c)

                    Next
                    Debug.WriteLine("")
                Next
            End If

        Next
        ' Dimen1 = "Dimension 0: " &
        'Dimen2 = "Dimension 1: " &
        'Finals = Dimen1 & Dimen2
        'TextBox1.Text = b
        ' Close.

        w.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = GetTable()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Function GetTable() As System.Data.DataTable
        ' Create new DataTable instance.
        Dim table As New System.Data.DataTable

        ' Create four typed columns in the DataTable.
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
End Class
