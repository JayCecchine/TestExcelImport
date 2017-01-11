Imports Microsoft.Office.Interop.Excel
Imports System.Data
Public Class Form1
    Dim returnSet As New DataSet
    Dim modSet As New DataSet

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

        'DataGridView1.DataSource = returnSet.Tables(0).DefaultView 'Or whatever

        thisFile.Close()
        xlBooks.Close()
        xl.Quit()

    End Function

    Function TestDataRetrieval()

        Button3.Show()
        Dim s = returnSet.Tables(0)

        'If s.Rows(1).Item(1) IsNot "" Then         'And s.Rows(1).Item(1) > 0 Then
        '    TextBox2.Text = String.Format("Update tblmenuitems Set price1 = '{0}' where itemnum in (100-111)", s.Rows(1).Item(1))
        'End If


        Dim modTable As New System.Data.DataTable
        Dim r = returnSet.Tables(0)
        Dim newCol = New DataColumn
        Dim newRow = modTable.NewRow()


        'Creating the new DataTable
        '   Creating the predefined columns
        modTable.Columns.Add("Regular")
        modTable.Columns.Add("Sub Total")
        modTable.Columns.Add("Delivery Fee")
        modTable.Columns.Add("Catering")
        modTable.Columns.Add("Cat Sub Total")
        modTable.Columns.Add("Cat Delivery Fee")

        'Filling the necessary data from the returnSet
        '   
        'Item Name / Subtotal / Del Fee  Catering Item / Subtotal / Del Fee
        modTable.Rows.Add(r.Rows(1).Item(0), r.Rows(1).Item(1), r.Rows(1).Item(2), r.Rows(19).Item(0), r.Rows(19).Item(1), r.Rows(19).Item(2))
        modTable.Rows.Add(r.Rows(2).Item(0), r.Rows(2).Item(1), r.Rows(2).Item(2), r.Rows(20).Item(0), r.Rows(20).Item(1), r.Rows(20).Item(2))
        modTable.Rows.Add(r.Rows(3).Item(0), r.Rows(3).Item(1), r.Rows(3).Item(2), r.Rows(21).Item(0), r.Rows(21).Item(1), r.Rows(21).Item(2))
        modTable.Rows.Add(r.Rows(4).Item(0), r.Rows(4).Item(1), r.Rows(4).Item(2), r.Rows(22).Item(0), r.Rows(22).Item(1), r.Rows(22).Item(2))
        modTable.Rows.Add(r.Rows(5).Item(0), r.Rows(5).Item(1), r.Rows(5).Item(2), r.Rows(23).Item(0), r.Rows(23).Item(1), r.Rows(23).Item(2))
        modTable.Rows.Add(r.Rows(6).Item(0), r.Rows(6).Item(1), r.Rows(6).Item(2), r.Rows(24).Item(0), r.Rows(24).Item(1), r.Rows(24).Item(2))
        modTable.Rows.Add(r.Rows(7).Item(0), r.Rows(7).Item(1), r.Rows(7).Item(2), r.Rows(25).Item(0), r.Rows(25).Item(1), r.Rows(25).Item(2))
        modTable.Rows.Add(r.Rows(8).Item(0), r.Rows(8).Item(1), r.Rows(8).Item(2), r.Rows(26).Item(0), r.Rows(26).Item(1), r.Rows(26).Item(2))
        modTable.Rows.Add(r.Rows(9).Item(0), r.Rows(9).Item(1), r.Rows(9).Item(2), r.Rows(27).Item(0), r.Rows(27).Item(1), r.Rows(27).Item(2))
        modTable.Rows.Add(r.Rows(10).Item(0), r.Rows(10).Item(1), r.Rows(10).Item(2), r.Rows(28).Item(0), r.Rows(28).Item(1), r.Rows(28).Item(2))
        modTable.Rows.Add(r.Rows(11).Item(0), r.Rows(11).Item(1), r.Rows(11).Item(2), r.Rows(29).Item(0), r.Rows(29).Item(1), r.Rows(29).Item(2))
        modTable.Rows.Add(r.Rows(12).Item(0), r.Rows(12).Item(1), r.Rows(12).Item(2))
        modTable.Rows.Add(r.Rows(13).Item(0), r.Rows(13).Item(1), r.Rows(13).Item(2), "Delivery Cap:", r.Rows(2).Item(10))
        modTable.Rows.Add(r.Rows(14).Item(0), r.Rows(14).Item(1), r.Rows(14).Item(2), "Flat Del Fee:", r.Rows(4).Item(10))
        modTable.Rows.Add(r.Rows(15).Item(0), r.Rows(15).Item(1), r.Rows(15).Item(2), "Use Cat Fees:", r.Rows(6).Item(10))
        modTable.Rows.Add(r.Rows(16).Item(0), r.Rows(16).Item(1), r.Rows(16).Item(2), "Tax Del Fees:", r.Rows(9).Item(10))
        modTable.Rows.Add(r.Rows(17).Item(0), r.Rows(17).Item(1), r.Rows(17).Item(2), "Tax Cat Fees:", r.Rows(11).Item(10))

        modSet.Tables.Add(modTable)





        'Dim Temp As New Form
        'Dim Grid As New DataGridView
        'Grid.DataSource = modSet.Tables(0).DefaultView

        'Temp.Show(Grid)


        DataGridView1.DataSource = modSet.Tables(0).DefaultView

        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        DataGridView1.AutoResizeColumns()

        '    modSet.Tables(0).GetChanges()
        '    modSet.Tables(0).AcceptChanges()

        ' Form2.Height = Height
        'Form2.Width = Width
        'Form2.Show()


        'Parse Excel File for values, add them if they're not null or 0
        'associate them with variables
        'compile variables into a data table that is user friendly
        '   ask user if data is correct, if not, allow for user to override pre-existing values
        '       Take new acquired data from data table, generate necessary scripts for end user
    End Function
    Function PriceScripts()
        modSet.GetChanges()
        modSet.AcceptChanges()
        Dim u As String = "update tblMenuItems where itemnum in"
        Dim m = modSet.Tables(0)


        'Regular Prices
        TextBox3.Text =
        m.Rows(0).Item(1).ToString()
        m.Rows(1).Item(1).ToString()
        m.Rows(2).Item(1).ToString()
        m.Rows(3).Item(1).ToString()
        m.Rows(4).Item(1).ToString()
        m.Rows(5).Item(1).ToString()
        m.Rows(6).Item(1).ToString()
        m.Rows(7).Item(1).ToString()
        m.Rows(8).Item(1).ToString()
        m.Rows(9).Item(1).ToString()
        m.Rows(10).Item(1).ToString()
        m.Rows(11).Item(1).ToString()
        m.Rows(12).Item(1).ToString()
        m.Rows(13).Item(1).ToString()
        m.Rows(14).Item(1).ToString()
        m.Rows(15).Item(1).ToString()
        m.Rows(16).Item(1).ToString()

        'Regular Del Fees
        m.Rows(0).Item(2).ToString()
        m.Rows(1).Item(2).ToString()
        m.Rows(2).Item(2).ToString()
        m.Rows(3).Item(2).ToString()
        m.Rows(4).Item(2).ToString()
        m.Rows(5).Item(2).ToString()
        m.Rows(6).Item(2).ToString()
        m.Rows(7).Item(2).ToString()
        m.Rows(8).Item(2).ToString()
        m.Rows(9).Item(2).ToString()
        m.Rows(10).Item(2).ToString()
        m.Rows(11).Item(2).ToString()
        m.Rows(12).Item(2).ToString()
        m.Rows(13).Item(2).ToString()
        m.Rows(14).Item(2).ToString()
        m.Rows(15).Item(2).ToString()
        m.Rows(16).Item(2).ToString()

        'Catering Prices
        m.Rows(0).Item(4).ToString()
        m.Rows(1).Item(4).ToString()
        m.Rows(2).Item(4).ToString()
        m.Rows(3).Item(4).ToString()
        m.Rows(4).Item(4).ToString()
        m.Rows(5).Item(4).ToString()
        m.Rows(6).Item(4).ToString()
        m.Rows(7).Item(4).ToString()
        m.Rows(8).Item(4).ToString()
        m.Rows(9).Item(4).ToString()
        m.Rows(10).Item(4).ToString()

        'Catering Del Fees
        m.Rows(0).Item(5).ToString()
        m.Rows(1).Item(5).ToString()
        m.Rows(2).Item(5).ToString()
        m.Rows(3).Item(5).ToString()
        m.Rows(4).Item(5).ToString()
        m.Rows(5).Item(5).ToString()
        m.Rows(6).Item(5).ToString()
        m.Rows(7).Item(5).ToString()
        m.Rows(8).Item(5).ToString()
        m.Rows(9).Item(5).ToString()
        m.Rows(10).Item(5).ToString()
    End Function



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        returnSet.Tables.Clear()
        modSet.Tables.Clear()

        Try
            TestExcelImport()
            TestDataRetrieval()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub


    Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button3.Hide()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            TestDataRetrieval()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PriceScripts()
    End Sub
End Class
