Imports Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Threading


Public Class Form1
    Dim returnSet As New DataSet
    Dim modSet As New DataSet
    Dim regPrice As String
    Dim catPrice As String
    Const WM_SETTEXT As Integer = &HC
    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, <MarshalAs(UnmanagedType.LPStr)> lParam As String) As IntPtr
    End Function
    Public Sub NoteStart(getString As String)
        'copy to clipboard

        Dim strCopyMe As String = getString
        Clipboard.SetDataObject(strCopyMe)

        'open notepad and wait for it to completely load
        'Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start("notepad.exe")
        Dim p As New Process()
        p.StartInfo.FileName = "notepad.exe"
        p.Start()
        'SetWindowText(p.MainWindowHandle, "Test")
        p.WaitForInputIdle()

        ' paste the data from the clipboard
        SendKeys.Send("^V")
    End Sub
    Shared Function SetWindowText(ByVal hwnd As IntPtr, ByVal windowName As String) As Boolean
    End Function
    Sub ExcelConverter()

        Dim openFile = New OpenFileDialog
        openFile.Title = "Select an Excel File"
        openFile.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"

        'If use does not hit ok when selecting file, function does nothing
        If openFile.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        'Selecting the excel application, using the workbooks, selecting an excel file
        Dim xl As New Microsoft.Office.Interop.Excel.Application
        Dim xlBooks As Workbooks = xl.Workbooks
        Dim thisFile As Workbook = xlBooks.Open(openFile.FileName)

        'The following is protection against incorrectly formatted AP sheets, since I don't know how
        'to automatically fix this yet, I will just have the parser stop and not create the data table
        Dim thisSheet = thisFile.Sheets(1).usedRange

        Console.WriteLine()

        If thisSheet.cells(1, 1).value.ToString <> "REGULAR" Then
            MsgBox("This excel sheet is not formatted properly.")
            Exit Sub
        End If
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
                'Console.WriteLine(String.Format("Read {0} row(s) from sheet {1}.", r - 1, s))
            Next
            returnSet.Tables.Add(returnTable)

        Next

        'DataGridView1.DataSource = returnSet.Tables(0).DefaultView 'Or whatever

        thisFile.Close()
        xlBooks.Close()
        xl.Quit()

    End Sub

    Sub APTable()

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
        'Catering groups added next to certain items for the purpose of the driver catering fees scripts
        modTable.Rows.Add(r.Rows(1).Item(0), r.Rows(1).Item(1), r.Rows(1).Item(2), String.Format("{0} (Platters)", r.Rows(19).Item(0)), r.Rows(19).Item(1), r.Rows(19).Item(2))
        modTable.Rows.Add(r.Rows(2).Item(0), r.Rows(2).Item(1), r.Rows(2).Item(2), r.Rows(20).Item(0), r.Rows(20).Item(1), r.Rows(20).Item(2))
        modTable.Rows.Add(r.Rows(3).Item(0), r.Rows(3).Item(1), r.Rows(3).Item(2), r.Rows(21).Item(0), r.Rows(21).Item(1), r.Rows(21).Item(2))
        modTable.Rows.Add(r.Rows(4).Item(0), r.Rows(4).Item(1), r.Rows(4).Item(2), String.Format("{0} (Boxes)", r.Rows(22).Item(0)), r.Rows(22).Item(1), r.Rows(22).Item(2))
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

        'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'DataGridView1.AutoResizeColumns()

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
    End Sub



    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub


    Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button3.Hide()
    End Sub



    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            PriceScripts()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub PriceScripts()
        modSet.GetChanges()
        modSet.AcceptChanges()
        Dim u As String = "update tblMenuItems set Price1= '{0}' "
        Dim d As String = "update tblMenuItemExtend set DelCharge='{0}'"
        Dim regPrice(52) As String
        Dim catPrice(52) As String
        Dim regFee(52) As String
        Dim catStore(52) As String
        Dim catDriver(52) As String
        Dim menuLevel(52) As String
        Dim catering As String
        Dim m = modSet.Tables(0)
        Dim regpHead, catpHead, regfHead, catfHead As String
        Dim test As Integer
        Dim count As Integer




        'Regular Prices
        If m.Rows(0).Item(1).ToString() IsNot "" Then
            regPrice(0) = String.Format("--Slims--" & vbNewLine & u & "where itemnum in (124,141,142,143,144,145,146)" & vbNewLine, m.Rows(0).Item(1).ToString())
        Else
            regPrice(0) = ""
        End If
        If m.Rows(1).Item(1).ToString() IsNot "" Then
            regPrice(1) = String.Format("--Subs--" & vbNewLine & u & "where itemnum in (110,111,112,113,114,115,116)" & vbNewLine, m.Rows(1).Item(1).ToString())
        Else
            regPrice(1) = ""
        End If
        If m.Rows(2).Item(1).ToString() IsNot "" Then
            regPrice(2) = String.Format("--Clubs--" & vbNewLine & u & "where itemnum in (125,126,127,128,129,130,131,132,133,134,650)" & vbNewLine, m.Rows(2).Item(1).ToString())
        Else
            regPrice(2) = ""
        End If
        If m.Rows(3).Item(1).ToString() IsNot "" Then
            regPrice(3) = String.Format("--Garg--" & vbNewLine & u & "where itemnum in (117)" & vbNewLine, m.Rows(3).Item(1).ToString())
        Else
            regPrice(3) = ""
        End If
        If m.Rows(4).Item(1).ToString() IsNot "" Then
            regPrice(4) = String.Format("--Fresh Bread--" & vbNewLine & u & "where itemnum in (244)" & vbNewLine, m.Rows(4).Item(1).ToString())
        Else
            regPrice(4) = ""
        End If
        If m.Rows(5).Item(1).ToString() IsNot "" Then
            regPrice(5) = String.Format("--Day Old Bread--" & vbNewLine & u & "where itemnum in (209)" & vbNewLine, m.Rows(5).Item(1).ToString())
        Else
            regPrice(5) = ""
        End If
        If m.Rows(6).Item(1).ToString() IsNot "" Then
            regPrice(6) = String.Format("--Meats--" & vbNewLine & u & "where itemnum in (247,450,451,452,453,454,455,456,119,120,121,122,123,139,147,179,180,212,213,245)" & vbNewLine, m.Rows(6).Item(1).ToString())
        Else
            regPrice(6) = ""
        End If
        'X Bacon... do I need a separate If Statment here?
        'It'd be m.row(7).item(1) if so, For now I am ignoring it
        If m.Rows(8).Item(1).ToString() IsNot "" Then
            regPrice(8) = String.Format("--Cheese--" & vbNewLine & u & "where itemnum in (118,212,219,211)" & vbNewLine, m.Rows(8).Item(1).ToString())
        Else
            regPrice(8) = ""
        End If
        If m.Rows(9).Item(1).ToString() IsNot "" Then
            regPrice(9) = String.Format("--Avocado--" & vbNewLine & u & "where itemnum in (102,201,213,220)" & vbNewLine, m.Rows(9).Item(1).ToString())
        Else
            regPrice(9) = ""
        End If
        If m.Rows(10).Item(1).ToString() IsNot "" Then
            regPrice(10) = String.Format("--Pickles--" & vbNewLine & u & "where itemnum in (162,182,183)" & vbNewLine, m.Rows(10).Item(1).ToString())
        Else
            regPrice(10) = ""
        End If
        If m.Rows(11).Item(1).ToString() IsNot "" Then
            regPrice(11) = String.Format("--Chips--" & vbNewLine & u & "where itemnum in (221,222,223,224,255,256,257,259,258,598,675,597)" & vbNewLine, m.Rows(11).Item(1).ToString())
        Else
            regPrice(11) = ""
        End If
        If m.Rows(12).Item(1).ToString() IsNot "" Then
            regPrice(12) = String.Format("--Medium Sodas--" & vbNewLine & u & "where [group]='6' and ItemDesc like '%med%'" & vbNewLine, m.Rows(12).Item(1).ToString())
        Else
            regPrice(12) = ""
        End If
        If m.Rows(13).Item(1).ToString() IsNot "" Then
            regPrice(13) = String.Format("--Large Sodas--" & vbNewLine & u & "where [group]='6' and ItemDesc like '%lg%'" & vbNewLine, m.Rows(13).Item(1).ToString())
        Else
            regPrice(13) = ""
        End If
        If m.Rows(14).Item(1).ToString() IsNot "" Then
            regPrice(14) = String.Format("--Canned Sodas--" & vbNewLine & u & "where [group]='6' and ItemDesc like '%can%'" & vbNewLine, m.Rows(14).Item(1).ToString())
        Else
            regPrice(14) = ""
        End If
        If m.Rows(15).Item(1).ToString() IsNot "" Then
            regPrice(15) = String.Format("--Bottled Water--" & vbNewLine & u & "where itemnum in (155)" & vbNewLine, m.Rows(15).Item(1).ToString())
        Else
            regPrice(15) = ""
        End If
        If m.Rows(16).Item(1).ToString() IsNot "" Then
            regPrice(16) = String.Format("--Cookies--" & vbNewLine & u & "where itemnum in (215,216,676)" & vbNewLine, m.Rows(16).Item(1).ToString())
        Else
            regPrice(16) = ""
        End If

        '
        '
        'Catering Prices
        '
        '

        If m.Rows(0).Item(4).ToString() IsNot "" Then
            catPrice(0) = String.Format("--15 piece platter--" & vbNewLine & u & "where itemnum in (300)" & vbNewLine, m.Rows(0).Item(4).ToString())
        Else
            catPrice(0) = ""
        End If
        If m.Rows(1).Item(4).ToString() IsNot "" Then
            catPrice(1) = String.Format("--30 piece platter--" & vbNewLine & u & "where itemnum in (301)" & vbNewLine, m.Rows(1).Item(4).ToString())
        Else
            catPrice(1) = ""
        End If
        If m.Rows(2).Item(4).ToString() IsNot "" Then
            catPrice(2) = String.Format("--Club Upcharge--" & vbNewLine & u & "where itemnum in (467,468,469,470,471,472,473,474,475,476,652)" & vbNewLine, m.Rows(2).Item(4).ToString())
        Else
            catPrice(2) = ""
        End If
        If m.Rows(3).Item(4).ToString() IsNot "" Then
            catPrice(3) = String.Format("--Slim Box Lunch--" & vbNewLine & u & "where itemnum in (290,291,292,293,294,295,296,297)" & vbNewLine, m.Rows(3).Item(4).ToString())
        Else
            catPrice(3) = ""
        End If
        If m.Rows(4).Item(4).ToString() IsNot "" Then
            catPrice(4) = String.Format("--Sub Box Lunch--" & vbNewLine & u & "where itemnum in (306,307,308,309,310,311,312)" & vbNewLine, m.Rows(4).Item(4).ToString())
        Else
            catPrice(4) = ""
        End If
        If m.Rows(5).Item(4).ToString() IsNot "" Then
            catPrice(5) = String.Format("--Club Box Lunch--" & vbNewLine & u & "where itemnum in (313,314,315,316,317,318,319,320,321,322,651)" & vbNewLine, m.Rows(5).Item(4).ToString())
        Else
            catPrice(5) = ""
        End If
        If m.Rows(6).Item(4).ToString() IsNot "" Then
            catPrice(6) = String.Format("--Garg Box Lunch--" & vbNewLine & u & "where itemnum in (323)" & vbNewLine, m.Rows(6).Item(4).ToString())
        Else
            catPrice(6) = ""
        End If
        If m.Rows(7).Item(4).ToString() IsNot "" Then
            catPrice(7) = String.Format("--Pickle Bucket--" & vbNewLine & u & "where itemnum in (100,109,154,184)" & vbNewLine, m.Rows(7).Item(4).ToString())
        Else
            catPrice(7) = ""
        End If
        If m.Rows(8).Item(4).ToString() IsNot "" Then
            catPrice(8) = String.Format("--Cookie Box--" & vbNewLine & u & "where itemnum in (869,863,867)" & vbNewLine, m.Rows(8).Item(4).ToString())
        Else
            catPrice(8) = ""
        End If
        If m.Rows(9).Item(4).ToString() IsNot "" Then
            catPrice(9) = String.Format("--12 Pack JJ Minis--" & vbNewLine & u & "where itemnum in (850)" & vbNewLine, m.Rows(9).Item(4).ToString())
        Else
            catPrice(9) = ""
        End If
        If m.Rows(10).Item(4).ToString() IsNot "" Then
            catPrice(10) = String.Format("--24 Pack JJ Minis--" & vbNewLine & u & "where itemnum in (864)" & vbNewLine, m.Rows(10).Item(4).ToString())
        Else
            catPrice(10) = ""
        End If

        '
        '
        'Regular Fees
        '
        '
        If m.Rows(0).Item(2).ToString() IsNot "" Then
            regFee(0) = String.Format("--Slims--" & vbNewLine & d & "where itemnum in (124,141,142,143,144,145,146) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(0).Item(2).ToString())
        Else
            regFee(0) = ""
        End If
        If m.Rows(1).Item(2).ToString() IsNot "" Then
            regFee(1) = String.Format("--Subs--" & vbNewLine & d & "where itemnum in (110,111,112,113,114,115,116) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(1).Item(2).ToString())
        Else
            regFee(1) = ""
        End If
        If m.Rows(2).Item(2).ToString() IsNot "" Then
            regFee(2) = String.Format("--Clubs--" & vbNewLine & d & "where itemnum in (125,126,127,128,129,130,131,132,133,134,650) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(2).Item(2).ToString())
        Else
            regFee(2) = ""
        End If
        If m.Rows(3).Item(2).ToString() IsNot "" Then
            regFee(3) = String.Format("--Garg--" & vbNewLine & d & "where itemnum in (117) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(3).Item(2).ToString())
        Else
            regFee(3) = ""
        End If
        If m.Rows(4).Item(2).ToString() IsNot "" Then
            regFee(4) = String.Format("--Fresh Bread--" & vbNewLine & d & "where itemnum in (244) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(4).Item(2).ToString())
        Else
            regFee(4) = ""
        End If
        If m.Rows(5).Item(2).ToString() IsNot "" Then
            regFee(5) = String.Format("--Day Old Bread--" & vbNewLine & d & "where itemnum in (209) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(5).Item(2).ToString())
        Else
            regFee(5) = ""
        End If
        If m.Rows(6).Item(2).ToString() IsNot "" Then
            regFee(6) = String.Format("--Meats--" & vbNewLine & d & "where itemnum in (247,450,451,452,453,454,455,456,119,120,121,122,123,139,147,179,180,212,213,245) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(6).Item(2).ToString())
        Else
            regFee(6) = ""
        End If
        If m.Rows(8).Item(2).ToString() IsNot "" Then
            regFee(8) = String.Format("--Cheese--" & vbNewLine & d & "where itemnum in (118,212,219,211) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(8).Item(2).ToString())
        Else
            regFee(8) = ""
        End If
        If m.Rows(9).Item(2).ToString() IsNot "" Then
            regFee(9) = String.Format("--Avocado--" & vbNewLine & d & "where itemnum in (102,201,213,220) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(9).Item(2).ToString())
        Else
            regFee(9) = ""
        End If
        If m.Rows(10).Item(2).ToString() IsNot "" Then
            regFee(10) = String.Format("--Pickles--" & vbNewLine & d & "where itemnum in (162,182,183) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(10).Item(2).ToString())
        Else
            regFee(10) = ""
        End If
        If m.Rows(11).Item(2).ToString() IsNot "" Then
            regFee(11) = String.Format("--Chips--" & vbNewLine & d & "where itemnum in (221,222,223,224,255,256,257,259,258,598,675,597) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(11).Item(2).ToString())
        Else
            regFee(11) = ""
        End If
        If m.Rows(12).Item(2).ToString() IsNot "" Then
            regFee(12) = String.Format("--Medium Sodas--" & vbNewLine & d & "where itemnum in(190,191,192,193,194,196,197,237,238,239,240,241,242,243,340,341,362) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(12).Item(2).ToString())
        Else
            regFee(12) = ""
        End If
        If m.Rows(13).Item(2).ToString() IsNot "" Then
            regFee(13) = String.Format("--Large Sodas--" & vbNewLine & d & "where itemnum in(198,202,203,204,205,206,207,208,342,358,359,380,381,382,383,384,385,388) AND TransTypeID IN (3,9)" & vbNewLine, m.Rows(13).Item(2).ToString())
        Else
            regFee(13) = ""
        End If
        If m.Rows(14).Item(2).ToString() IsNot "" Then
            regFee(14) = String.Format("--Canned Sodas--" & vbNewLine & d & "where itemnum in (344,345,346,360,361,364,365,366,862,444,921,922,923,924,925,926,927,928,929,930,931) and TransTypeID IN (3,9)" & vbNewLine, m.Rows(14).Item(2).ToString())
        Else
            regFee(14) = ""
        End If
        If m.Rows(15).Item(2).ToString() IsNot "" Then
            regFee(15) = String.Format("--Bottled Water--" & vbNewLine & d & "where itemnum in (155) and TransTypeID IN (3,9)" & vbNewLine, m.Rows(15).Item(2).ToString())
        Else
            regFee(15) = ""
        End If
        If m.Rows(16).Item(2).ToString() IsNot "" Then
            regFee(16) = String.Format("--Cookies--" & vbNewLine & d & "where itemnum in (215,216,676) and TransTypeID IN (3,9)" & vbNewLine, m.Rows(16).Item(2).ToString())
            'Else
            regFee(16) = ""
        End If

        'Catering Fees to Store or Driver
        Select Case m.Rows(14).Item(4).ToString
            Case ""
                catering = ""
            Case = "Driver"
                If m.Rows(0).Item(5).ToString IsNot "" Then
                    catDriver(0) = String.Format("--Platters & Minis--" & vbNewLine & "update tblmenugroups set cateringcharge = '{0}', CateringGroup = 'True'" & vbNewLine & "where groupid = 24" & vbNewLine, m.Rows(0).Item(5).ToString())
                Else
                    catDriver(0) = ""
                End If
                If m.Rows(3).Item(5).ToString IsNot "" Then
                    catDriver(1) = String.Format("--Box Lunches--" & vbNewLine & "update tblmenugroups set cateringcharge = '{0}', CateringGroup = 'True'" & vbNewLine & "where groupid = 12" & vbNewLine, m.Rows(3).Item(5).ToString())
                Else
                    catDriver(1) = ""
                End If
                If m.Rows(8).Item(5).ToString IsNot "" Then
                    catDriver(2) = String.Format("--Cookie Boxes--" & vbNewLine & "update tblmenugroups set cateringcharge = '{0}', CateringGroup = 'True'" & vbNewLine & "where groupid = 21" & vbNewLine, m.Rows(8).Item(5).ToString())
                Else
                    catDriver(2) = ""
                End If
                If m.Rows(7).Item(5).ToString IsNot "" Then
                    catDriver(3) = String.Format("--Pickle Bucket--" & vbNewLine & "update tblmenugroups set cateringcharge = '{0}', CateringGroup = 'True'" & vbNewLine & "where groupid = 11" & vbNewLine, m.Rows(7).Item(5).ToString())
                Else
                    catDriver(3) = ""
                End If
                catering = catDriver(0) & catDriver(1) & catDriver(2) & catDriver(3)
            Case = "Store"
                Dim c As String = "update tblMenuItemExtend SET DelCharge = "
                If m.Rows(0).Item(5).ToString IsNot "" Then
                    catStore(0) = String.Format("--15 Piece Platter--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (300) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(0).Item(5).ToString)
                Else
                    catStore(0) = ""
                End If
                If m.Rows(1).Item(5).ToString IsNot "" Then
                    catStore(1) = String.Format("--30 Piece Platter--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (301) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(1).Item(5).ToString)
                Else
                    catStore(1) = ""
                End If
                If m.Rows(2).Item(5).ToString IsNot "" Then
                    catStore(2) = String.Format("--Club Upcharge--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (467,468,469,470,471,472,473,474,475,476,652) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(2).Item(5).ToString)
                Else
                    catStore(2) = ""
                End If
                If m.Rows(3).Item(5).ToString IsNot "" Then
                    catStore(3) = String.Format("--Slim Box--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (290,291,292,293,294,295,296,297) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(3).Item(5).ToString)
                Else
                    catStore(3) = ""
                End If
                If m.Rows(4).Item(5).ToString IsNot "" Then
                    catStore(4) = String.Format("--Sub Box--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (306,307,308,309,310,311,312) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(4).Item(5).ToString)
                Else
                    catStore(4) = ""
                End If
                If m.Rows(5).Item(5).ToString IsNot "" Then
                    catStore(5) = String.Format("--Club Box--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (313,314,315,316,317,318,319,320,321,322,651) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(5).Item(5).ToString)
                Else
                    catStore(5) = ""
                End If
                If m.Rows(6).Item(5).ToString IsNot "" Then
                    catStore(6) = String.Format("--Garg Box--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (323) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(6).Item(5).ToString)
                Else
                    catStore(6) = ""
                End If
                If m.Rows(7).Item(5).ToString IsNot "" Then
                    catStore(7) = String.Format("--Pickle Bucket--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (100,109,154,184) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(7).Item(5).ToString)
                Else
                    catStore(7) = ""
                End If
                If m.Rows(8).Item(5).ToString IsNot "" Then
                    catStore(8) = String.Format("--Cookie Box--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (869,863,867) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(8).Item(5).ToString)
                Else
                    catStore(8) = ""
                End If
                If m.Rows(9).Item(5).ToString IsNot "" Then
                    catStore(9) = String.Format("--12 Pack Mini Jimmy--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (850) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(9).Item(5).ToString)
                Else
                    catStore(9) = ""
                End If
                If m.Rows(10).Item(5).ToString IsNot "" Then
                    catStore(10) = String.Format("--24 Pack Mini Jimmy--" & vbNewLine & "{0}'{1}' WHERE ItemNum IN (864) and TransTypeID IN (3,9)" & vbNewLine, c, m.Rows(10).Item(5).ToString)
                Else
                    catStore(10) = ""
                End If
                catering = catStore(0) & catStore(1) & catStore(2) & catStore(3) & catStore(4) & catStore(5) & catStore(6) & catStore(7) & catStore(8) & catStore(9) & catStore(10)



        End Select

        'Determining if there are values in the array, if there are no values, it will not generate a heading
        For i As Integer = 0 To regPrice.GetUpperBound(0)
            If regPrice(i) = "" Then
                count = count + 1
            Else
                count = count + 0
            End If
        Next
        If count = 53 Then
            regpHead = ""
        Else
            regpHead = "" & vbNewLine &
                "-------------------------------------------------------------------------------------------------------------" & vbNewLine &
                "-------------------------------------------Regular Prices----------------------------------------------------" & vbNewLine &
                "-------------------------------------------------------------------------------------------------------------" & vbNewLine
        End If
        count = 0
        For i As Integer = 0 To catPrice.GetUpperBound(0)
            If catPrice(i) = "" Then
                count = count + 1
            Else
                count = count + 0
            End If

        Next
        If count = 53 Then
            catpHead = ""
        Else
            catpHead = "" & vbNewLine &
                "-------------------------------------------------------------------------------------------------------------" & vbNewLine &
                "-------------------------------------------Catering Prices---------------------------------------------------" & vbNewLine &
                "-------------------------------------------------------------------------------------------------------------" & vbNewLine
        End If
        count = 0
        For i As Integer = 0 To regFee.GetUpperBound(0)
            If regFee(i) = "" Then
                count = count + 1
            Else
                count = count + 0
            End If
        Next
        If count = 53 Then
            regfHead = ""
        Else
            regfHead = "" & vbNewLine &
                "-------------------------------------------------------------------------------------------------------------" & vbNewLine &
                "--------------------------------------------Delivery Fees----------------------------------------------------" & vbNewLine &
                "-------------------------------------------------------------------------------------------------------------" & vbNewLine
        End If

        If catering = "" Then
            catfHead = ""
        Else
            catfHead = "" & vbNewLine &
                "-------------------------------------------------------------------------------------------------------------" & vbNewLine &
                "--------------------------------------------Catering Fees----------------------------------------------------" & vbNewLine &
                "-------------------------------------------------------------------------------------------------------------" & vbNewLine
        End If

        Me.regPrice = "use PDQPOS" & vbNewLine & "go" & vbNewLine &
        regpHead &
        regPrice(0) & regPrice(1) & regPrice(2) & regPrice(3) & regPrice(4) & regPrice(5) & regPrice(6) & regPrice(7) &
        regPrice(8) & regPrice(9) & regPrice(10) & regPrice(11) & regPrice(12) & regPrice(13) & regPrice(14) & regPrice(15) & regPrice(16) &
        regPrice(17) & vbNewLine &
        catpHead &
        catPrice(0) & catPrice(1) & catPrice(2) & catPrice(3) & catPrice(4) & catPrice(5) & catPrice(6) &
        catPrice(7) & catPrice(8) & catPrice(9) & catPrice(10) & vbNewLine &
        regfHead &
        regFee(0) & regFee(1) & regFee(2) & regFee(3) & regFee(4) & regFee(5) & regFee(6) &
        regFee(7) & regFee(8) & regFee(9) & regFee(10) & regFee(11) & regFee(12) & regFee(13) & regFee(14) &
        regFee(15) & regFee(16) & vbNewLine &
        catfHead & catering
        NoteStart(Me.regPrice)
    End Sub
    Public Sub FlatDeliveryFee()
        Try
            Dim flatFee As String
            Dim m = modSet.Tables(0)
            Dim flatFeeVal = m.Rows(13).Item(4).ToString

            If flatFeeVal.Length > 0 Then
                Dim v = vbNewLine
                flatFee = String.Format("use pdqpos" & v & "go" & v & "--Resets all Del charges to 0--" &
                    v & "UPDATE tblMenuItemExtend SET DelCharge='0.00'" & v & "--UPDATE TAX ON DELIVERY FEE--" &
                    v & "UPDATE tblMenuItemExtend SET TaxDelCharge='False'" & v & "WHERE TransTypeID IN (3,9)" & v & "--Set Flat Fee below--" & v &
                    "UPDATE tblMenuItemExtend SET DelCharge= '{0}'" & v & "WHERE ItemNum IN (124,141,142,143,144,145,146,140,110,111,112,113,114,115,116,125,126,127,128,129,130,131,132,133,134,650,117,244,209,247,450,451,452,453,454,455,456,119,120,121,122,123,139,147,179,180,212,213,245,201,211,102,212,213,118,219,102,220,162,182,183,221,222,223,224,255,256,257,259,258,598,675,597,190,191,192,193,194,196,197,237,238,239,240,241,242,243,340,341,362,198,202,203,204,205,206,207,208,342,358,359,380,381,382,383,384,385,388,344,345,346,360,361,364,365,366,862,444,155,215,216,676,921,922,923,924,925,926,927,928,929,930,931)
                    AND TransTypeID IN (3,9)", flatFeeVal)
                NoteStart(flatFee)

            Else
                Form3.Show()
                Form3.Button1.Hide()
                Form3.TextBox2.Hide()
                Form3.Label2.Hide()
                Form3.Label1.Text = "FlatFee"
            End If

        Catch
            MsgBox("You did something stupid")
        End Try


    End Sub
    Public Sub Wheat(getPrice1 As String)
        Dim wheat As String
        Dim v = vbNewLine
        Dim w = "where itemnum = 135 and itemlevel "
        Dim u = "update tblmenuitemlevel"
        Dim s = "set itemprice1 = "

        wheat = String.Format("Use PDQPOS" & v & "go" & v & "--slims & slim boxes--" & v &
                u & v & s & "'{0}' " & w & "between 1 and 7" &
                v & "--subs & sub boxes--" & v & u & v & s & "'{0}' " & w & "between 8 and 14" &
                v & "--garg & garg box--" & v & u & v & s & "'{0}' " & w & "= " & "26" &
                v & "--club & club boxes--" & v & u & v & s & "'{0}' " & w & "between 15 and 25" &
                v & "--doublegut--" & v & u & v & s & "'{0}' " & w & "= " & "34", getPrice1)
        NoteStart(wheat)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            APTable()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub Button2_Click_1(sender As Object, e As EventArgs)
        Try
            APTable()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button2_Click_2(sender As Object, e As EventArgs) Handles Button2.Click
        Form2.Show()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim returnset As New DataSet
            Dim modset As New DataSet
        returnset.Clear()
        modset.Clear()


        Try
            ExcelConverter()
        Catch ex As Exception

        End Try

    End Sub
End Class
