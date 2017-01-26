Public Class Form3
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim price1, price2 As String
        price1 = TextBox1.Text.ToString

        ' Try
        Select Case price1
            Case ""
                MsgBox("Enter something, Dumbass")
            Case price1.Length > 0
                price1 = TextBox1.Text.ToString
                price2 = TextBox2.Text.ToString
                Form1.Wheat(price1)
                Hide()
        End Select
        'Catch ex(messagebox(ex)
    End Sub
    Public Sub FlatFeeManual()
        Dim v = vbNewLine
        Dim flatFee As String
        Dim flatFeeVal As String = TextBox1.Text.ToString
        'Select Case flatFeeVal
        '    Case ""
        '        MsgBox("Enter something, Dumbass")
        If flatFeeVal.Length > 0 Then
            flatFee = String.Format("use pdqpos" & v & "go" & v & "--Resets all Del charges to 0--" &
                 v & "UPDATE tblMenuItemExtend SET DelCharge='0.00'" & v & "--UPDATE TAX ON DELIVERY FEE--" &
                 v & "UPDATE tblMenuItemExtend SET TaxDelCharge='False'" & v & "WHERE TransTypeID IN (3,9)" & v & "--Set Flat Fee below--" & v &
                 "UPDATE tblMenuItemExtend SET DelCharge= '{0}'" & v & "WHERE ItemNum IN (124,141,142,143,144,145,146,140,110,111,112,113,114,115,116,125,126,127,128,129,130,131,132,133,134,650,117,244,209,247,450,451,452,453,454,455,456,119,120,121,122,123,139,147,179,180,212,213,245,201,211,102,212,213,118,219,102,220,162,182,183,221,222,223,224,255,256,257,259,258,598,675,597,190,191,192,193,194,196,197,237,238,239,240,241,242,243,340,341,362,198,202,203,204,205,206,207,208,342,358,359,380,381,382,383,384,385,388,344,345,346,360,361,364,365,366,862,444,155,215,216,676,921,922,923,924,925,926,927,928,929,930,931)
                 AND TransTypeID IN (3,9)", flatFeeVal)
            Form1.NoteStart(flatFee)
        Else
            MsgBox("Enter a flat fee.")
        End If


        'End Select

    End Sub
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            FlatFeeManual()
        Catch exception As Exception
            MsgBox(exception)
        End Try
    End Sub
End Class