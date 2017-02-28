Public Class Form3
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim price1, price2 As String
        price1 = TextBox1.Text.ToString

        If price1.Length > 0 Then
            price1 = TextBox1.Text.ToString
            Form1.Wheat(price1)
            Close()
        Else
            MsgBox("Enter a value.")
            Exit Sub
        End If

    End Sub
    Public Sub FlatFeeManual()
        Dim v = vbNewLine
        Dim flatFee As String
        Dim flatFeeVal As String = TextBox1.Text.ToString

        If flatFeeVal.Length > 0 Then
            flatFee = String.Format("use pdqpos" & v & "go" & v & "--Resets all Del charges to 0--" &
                 v & "UPDATE tblMenuItemExtend SET DelCharge='0.00'" & v & "--UPDATE TAX ON DELIVERY FEE--" &
                 v & "UPDATE tblMenuItemExtend SET TaxDelCharge='False'" & v & "WHERE TransTypeID IN (3,9)" & v & "--Set Flat Fee below--" & v &
                 "UPDATE tblMenuItemExtend SET DelCharge= '{0}'" & v & "WHERE ItemNum IN (124,141,142,143,144,145,146,140,110,111,112,113,114,115,116,125,126,127,128,129,130,131,132,133,134,650,117,244,209,247,450,451,452,453,454,455,456,119,120,121,122,123,139,147,179,180,212,213,245,201,211,102,212,213,118,219,102,220,162,182,183,221,222,223,224,255,256,257,259,258,598,675,597,190,191,192,193,194,196,197,237,238,239,240,241,242,243,340,341,362,198,202,203,204,205,206,207,208,342,358,359,380,381,382,383,384,385,388,344,345,346,360,361,364,365,366,862,444,155,215,216,676,921,922,923,924,925,926,927,928,929,930,931)
                 AND TransTypeID IN (3,9)", flatFeeVal)
            Form1.NoteStart(flatFee)
        Else
            MsgBox("Enter a value.")
            Exit Sub
        End If

        Close()

        'End Select

    End Sub


    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Hide()
        Button2.Hide()
        Button3.Hide()
        Label2.Hide()
        'TextBox2.Hide()
        TextBox3.Hide()
        'TextBox4.Hide()
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
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



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim empDiscount = TextBox1.Text.ToString
        If empDiscount.Length > 0 Then
            empDiscount = TextBox1.Text.ToString
            Form1.EmpDisc(empDiscount)
            Close()
        Else
            MsgBox("Enter a value.")
            Exit Sub
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            BuildManual()
        Catch ex As Exception
            MsgBox(ex)
        End Try
    End Sub
    Private Sub BuildManual()
        Dim v = vbNewLine
        Dim buildThis As String
        Dim delCapVal As String = TextBox1.Text.ToString
        Dim taxDelVal As String
        Dim catCapVal As String = ""
        Dim Taxcatval As String
        Dim pin As String = TextBox5.Text.ToString

        If RadioButton1.Checked = True Then
            taxDelVal = "True"
        ElseIf RadioButton2.Checked = True Then
            taxDelVal = "False"
        End If
        If RadioButton3.Checked = True Then
            Taxcatval = "True"
        ElseIf RadioButton4.Checked = True Then
            Taxcatval = "False"
        End If

        If pin = "4879" Then
            If delCapVal.Length > 0 And taxDelVal.Length > 0 And Taxcatval.Length > 0 Then
                buildThis = String.Format("USE PDQPOS" & v &
            "GO" & v &
            "UPDATE tblconfigmain SET optvalue='False' WHERE optname='LegacyOrderCalc' AND optgroup='AdvOrderCalc'" & v &
            "UPDATE tblconfigmain SET optvalue='True' WHERE optname='AdvDelChargePerItem' AND optgroup='AdvOrderCalc'" & v &
            "UPDATE tblconfigmain SET optvalue='{0}' WHERE optname='AdvDelChargeMax' AND optgroup='AdvOrderCalc'" & v &
            "--UPDATE tblconfigmain SET optvalue='{1}' WHERE optname='AdvDelCaterChargeMax' AND" & v &
            "--optgroup='AdvOrderCalc'" & v &
            "UPDATE tblconfigmain SET optvalue='True' WHERE optname='JJAdvancedView' AND optgroup='MenuEntry'" & v &
            "UPDATE tblConfigMain SET OptValue = 'True'  WHERE OptName = 'AdvancedOOAPI' AND OptGroup='OnlineOrdering'" & v &
            "UPDATE [PDQPOS].[dbo].[tblDescTransType] SET [DescItemPriceIndex]=1" & v &
            "UPDATE [PDQPOS].[dbo].[tblConfigMain] SET [OptValue]= '{2}' WHERE [OptName] = 'TaxDeliveryCharge'" & v &
            "UPDATE [PDQPOS].[dbo].[tblConfigMain] SET [OptValue]= '{3}' WHERE [OptName] = 'TaxCateringCharge'" & v &
            "DECLARE @TaxDelCharge BIT" & v &
                "SELECT @TaxDelCharge =COUNT(*)" & v &
                "FROM tblConfigMain" & v &
                "WHERE tblConfigMain.OptGroup='DeliveryCharges' AND tblConfigMain.OptName ='TaxDeliveryCharge' AND OptValue='true'" & v &
            "INSERT INTO [tblOrderItemExtend]" & v &
                       "([OrderID]" & v &
                       ",[ItemUID]" & v &
                       ",[TransTypeID]" & v &
                       ",[PriceIndex]" & v &
                       ",[ItemPrice]" & v &
                       ",[AltItemPrice]" & v &
                       ",[DelCharge]" & v &
                       ",[TaxItem]" & v &
                       ",[TaxDelCharge]" & v &
                       ",[TaxRateID]" & v &
                       ",[TaxRate]" & v &
                       ",[AdvTaxAmountItem]" & v &
                       ",[AdvTaxAmountDelCharge]" & v &
                       ",[LinkedSmartCoupID])" & v &
                 "SELECT ITM.OrderID,ITM.ItemID ," & v &
                 "transtypeid, PriceIndex,ItemPrice,AltPrice,ITM.DelCharge,ChargeTax" & v &
                 ", @TaxDelCharge AS 'TaxDelCharge',ITM.TaxRateID," & v &
                 "ISNULL((SELECT TaxRate FROM tblConfigTaxRate" & v &
                 "WHERE TaxRateID=ITM.TaxRateID),0)" & v &
                 ",ITM.ItemTax, ITM.DelChargeTax" & v &
                 ",'00000000-0000-0000-0000-000000000000'" & v &
                  "FROM tblOrderItems AS ITM" & v &
                 "INNER JOIN tblOrders ON tblOrders.OrderID=ITM.OrderID" & v &
                 "WHERE ITM.ItemID NOT IN (SELECT [ItemUID] FROM tblOrderItemExtend)" & v &
            "------------------------------------------------------------------------------------------------------------------------------------------------------------------" & v &
            "------------------------------------------------------------------------------------------------------------------------------------------------------------------" & v &
            "------------------------------------------------------------------------------------------------------------------------------------------------------------------" & v &
            "------------------------------------------------------------------------------------------------------------------------------------------------------------------" & v &
            "INSERT INTO [tblDelayOrderItemExtend]" & v &
                       "([OrderID]" & v &
                       ",[ItemUID]" & v &
                       ",[TransTypeID]" & v &
                       ",[PriceIndex]" & v &
                       ",[ItemPrice]" & v &
                       ",[AltItemPrice]" & v &
                       ",[DelCharge]" & v &
                       ",[TaxItem]" & v &
                       ",[TaxDelCharge]" & v &
                       ",[TaxRateID]" & v &
                       ",[TaxRate]" & v &
                       ",[AdvTaxAmountItem]" & v &
                       ",[AdvTaxAmountDelCharge]" & v &
                       ",[LinkedSmartCoupID])" & v &
                 "SELECT ITM.OrderID,ITM.ItemID ," & v &
                 "transtypeid, PriceIndex,ItemPrice,AltPrice,ITM.DelCharge,ChargeTax" & v &
                 ", @TaxDelCharge AS 'TaxDelCharge',ITM.TaxRateID," & v &
                 "ISNULL((SELECT TaxRate FROM tblConfigTaxRate" & v &
                 "WHERE TaxRateID=ITM.TaxRateID),0)" & v &
                 ",ITM.ItemTax, ITM.DelChargeTax" & v &
                 ",'00000000-0000-0000-0000-000000000000'" & v &
                 "FROM tblDelayOrderItems AS ITM" & v &
                 "INNER JOIN tblDelayOrders ON tblDelayOrders.OrderID=ITM.OrderID" & v &
                 "WHERE ITM.ItemID NOT IN (SELECT [ItemUID] FROM tblDelayOrderItemExtend) AND tblDelayOrders.Settled=0", delCapVal, catCapVal, taxDelVal, Taxcatval)

                Form1.NoteStart(buildThis)
            Else
                MsgBox("Enter a value.")
                Exit Sub
            End If
        Else
            MsgBox("Access Denied")
            Exit Sub
        End If


        Close()
    End Sub
    Private Sub TaxDelCharge()
        Dim v = vbNewLine
        Dim taxThis As String
        Dim taxDelVal As String
        Dim pin As String = TextBox5.Text.ToString

        If RadioButton1.Checked = True Then
            taxDelVal = "True"
        ElseIf RadioButton2.Checked = True Then
            taxDelVal = "False"
        End If

        taxThis = String.Format("use pdqpos" & v &
        "go" & v &
        "Update tblconfigmain set optvalue = '{0}'" & v &
        "where optname = 'taxdeliverycharge'" & v &
        "Update tblmenuitemextend set taxdelcharge = '{0}'" & v &
        "Where TransTypeID In(3, 9)", taxDelVal)

        Form1.NoteStart(taxThis)
        Close()
    End Sub
    Sub FormAdjust(getString As String)
        Select Case getString
            Case = "Flat"
                Show()
                Text = "Flat Fee"
                Button1.Hide()
                Button2.Show()
                Button3.Hide()
                Button4.Hide()
                Button5.Hide()
                AcceptButton = Button2
                Label1.Show()
                Label2.Hide()
                Label3.Hide()
                Label4.Hide()
                Label5.Hide()
                RadioButton1.Hide()
                RadioButton2.Hide()
                RadioButton3.Hide()
                RadioButton4.Hide()
                TextBox1.Show()
                TextBox3.Hide()
                TextBox5.Hide()
                GroupBox1.Hide()
                GroupBox2.Hide()

                Label1.Text = "Flat Fee:"
                Label2.Location = New Point(25, 96)
                Label1.Location = New Point(107, 58)
                Label2.Location = New Point(25, 96)
                Button1.Location = New Point(94, 118)
                Button2.Location = New Point(94, 118)
                Button3.Location = New Point(94, 118)
                Button4.Location = New Point(94, 118)
                TextBox1.Location = New Point(82, 74)

            Case = "Build"
                Show()
                Label1.Show()
                Label2.Hide()
                Label3.Hide()
                Label4.Hide()
                Label5.Show()
                RadioButton1.Show()
                RadioButton2.Show()
                RadioButton3.Show()
                RadioButton4.Show()
                TextBox1.Show()
                TextBox3.Hide()
                TextBox5.Show()
                Button1.Hide()
                Button2.Hide()
                Button3.Hide()
                Button4.Show()
                Button5.Hide()
                GroupBox1.Show()
                GroupBox2.Show()

                RadioButton1.Checked = True
                RadioButton3.Checked = True

                Text = "Build MenuItemExtend"
                Label1.Text = "Delivery Cap:"
                Label2.Text = "Tax Delivery (Yes/No):"
                Label4.Text = "Tax Catering (Yes/No):"
                Label5.Text = "PIN:"
                Label1.Location = New Point(56, 15)
                Label2.Location = New Point(10, 73)
                Label4.Location = New Point(10, 91)
                Label5.Location = New Point(18, 135)
                TextBox1.Location = New Point(126, 12)
                GroupBox1.Location = New Point(94, 36)
                GroupBox2.Location = New Point(94, 81)
                'TextBox2.Location = New Point(124, 64)
                'TextBox4.Location = New Point(124, 89)
                TextBox5.Location = New Point(52, 132)
                Button1.Location = New Point(94, 130)
                Button2.Location = New Point(94, 130)
                Button3.Location = New Point(94, 130)
                Button4.Location = New Point(94, 130)

                AcceptButton = Button4

            Case = "EmpDisc"
                Show()
                Text = "Emp Discount"
                Label1.Text = "Discount:"
                Label1.Show()
                Label2.Show()
                Label2.Text = "Set percentage as a decimal (50% = .5)"
                Label3.Hide()
                Label4.Hide()
                Label5.Hide()
                RadioButton1.Hide()
                RadioButton2.Hide()
                RadioButton3.Hide()
                RadioButton4.Hide()
                TextBox1.Show()
                TextBox3.Hide()
                TextBox5.Hide()
                Button1.Hide()
                Button2.Hide()
                Button4.Hide()
                Button3.Show()
                Button5.Hide()
                GroupBox1.Hide()
                GroupBox2.Hide()
                Label2.Location = New Point(25, 96)
                Label1.Location = New Point(107, 58)
                Label2.Location = New Point(25, 96)
                Button1.Location = New Point(94, 118)
                Button2.Location = New Point(94, 118)
                Button3.Location = New Point(94, 118)
                Button4.Location = New Point(94, 118)
                TextBox1.Location = New Point(82, 74)
                AcceptButton = Button3

            Case = "Wheat"
                Show()
                Text = "Wheat"
                Label1.Text = "Price 1:"
                Label1.Location = New Point(107, 58)
                Label2.Location = New Point(107, 97)
                Label1.Show()
                Label2.Hide()
                Label3.Hide()
                Label4.Hide()
                Label5.Hide()
                RadioButton1.Hide()
                RadioButton2.Hide()
                RadioButton3.Hide()
                RadioButton4.Hide()
                TextBox1.Show()
                TextBox3.Hide()
                TextBox5.Hide()
                Button1.Show()
                Button2.Hide()
                Button3.Hide()
                Button4.Hide()
                Button5.Hide()
                GroupBox1.Hide()
                GroupBox2.Hide()

                Button1.Location = New Point(94, 118)
                Button2.Location = New Point(94, 118)
                Button3.Location = New Point(94, 118)
                Button4.Location = New Point(94, 118)
                TextBox1.Location = New Point(82, 74)


                AcceptButton = Button1
            Case = "TaxDelCap"
                Show()
                Text = "Tax Delivery Cap"
                Button1.hide()
                Button2.Hide()
                Button3.Hide()
                Button4.Hide()
                Button5.Show()
                RadioButton1.Show()
                RadioButton2.Show()
                GroupBox1.Show()
                GroupBox2.Hide()
                TextBox1.Hide()
                TextBox3.Hide()
                TextBox5.Hide()
                Label1.Hide()
                Label2.Hide()
                Label3.Hide()
                Label4.Hide()
                Label5.Hide()
                GroupBox1.Location = New Point(100, 52)
                Button5.Location = New Point(100, 112)
        End Select
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TaxDelCharge()
    End Sub
End Class