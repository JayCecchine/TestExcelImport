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
        TextBox2.Hide()
        TextBox3.Hide()
        TextBox4.Hide()
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
        Dim catCapVal As String = ""
        Dim taxDelVal As String = TextBox3.text.ToString
        Dim Taxcatval As String = TextBox4.text.ToString

        If delCapVal.Length > 0 And taxDelVal.Length > 0 And Taxcatval.Length > 0 Then
            buildThis = String.Format("USE PDQPOS" & v &
            "GO" & v &
            "UPDATE tblconfigmain SET optvalue='False' WHERE optname='LegacyOrderCalc' AND optgroup='AdvOrderCalc'" & v &
            "UPDATE tblconfigmain SET optvalue='True' WHERE optname='AdvDelChargePerItem' AND optgroup='AdvOrderCalc'" & v &
            "UPDATE tblconfigmain SET optvalue='{0}' WHERE optname='AdvDelChargeMax' AND optgroup='AdvOrderCalc'" & v &
            "--UPDATE tblconfigmain SET optvalue='{1}' WHERE optname='AdvDelCaterChargeMax' AND" & v &
            "optgroup='AdvOrderCalc'" & v &
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
                TextBox2.Hide()
                TextBox3.Hide()
                TextBox4.Hide()
                AcceptButton = Button2
                Label2.Hide()
                Label1.Text = "Flat Fee:"
                Label2.Location = New Point(25, 96)
                Label1.Location = New Point(107, 58)
                Label2.Location = New Point(25, 96)
                Button1.Location = New Point(94, 118)
                Button2.Location = New Point(94, 118)
                Button3.Location = New Point(94, 118)
                Button4.Location = New Point(94, 118)
                TextBox1.Location = New Point(82, 74)
        End Select
    End Sub




End Class