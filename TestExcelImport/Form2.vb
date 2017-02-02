Public Class Form2

    Inherits System.Windows.Forms.Form
    Public myCaller As Form1

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.FlatDeliveryFee()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Hide()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Wheat.Click
        Form3.Show()
        Form3.Text = "Wheat"
        Form3.Label1.Text = "Price 1:"
        Form3.Label1.Location = New Point(107, 58)
        Form3.Label2.Location = New Point(107, 97)
        Form3.Label2.Hide()
        Form3.Button1.Location = New Point(94, 118)
        Form3.Button2.Location = New Point(94, 118)
        Form3.Button3.Location = New Point(94, 118)
        Form3.Button4.Location = New Point(94, 118)
        Form3.TextBox1.Location = New Point(82, 74)
        Form3.TextBox2.Hide()
        Form3.TextBox3.Hide()
        Form3.TextBox4.Hide()
        Form3.Button1.Show()
        Form3.Button2.Hide()
        Form3.Button3.Hide()
        Form3.Button4.Hide()
        Form3.AcceptButton = Button1
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form3.Show()
        Form3.Text = "Emp Discount"

        Form3.Label1.Text = "Discount:"
        Form3.Label2.Show()
        Form3.Label2.Text = "Set percentage as a decimal (50% = .5)"
        Form3.Label2.Location = New Point(25, 96)
        Form3.Label1.Location = New Point(107, 58)
        Form3.Label2.Location = New Point(25, 96)
        Form3.Button1.Location = New Point(94, 118)
        Form3.Button2.Location = New Point(94, 118)
        Form3.Button3.Location = New Point(94, 118)
        Form3.Button4.Location = New Point(94, 118)
        Form3.TextBox1.Location = New Point(82, 74)

        Form3.TextBox2.Hide()
        Form3.TextBox3.Hide()
        Form3.TextBox4.Hide()

        Form3.Button1.Hide()
        Form3.Button2.Hide()
        Form3.Button4.Hide()
        Form3.Button3.Show()
        Form3.AcceptButton = Button3
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Form1.CorrectItemLevel()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class