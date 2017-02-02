<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Wheat = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 26)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(141, 25)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Delivery Cap "
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 88)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(141, 22)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "Flat Delivery Fee"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(12, 116)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(141, 23)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "Remove Catering Groups"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(12, 145)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(141, 23)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "Build Menu Item Extend"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(12, 57)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(141, 25)
        Me.Button3.TabIndex = 8
        Me.Button3.Text = "Employee Discount"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button6.Location = New System.Drawing.Point(98, 174)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(75, 23)
        Me.Button6.TabIndex = 9
        Me.Button6.Text = "Done"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Wheat
        '
        Me.Wheat.Location = New System.Drawing.Point(159, 57)
        Me.Wheat.Name = "Wheat"
        Me.Wheat.Size = New System.Drawing.Size(114, 25)
        Me.Wheat.TabIndex = 10
        Me.Wheat.Text = "Wheat Leveling"
        Me.Wheat.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(160, 86)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(113, 23)
        Me.Button7.TabIndex = 11
        Me.Button7.Text = "Correct Item Levels"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button6
        Me.ClientSize = New System.Drawing.Size(285, 203)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Wheat)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(12, 45)
        Me.Name = "Form2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Miscellaneous Scripts"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Wheat As Button
    Friend WithEvents Button7 As Button
End Class
