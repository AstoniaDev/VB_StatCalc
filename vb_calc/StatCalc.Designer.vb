<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStatCalc
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
        Me.WarCalc_Btn = New System.Windows.Forms.Button()
        Me.MageCalc_Btn = New System.Windows.Forms.Button()
        Me.SeyanCalc_Btn = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'WarCalc_Btn
        '
        Me.WarCalc_Btn.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.WarCalc_Btn.Location = New System.Drawing.Point(12, 29)
        Me.WarCalc_Btn.Name = "WarCalc_Btn"
        Me.WarCalc_Btn.Size = New System.Drawing.Size(117, 37)
        Me.WarCalc_Btn.TabIndex = 0
        Me.WarCalc_Btn.Text = "Warrior Calc"
        Me.WarCalc_Btn.UseVisualStyleBackColor = True
        '
        'MageCalc_Btn
        '
        Me.MageCalc_Btn.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.MageCalc_Btn.Location = New System.Drawing.Point(135, 29)
        Me.MageCalc_Btn.Name = "MageCalc_Btn"
        Me.MageCalc_Btn.Size = New System.Drawing.Size(117, 37)
        Me.MageCalc_Btn.TabIndex = 1
        Me.MageCalc_Btn.Text = "Mage Calc"
        Me.MageCalc_Btn.UseVisualStyleBackColor = True
        '
        'SeyanCalc_Btn
        '
        Me.SeyanCalc_Btn.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.SeyanCalc_Btn.Location = New System.Drawing.Point(258, 29)
        Me.SeyanCalc_Btn.Name = "SeyanCalc_Btn"
        Me.SeyanCalc_Btn.Size = New System.Drawing.Size(117, 37)
        Me.SeyanCalc_Btn.TabIndex = 2
        Me.SeyanCalc_Btn.Text = "Seyan Calc"
        Me.SeyanCalc_Btn.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(15, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(360, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Old World Stat Calculators"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(9, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(366, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Originally made by Slayin and modified by Whitestar."
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial Narrow", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.Label5.Location = New System.Drawing.Point(332, 82)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 14)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Ver 1.0.001"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(12, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(363, 14)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Modified for use in Old World by Valadez"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmStatCalc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(387, 99)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SeyanCalc_Btn)
        Me.Controls.Add(Me.MageCalc_Btn)
        Me.Controls.Add(Me.WarCalc_Btn)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(1000, 1000)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(16, 38)
        Me.Name = "frmStatCalc"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stat Calculator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents WarCalc_Btn As System.Windows.Forms.Button
    Friend WithEvents MageCalc_Btn As System.Windows.Forms.Button
    Friend WithEvents SeyanCalc_Btn As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
