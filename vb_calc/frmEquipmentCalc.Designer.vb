<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEquipmentCalc
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
        Me.Strength_TxtBx = New System.Windows.Forms.TextBox
        Me.Complexity_TxtBx = New System.Windows.Forms.TextBox
        Me.Prof_TxtBx = New System.Windows.Forms.TextBox
        Me.Clan_TxtBx = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.SU_TxtBx = New System.Windows.Forms.TextBox
        Me.GU_TxtBx = New System.Windows.Forms.TextBox
        Me.SUGU_TxtBx = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Repair_Btn = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.sType1_Combo = New System.Windows.Forms.ComboBox
        Me.cVal1_TxtBx = New System.Windows.Forms.TextBox
        Me.rStr1_TxtBx = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Enhance1_Btn = New System.Windows.Forms.Button
        Me.Enhance2_Btn = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.rStr2_TxtBx = New System.Windows.Forms.TextBox
        Me.cVal2_TxtBx = New System.Windows.Forms.TextBox
        Me.sType2_Combo = New System.Windows.Forms.ComboBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.eVal1_TxtBx = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Strength_TxtBx
        '
        Me.Strength_TxtBx.Location = New System.Drawing.Point(116, 66)
        Me.Strength_TxtBx.Name = "Strength_TxtBx"
        Me.Strength_TxtBx.Size = New System.Drawing.Size(100, 20)
        Me.Strength_TxtBx.TabIndex = 0
        Me.Strength_TxtBx.Text = "100"
        '
        'Complexity_TxtBx
        '
        Me.Complexity_TxtBx.Location = New System.Drawing.Point(116, 92)
        Me.Complexity_TxtBx.Name = "Complexity_TxtBx"
        Me.Complexity_TxtBx.Size = New System.Drawing.Size(100, 20)
        Me.Complexity_TxtBx.TabIndex = 1
        Me.Complexity_TxtBx.Text = "100"
        '
        'Prof_TxtBx
        '
        Me.Prof_TxtBx.Location = New System.Drawing.Point(177, 118)
        Me.Prof_TxtBx.Name = "Prof_TxtBx"
        Me.Prof_TxtBx.Size = New System.Drawing.Size(39, 20)
        Me.Prof_TxtBx.TabIndex = 2
        Me.Prof_TxtBx.Text = "0"
        '
        'Clan_TxtBx
        '
        Me.Clan_TxtBx.Location = New System.Drawing.Point(177, 144)
        Me.Clan_TxtBx.Name = "Clan_TxtBx"
        Me.Clan_TxtBx.Size = New System.Drawing.Size(39, 20)
        Me.Clan_TxtBx.TabIndex = 3
        Me.Clan_TxtBx.Text = "0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(221, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 20)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Repairing"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(55, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Strength"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(43, 95)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Complexity"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(17, 121)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(154, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Smith Profession (Max:25)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(12, 147)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(159, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Smith Clan Bonus (Max:20)"
        '
        'SU_TxtBx
        '
        Me.SU_TxtBx.Location = New System.Drawing.Point(419, 66)
        Me.SU_TxtBx.Name = "SU_TxtBx"
        Me.SU_TxtBx.ReadOnly = True
        Me.SU_TxtBx.Size = New System.Drawing.Size(100, 20)
        Me.SU_TxtBx.TabIndex = 9
        '
        'GU_TxtBx
        '
        Me.GU_TxtBx.Location = New System.Drawing.Point(419, 92)
        Me.GU_TxtBx.Name = "GU_TxtBx"
        Me.GU_TxtBx.ReadOnly = True
        Me.GU_TxtBx.Size = New System.Drawing.Size(100, 20)
        Me.GU_TxtBx.TabIndex = 10
        '
        'SUGU_TxtBx
        '
        Me.SUGU_TxtBx.Location = New System.Drawing.Point(419, 118)
        Me.SUGU_TxtBx.Name = "SUGU_TxtBx"
        Me.SUGU_TxtBx.ReadOnly = True
        Me.SUGU_TxtBx.Size = New System.Drawing.Size(100, 20)
        Me.SUGU_TxtBx.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(285, 69)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(128, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Required SU to 110%"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(284, 95)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(129, 13)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Required GU to 120%"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(222, 121)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(191, 13)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Required GU from 110% to 120%"
        '
        'Repair_Btn
        '
        Me.Repair_Btn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Repair_Btn.Location = New System.Drawing.Point(203, 184)
        Me.Repair_Btn.Name = "Repair_Btn"
        Me.Repair_Btn.Size = New System.Drawing.Size(115, 27)
        Me.Repair_Btn.TabIndex = 15
        Me.Repair_Btn.Text = "Calculate"
        Me.Repair_Btn.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(216, 237)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 20)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "Enhancing"
        '
        'sType1_Combo
        '
        Me.sType1_Combo.Cursor = System.Windows.Forms.Cursors.Default
        Me.sType1_Combo.FormattingEnabled = True
        Me.sType1_Combo.Items.AddRange(New Object() {"Single", "Double", "Triple"})
        Me.sType1_Combo.Location = New System.Drawing.Point(220, 283)
        Me.sType1_Combo.MaxDropDownItems = 3
        Me.sType1_Combo.Name = "sType1_Combo"
        Me.sType1_Combo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.sType1_Combo.Size = New System.Drawing.Size(90, 21)
        Me.sType1_Combo.TabIndex = 17
        Me.sType1_Combo.Text = "Single"
        '
        'cVal1_TxtBx
        '
        Me.cVal1_TxtBx.Location = New System.Drawing.Point(244, 310)
        Me.cVal1_TxtBx.Name = "cVal1_TxtBx"
        Me.cVal1_TxtBx.Size = New System.Drawing.Size(66, 20)
        Me.cVal1_TxtBx.TabIndex = 18
        Me.cVal1_TxtBx.Text = "1"
        '
        'rStr1_TxtBx
        '
        Me.rStr1_TxtBx.Location = New System.Drawing.Point(244, 336)
        Me.rStr1_TxtBx.Name = "rStr1_TxtBx"
        Me.rStr1_TxtBx.Size = New System.Drawing.Size(66, 20)
        Me.rStr1_TxtBx.TabIndex = 19
        Me.rStr1_TxtBx.Text = "500"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(154, 313)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(84, 13)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "Current Value"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(128, 339)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(110, 13)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "Strength Required"
        '
        'Enhance1_Btn
        '
        Me.Enhance1_Btn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Enhance1_Btn.Location = New System.Drawing.Point(203, 381)
        Me.Enhance1_Btn.Name = "Enhance1_Btn"
        Me.Enhance1_Btn.Size = New System.Drawing.Size(115, 27)
        Me.Enhance1_Btn.TabIndex = 22
        Me.Enhance1_Btn.Text = "Calculate"
        Me.Enhance1_Btn.UseVisualStyleBackColor = True
        '
        'Enhance2_Btn
        '
        Me.Enhance2_Btn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Enhance2_Btn.Location = New System.Drawing.Point(203, 573)
        Me.Enhance2_Btn.Name = "Enhance2_Btn"
        Me.Enhance2_Btn.Size = New System.Drawing.Size(115, 27)
        Me.Enhance2_Btn.TabIndex = 28
        Me.Enhance2_Btn.Text = "Calculate"
        Me.Enhance2_Btn.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(128, 531)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(110, 13)
        Me.Label12.TabIndex = 27
        Me.Label12.Text = "Strength Required"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(154, 479)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(84, 13)
        Me.Label13.TabIndex = 26
        Me.Label13.Text = "Current Value"
        '
        'rStr2_TxtBx
        '
        Me.rStr2_TxtBx.Location = New System.Drawing.Point(244, 528)
        Me.rStr2_TxtBx.Name = "rStr2_TxtBx"
        Me.rStr2_TxtBx.Size = New System.Drawing.Size(66, 20)
        Me.rStr2_TxtBx.TabIndex = 25
        Me.rStr2_TxtBx.Text = "500"
        '
        'cVal2_TxtBx
        '
        Me.cVal2_TxtBx.Location = New System.Drawing.Point(244, 476)
        Me.cVal2_TxtBx.Name = "cVal2_TxtBx"
        Me.cVal2_TxtBx.Size = New System.Drawing.Size(66, 20)
        Me.cVal2_TxtBx.TabIndex = 24
        Me.cVal2_TxtBx.Text = "1"
        '
        'sType2_Combo
        '
        Me.sType2_Combo.Cursor = System.Windows.Forms.Cursors.Default
        Me.sType2_Combo.FormattingEnabled = True
        Me.sType2_Combo.Items.AddRange(New Object() {"Single", "Double", "Triple"})
        Me.sType2_Combo.Location = New System.Drawing.Point(220, 449)
        Me.sType2_Combo.MaxDropDownItems = 3
        Me.sType2_Combo.Name = "sType2_Combo"
        Me.sType2_Combo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.sType2_Combo.Size = New System.Drawing.Size(90, 21)
        Me.sType2_Combo.TabIndex = 23
        Me.sType2_Combo.Text = "Single"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(173, 505)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(65, 13)
        Me.Label14.TabIndex = 30
        Me.Label14.Text = "End Value"
        '
        'eVal1_TxtBx
        '
        Me.eVal1_TxtBx.Location = New System.Drawing.Point(244, 502)
        Me.eVal1_TxtBx.Name = "eVal1_TxtBx"
        Me.eVal1_TxtBx.Size = New System.Drawing.Size(66, 20)
        Me.eVal1_TxtBx.TabIndex = 29
        Me.eVal1_TxtBx.Text = "2"
        '
        'frmEquipmentCalc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(537, 629)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.eVal1_TxtBx)
        Me.Controls.Add(Me.Enhance2_Btn)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.rStr2_TxtBx)
        Me.Controls.Add(Me.cVal2_TxtBx)
        Me.Controls.Add(Me.sType2_Combo)
        Me.Controls.Add(Me.Enhance1_Btn)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.rStr1_TxtBx)
        Me.Controls.Add(Me.cVal1_TxtBx)
        Me.Controls.Add(Me.sType1_Combo)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Repair_Btn)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.SUGU_TxtBx)
        Me.Controls.Add(Me.GU_TxtBx)
        Me.Controls.Add(Me.SU_TxtBx)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Clan_TxtBx)
        Me.Controls.Add(Me.Prof_TxtBx)
        Me.Controls.Add(Me.Complexity_TxtBx)
        Me.Controls.Add(Me.Strength_TxtBx)
        Me.Name = "frmEquipmentCalc"
        Me.Text = "Equipment Calc"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Strength_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents Complexity_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents Prof_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents Clan_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents SU_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents GU_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents SUGU_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Repair_Btn As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents sType1_Combo As System.Windows.Forms.ComboBox
    Friend WithEvents cVal1_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents rStr1_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Enhance1_Btn As System.Windows.Forms.Button
    Friend WithEvents Enhance2_Btn As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents rStr2_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents cVal2_TxtBx As System.Windows.Forms.TextBox
    Friend WithEvents sType2_Combo As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents eVal1_TxtBx As System.Windows.Forms.TextBox
End Class
