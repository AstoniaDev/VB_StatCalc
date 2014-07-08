Public Class frmStatCalc

    Private Sub WarCalc_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WarCalc_Btn.Click
        frmWarriorCalc.Show()
    End Sub

    Private Sub MageCalc_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MageCalc_Btn.Click
        frmMageCalc.Show()
    End Sub

    Private Sub SeyanCalc_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SeyanCalc_Btn.Click
        frmSeyanCalc.Show()
    End Sub

    Private Sub EquipCalc_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmEquipmentCalc.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmNPC_Calc.Show()
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub
End Class
