Public Class frmEquipmentCalc

    Private Sub Repair_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Repair_Btn.Click
        Dim mSU, mGU, mSUGU As Double
        Dim vStr, vComp As Integer
        Dim bProf, bClan, bTotal As Double

        vStr = Int(Strength_TxtBx.Text)
        vComp = Int(Complexity_TxtBx.Text)

        bProf = Int(Prof_TxtBx.Text) / 100
        bClan = Int(Clan_TxtBx.Text) / 100

        bTotal = 1 - bProf - bClan

        mSU = vComp * 1.1
        mGU = vComp * 1.2

        If vStr <= mSU Then mSUGU = mGU - mSU Else mSUGU = mGU - vStr

        SU_TxtBx.Text = Int((mSU - vStr) * bTotal)
        GU_TxtBx.Text = Int((mGU - vStr) * bTotal)

        SUGU_TxtBx.Text = Int(mSUGU * bTotal)

    End Sub

    Private Sub Enhance1_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Enhance1_Btn.Click
        Select Case sType1_Combo.Text
            Case "Single"
                rStr1_TxtBx.Text = (Int(cVal1_TxtBx.Text) * 250) + 250
            Case "Double"
                rStr1_TxtBx.Text = (Int(cVal1_TxtBx.Text) * 1000) + 1000
            Case "Triple"
                rStr1_TxtBx.Text = (Int(cVal1_TxtBx.Text) * 3000) + 3000
        End Select
    End Sub

    Private Sub Enhance2_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Enhance2_Btn.Click
        Dim x, y As Integer
        Dim sVal As Integer

        x = Int(cVal2_TxtBx.Text)
        y = Int(eVal1_TxtBx.Text)

        Select Case sType2_Combo.Text
            Case "Single"
                Do While x < y
                    sVal = sVal + ((250 * x) + 250)
                    x = x + 1
                Loop
            Case "Double"
                Do While x < y
                    sVal = sVal + ((1000 * x) + 1000)
                    x = x + 1
                Loop
            Case "Triple"
                Do While x < y
                    sVal = sVal + ((3000 * x) + 3000)
                    x = x + 1
                Loop
        End Select

        rStr2_TxtBx.Text = sVal

    End Sub
End Class