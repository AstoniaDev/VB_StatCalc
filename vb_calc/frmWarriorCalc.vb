Imports System.IO

Public Class frmWarriorCalc

    Public gLevel As Double
    Private Sub SetExpValues()
        Dim sb As Integer
        Dim sf As Integer
        Dim cf As Double
        Dim ExpReq As Integer

        sb = 25
        sf = 3
        cf = 0.1

        ExpReq = Int((((Int(HP_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        HP_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(End_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        End_ExpLbl.Text = ExpReq

        sb = 10
        sf = 2

        ExpReq = Int((((Int(Wis_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Wis_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Int_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Int_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Agi_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Agi_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Str_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Str_ExpLbl.Text = ExpReq

        sb = 1
        sf = 1

        ExpReq = Int((((Int(Dagger_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Dagger_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(H2H_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        H2H_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Sword_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Sword_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(TwoHand_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        TwoHand_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Rage_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Rage_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Attack_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Attack_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Parry_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Parry_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Tactics_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Tactics_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Warcry_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Warcry_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Surround_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Surround_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Speed_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Speed_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Immunity_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Immunity_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Barter_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Barter_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Perc_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Perc_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Stealth_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Stealth_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Regen_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Regen_ExpLbl.Text = ExpReq

        sf = 3

        ExpReq = Int((((Int(Prof_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Prof_ExpLbl.Text = ExpReq

        SetMods()

    End Sub

    Private Sub SetMods()
        Dim WisMod, IntMod, AgiMod, StrMod As Integer
        Dim AAS, IAS, SSA, SSI, WIS, IIW, AAI, SSS As Integer
        Dim MaxStatVal, StatVal, MaxEquipVal, EquipVal, ModVal As Integer
        Dim TactBonus, RageBonus As Integer

        If ShowTact_Chk.Checked = True Then TactBonus = Int(TactBonus_TxtBx.Text) Else TactBonus = 0
        If ShowRage_Chk.Checked = True Then RageBonus = Int(RageBonus_TxtBx.Text) Else RageBonus = 0

        MaxEquipVal = Math.Floor(Int(HP_TxtBx.Text) / 2)
        EquipVal = Int(HP_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        HP_Mod_Lbl.Text = Int(HP_TxtBx.Text) + ModVal

        MaxEquipVal = Math.Floor(Int(End_TxtBx.Text) / 2)
        EquipVal = Int(End_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        End_Mod_Lbl.Text = Int(End_TxtBx.Text) + ModVal

        MaxEquipVal = Math.Floor(Int(Wis_TxtBx.Text) / 2)
        EquipVal = Int(Wis_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        Wis_Mod_Lbl.Text = Int(Wis_TxtBx.Text) + ModVal

        MaxEquipVal = Math.Floor(Int(Int_TxtBx.Text) / 2)
        EquipVal = Int(Int_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        Int_Mod_Lbl.Text = Int(Int_TxtBx.Text) + ModVal

        MaxEquipVal = Math.Floor(Int(Agi_TxtBx.Text) / 2)
        EquipVal = Int(Agi_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        Agi_Mod_Lbl.Text = Int(Agi_TxtBx.Text) + ModVal

        MaxEquipVal = Math.Floor(Int(Str_TxtBx.Text) / 2)
        EquipVal = Int(Str_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        Str_Mod_Lbl.Text = Int(Str_TxtBx.Text) + ModVal

        WisMod = Int(Wis_Mod_Lbl.Text)
        IntMod = Int(Int_Mod_Lbl.Text)
        AgiMod = Int(Agi_Mod_Lbl.Text)
        StrMod = Int(Str_Mod_Lbl.Text)

        AAS = Math.Floor((AgiMod + AgiMod + StrMod) / 5)
        IAS = Math.Floor((IntMod + AgiMod + StrMod) / 5)
        SSA = Math.Floor((StrMod + StrMod + AgiMod) / 5)
        SSI = Math.Floor((StrMod + StrMod + IntMod) / 5)
        WIS = Math.Floor((WisMod + IntMod + StrMod) / 5)
        IIW = Math.Floor((IntMod + IntMod + WisMod) / 5)
        AAI = Math.Floor((AgiMod + AgiMod + IntMod) / 5)
        SSS = Math.Floor((StrMod + StrMod + StrMod) / 5)

        If Int(Dagger_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Dagger_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Dagger_Mod_TxtBx.Text) > Int(Int(Dagger_TxtBx.Text) / 2) Then EquipVal = Int(Int(Dagger_TxtBx.Text) / 2) Else EquipVal = Int(Dagger_Mod_TxtBx.Text)
        ModVal = Int(Dagger_TxtBx.Text) + StatVal + EquipVal
        Dagger_Mod_Lbl.Text = ModVal + TactBonus

        If Int(H2H_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(H2H_TxtBx.Text) * 2
        If SSA > MaxStatVal Then StatVal = MaxStatVal Else StatVal = SSA
        If Int(H2H_Mod_TxtBx.Text) > Int(Int(H2H_TxtBx.Text) / 2) Then EquipVal = Int(Int(H2H_TxtBx.Text) / 2) Else EquipVal = Int(H2H_Mod_TxtBx.Text)
        ModVal = Int(H2H_TxtBx.Text) + StatVal + EquipVal
        H2H_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Sword_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Sword_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Sword_Mod_TxtBx.Text) > Int(Int(Sword_TxtBx.Text) / 2) Then EquipVal = Int(Int(Sword_TxtBx.Text) / 2) Else EquipVal = Int(Sword_Mod_TxtBx.Text)
        ModVal = Int(Sword_TxtBx.Text) + StatVal + EquipVal
        Sword_Mod_Lbl.Text = ModVal + TactBonus


        If Int(TwoHand_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(TwoHand_TxtBx.Text) * 2
        If SSA > MaxStatVal Then StatVal = MaxStatVal Else StatVal = SSA
        If Int(TwoHand_Mod_TxtBx.Text) > Int(Int(TwoHand_TxtBx.Text) / 2) Then EquipVal = Int(Int(TwoHand_TxtBx.Text) / 2) Else EquipVal = Int(TwoHand_Mod_TxtBx.Text)
        ModVal = Int(TwoHand_TxtBx.Text) + StatVal + EquipVal
        TwoHand_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Rage_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Rage_TxtBx.Text) * 2
        If SSI > MaxStatVal Then StatVal = MaxStatVal Else StatVal = SSI
        If Int(Rage_Mod_TxtBx.Text) > Int(Int(Rage_TxtBx.Text) / 2) Then EquipVal = Int(Int(Rage_TxtBx.Text) / 2) Else EquipVal = Int(Rage_Mod_TxtBx.Text)
        ModVal = Int(Rage_TxtBx.Text) + StatVal + EquipVal
        Rage_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Attack_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Attack_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Attack_Mod_TxtBx.Text) > Int(Int(Attack_TxtBx.Text) / 2) Then EquipVal = Int(Int(Attack_TxtBx.Text) / 2) Else EquipVal = Int(Attack_Mod_TxtBx.Text)
        ModVal = Int(Attack_TxtBx.Text) + StatVal + EquipVal
        Attack_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Parry_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Parry_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Parry_Mod_TxtBx.Text) > Int(Int(Parry_TxtBx.Text) / 2) Then EquipVal = Int(Int(Parry_TxtBx.Text) / 2) Else EquipVal = Int(Parry_Mod_TxtBx.Text)
        ModVal = Int(Parry_TxtBx.Text) + StatVal + EquipVal
        Parry_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Tactics_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Tactics_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Tactics_Mod_TxtBx.Text) > Int(Int(Tactics_TxtBx.Text) / 2) Then EquipVal = Int(Int(Tactics_TxtBx.Text) / 2) Else EquipVal = Int(Tactics_Mod_TxtBx.Text)
        ModVal = Int(Tactics_TxtBx.Text) + StatVal + EquipVal
        Tactics_Mod_Lbl.Text = ModVal


        If Int(Warcry_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Warcry_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Warcry_Mod_TxtBx.Text) > Int(Int(Warcry_TxtBx.Text) / 2) Then EquipVal = Int(Int(Warcry_TxtBx.Text) / 2) Else EquipVal = Int(Warcry_Mod_TxtBx.Text)
        ModVal = Int(Warcry_TxtBx.Text) + StatVal + EquipVal
        Warcry_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Surround_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Surround_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Surround_Mod_TxtBx.Text) > Int(Int(Surround_TxtBx.Text) / 2) Then EquipVal = Int(Int(Surround_TxtBx.Text) / 2) Else EquipVal = Int(Surround_Mod_TxtBx.Text)
        ModVal = Int(Surround_TxtBx.Text) + StatVal + EquipVal
        Surround_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Speed_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Speed_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Speed_Mod_TxtBx.Text) > Int(Int(Speed_TxtBx.Text) / 2) Then EquipVal = Int(Int(Speed_TxtBx.Text) / 2) Else EquipVal = Int(Speed_Mod_TxtBx.Text)
        ModVal = Int(Speed_TxtBx.Text) + StatVal + EquipVal
        Speed_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Immunity_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Immunity_TxtBx.Text) * 2
        If WIS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = WIS
        If Int(Immunity_Mod_TxtBx.Text) > Int(Int(Immunity_TxtBx.Text) / 2) Then EquipVal = Int(Int(Immunity_TxtBx.Text) / 2) Else EquipVal = Int(Immunity_Mod_TxtBx.Text)
        ModVal = Int(Immunity_TxtBx.Text) + StatVal + EquipVal
        Immunity_Mod_Lbl.Text = ModVal + TactBonus + RageBonus


        If Int(Barter_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Barter_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Barter_Mod_TxtBx.Text) > Int(Int(Barter_TxtBx.Text) / 2) Then EquipVal = Int(Int(Barter_TxtBx.Text) / 2) Else EquipVal = Int(Barter_Mod_TxtBx.Text)
        ModVal = Int(Barter_TxtBx.Text) + StatVal + EquipVal
        Barter_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Perc_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Perc_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Perc_Mod_TxtBx.Text) > Int(Int(Perc_TxtBx.Text) / 2) Then EquipVal = Int(Int(Perc_TxtBx.Text) / 2) Else EquipVal = Int(Perc_Mod_TxtBx.Text)
        ModVal = Int(Perc_TxtBx.Text) + StatVal + EquipVal
        Perc_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Stealth_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Stealth_TxtBx.Text) * 2
        If AAI > MaxStatVal Then StatVal = MaxStatVal Else StatVal = AAI
        If Int(Stealth_Mod_TxtBx.Text) > Int(Int(Stealth_TxtBx.Text) / 2) Then EquipVal = Int(Int(Stealth_TxtBx.Text) / 2) Else EquipVal = Int(Stealth_Mod_TxtBx.Text)
        ModVal = Int(Stealth_TxtBx.Text) + StatVal + EquipVal
        Stealth_Mod_Lbl.Text = ModVal + TactBonus


        If Int(Regen_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Regen_TxtBx.Text) * 2
        If SSS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = SSS
        If Int(Regen_Mod_TxtBx.Text) > Int(Int(Regen_TxtBx.Text) / 2) Then EquipVal = Int(Int(Regen_TxtBx.Text) / 2) Else EquipVal = Int(Regen_Mod_TxtBx.Text)
        ModVal = Int(Regen_TxtBx.Text) + StatVal + EquipVal
        Regen_Mod_Lbl.Text = ModVal + TactBonus

        SetLevel()

    End Sub

    Private Sub SetLevel()
        Dim sVal As Integer
        Dim TotalExp, ReqLevel As Double
        Dim sb1, sb2, sb3, sb4 As Integer 'Base Values
        Dim sf1, sf2, sf3, sf4 As Integer 'Scale Factors
        Dim cf1, cf2, cf3 As Double 'Class Factors
        'Set Base Values for Stats
        sb1 = 25
        sb2 = 10
        sb3 = 1
        sb4 = 1
        'Set Scale Factors for groups
        sf1 = 3
        sf2 = 2
        sf3 = 1
        sf4 = 3
        'Set Class Factor Values
        cf1 = 0.1333333333333
        cf2 = 0.1
        cf3 = 0.1

        sVal = 25
        Do Until sVal = Int(HP_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb1) + 6) ^ 3) * (sf1 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 25
        Do Until sVal = Int(End_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb1) + 6) ^ 3) * (sf1 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 10
        Do Until sVal = Int(Wis_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb2) + 6) ^ 3) * (sf2 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 10
        Do Until sVal = Int(Int_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb2) + 6) ^ 3) * (sf2 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 10
        Do Until sVal = Int(Agi_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb2) + 6) ^ 3) * (sf2 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 10
        Do Until sVal = Int(Str_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb2) + 6) ^ 3) * (sf2 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Dagger_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(H2H_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Sword_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(TwoHand_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Rage_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Attack_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Parry_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Tactics_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Warcry_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Surround_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Speed_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Immunity_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Barter_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Perc_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Stealth_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Regen_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf2))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Prof_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb4) + 6) ^ 3) * (sf4 * cf2))
            sVal = sVal + 1
        Loop

        ReqLevel = TotalExp ^ 0.25
        If ReqLevel = 0 Then Level_TxtBx.Text = 1 Else Level_TxtBx.Text = ReqLevel

        'Whitestar Mod
        gLevel = ReqLevel
        SetExtraValues()

    End Sub

    Private Sub SetExtraValues()
        Dim wv_add, x As Integer
        Dim wv As Double
        Dim ws1, ws2, ws_max As Integer
        Dim AAS_VAL As Double
        Dim RageBonus As Integer

        wv = 0
        wv_add = 1
        x = 0

        Do While x < Int(Str_Mod_Lbl.Text)

            If wv_add = 1 Then
                wv = wv + 0.1
                wv_add = 0
            Else
                wv = wv + 0.15
                wv_add = 1
            End If

            x = x + 1
        Loop

        wv = ((Math.Floor(Int(Str_Mod_Lbl.Text) / 2) * 0.15) + (Math.Ceiling(Int(Str_Mod_Lbl.Text) / 2) * 0.1))

        wv = wv + CDbl(WWV_TxtBx.Text)

        WV_TxtBx.Text = wv

        'TactBonus_TxtBx.Text = Int(Int(Tactics_Mod_Lbl.Text) * 0.15)
        'NxtTact.Text = Int(((Int(TactBonus_TxtBx.Text) + 1) / 0.15) + 0.99)
        ' Calculate tactics 
        'TactBonus_TxtBx.Text = Int(Int(Tactics_Mod_Lbl.Text) * 0.15)
        TactBonus_TxtBx.Text = Int(Int(Tactics_Mod_Lbl.Text) / 6.667)
        'NxtTact.Text = Int(((Int(TactBonus_TxtBx.Text) + 1) / 0.15) + 0.99)
        NxtTact.Text = Int(((Int(TactBonus_TxtBx.Text) + 1) * 6.667) + 0.99)

        ws1 = Math.Max(Int(Sword_Mod_Lbl.Text), Int(TwoHand_Mod_Lbl.Text))
        ws2 = Math.Max(Int(Dagger_Mod_Lbl.Text), Int(H2H_Mod_Lbl.Text))
        ws_max = Math.Max(ws1, ws2)

        AAS_VAL = ((Int(Agi_Mod_Lbl.Text) * 2) + Int(Str_Mod_Lbl.Text)) / 5

        RageBonus = Math.Floor(Int(Int(Rage_Mod_Lbl.Text) / 8))

        RageBonus_TxtBx.Text = RageBonus
        NxtRage.Text = Int((RageBonus + 1) * 8)

        If ShowTact_Chk.Checked = True Then
            Offense_TxtBx.Text = Int((Int(Attack_Mod_Lbl.Text) * 2) + Int(ws_max))
            Defense_TxtBx.Text = Int((Int(Parry_Mod_Lbl.Text) * 2) + Int(ws_max)) - CDbl(DefReduct_TxtBx.Text)
            RawSpeed_TxtBx.Text = Int(Int(AAS_VAL) + Int(((Int(Speed_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)) / 2) + Int(Int(TactBonus_TxtBx.Text) / 2))) - CDbl(SpeedReduct_TxtBx.Text)

            RageBonus = Int(Int(Rage_Mod_Lbl.Text) / 8)
        Else
            Offense_TxtBx.Text = Int(((Int(Attack_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)) * 2) + Int((ws_max + Int(TactBonus_TxtBx.Text))))
            Defense_TxtBx.Text = Int(((Int(Parry_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)) * 2) + Int((ws_max + Int(TactBonus_TxtBx.Text)))) - CDbl(DefReduct_TxtBx.Text)
            RawSpeed_TxtBx.Text = Int(Int(AAS_VAL) + Int((Int(Speed_Mod_Lbl.Text) / 2) + Int(Int(TactBonus_TxtBx.Text) / 2))) - CDbl(SpeedReduct_TxtBx.Text)

            RageBonus = Int((Int(Rage_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)) / 8)
        End If

        RageBonus_TxtBx.Text = RageBonus
        NxtRage.Text = Int((RageBonus + 1) * 8)

        If ShowRage_Chk.Checked = True Then Offense_TxtBx.Text = Int(Offense_TxtBx.Text) + RageBonus
        If ShowRage_Chk.Checked = True Then Defense_TxtBx.Text = Int(Defense_TxtBx.Text) + RageBonus

        ValidateData()

    End Sub

    Private Sub ValidateData()
        If Int(HP_Mod_TxtBx.Text) > Math.Floor(Int(HP_TxtBx.Text) / 2) Then HP_Mod_TxtBx.ForeColor = Color.Red Else HP_Mod_TxtBx.ForeColor = Color.Black
        If Int(End_Mod_TxtBx.Text) > Math.Floor(Int(End_TxtBx.Text) / 2) Then End_Mod_TxtBx.ForeColor = Color.Red Else End_Mod_TxtBx.ForeColor = Color.Black
        If Int(Wis_Mod_TxtBx.Text) > Math.Floor(Int(Wis_TxtBx.Text) / 2) Then Wis_Mod_TxtBx.ForeColor = Color.Red Else Wis_Mod_TxtBx.ForeColor = Color.Black
        If Int(Int_Mod_TxtBx.Text) > Math.Floor(Int(Int_TxtBx.Text) / 2) Then Int_Mod_TxtBx.ForeColor = Color.Red Else Int_Mod_TxtBx.ForeColor = Color.Black
        If Int(Agi_Mod_TxtBx.Text) > Math.Floor(Int(Agi_TxtBx.Text) / 2) Then Agi_Mod_TxtBx.ForeColor = Color.Red Else Agi_Mod_TxtBx.ForeColor = Color.Black
        If Int(Str_Mod_TxtBx.Text) > Math.Floor(Int(Str_TxtBx.Text) / 2) Then Str_Mod_TxtBx.ForeColor = Color.Red Else Str_Mod_TxtBx.ForeColor = Color.Black
        If Int(Dagger_Mod_TxtBx.Text) > Math.Floor(Int(Dagger_TxtBx.Text) / 2) Then Dagger_Mod_TxtBx.ForeColor = Color.Red Else Dagger_Mod_TxtBx.ForeColor = Color.Black
        If Int(H2H_Mod_TxtBx.Text) > Math.Floor(Int(H2H_TxtBx.Text) / 2) Then H2H_Mod_TxtBx.ForeColor = Color.Red Else H2H_Mod_TxtBx.ForeColor = Color.Black
        If Int(Sword_Mod_TxtBx.Text) > Math.Floor(Int(Sword_TxtBx.Text) / 2) Then Sword_Mod_TxtBx.ForeColor = Color.Red Else Sword_Mod_TxtBx.ForeColor = Color.Black
        If Int(TwoHand_Mod_TxtBx.Text) > Math.Floor(Int(TwoHand_TxtBx.Text) / 2) Then TwoHand_Mod_TxtBx.ForeColor = Color.Red Else TwoHand_Mod_TxtBx.ForeColor = Color.Black
        If Int(Rage_Mod_TxtBx.Text) > Math.Floor(Int(Rage_TxtBx.Text) / 2) Then Rage_Mod_TxtBx.ForeColor = Color.Red Else Rage_Mod_TxtBx.ForeColor = Color.Black
        If Int(Attack_Mod_TxtBx.Text) > Math.Floor(Int(Attack_TxtBx.Text) / 2) Then Attack_Mod_TxtBx.ForeColor = Color.Red Else Attack_Mod_TxtBx.ForeColor = Color.Black
        If Int(Parry_Mod_TxtBx.Text) > Math.Floor(Int(Parry_TxtBx.Text) / 2) Then Parry_Mod_TxtBx.ForeColor = Color.Red Else Parry_Mod_TxtBx.ForeColor = Color.Black
        If Int(Tactics_Mod_TxtBx.Text) > Math.Floor(Int(Tactics_TxtBx.Text) / 2) Then Tactics_Mod_TxtBx.ForeColor = Color.Red Else Tactics_Mod_TxtBx.ForeColor = Color.Black
        If Int(Warcry_Mod_TxtBx.Text) > Math.Floor(Int(Warcry_TxtBx.Text) / 2) Then Warcry_Mod_TxtBx.ForeColor = Color.Red Else Warcry_Mod_TxtBx.ForeColor = Color.Black
        If Int(Surround_Mod_TxtBx.Text) > Math.Floor(Int(Surround_TxtBx.Text) / 2) Then Surround_Mod_TxtBx.ForeColor = Color.Red Else Surround_Mod_TxtBx.ForeColor = Color.Black
        If Int(Speed_Mod_TxtBx.Text) > Math.Floor(Int(Speed_TxtBx.Text) / 2) Then Speed_Mod_TxtBx.ForeColor = Color.Red Else Speed_Mod_TxtBx.ForeColor = Color.Black
        If Int(Immunity_Mod_TxtBx.Text) > Math.Floor(Int(Immunity_TxtBx.Text) / 2) Then Immunity_Mod_TxtBx.ForeColor = Color.Red Else Immunity_Mod_TxtBx.ForeColor = Color.Black
        If Int(Barter_Mod_TxtBx.Text) > Math.Floor(Int(Barter_TxtBx.Text) / 2) Then Barter_Mod_TxtBx.ForeColor = Color.Red Else Barter_Mod_TxtBx.ForeColor = Color.Black
        If Int(Perc_Mod_TxtBx.Text) > Math.Floor(Int(Perc_TxtBx.Text) / 2) Then Perc_Mod_TxtBx.ForeColor = Color.Red Else Perc_Mod_TxtBx.ForeColor = Color.Black
        If Int(Stealth_Mod_TxtBx.Text) > Math.Floor(Int(Stealth_TxtBx.Text) / 2) Then Stealth_Mod_TxtBx.ForeColor = Color.Red Else Stealth_Mod_TxtBx.ForeColor = Color.Black
        If Int(Regen_Mod_TxtBx.Text) > Math.Floor(Int(Regen_TxtBx.Text) / 2) Then Regen_Mod_TxtBx.ForeColor = Color.Red Else Regen_Mod_TxtBx.ForeColor = Color.Black
        'Whitestar Mod
        If gLevel > Int(MaxLevel_txtbox.Text) + 1 Then
            MsgBox("The last stat increase exceeded your MaxLevel Setting")
        End If


    End Sub

   
    Private Sub HP_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles HP_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(HP_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                HP_TxtBx.Text = Int(HP_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(HP_TxtBx.Text) = 25 Then
                    MsgBox("Can not reduce below 25")
                    Exit Sub
                End If

                HP_TxtBx.Text = Int(HP_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub End_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles End_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(End_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                End_TxtBx.Text = Int(End_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(End_TxtBx.Text) = 25 Then
                    MsgBox("Can not reduce below 25")
                    Exit Sub
                End If

                End_TxtBx.Text = Int(End_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Wis_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Wis_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Wis_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Wis_TxtBx.Text = Int(Wis_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Wis_TxtBx.Text) = 10 Then
                    MsgBox("Can not reduce below 10")
                    Exit Sub
                End If

                Wis_TxtBx.Text = Int(Wis_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Int_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Int_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Int_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Int_TxtBx.Text = Int(Int_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Int_TxtBx.Text) = 10 Then
                    MsgBox("Can not reduce below 10")
                    Exit Sub
                End If

                Int_TxtBx.Text = Int(Int_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Agi_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Agi_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Agi_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Agi_TxtBx.Text = Int(Agi_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Agi_TxtBx.Text) = 10 Then
                    MsgBox("Can not reduce below 10")
                    Exit Sub
                End If

                Agi_TxtBx.Text = Int(Agi_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Str_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Str_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Str_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Str_TxtBx.Text = Int(Str_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Str_TxtBx.Text) = 10 Then
                    MsgBox("Can not reduce below 10")
                    Exit Sub
                End If

                Str_TxtBx.Text = Int(Str_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Dagger_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Dagger_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Dagger_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Dagger_TxtBx.Text = Int(Dagger_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Dagger_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Dagger_TxtBx.Text = Int(Dagger_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub H2H_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles H2H_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(H2H_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                H2H_TxtBx.Text = Int(H2H_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(H2H_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                H2H_TxtBx.Text = Int(H2H_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Sword_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Sword_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Sword_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Sword_TxtBx.Text = Int(Sword_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Sword_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Sword_TxtBx.Text = Int(Sword_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub TwoHand_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TwoHand_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(TwoHand_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                TwoHand_TxtBx.Text = Int(TwoHand_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(TwoHand_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                TwoHand_TxtBx.Text = Int(TwoHand_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Rage_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Rage_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Rage_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Rage_TxtBx.Text = Int(Rage_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Rage_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Rage_TxtBx.Text = Int(Rage_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Attack_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Attack_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Attack_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Attack_TxtBx.Text = Int(Attack_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Attack_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Attack_TxtBx.Text = Int(Attack_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Parry_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Parry_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Parry_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Parry_TxtBx.Text = Int(Parry_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Parry_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Parry_TxtBx.Text = Int(Parry_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Tactics_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Tactics_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Tactics_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Tactics_TxtBx.Text = Int(Tactics_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Tactics_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Tactics_TxtBx.Text = Int(Tactics_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Warcry_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Warcry_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Warcry_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Warcry_TxtBx.Text = Int(Warcry_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Warcry_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Warcry_TxtBx.Text = Int(Warcry_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Surround_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Surround_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Surround_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Surround_TxtBx.Text = Int(Surround_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Surround_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Surround_TxtBx.Text = Int(Surround_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Speed_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Speed_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Speed_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Speed_TxtBx.Text = Int(Speed_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Speed_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Speed_TxtBx.Text = Int(Speed_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Immunity_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Immunity_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Immunity_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Immunity_TxtBx.Text = Int(Immunity_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Immunity_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Immunity_TxtBx.Text = Int(Immunity_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Barter_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Barter_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Barter_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Barter_TxtBx.Text = Int(Barter_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Barter_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Barter_TxtBx.Text = Int(Barter_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Perc_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Perc_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Perc_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Perc_TxtBx.Text = Int(Perc_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Perc_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Perc_TxtBx.Text = Int(Perc_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Stealth_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Stealth_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Stealth_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Stealth_TxtBx.Text = Int(Stealth_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Stealth_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Stealth_TxtBx.Text = Int(Stealth_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Regen_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Regen_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Regen_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Regen_TxtBx.Text = Int(Regen_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Regen_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Regen_TxtBx.Text = Int(Regen_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub Prof_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Prof_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Prof_TxtBx.Text) = 50 Then
                    MsgBox("Can not raise beyond 50")
                    Exit Sub
                End If

                Prof_TxtBx.Text = Int(Prof_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Prof_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Prof_TxtBx.Text = Int(Prof_TxtBx.Text) - 1

        End Select

        SetExpValues()

    End Sub

    Private Sub HP_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HP_TxtBx.Leave
        If Int(HP_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then HP_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(HP_TxtBx.Text) < 25 Then HP_TxtBx.Text = 25

        SetExpValues()
    End Sub

    Private Sub End_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles End_TxtBx.Leave
        If Int(End_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then End_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(End_TxtBx.Text) < 25 Then End_TxtBx.Text = 25

        SetExpValues()
    End Sub

    Private Sub Wis_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Wis_TxtBx.Leave
        If Int(Wis_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Wis_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Wis_TxtBx.Text) < 10 Then Wis_TxtBx.Text = 10

        SetExpValues()
    End Sub

    Private Sub Int_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Int_TxtBx.Leave
        If Int(Int_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Int_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Int_TxtBx.Text) < 10 Then Int_TxtBx.Text = 10

        SetExpValues()
    End Sub

    Private Sub Agi_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Agi_TxtBx.Leave
        If Int(Agi_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Agi_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Agi_TxtBx.Text) < 10 Then Agi_TxtBx.Text = 10

        SetExpValues()
    End Sub

    Private Sub Str_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Str_TxtBx.Leave
        If Int(Str_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Str_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Str_TxtBx.Text) < 10 Then Str_TxtBx.Text = 10

        SetExpValues()
    End Sub

    Private Sub Dagger_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dagger_TxtBx.Leave
        If Int(Dagger_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Dagger_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Dagger_TxtBx.Text) < 1 Then Dagger_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub H2H_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles H2H_TxtBx.Leave
        If Int(H2H_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then H2H_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(H2H_TxtBx.Text) < 1 Then H2H_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Sword_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sword_TxtBx.Leave
        If Int(Sword_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Sword_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Sword_TxtBx.Text) < 1 Then Sword_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub TwoHand_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TwoHand_TxtBx.Leave
        If Int(TwoHand_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then TwoHand_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(TwoHand_TxtBx.Text) < 1 Then TwoHand_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Rage_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rage_TxtBx.Leave
        If Int(Rage_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Rage_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Rage_TxtBx.Text) < 1 Then Rage_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Attack_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Attack_TxtBx.Leave
        If Int(Attack_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Attack_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Attack_TxtBx.Text) < 1 Then Attack_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Parry_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Parry_TxtBx.Leave
        If Int(Parry_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Parry_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Parry_TxtBx.Text) < 1 Then Parry_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Tactics_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tactics_TxtBx.Leave
        If Int(Tactics_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Tactics_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Tactics_TxtBx.Text) < 1 Then Tactics_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Warcry_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Warcry_TxtBx.Leave
        If Int(Warcry_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Warcry_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Warcry_TxtBx.Text) < 1 Then Warcry_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Surround_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Surround_TxtBx.Leave
        If Int(Surround_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Surround_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Surround_TxtBx.Text) < 1 Then Surround_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Speed_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Speed_TxtBx.Leave
        If Int(Speed_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Speed_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Speed_TxtBx.Text) < 1 Then Speed_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Immunity_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Immunity_TxtBx.Leave
        If Int(Immunity_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Immunity_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Immunity_TxtBx.Text) < 1 Then Immunity_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Barter_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Barter_TxtBx.Leave
        If Int(Barter_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Barter_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Barter_TxtBx.Text) < 1 Then Barter_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Perc_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Perc_TxtBx.Leave
        If Int(Perc_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Perc_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Perc_TxtBx.Text) < 1 Then Perc_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Stealth_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Stealth_TxtBx.Leave
        If Int(Stealth_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Stealth_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Stealth_TxtBx.Text) < 1 Then Stealth_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Regen_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Regen_TxtBx.Leave
        If Int(Regen_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Regen_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Regen_TxtBx.Text) < 1 Then Regen_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub Prof_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Prof_TxtBx.Leave
        If Int(Prof_TxtBx.Text) > 50 Then Prof_TxtBx.Text = 50
        If Int(Prof_TxtBx.Text) < 1 Then Prof_TxtBx.Text = 1

        SetExpValues()
    End Sub

    Private Sub HP_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HP_Mod_TxtBx.Leave
        If Int(HP_Mod_TxtBx.Text) > Math.Floor(Int(HP_TxtBx.Text) / 2) Then HP_Mod_TxtBx.ForeColor = Color.Red Else HP_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub End_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles End_Mod_TxtBx.Leave
        If Int(End_Mod_TxtBx.Text) > Math.Floor(Int(End_TxtBx.Text) / 2) Then End_Mod_TxtBx.ForeColor = Color.Red Else End_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Wis_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Wis_Mod_TxtBx.Leave
        If Int(Wis_Mod_TxtBx.Text) > Math.Floor(Int(Wis_TxtBx.Text) / 2) Then Wis_Mod_TxtBx.ForeColor = Color.Red Else Wis_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Int_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Int_Mod_TxtBx.Leave
        If Int(Int_Mod_TxtBx.Text) > Math.Floor(Int(Int_TxtBx.Text) / 2) Then Int_Mod_TxtBx.ForeColor = Color.Red Else Int_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Agi_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Agi_Mod_TxtBx.Leave
        If Int(Agi_Mod_TxtBx.Text) > Math.Floor(Int(Agi_TxtBx.Text) / 2) Then Agi_Mod_TxtBx.ForeColor = Color.Red Else Agi_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Str_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Str_Mod_TxtBx.Leave
        If Int(Str_Mod_TxtBx.Text) > Math.Floor(Int(Str_TxtBx.Text) / 2) Then Str_Mod_TxtBx.ForeColor = Color.Red Else Str_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Dagger_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dagger_Mod_TxtBx.Leave
        If Int(Dagger_Mod_TxtBx.Text) > Math.Floor(Int(Dagger_TxtBx.Text) / 2) Then Dagger_Mod_TxtBx.ForeColor = Color.Red Else Dagger_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub H2H_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles H2H_Mod_TxtBx.Leave
        If Int(H2H_Mod_TxtBx.Text) > Math.Floor(Int(H2H_TxtBx.Text) / 2) Then H2H_Mod_TxtBx.ForeColor = Color.Red Else H2H_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Sword_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sword_Mod_TxtBx.Leave
        If Int(Sword_Mod_TxtBx.Text) > Math.Floor(Int(Sword_TxtBx.Text) / 2) Then Sword_Mod_TxtBx.ForeColor = Color.Red Else Sword_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub TwoHand_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TwoHand_Mod_TxtBx.Leave
        If Int(TwoHand_Mod_TxtBx.Text) > Math.Floor(Int(TwoHand_TxtBx.Text) / 2) Then TwoHand_Mod_TxtBx.ForeColor = Color.Red Else TwoHand_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Rage_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rage_Mod_TxtBx.Leave
        If Int(Rage_Mod_TxtBx.Text) > Math.Floor(Int(Rage_TxtBx.Text) / 2) Then Rage_Mod_TxtBx.ForeColor = Color.Red Else Rage_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Attack_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Attack_Mod_TxtBx.Leave
        If Int(Attack_Mod_TxtBx.Text) > Math.Floor(Int(Attack_TxtBx.Text) / 2) Then Attack_Mod_TxtBx.ForeColor = Color.Red Else Attack_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Parry_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Parry_Mod_TxtBx.Leave
        If Int(Parry_Mod_TxtBx.Text) > Math.Floor(Int(Parry_TxtBx.Text) / 2) Then Parry_Mod_TxtBx.ForeColor = Color.Red Else Parry_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Tactics_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tactics_Mod_TxtBx.Leave
        If Int(Tactics_Mod_TxtBx.Text) > Math.Floor(Int(Tactics_TxtBx.Text) / 2) Then Tactics_Mod_TxtBx.ForeColor = Color.Red Else Tactics_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Warcry_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Warcry_Mod_TxtBx.Leave
        If Int(Warcry_Mod_TxtBx.Text) > Math.Floor(Int(Warcry_TxtBx.Text) / 2) Then Warcry_Mod_TxtBx.ForeColor = Color.Red Else Warcry_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Surround_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Surround_Mod_TxtBx.Leave
        If Int(Surround_Mod_TxtBx.Text) > Math.Floor(Int(Surround_TxtBx.Text) / 2) Then Surround_Mod_TxtBx.ForeColor = Color.Red Else Surround_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Speed_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Speed_Mod_TxtBx.Leave
        If Int(Speed_Mod_TxtBx.Text) > Math.Floor(Int(Speed_TxtBx.Text) / 2) Then Speed_Mod_TxtBx.ForeColor = Color.Red Else Speed_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Immunity_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Immunity_Mod_TxtBx.Leave
        If Int(Immunity_Mod_TxtBx.Text) > Math.Floor(Int(Immunity_TxtBx.Text) / 2) Then Immunity_Mod_TxtBx.ForeColor = Color.Red Else Immunity_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Barter_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Barter_Mod_TxtBx.Leave
        If Int(Barter_Mod_TxtBx.Text) > Math.Floor(Int(Barter_TxtBx.Text) / 2) Then Barter_Mod_TxtBx.ForeColor = Color.Red Else Barter_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Perc_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Perc_Mod_TxtBx.Leave
        If Int(Perc_Mod_TxtBx.Text) > Math.Floor(Int(Perc_TxtBx.Text) / 2) Then Perc_Mod_TxtBx.ForeColor = Color.Red Else Perc_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Stealth_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Stealth_Mod_TxtBx.Leave
        If Int(Stealth_Mod_TxtBx.Text) > Math.Floor(Int(Stealth_TxtBx.Text) / 2) Then Stealth_Mod_TxtBx.ForeColor = Color.Red Else Stealth_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub Regen_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Regen_Mod_TxtBx.Leave
        If Int(Regen_Mod_TxtBx.Text) > Math.Floor(Int(Regen_TxtBx.Text) / 2) Then Regen_Mod_TxtBx.ForeColor = Color.Red Else Regen_Mod_TxtBx.ForeColor = Color.Black

        SetExpValues()
    End Sub

    Private Sub WWV_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WWV_TxtBx.Leave
        SetExpValues()
    End Sub

    Private Sub DefReduct_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DefReduct_TxtBx.Leave
        SetExpValues()
    End Sub

    Private Sub SpeedReduct_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpeedReduct_TxtBx.Leave
        SetExpValues()
    End Sub

    Private Sub ShowTact_Chk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowTact_Chk.CheckedChanged
        Select Case ShowTact_Chk.Checked
            Case True
                Dagger_Mod_Lbl.Text = Int(Dagger_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                H2H_Mod_Lbl.Text = Int(H2H_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Sword_Mod_Lbl.Text = Int(Sword_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                TwoHand_Mod_Lbl.Text = Int(TwoHand_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Rage_Mod_Lbl.Text = Int(Rage_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Attack_Mod_Lbl.Text = Int(Attack_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Parry_Mod_Lbl.Text = Int(Parry_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Warcry_Mod_Lbl.Text = Int(Warcry_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Surround_Mod_Lbl.Text = Int(Surround_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Speed_Mod_Lbl.Text = Int(Speed_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Immunity_Mod_Lbl.Text = Int(Immunity_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Barter_Mod_Lbl.Text = Int(Barter_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Perc_Mod_Lbl.Text = Int(Perc_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Stealth_Mod_Lbl.Text = Int(Stealth_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
                Regen_Mod_Lbl.Text = Int(Regen_Mod_Lbl.Text) + Int(TactBonus_TxtBx.Text)
            Case False
                Dagger_Mod_Lbl.Text = Int(Dagger_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                H2H_Mod_Lbl.Text = Int(H2H_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Sword_Mod_Lbl.Text = Int(Sword_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                TwoHand_Mod_Lbl.Text = Int(TwoHand_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Rage_Mod_Lbl.Text = Int(Rage_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Attack_Mod_Lbl.Text = Int(Attack_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Parry_Mod_Lbl.Text = Int(Parry_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Warcry_Mod_Lbl.Text = Int(Warcry_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Surround_Mod_Lbl.Text = Int(Surround_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Speed_Mod_Lbl.Text = Int(Speed_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Immunity_Mod_Lbl.Text = Int(Immunity_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Barter_Mod_Lbl.Text = Int(Barter_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Perc_Mod_Lbl.Text = Int(Perc_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Stealth_Mod_Lbl.Text = Int(Stealth_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
                Regen_Mod_Lbl.Text = Int(Regen_Mod_Lbl.Text) - Int(TactBonus_TxtBx.Text)
        End Select

        SetExpValues()
    End Sub

    Private Sub ShowRage_Chk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowRage_Chk.CheckedChanged
        Select Case ShowRage_Chk.Checked
            Case True
                Immunity_Mod_Lbl.Text = Int(Immunity_Mod_Lbl.Text) + Int(RageBonus_TxtBx.Text)
            Case False
                Immunity_Mod_Lbl.Text = Int(Immunity_Mod_Lbl.Text) - Int(RageBonus_TxtBx.Text)
        End Select

        SetExpValues()
    End Sub

    Private Sub Save_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save_Btn.Click
        Dim sFD As New SaveFileDialog

        sFD.Filter = "Text File|*.txt"
        sFD.ShowDialog()

        If sFD.FileName = "" Then MsgBox("File not saved, Invalid file name.") : Exit Sub

        Using sw As StreamWriter = File.CreateText(sFD.FileName)
            sw.WriteLine(HP_TxtBx.Text)
            sw.WriteLine(End_TxtBx.Text)
            sw.WriteLine(Wis_TxtBx.Text)
            sw.WriteLine(Int_TxtBx.Text)
            sw.WriteLine(Agi_TxtBx.Text)
            sw.WriteLine(Str_TxtBx.Text)
            sw.WriteLine(Dagger_TxtBx.Text)
            sw.WriteLine(H2H_TxtBx.Text)
            sw.WriteLine(Sword_TxtBx.Text)
            sw.WriteLine(TwoHand_TxtBx.Text)
            sw.WriteLine(Rage_TxtBx.Text)
            sw.WriteLine(Attack_TxtBx.Text)
            sw.WriteLine(Parry_TxtBx.Text)
            sw.WriteLine(Tactics_TxtBx.Text)
            sw.WriteLine(Warcry_TxtBx.Text)
            sw.WriteLine(Surround_TxtBx.Text)
            sw.WriteLine(Speed_TxtBx.Text)
            sw.WriteLine(Immunity_TxtBx.Text)
            sw.WriteLine(Barter_TxtBx.Text)
            sw.WriteLine(Perc_TxtBx.Text)
            sw.WriteLine(Stealth_TxtBx.Text)
            sw.WriteLine(Regen_TxtBx.Text)
            sw.WriteLine(Prof_TxtBx.Text)
            sw.WriteLine(HP_Mod_TxtBx.Text)
            sw.WriteLine(End_Mod_TxtBx.Text)
            sw.WriteLine(Wis_Mod_TxtBx.Text)
            sw.WriteLine(Int_Mod_TxtBx.Text)
            sw.WriteLine(Agi_Mod_TxtBx.Text)
            sw.WriteLine(Str_Mod_TxtBx.Text)
            sw.WriteLine(Dagger_Mod_TxtBx.Text)
            sw.WriteLine(H2H_Mod_TxtBx.Text)
            sw.WriteLine(Sword_Mod_TxtBx.Text)
            sw.WriteLine(TwoHand_Mod_TxtBx.Text)
            sw.WriteLine(Rage_Mod_TxtBx.Text)
            sw.WriteLine(Attack_Mod_TxtBx.Text)
            sw.WriteLine(Parry_Mod_TxtBx.Text)
            sw.WriteLine(Tactics_Mod_TxtBx.Text)
            sw.WriteLine(Warcry_Mod_TxtBx.Text)
            sw.WriteLine(Surround_Mod_TxtBx.Text)
            sw.WriteLine(Speed_Mod_TxtBx.Text)
            sw.WriteLine(Immunity_Mod_TxtBx.Text)
            sw.WriteLine(Barter_Mod_TxtBx.Text)
            sw.WriteLine(Perc_Mod_TxtBx.Text)
            sw.WriteLine(Stealth_Mod_TxtBx.Text)
            sw.WriteLine(Regen_Mod_TxtBx.Text)
            sw.WriteLine(WWV_TxtBx.Text)
            sw.WriteLine(DefReduct_TxtBx.Text)
            sw.WriteLine(SpeedReduct_TxtBx.Text)
            sw.Close()
        End Using

    End Sub

    Private Sub Load_Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Load_Btn.Click
        Dim oFD As New OpenFileDialog

        oFD.Filter = "Text File|*.txt"
        oFD.ShowDialog()

        If oFD.FileName = "" Then MsgBox("File not Loaded, Invalid File Name") : Exit Sub

        Using sr As StreamReader = File.OpenText(oFD.FileName)
            HP_TxtBx.Text = sr.ReadLine()
            End_TxtBx.Text = sr.ReadLine()
            Wis_TxtBx.Text = sr.ReadLine()
            Int_TxtBx.Text = sr.ReadLine()
            Agi_TxtBx.Text = sr.ReadLine()
            Str_TxtBx.Text = sr.ReadLine()
            Dagger_TxtBx.Text = sr.ReadLine()
            H2H_TxtBx.Text = sr.ReadLine()
            Sword_TxtBx.Text = sr.ReadLine()
            TwoHand_TxtBx.Text = sr.ReadLine()
            Rage_TxtBx.Text = sr.ReadLine()
            Attack_TxtBx.Text = sr.ReadLine()
            Parry_TxtBx.Text = sr.ReadLine()
            Tactics_TxtBx.Text = sr.ReadLine()
            Warcry_TxtBx.Text = sr.ReadLine()
            Surround_TxtBx.Text = sr.ReadLine()
            Speed_TxtBx.Text = sr.ReadLine()
            Immunity_TxtBx.Text = sr.ReadLine()
            Barter_TxtBx.Text = sr.ReadLine()
            Perc_TxtBx.Text = sr.ReadLine()
            Stealth_TxtBx.Text = sr.ReadLine()
            Regen_TxtBx.Text = sr.ReadLine()
            Prof_TxtBx.Text = sr.ReadLine()
            HP_Mod_TxtBx.Text = sr.ReadLine()
            End_Mod_TxtBx.Text = sr.ReadLine()
            Wis_Mod_TxtBx.Text = sr.ReadLine()
            Int_Mod_TxtBx.Text = sr.ReadLine()
            Agi_Mod_TxtBx.Text = sr.ReadLine()
            Str_Mod_TxtBx.Text = sr.ReadLine()
            Dagger_Mod_TxtBx.Text = sr.ReadLine()
            H2H_Mod_TxtBx.Text = sr.ReadLine()
            Sword_Mod_TxtBx.Text = sr.ReadLine()
            TwoHand_Mod_TxtBx.Text = sr.ReadLine()
            Rage_Mod_TxtBx.Text = sr.ReadLine()
            Attack_Mod_TxtBx.Text = sr.ReadLine()
            Parry_Mod_TxtBx.Text = sr.ReadLine()
            Tactics_Mod_TxtBx.Text = sr.ReadLine()
            Warcry_Mod_TxtBx.Text = sr.ReadLine()
            Surround_Mod_TxtBx.Text = sr.ReadLine()
            Speed_Mod_TxtBx.Text = sr.ReadLine()
            Immunity_Mod_TxtBx.Text = sr.ReadLine()
            Barter_Mod_TxtBx.Text = sr.ReadLine()
            Perc_Mod_TxtBx.Text = sr.ReadLine()
            Stealth_Mod_TxtBx.Text = sr.ReadLine()
            Regen_Mod_TxtBx.Text = sr.ReadLine()
            WWV_TxtBx.Text = sr.ReadLine()
            DefReduct_TxtBx.Text = sr.ReadLine()
            SpeedReduct_TxtBx.Text = sr.ReadLine()
            sr.Close()
        End Using

        SetExpValues()

    End Sub

   
    Private Sub frmWarriorCalc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetExpValues()
    End Sub
End Class