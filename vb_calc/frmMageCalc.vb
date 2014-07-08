Imports System.IO

Public Class frmMageCalc
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

        ExpReq = Int((((Int(Mana_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Mana_ExpLbl.Text = ExpReq

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

        ExpReq = Int((((Int(Staff_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Staff_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Duration_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Duration_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Immunity_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Immunity_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Bless_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Bless_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Heal_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Heal_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Freeze_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Freeze_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(MagicShield_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        MagicShield_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Lightning_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Lightning_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Fire_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Fire_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Barter_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Barter_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Perc_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Perc_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Stealth_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Stealth_ExpLbl.Text = ExpReq

        ExpReq = Int((((Int(Meditate_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Meditate_ExpLbl.Text = ExpReq

        sf = 3

        ExpReq = Int((((Int(Prof_TxtBx.Text) - sb) + 6) ^ 3) * (sf * cf))
        Prof_ExpLbl.Text = ExpReq

        SetMods()
    End Sub

    Private Sub SetMods()
        Dim WisMod, IntMod, AgiMod, StrMod As Integer
        Dim IAS, SSA, WIS, IIW, AAI, WWW As Integer
        Dim MaxStatVal, StatVal, MaxEquipVal, EquipVal, ModVal As Integer
        Dim BlessBonus, MaxBlessBonus As Integer

        If ShowBless_Chk.Checked = True Then BlessBonus = Int(BlessBonus_TxtBx.Text) Else BlessBonus = 0

        MaxEquipVal = Math.Floor(Int(HP_TxtBx.Text) / 2)
        EquipVal = Int(HP_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        HP_Mod_Lbl.Text = Int(HP_TxtBx.Text) + ModVal

        MaxEquipVal = Math.Floor(Int(End_TxtBx.Text) / 2)
        EquipVal = Int(End_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        End_Mod_Lbl.Text = Int(End_TxtBx.Text) + ModVal

        MaxEquipVal = Math.Floor(Int(Mana_TxtBx.Text) / 2)
        EquipVal = Int(Mana_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        Mana_Mod_Lbl.Text = Int(Mana_TxtBx.Text) + ModVal

        MaxEquipVal = Math.Floor(Int(Wis_TxtBx.Text) / 2)
        EquipVal = Int(Wis_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        If BlessBonus > Math.Floor(Int(Wis_TxtBx.Text) / 2) Then MaxBlessBonus = Math.Floor(Int(Wis_TxtBx.Text) / 2) Else MaxBlessBonus = BlessBonus
        Wis_Mod_Lbl.Text = Int(Wis_TxtBx.Text) + ModVal + MaxBlessBonus

        MaxEquipVal = Math.Floor(Int(Int_TxtBx.Text) / 2)
        EquipVal = Int(Int_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        If BlessBonus > Math.Floor(Int(Int_TxtBx.Text) / 2) Then MaxBlessBonus = Math.Floor(Int(Int_TxtBx.Text) / 2) Else MaxBlessBonus = BlessBonus
        Int_Mod_Lbl.Text = Int(Int_TxtBx.Text) + ModVal + MaxBlessBonus

        MaxEquipVal = Math.Floor(Int(Agi_TxtBx.Text) / 2)
        EquipVal = Int(Agi_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        If BlessBonus > Math.Floor(Int(Agi_TxtBx.Text) / 2) Then MaxBlessBonus = Math.Floor(Int(Agi_TxtBx.Text) / 2) Else MaxBlessBonus = BlessBonus
        Agi_Mod_Lbl.Text = Int(Agi_TxtBx.Text) + ModVal + MaxBlessBonus

        MaxEquipVal = Math.Floor(Int(Str_TxtBx.Text) / 2)
        EquipVal = Int(Str_Mod_TxtBx.Text)
        If EquipVal > MaxEquipVal Then ModVal = MaxEquipVal Else ModVal = EquipVal
        If BlessBonus > Math.Floor(Int(Str_TxtBx.Text) / 2) Then MaxBlessBonus = Math.Floor(Int(Str_TxtBx.Text) / 2) Else MaxBlessBonus = BlessBonus
        Str_Mod_Lbl.Text = Int(Str_TxtBx.Text) + ModVal + MaxBlessBonus

        WisMod = Int(Wis_Mod_Lbl.Text)
        IntMod = Int(Int_Mod_Lbl.Text)
        AgiMod = Int(Agi_Mod_Lbl.Text)
        StrMod = Int(Str_Mod_Lbl.Text)

        IAS = Math.Floor((IntMod + AgiMod + StrMod) / 5)
        SSA = Math.Floor((StrMod + StrMod + AgiMod) / 5)
        WIS = Math.Floor((WisMod + IntMod + StrMod) / 5)
        IIW = Math.Floor((IntMod + IntMod + WisMod) / 5)
        AAI = Math.Floor((AgiMod + AgiMod + IntMod) / 5)
        WWW = Math.Floor((WisMod + WisMod + WisMod) / 5)

        If Int(Dagger_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Dagger_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Dagger_Mod_TxtBx.Text) > Int(Int(Dagger_TxtBx.Text) / 2) Then EquipVal = Int(Int(Dagger_TxtBx.Text) / 2) Else EquipVal = Int(Dagger_Mod_TxtBx.Text)
        ModVal = Int(Dagger_TxtBx.Text) + StatVal + EquipVal
        Dagger_Mod_Lbl.Text = ModVal

        If Int(H2H_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(H2H_TxtBx.Text) * 2
        If SSA > MaxStatVal Then StatVal = MaxStatVal Else StatVal = SSA
        If Int(H2H_Mod_TxtBx.Text) > Int(Int(H2H_TxtBx.Text) / 2) Then EquipVal = Int(Int(H2H_TxtBx.Text) / 2) Else EquipVal = Int(H2H_Mod_TxtBx.Text)
        ModVal = Int(H2H_TxtBx.Text) + StatVal + EquipVal
        H2H_Mod_Lbl.Text = ModVal

        If Int(Staff_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Staff_TxtBx.Text) * 2
        If IAS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IAS
        If Int(Staff_Mod_TxtBx.Text) > Int(Int(Staff_TxtBx.Text) / 2) Then EquipVal = Int(Int(Staff_TxtBx.Text) / 2) Else EquipVal = Int(Staff_Mod_TxtBx.Text)
        ModVal = Int(Staff_TxtBx.Text) + StatVal + EquipVal
        Staff_Mod_Lbl.Text = ModVal

        If Int(Duration_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Duration_TxtBx.Text) * 2
        If WIS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = WIS
        If Int(Duration_Mod_TxtBx.Text) > Int(Int(Duration_TxtBx.Text) / 2) Then EquipVal = Int(Int(Duration_TxtBx.Text) / 2) Else EquipVal = Int(Duration_Mod_TxtBx.Text)
        ModVal = Int(Duration_TxtBx.Text) + StatVal + EquipVal
        Duration_Mod_Lbl.Text = ModVal

        If Int(Immunity_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Immunity_TxtBx.Text) * 2
        If WIS > MaxStatVal Then StatVal = MaxStatVal Else StatVal = WIS
        If Int(Immunity_Mod_TxtBx.Text) > Int(Int(Immunity_TxtBx.Text) / 2) Then EquipVal = Int(Int(Immunity_TxtBx.Text) / 2) Else EquipVal = Int(Immunity_Mod_TxtBx.Text)
        ModVal = Int(Immunity_TxtBx.Text) + StatVal + EquipVal
        Immunity_Mod_Lbl.Text = ModVal

        If Int(Bless_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Bless_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Bless_Mod_TxtBx.Text) > Int(Int(Bless_TxtBx.Text) / 2) Then EquipVal = Int(Int(Bless_TxtBx.Text) / 2) Else EquipVal = Int(Bless_Mod_TxtBx.Text)
        ModVal = Int(Bless_TxtBx.Text) + StatVal + EquipVal
        Bless_Mod_Lbl.Text = ModVal

        If Int(Heal_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Heal_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Heal_Mod_TxtBx.Text) > Int(Int(Heal_TxtBx.Text) / 2) Then EquipVal = Int(Int(Heal_TxtBx.Text) / 2) Else EquipVal = Int(Heal_Mod_TxtBx.Text)
        ModVal = Int(Heal_TxtBx.Text) + StatVal + EquipVal
        Heal_Mod_Lbl.Text = ModVal

        If Int(Freeze_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Freeze_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Freeze_Mod_TxtBx.Text) > Int(Int(Freeze_TxtBx.Text) / 2) Then EquipVal = Int(Int(Freeze_TxtBx.Text) / 2) Else EquipVal = Int(Freeze_Mod_TxtBx.Text)
        ModVal = Int(Freeze_TxtBx.Text) + StatVal + EquipVal
        Freeze_Mod_Lbl.Text = ModVal

        If Int(MagicShield_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(MagicShield_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(MagicShield_Mod_TxtBx.Text) > Int(Int(MagicShield_TxtBx.Text) / 2) Then EquipVal = Int(Int(MagicShield_TxtBx.Text) / 2) Else EquipVal = Int(MagicShield_Mod_TxtBx.Text)
        ModVal = Int(MagicShield_TxtBx.Text) + StatVal + EquipVal
        MagicShield_Mod_Lbl.Text = ModVal

        If Int(Lightning_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Lightning_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Lightning_Mod_TxtBx.Text) > Int(Int(Lightning_TxtBx.Text) / 2) Then EquipVal = Int(Int(Lightning_TxtBx.Text) / 2) Else EquipVal = Int(Lightning_Mod_TxtBx.Text)
        ModVal = Int(Lightning_TxtBx.Text) + StatVal + EquipVal
        Lightning_Mod_Lbl.Text = ModVal

        If Int(Fire_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Fire_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Fire_Mod_TxtBx.Text) > Int(Int(Fire_TxtBx.Text) / 2) Then EquipVal = Int(Int(Fire_TxtBx.Text) / 2) Else EquipVal = Int(Fire_Mod_TxtBx.Text)
        ModVal = Int(Fire_TxtBx.Text) + StatVal + EquipVal
        Fire_Mod_Lbl.Text = ModVal

        If Int(Barter_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Barter_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Barter_Mod_TxtBx.Text) > Int(Int(Barter_TxtBx.Text) / 2) Then EquipVal = Int(Int(Barter_TxtBx.Text) / 2) Else EquipVal = Int(Barter_Mod_TxtBx.Text)
        ModVal = Int(Barter_TxtBx.Text) + StatVal + EquipVal
        Barter_Mod_Lbl.Text = ModVal


        If Int(Perc_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Perc_TxtBx.Text) * 2
        If IIW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = IIW
        If Int(Perc_Mod_TxtBx.Text) > Int(Int(Perc_TxtBx.Text) / 2) Then EquipVal = Int(Int(Perc_TxtBx.Text) / 2) Else EquipVal = Int(Perc_Mod_TxtBx.Text)
        ModVal = Int(Perc_TxtBx.Text) + StatVal + EquipVal
        Perc_Mod_Lbl.Text = ModVal


        If Int(Stealth_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Stealth_TxtBx.Text) * 2
        If AAI > MaxStatVal Then StatVal = MaxStatVal Else StatVal = AAI
        If Int(Stealth_Mod_TxtBx.Text) > Int(Int(Stealth_TxtBx.Text) / 2) Then EquipVal = Int(Int(Stealth_TxtBx.Text) / 2) Else EquipVal = Int(Stealth_Mod_TxtBx.Text)
        ModVal = Int(Stealth_TxtBx.Text) + StatVal + EquipVal
        Stealth_Mod_Lbl.Text = ModVal


        If Int(Meditate_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Meditate_TxtBx.Text) * 2
        If WWW > MaxStatVal Then StatVal = MaxStatVal Else StatVal = WWW
        If Int(Meditate_Mod_TxtBx.Text) > Int(Int(Meditate_TxtBx.Text) / 2) Then EquipVal = Int(Int(Meditate_TxtBx.Text) / 2) Else EquipVal = Int(Meditate_Mod_TxtBx.Text)
        ModVal = Int(Meditate_TxtBx.Text) + StatVal + EquipVal
        Meditate_Mod_Lbl.Text = ModVal

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
        'cf1 = 0.1333333333333
        cf2 = 0.1
        cf3 = 0.1

        sVal = 25
        Do Until sVal = Int(HP_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb1) + 6) ^ 3) * (sf1 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 25
        Do Until sVal = Int(End_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb1) + 6) ^ 3) * (sf1 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 25
        Do Until sVal = Int(Mana_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb1) + 6) ^ 3) * (sf1 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 10
        Do Until sVal = Int(Wis_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb2) + 6) ^ 3) * (sf2 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 10
        Do Until sVal = Int(Int_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb2) + 6) ^ 3) * (sf2 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 10
        Do Until sVal = Int(Agi_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb2) + 6) ^ 3) * (sf2 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 10
        Do Until sVal = Int(Str_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb2) + 6) ^ 3) * (sf2 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Dagger_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(H2H_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Staff_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Duration_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Immunity_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Bless_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Heal_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Freeze_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(MagicShield_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Lightning_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Fire_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Barter_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Perc_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Stealth_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Meditate_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb3) + 6) ^ 3) * (sf3 * cf3))
            sVal = sVal + 1
        Loop

        sVal = 1
        Do Until sVal = Int(Prof_TxtBx.Text)
            TotalExp = TotalExp + Int((((sVal - sb4) + 6) ^ 3) * (sf4 * cf3))
            sVal = sVal + 1
        Loop

        ReqLevel = TotalExp ^ 0.25
        If ReqLevel = 0 Then Level_TxtBx.Text = 1 Else Level_TxtBx.Text = ReqLevel

        gLevel = ReqLevel
        SetExtraValues()
    End Sub

    Private Sub SetExtraValues()
        Dim ws, ws_max As Integer
        Dim WisMod, IntMod, AgiMod, StrMod As Integer
        Dim MaxBlessMod As Integer
        Dim offVal, def_red As Integer
        Dim wv As Double


        If ShowBless_Chk.Checked = True Then
            WisMod = Int(Wis_Mod_Lbl.Text) - Math.Min(Int(BlessBonus_TxtBx.Text), (Int(Int(Wis_TxtBx.Text) / 2)))
            IntMod = Int(Int_Mod_Lbl.Text) - Math.Min(Int(BlessBonus_TxtBx.Text), (Int(Int(Int_TxtBx.Text) / 2)))
        Else
            WisMod = Int(Wis_Mod_Lbl.Text)
            IntMod = Int(Int_Mod_Lbl.Text)
        End If

        If Math.Floor(Int(Bless_TxtBx.Text) / 2) < Int(Bless_Mod_TxtBx.Text) Then MaxBlessMod = Math.Floor(Int(Bless_TxtBx.Text) / 2) Else MaxBlessMod = Int(Bless_Mod_TxtBx.Text)

        BlessBonus_TxtBx.Text = Int((Int(Bless_TxtBx.Text) + ((WisMod + IntMod + IntMod) / 5) + MaxBlessMod) / 4)

        NxtBless.Text = Int((Int(BlessBonus_TxtBx.Text) + 1) * 4)


        StrMod = Int(Str_Mod_Lbl.Text)
        AgiMod = Int(Agi_Mod_Lbl.Text)


        ws = Math.Max(Int(Dagger_Mod_Lbl.Text), Int(H2H_Mod_Lbl.Text))
        ws_max = Math.Max(ws, Int(Staff_Mod_Lbl.Text))

        offVal = Math.Floor(Int(ws_max) + (Int(Lightning_Mod_Lbl.Text) + Int(Fire_Mod_Lbl.Text) + Int(Freeze_Mod_Lbl.Text) + Int(Heal_Mod_Lbl.Text) + Int(Bless_Mod_Lbl.Text) + Int(MagicShield_Mod_Lbl.Text)) / 3 - 3 * Int(Level_TxtBx.Text) - 20)
        If offVal < 0 Then offVal = 0 Else 
        Offense_TxtBx.Text = offVal
        def_red = Math.Floor(Int(Level_TxtBx.Text) / 2)
        If def_red > 10 Then def_red = 10
        Defense_TxtBx.Text = Int(((ws_max + (Int(MagicShield_Mod_Lbl.Text) * 2)) - Int(DefReduct_TxtBx.Text) - def_red))

        wv = CDbl(WWV_TxtBx.Text) + (Math.Floor(StrMod / 2) * 0.15) + (Math.Ceiling(StrMod / 2) * 0.1)

        WV_TxtBx.Text = wv

        RawSpeed_TxtBx.Text = Int((AgiMod + AgiMod + StrMod) / 5)



        PostValidateData()

    End Sub
    Private Sub PreValidateDate()
        If Int(HP_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then HP_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(HP_TxtBx.Text) < 25 Then HP_TxtBx.Text = 25

        If Int(End_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then End_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(End_TxtBx.Text) < 25 Then End_TxtBx.Text = 25

        If Int(Mana_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Mana_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Mana_TxtBx.Text) < 25 Then Mana_TxtBx.Text = 25

        If Int(Wis_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Wis_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Wis_TxtBx.Text) < 10 Then Wis_TxtBx.Text = 10

        If Int(Int_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Int_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Int_TxtBx.Text) < 10 Then Int_TxtBx.Text = 10

        If Int(Agi_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Agi_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Agi_TxtBx.Text) < 10 Then Agi_TxtBx.Text = 10

        If Int(Str_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Str_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Str_TxtBx.Text) < 10 Then Str_TxtBx.Text = 10

        If Int(Dagger_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Dagger_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Dagger_TxtBx.Text) < 1 Then Dagger_TxtBx.Text = 1

        If Int(H2H_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then H2H_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(H2H_TxtBx.Text) < 1 Then H2H_TxtBx.Text = 1

        If Int(Staff_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Staff_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Staff_TxtBx.Text) < 1 Then Staff_TxtBx.Text = 1

        If Int(Duration_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Duration_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Duration_TxtBx.Text) < 1 Then Duration_TxtBx.Text = 1

        If Int(Immunity_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Immunity_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Immunity_TxtBx.Text) < 1 Then Immunity_TxtBx.Text = 1

        If Int(Bless_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Bless_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Bless_TxtBx.Text) < 1 Then Bless_TxtBx.Text = 1

        If Int(Heal_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Heal_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Heal_TxtBx.Text) < 1 Then Heal_TxtBx.Text = 1

        If Int(Freeze_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Freeze_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Freeze_TxtBx.Text) < 1 Then Freeze_TxtBx.Text = 1

        If Int(MagicShield_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then MagicShield_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(MagicShield_TxtBx.Text) < 1 Then MagicShield_TxtBx.Text = 1

        If Int(Lightning_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Lightning_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Lightning_TxtBx.Text) < 1 Then Lightning_TxtBx.Text = 1

        If Int(Fire_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Fire_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Fire_TxtBx.Text) < 1 Then Fire_TxtBx.Text = 1

        If Int(Barter_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Barter_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Barter_TxtBx.Text) < 1 Then Barter_TxtBx.Text = 1

        If Int(Perc_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Perc_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Perc_TxtBx.Text) < 1 Then Perc_TxtBx.Text = 1

        If Int(Stealth_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Stealth_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Stealth_TxtBx.Text) < 1 Then Stealth_TxtBx.Text = 1

        If Int(Meditate_TxtBx.Text) > Int(MaxBase_txtbox.Text) Then Meditate_TxtBx.Text = Int(MaxBase_txtbox.Text)
        If Int(Meditate_TxtBx.Text) < 1 Then Meditate_TxtBx.Text = 1

        If Int(Prof_TxtBx.Text) > 50 Then Prof_TxtBx.Text = 50
        If Int(Prof_TxtBx.Text) < 1 Then Prof_TxtBx.Text = 1

        SetExpValues()
    End Sub
    Private Sub PostValidateData()
        If Int(HP_Mod_TxtBx.Text) > Math.Floor(Int(HP_TxtBx.Text) / 2) Then HP_Mod_TxtBx.ForeColor = Color.Red Else HP_Mod_TxtBx.ForeColor = Color.Black
        If Int(End_Mod_TxtBx.Text) > Math.Floor(Int(End_TxtBx.Text) / 2) Then End_Mod_TxtBx.ForeColor = Color.Red Else End_Mod_TxtBx.ForeColor = Color.Black
        If Int(Mana_Mod_TxtBx.Text) > Math.Floor(Int(Mana_TxtBx.Text) / 2) Then Mana_Mod_TxtBx.ForeColor = Color.Red Else Mana_Mod_TxtBx.ForeColor = Color.Black
        If Int(Wis_Mod_TxtBx.Text) > Math.Floor(Int(Wis_TxtBx.Text) / 2) Then Wis_Mod_TxtBx.ForeColor = Color.Red Else Wis_Mod_TxtBx.ForeColor = Color.Black
        If Int(Int_Mod_TxtBx.Text) > Math.Floor(Int(Int_TxtBx.Text) / 2) Then Int_Mod_TxtBx.ForeColor = Color.Red Else Int_Mod_TxtBx.ForeColor = Color.Black
        If Int(Agi_Mod_TxtBx.Text) > Math.Floor(Int(Agi_TxtBx.Text) / 2) Then Agi_Mod_TxtBx.ForeColor = Color.Red Else Agi_Mod_TxtBx.ForeColor = Color.Black
        If Int(Str_Mod_TxtBx.Text) > Math.Floor(Int(Str_TxtBx.Text) / 2) Then Str_Mod_TxtBx.ForeColor = Color.Red Else Str_Mod_TxtBx.ForeColor = Color.Black
        If Int(Dagger_Mod_TxtBx.Text) > Math.Floor(Int(Dagger_TxtBx.Text) / 2) Then Dagger_Mod_TxtBx.ForeColor = Color.Red Else Dagger_Mod_TxtBx.ForeColor = Color.Black
        If Int(H2H_Mod_TxtBx.Text) > Math.Floor(Int(H2H_TxtBx.Text) / 2) Then H2H_Mod_TxtBx.ForeColor = Color.Red Else H2H_Mod_TxtBx.ForeColor = Color.Black
        If Int(Staff_Mod_TxtBx.Text) > Math.Floor(Int(Staff_TxtBx.Text) / 2) Then Staff_Mod_TxtBx.ForeColor = Color.Red Else Staff_Mod_TxtBx.ForeColor = Color.Black
        If Int(Duration_Mod_TxtBx.Text) > Math.Floor(Int(Duration_TxtBx.Text) / 2) Then Duration_Mod_TxtBx.ForeColor = Color.Red Else Duration_Mod_TxtBx.ForeColor = Color.Black
        If Int(Immunity_Mod_TxtBx.Text) > Math.Floor(Int(Immunity_TxtBx.Text) / 2) Then Immunity_Mod_TxtBx.ForeColor = Color.Red Else Immunity_Mod_TxtBx.ForeColor = Color.Black
        If Int(Bless_Mod_TxtBx.Text) > Math.Floor(Int(Bless_TxtBx.Text) / 2) Then Bless_Mod_TxtBx.ForeColor = Color.Red Else Bless_Mod_TxtBx.ForeColor = Color.Black
        If Int(Heal_Mod_TxtBx.Text) > Math.Floor(Int(Heal_TxtBx.Text) / 2) Then Heal_Mod_TxtBx.ForeColor = Color.Red Else Heal_Mod_TxtBx.ForeColor = Color.Black
        If Int(Freeze_Mod_TxtBx.Text) > Math.Floor(Int(Freeze_TxtBx.Text) / 2) Then Freeze_Mod_TxtBx.ForeColor = Color.Red Else Freeze_Mod_TxtBx.ForeColor = Color.Black
        If Int(MagicShield_Mod_TxtBx.Text) > Math.Floor(Int(MagicShield_TxtBx.Text) / 2) Then MagicShield_Mod_TxtBx.ForeColor = Color.Red Else MagicShield_Mod_TxtBx.ForeColor = Color.Black
        If Int(Lightning_Mod_TxtBx.Text) > Math.Floor(Int(Lightning_TxtBx.Text) / 2) Then Lightning_Mod_TxtBx.ForeColor = Color.Red Else Lightning_Mod_TxtBx.ForeColor = Color.Black
        If Int(Fire_Mod_TxtBx.Text) > Math.Floor(Int(Fire_TxtBx.Text) / 2) Then Fire_Mod_TxtBx.ForeColor = Color.Red Else Fire_Mod_TxtBx.ForeColor = Color.Black
        If Int(Barter_Mod_TxtBx.Text) > Math.Floor(Int(Barter_TxtBx.Text) / 2) Then Barter_Mod_TxtBx.ForeColor = Color.Red Else Barter_Mod_TxtBx.ForeColor = Color.Black
        If Int(Perc_Mod_TxtBx.Text) > Math.Floor(Int(Perc_TxtBx.Text) / 2) Then Perc_Mod_TxtBx.ForeColor = Color.Red Else Perc_Mod_TxtBx.ForeColor = Color.Black
        If Int(Stealth_Mod_TxtBx.Text) > Math.Floor(Int(Stealth_TxtBx.Text) / 2) Then Stealth_Mod_TxtBx.ForeColor = Color.Red Else Stealth_Mod_TxtBx.ForeColor = Color.Black
        If Int(Meditate_Mod_TxtBx.Text) > Math.Floor(Int(Meditate_TxtBx.Text) / 2) Then Meditate_Mod_TxtBx.ForeColor = Color.Red Else Meditate_Mod_TxtBx.ForeColor = Color.Black

        'Whitestar Mod
        If gLevel > Int(MaxLevel_txtbox.Text) + 1 Then
            MsgBox("The last stat increase exceeded your MaxLevel Setting")
        End If

    End Sub
    Private Sub frmMageCalc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetExpValues()
    End Sub

    Private Sub HP_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HP_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub End_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles End_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Mana_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Mana_TxtBx.Leave
        PreValidateDate()
    End Sub


    Private Sub Wis_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Wis_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Int_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Int_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Agi_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Agi_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Str_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Str_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Dagger_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dagger_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub H2H_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles H2H_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Staff_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Staff_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Duration_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Duration_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Immunity_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Immunity_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Bless_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bless_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Heal_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Heal_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Freeze_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Freeze_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub MagicShield_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MagicShield_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Lightning_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Lightning_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Fire_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fire_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Barter_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Barter_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Perc_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Perc_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Stealth_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Stealth_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Meditate_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Meditate_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Prof_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Prof_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub HP_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HP_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub End_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles End_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Mana_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Mana_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Wis_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Wis_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Int_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Int_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Agi_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Agi_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Str_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Str_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Dagger_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dagger_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub H2H_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles H2H_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Staff_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Staff_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Duration_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Duration_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Immunity_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Immunity_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Bless_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bless_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Heal_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Heal_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Freeze_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Freeze_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub MagicShield_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MagicShield_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Lightning_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Lightning_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Fire_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fire_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Barter_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Barter_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Perc_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Perc_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Stealth_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Stealth_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Meditate_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Meditate_Mod_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub WWV_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WWV_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub DefReduct_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DefReduct_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub ShowBless_Chk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowBless_Chk.CheckedChanged
        PreValidateDate()
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

    Private Sub Mana_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Mana_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Mana_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Mana_TxtBx.Text = Int(Mana_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Mana_TxtBx.Text) = 25 Then
                    MsgBox("Can not reduce below 25")
                    Exit Sub
                End If

                Mana_TxtBx.Text = Int(Mana_TxtBx.Text) - 1

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

    Private Sub Staff_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Staff_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Staff_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Staff_TxtBx.Text = Int(Staff_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Staff_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Staff_TxtBx.Text = Int(Staff_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub Duration_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Duration_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Duration_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Duration_TxtBx.Text = Int(Duration_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Duration_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Duration_TxtBx.Text = Int(Duration_TxtBx.Text) - 1

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

    Private Sub Bless_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Bless_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Bless_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Bless_TxtBx.Text = Int(Bless_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Bless_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Bless_TxtBx.Text = Int(Bless_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub Heal_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Heal_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Heal_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Heal_TxtBx.Text = Int(Heal_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Heal_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Heal_TxtBx.Text = Int(Heal_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub Freeze_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Freeze_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Freeze_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Freeze_TxtBx.Text = Int(Freeze_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Freeze_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Freeze_TxtBx.Text = Int(Freeze_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub MagicShield_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MagicShield_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(MagicShield_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                MagicShield_TxtBx.Text = Int(MagicShield_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(MagicShield_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                MagicShield_TxtBx.Text = Int(MagicShield_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub Lightning_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Lightning_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Lightning_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Lightning_TxtBx.Text = Int(Lightning_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Lightning_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Lightning_TxtBx.Text = Int(Lightning_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub Fire_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Fire_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Fire_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Fire_TxtBx.Text = Int(Fire_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Fire_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Fire_TxtBx.Text = Int(Fire_TxtBx.Text) - 1

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

    Private Sub Meditate_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Meditate_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Meditate_TxtBx.Text) >= Int(MaxBase_txtbox.Text) Then
                    MsgBox("Cannot raise stats beyond MaxBase, Please Check your MaxBase settings")
                    Exit Sub
                End If

                Meditate_TxtBx.Text = Int(Meditate_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Meditate_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Meditate_TxtBx.Text = Int(Meditate_TxtBx.Text) - 1

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
            sw.WriteLine(Staff_TxtBx.Text)
            sw.WriteLine(Duration_TxtBx.Text)
            sw.WriteLine(Immunity_TxtBx.Text)
            sw.WriteLine(Bless_TxtBx.Text)
            sw.WriteLine(Heal_TxtBx.Text)
            sw.WriteLine(Freeze_TxtBx.Text)
            sw.WriteLine(MagicShield_TxtBx.Text)
            sw.WriteLine(Lightning_TxtBx.Text)
            sw.WriteLine(Fire_TxtBx.Text)
            sw.WriteLine(Barter_TxtBx.Text)
            sw.WriteLine(Perc_TxtBx.Text)
            sw.WriteLine(Stealth_TxtBx.Text)
            sw.WriteLine(Meditate_TxtBx.Text)
            sw.WriteLine(Prof_TxtBx.Text)
            sw.WriteLine(HP_Mod_TxtBx.Text)
            sw.WriteLine(End_Mod_TxtBx.Text)
            sw.WriteLine(Wis_Mod_TxtBx.Text)
            sw.WriteLine(Int_Mod_TxtBx.Text)
            sw.WriteLine(Agi_Mod_TxtBx.Text)
            sw.WriteLine(Str_Mod_TxtBx.Text)
            sw.WriteLine(Dagger_Mod_TxtBx.Text)
            sw.WriteLine(H2H_Mod_TxtBx.Text)
            sw.WriteLine(Staff_Mod_TxtBx.Text)
            sw.WriteLine(Duration_Mod_TxtBx.Text)
            sw.WriteLine(Immunity_Mod_TxtBx.Text)
            sw.WriteLine(Bless_Mod_TxtBx.Text)
            sw.WriteLine(Heal_Mod_TxtBx.Text)
            sw.WriteLine(Freeze_Mod_TxtBx.Text)
            sw.WriteLine(MagicShield_Mod_TxtBx.Text)
            sw.WriteLine(Lightning_Mod_TxtBx.Text)
            sw.WriteLine(Fire_Mod_TxtBx.Text)
            sw.WriteLine(Barter_Mod_TxtBx.Text)
            sw.WriteLine(Perc_Mod_TxtBx.Text)
            sw.WriteLine(Stealth_Mod_TxtBx.Text)
            sw.WriteLine(Meditate_Mod_TxtBx.Text)
            sw.WriteLine(WWV_TxtBx.Text)
            sw.WriteLine(DefReduct_TxtBx.Text)
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
            Staff_TxtBx.Text = sr.ReadLine()
            Duration_TxtBx.Text = sr.ReadLine()
            Immunity_TxtBx.Text = sr.ReadLine()
            Bless_TxtBx.Text = sr.ReadLine()
            Heal_TxtBx.Text = sr.ReadLine()
            Freeze_TxtBx.Text = sr.ReadLine()
            MagicShield_TxtBx.Text = sr.ReadLine()
            Lightning_TxtBx.Text = sr.ReadLine()
            Fire_TxtBx.Text = sr.ReadLine()
            Barter_TxtBx.Text = sr.ReadLine()
            Perc_TxtBx.Text = sr.ReadLine()
            Stealth_TxtBx.Text = sr.ReadLine()
            Meditate_TxtBx.Text = sr.ReadLine()
            Prof_TxtBx.Text = sr.ReadLine()
            HP_Mod_TxtBx.Text = sr.ReadLine()
            End_Mod_TxtBx.Text = sr.ReadLine()
            Wis_Mod_TxtBx.Text = sr.ReadLine()
            Int_Mod_TxtBx.Text = sr.ReadLine()
            Agi_Mod_TxtBx.Text = sr.ReadLine()
            Str_Mod_TxtBx.Text = sr.ReadLine()
            Dagger_Mod_TxtBx.Text = sr.ReadLine()
            H2H_Mod_TxtBx.Text = sr.ReadLine()
            Staff_Mod_TxtBx.Text = sr.ReadLine()
            Duration_Mod_TxtBx.Text = sr.ReadLine()
            Immunity_Mod_TxtBx.Text = sr.ReadLine()
            Bless_Mod_TxtBx.Text = sr.ReadLine()
            Heal_Mod_TxtBx.Text = sr.ReadLine()
            Freeze_Mod_TxtBx.Text = sr.ReadLine()
            MagicShield_Mod_TxtBx.Text = sr.ReadLine()
            Lightning_Mod_TxtBx.Text = sr.ReadLine()
            Fire_Mod_TxtBx.Text = sr.ReadLine()
            Barter_Mod_TxtBx.Text = sr.ReadLine()
            Perc_Mod_TxtBx.Text = sr.ReadLine()
            Stealth_Mod_TxtBx.Text = sr.ReadLine()
            Meditate_Mod_TxtBx.Text = sr.ReadLine()
            WWV_TxtBx.Text = sr.ReadLine()
            DefReduct_TxtBx.Text = sr.ReadLine()
            sr.Close()
        End Using

        SetExpValues()
    End Sub


End Class