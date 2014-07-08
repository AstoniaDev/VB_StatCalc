Imports System.IO

Public Class frmMageCalc
    Public gLevel As Double
    Public FileName As String = ""
    Public MaxBase As Integer = 115
    Public Loading As Integer

    Private Function setEquipMod(ByVal EquipVal As Integer, ByVal MaxEquipVal As Integer) As Integer
        If EquipVal = 0 Then Return 0

        If EquipVal > Math.Floor(MaxEquipVal * 0.5) Then
            Return Math.Floor(MaxEquipVal * 0.5)
        Else
            Return EquipVal
        End If
    End Function

    Private Function setBlessMod(ByVal BlessVal As Integer, ByVal MaxEquipVal As Integer) As Integer

        If ShowBless_Chk.Checked = True Then
            If BlessVal > Math.Floor(MaxEquipVal * 0.5) Or Max_Bless_Chk.Checked = True Then
                Return Math.Floor(MaxEquipVal * 0.5)
            Else
                Return BlessVal
            End If
        End If
        Return 0
    End Function

    Private Function setStatVal(ByVal BaseStat As Integer, ByVal WIAS As Integer) As Integer
        Dim MaxStat As Integer
        'set silly low stats
        If BaseStat <= 7 Then
            MaxStat = 15
        Else
            MaxStat = BaseStat * 2
        End If
        'set IAS mods
        If WIAS > MaxStat Then
            Return MaxStat
        Else
            Return WIAS
        End If
    End Function

    Private Function GetExpValues(ByVal TxtBx As Integer, ByVal StartBase As Integer, ByVal StatFactor As Integer, ByVal ClassFactor As Double, ByVal PTMFactor As Integer) As Integer
        Dim Exp As Integer

        Exp = 0
        Exp = Int((((TxtBx - StartBase) + 6) ^ 3) * (StatFactor * ClassFactor))
        If TxtBx >= MaxBase Then Exp += 3000000 * PTMFactor
        Return Exp

    End Function

    Private Function GetLevelExp(ByVal TxtBx As Integer, ByVal StartBase As Integer, ByVal StatFactor As Integer, ByVal ClassFactor As Double, ByVal PTMFactor As Integer) As Integer
        Dim sVal, Exp As Integer

        Exp = 0
        sVal = StartBase
        Do Until sVal >= TxtBx
            Exp += GetExpValues(sVal, StartBase, StatFactor, ClassFactor, PTMFactor)
            sVal += 1
        Loop
        Return Exp

    End Function

    Private Sub SetExpValues()
        Dim StartBase, StatFactor, PTMFactor As Integer
        Dim ClassFactor As Double

        If Hardcore_Chk.Checked = True Then MaxBase = 122 Else MaxBase = 115

        StartBase = 10
        StatFactor = 3
        ClassFactor = 0.1
        PTMFactor = 0


        HP_ExpLbl.Text = GetExpValues(Int(HP_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        End_ExpLbl.Text = GetExpValues(Int(End_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Mana_ExpLbl.Text = GetExpValues(Int(Mana_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        ' WIAS EXP Calculations
        StatFactor = 2
        PTMFactor = 2
        Wis_ExpLbl.Text = GetExpValues(Int(Wis_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Int_ExpLbl.Text = GetExpValues(Int(Int_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Agi_ExpLbl.Text = GetExpValues(Int(Agi_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Str_ExpLbl.Text = GetExpValues(Int(Str_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        ' Skills and Spells EXP Calulations
        StartBase = 1
        StatFactor = 1
        PTMFactor = 1

        Dagger_ExpLbl.Text = GetExpValues(Int(Dagger_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        H2H_ExpLbl.Text = GetExpValues(Int(H2H_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Staff_ExpLbl.Text = GetExpValues(Int(Staff_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        Bless_ExpLbl.Text = GetExpValues(Int(Bless_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Heal_ExpLbl.Text = GetExpValues(Int(Heal_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Freeze_ExpLbl.Text = GetExpValues(Int(Freeze_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        MagicShield_ExpLbl.Text = GetExpValues(Int(MagicShield_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Lightning_ExpLbl.Text = GetExpValues(Int(Lightning_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Fire_ExpLbl.Text = GetExpValues(Int(Fire_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Pulse_ExpLbl.Text = GetExpValues(Int(Pulse_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Duration_ExpLbl.Text = GetExpValues(Int(Duration_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        Barter_ExpLbl.Text = GetExpValues(Int(Barter_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Perc_ExpLbl.Text = GetExpValues(Int(Perc_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Stealth_ExpLbl.Text = GetExpValues(Int(Stealth_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Meditate_ExpLbl.Text = GetExpValues(Int(Meditate_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        Immunity_ExpLbl.Text = GetExpValues(Int(Immunity_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        ' Profession EXP Calculations
        StatFactor = 3
        PTMFactor = 0
        Prof_ExpLbl.Text = GetExpValues(Int(Prof_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        SetMods()

    End Sub

    Private Sub SetMods()
        Dim WisMod, IntMod, AgiMod, StrMod As Integer
        Dim IAS, SSA, WIS, IIW, WWW As Integer
        Dim BlessBonus As Integer

        If ShowBless_Chk.Checked = True Then
            If Max_Bless_Chk.Checked = True Then BlessBonus = 1000 Else BlessBonus = Math.Floor(Int(BlessBonus_TxtBx.Text))
        End If

        HP_Mod_Lbl.Text = Int(HP_TxtBx.Text) + setEquipMod(Int(HP_Mod_TxtBx.Text), Int(HP_TxtBx.Text))
        End_Mod_Lbl.Text = Int(End_TxtBx.Text) + setEquipMod(Int(End_Mod_TxtBx.Text), Int(End_TxtBx.Text))
        Mana_Mod_Lbl.Text = Int(Mana_TxtBx.Text) + setEquipMod(Int(Mana_Mod_TxtBx.Text), Int(Mana_TxtBx.Text))

        Wis_Mod_Lbl.Text = Int(Wis_TxtBx.Text) + setEquipMod(Int(Wis_Mod_TxtBx.Text), Int(Wis_TxtBx.Text)) + Int(LDW_TxtBx.Text / 2) + Int(CW_TxtBx.Text)
        Int_Mod_Lbl.Text = Int(Int_TxtBx.Text) + setEquipMod(Int(Int_Mod_TxtBx.Text), Int(Int_TxtBx.Text)) + Int(LDW_TxtBx.Text / 2) + Int(CW_TxtBx.Text)
        Agi_Mod_Lbl.Text = Int(Agi_TxtBx.Text) + setEquipMod(Int(Agi_Mod_TxtBx.Text), Int(Agi_TxtBx.Text)) + Int(LDW_TxtBx.Text / 2) + Int(CW_TxtBx.Text)
        Str_Mod_Lbl.Text = Int(Str_TxtBx.Text) + setEquipMod(Int(Str_Mod_TxtBx.Text), Int(Str_TxtBx.Text)) + Int(LDW_TxtBx.Text / 2) + Int(CW_TxtBx.Text)

        WisMod = Int(Wis_Mod_Lbl.Text)
        IntMod = Int(Int_Mod_Lbl.Text)

        IIW = Math.Floor((IntMod + IntMod + WisMod) / 5)

        Bless_Mod_Lbl.Text = Int(Bless_TxtBx.Text) + setStatVal(Int(Bless_TxtBx.Text), IIW) + setEquipMod(Int(Bless_Mod_TxtBx.Text), Int(Bless_TxtBx.Text))
        Wis_Mod_Lbl.Text += setBlessMod(Int(Bless_Mod_Lbl.Text / 4), Int(Wis_TxtBx.Text))
        Int_Mod_Lbl.Text += setBlessMod(Int(Bless_Mod_Lbl.Text / 4), Int(Int_TxtBx.Text))
        Agi_Mod_Lbl.Text += setBlessMod(Int(Bless_Mod_Lbl.Text / 4), Int(Agi_TxtBx.Text))
        Str_Mod_Lbl.Text += setBlessMod(Int(Bless_Mod_Lbl.Text / 4), Int(Str_TxtBx.Text))

        WisMod = Int(Wis_Mod_Lbl.Text)
        IntMod = Int(Int_Mod_Lbl.Text)
        AgiMod = Int(Agi_Mod_Lbl.Text)
        StrMod = Int(Str_Mod_Lbl.Text)

        IAS = Math.Floor((IntMod + AgiMod + StrMod) / 5)
        SSA = Math.Floor((StrMod + StrMod + AgiMod) / 5)
        WIS = Math.Floor((WisMod + IntMod + StrMod) / 5)
        IIW = Math.Floor((IntMod + IntMod + WisMod) / 5)
        WWW = Math.Floor((WisMod + WisMod + WisMod) / 5)

        Dagger_Mod_Lbl.Text = Int(Dagger_TxtBx.Text) + setStatVal(Int(Dagger_TxtBx.Text), IAS) + setEquipMod(Int(Dagger_Mod_TxtBx.Text), Int(Dagger_TxtBx.Text))
        H2H_Mod_Lbl.Text = Int(H2H_TxtBx.Text) + setStatVal(Int(H2H_TxtBx.Text), SSA) + setEquipMod(Int(H2H_Mod_TxtBx.Text), Int(H2H_TxtBx.Text))
        Staff_Mod_Lbl.Text = Int(Staff_TxtBx.Text) + setStatVal(Int(Staff_TxtBx.Text), IAS) + setEquipMod(Int(Staff_Mod_TxtBx.Text), Int(Staff_TxtBx.Text))

        Bless_Mod_Lbl.Text = Int(Bless_TxtBx.Text) + setStatVal(Int(Bless_TxtBx.Text), IIW) + setEquipMod(Int(Bless_Mod_TxtBx.Text), Int(Bless_TxtBx.Text))
        Heal_Mod_Lbl.Text = Int(Heal_TxtBx.Text) + setStatVal(Int(Heal_TxtBx.Text), IIW) + setEquipMod(Int(Heal_Mod_TxtBx.Text), Int(Heal_TxtBx.Text))
        Freeze_Mod_Lbl.Text = Int(Freeze_TxtBx.Text) + setStatVal(Int(Freeze_TxtBx.Text), IIW) + setEquipMod(Int(Freeze_Mod_TxtBx.Text), Int(Freeze_TxtBx.Text))
        MagicShield_Mod_Lbl.Text = Int(MagicShield_TxtBx.Text) + setStatVal(Int(MagicShield_TxtBx.Text), IIW) + setEquipMod(Int(MagicShield_Mod_TxtBx.Text), Int(MagicShield_TxtBx.Text))
        Lightning_Mod_Lbl.Text = Int(Lightning_TxtBx.Text) + setStatVal(Int(Lightning_TxtBx.Text), IIW) + setEquipMod(Int(Lightning_Mod_TxtBx.Text), Int(Lightning_TxtBx.Text))
        Fire_Mod_Lbl.Text = Int(Fire_TxtBx.Text) + setStatVal(Int(Fire_TxtBx.Text), IIW) + setEquipMod(Int(Fire_Mod_TxtBx.Text), Int(Fire_TxtBx.Text))
        Pulse_Mod_Lbl.Text = Int(Pulse_TxtBx.Text) + setStatVal(Int(Pulse_TxtBx.Text), IIW) + setEquipMod(Int(Pulse_Mod_TxtBx.Text), Int(Pulse_TxtBx.Text))
        If Int(Duration_TxtBx.Text) > 0 Then Duration_Mod_Lbl.Text = Int(Duration_TxtBx.Text) + setStatVal(Int(Duration_TxtBx.Text), WIS) + setEquipMod(Int(Duration_Mod_TxtBx.Text), Int(Duration_TxtBx.Text)) Else Duration_Mod_Lbl.Text = 0

        Barter_Mod_Lbl.Text = Int(Barter_TxtBx.Text) + setStatVal(Int(Barter_TxtBx.Text), IIW) + setEquipMod(Int(Barter_Mod_TxtBx.Text), Int(Barter_TxtBx.Text))
        Perc_Mod_Lbl.Text = Int(Perc_TxtBx.Text) + setStatVal(Int(Perc_TxtBx.Text), IIW) + setEquipMod(Int(Perc_Mod_TxtBx.Text), Int(Perc_TxtBx.Text)) + Int(Thief_TxtBx.Text / 2)
        Stealth_Mod_Lbl.Text = Int(Stealth_TxtBx.Text) + setStatVal(Int(Stealth_TxtBx.Text), IIW) + setEquipMod(Int(Stealth_Mod_TxtBx.Text), Int(Perc_TxtBx.Text)) + Int(Thief_TxtBx.Text)
        Meditate_Mod_Lbl.Text = Int(Meditate_TxtBx.Text) + setStatVal(Int(Meditate_TxtBx.Text), WWW) + setEquipMod(Int(Meditate_Mod_TxtBx.Text), Int(Meditate_TxtBx.Text))
        Immunity_Mod_Lbl.Text = Int(Immunity_TxtBx.Text) + setStatVal(Int(Immunity_TxtBx.Text), WIS) + setEquipMod(Int(Immunity_Mod_TxtBx.Text), Int(Immunity_TxtBx.Text))

        SetLevel()
    End Sub

    Private Sub SetLevel()
        Dim TotalExp, ReqLevel As Double
        Dim StartBase, StatFactor, PTMFactor As Integer
        Dim ClassFactor As Double

        TotalExp = 0
        StartBase = 10
        StatFactor = 3
        ClassFactor = 0.1
        PTMFactor = 0

        TotalExp += GetLevelExp(Int(HP_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(End_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Mana_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        ' WIAS EXP Calculations
        StatFactor = 2
        PTMFactor = 2
        TotalExp += GetLevelExp(Int(Wis_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Int_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Agi_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Str_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        ' Skills and Spells EXP Calulations
        StartBase = 1
        StatFactor = 1
        PTMFactor = 1

        TotalExp += GetLevelExp(Int(Dagger_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(H2H_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Staff_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        TotalExp += GetLevelExp(Int(Bless_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Heal_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Freeze_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(MagicShield_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Lightning_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Fire_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Pulse_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Duration_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        TotalExp += GetLevelExp(Int(Barter_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Perc_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Stealth_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Meditate_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)
        TotalExp += GetLevelExp(Int(Immunity_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        ' Profession EXP Calculations
        StatFactor = 3
        PTMFactor = 0
        TotalExp += GetLevelExp(Int(Prof_TxtBx.Text), StartBase, StatFactor, ClassFactor, PTMFactor)

        ReqLevel = TotalExp ^ 0.25
        If ReqLevel = 0 Then Level_TxtBx.Text = 1 Else Level_TxtBx.Text = ReqLevel

        gLevel = ReqLevel
        SetExtraValues()
    End Sub

    Private Sub SetExtraValues()
        Dim ws, ws_max, ws_type As Integer
        Dim WisMod, IntMod, StrMod As Integer
        Dim MaxBlessMod, NoBlessDuration, MaxStatVal, BlessBonus As Integer
        Dim offVal, def_red As Integer
        Dim wv, av, spellaverage As Double

        'Calculate Bless
        If Max_Bless_Chk.Checked = True Then BlessBonus_TxtBx.Text = 1000

        If ShowBless_Chk.Checked = True Then
            WisMod = Int(Wis_Mod_Lbl.Text) - Math.Min(Int(BlessBonus_TxtBx.Text), (Int(Int(Wis_TxtBx.Text) / 2)))
            IntMod = Int(Int_Mod_Lbl.Text) - Math.Min(Int(BlessBonus_TxtBx.Text), (Int(Int(Int_TxtBx.Text) / 2)))
            StrMod = Int(Str_Mod_Lbl.Text) - Math.Min(Int(BlessBonus_TxtBx.Text), (Int(Int(Str_TxtBx.Text) / 2)))
        Else
            WisMod = Int(Wis_Mod_Lbl.Text)
            IntMod = Int(Int_Mod_Lbl.Text)
            StrMod = Int(Str_Mod_Lbl.Text)
        End If

        If Math.Floor(Int(Bless_TxtBx.Text) / 2) < Int(Bless_Mod_TxtBx.Text) Then MaxBlessMod = Math.Floor(Int(Bless_TxtBx.Text) / 2) Else MaxBlessMod = Int(Bless_Mod_TxtBx.Text)
        If Int(Bless_TxtBx.Text) <= 7 Then MaxStatVal = 15 Else MaxStatVal = Int(Bless_TxtBx.Text) * 2
        BlessBonus = Int((WisMod + IntMod + IntMod) / 5)
        If MaxStatVal < BlessBonus Then BlessBonus = MaxStatVal
        BlessBonus_TxtBx.Text = Int((Int(Bless_TxtBx.Text) + BlessBonus + MaxBlessMod) / 4)
        NxtBless.Text = Int((Int(BlessBonus_TxtBx.Text) + 1) * 4)
        OtherBlessBonus_TxtBx.Text = Int(Int(Bless_Mod_Lbl.Text) / 4)
        OtherNxtBless.Text = Int((Int(OtherBlessBonus_TxtBx.Text) + 1) * 4)

        'Calculate Duration
        If Int(Duration_TxtBx.Text) > 0 Then NoBlessDuration = Int(Int(Duration_TxtBx.Text) + ((WisMod + IntMod + StrMod) / 5) + MaxBlessMod) Else NoBlessDuration = 0
        SelfBless_TxtBx.Text = Math.Round((120 + (120 * NoBlessDuration / 35)) * 100) / 100
        OtherBless_TxtBx.Text = Math.Round((120 + (120 * Int(Duration_Mod_Lbl.Text) / 35)) * 100) / 100
        FlashDur_TxtBx.Text = Math.Round((2 + (2 * Int(Duration_Mod_Lbl.Text) / 35)) * 100) / 100
        FreezeDur_TxtBx.Text = Math.Round((4 + (4 * Int(Duration_Mod_Lbl.Text) / 35)) * 100) / 100

        'Determine Highest Weapon Skill
        ws_type = 0
        ws = Math.Max(Int(Dagger_Mod_Lbl.Text), Int(H2H_Mod_Lbl.Text))
        ws_max = Math.Max(ws, Int(Staff_Mod_Lbl.Text))
        If ws_max = Int(Dagger_Mod_Lbl.Text) Then
            ws_type = 0
        ElseIf ws_max = Int(H2H_Mod_Lbl.Text) Then
            ws_type = 1
        ElseIf ws_max = Int(Staff_Mod_Lbl.Text) Then
            ws_type = 2
        End If

        'Calculate Offence
        spellaverage = (Int(Lightning_Mod_Lbl.Text) + Int(Fire_Mod_Lbl.Text) + Int(Freeze_Mod_Lbl.Text) + Int(Heal_Mod_Lbl.Text) + Int(Bless_Mod_Lbl.Text) + Int(MagicShield_Mod_Lbl.Text) + Int(Pulse_Mod_Lbl.Text)) / 8
        offVal = Math.Floor(Int(ws_max) + spellaverage * 2 - Int(Level_TxtBx.Text))
        If offVal < 0 Then offVal = 0
        Offense_TxtBx.Text = offVal

        'Calculate Defence
        def_red = Math.Floor(Int(Level_TxtBx.Text) / 2)
        If def_red > 10 Then def_red = 10
        Defense_TxtBx.Text = Int(ws_max + (Int(MagicShield_Mod_Lbl.Text) * 2))

        'Calculate Weapon Value
        If ws_type = 0 Then
            If Int(Dagger_TxtBx.Text) >= 110 Then
                wv = 121
            Else
                wv = 1 + (Math.Floor(Int(Dagger_TxtBx.Text) / 10 + 1) * 10)
            End If
        ElseIf ws_type = 1 Then
            wv = 0
        ElseIf ws_type = 2 Then
            If Int(Staff_TxtBx.Text) >= 110 Then
                wv = 122
            Else
                wv = 2 + (Math.Floor(Int(Staff_TxtBx.Text) / 10 + 1) * 10)
            End If
        End If

        WV_TxtBx.Text = wv

        'Calculate Armour Value
        av = spellaverage * 17.5 / 20
        av = Math.Floor(av * 20) / 20
        AV_TxtBx.Text = av

        'Calculate Speed
        RawSpeed_TxtBx.Text = Int((Int(Agi_Mod_Lbl.Text) + Int(Agi_Mod_Lbl.Text) + Int(Str_Mod_Lbl.Text)) / 5) + (Int(Ath_TxtBx.Text) * 3)



        PostValidateData()

    End Sub
    Private Sub PreValidateDate()

        If Loading = 1 Then Exit Sub

        If Hardcore_Chk.Checked = True Then MaxBase = 122 Else MaxBase = 115

        If Int(HP_TxtBx.Text) > MaxBase Then HP_TxtBx.Text = MaxBase
        If Int(HP_TxtBx.Text) < 10 Then HP_TxtBx.Text = 10

        If Int(End_TxtBx.Text) > MaxBase Then End_TxtBx.Text = MaxBase
        If Int(End_TxtBx.Text) < 10 Then End_TxtBx.Text = 10

        If Int(Mana_TxtBx.Text) > MaxBase Then Mana_TxtBx.Text = MaxBase
        If Int(Mana_TxtBx.Text) < 10 Then Mana_TxtBx.Text = 10

        If Int(Wis_TxtBx.Text) > 250 Then Wis_TxtBx.Text = 250
        If Int(Wis_TxtBx.Text) < 10 Then Wis_TxtBx.Text = 10

        If Int(Int_TxtBx.Text) > 250 Then Int_TxtBx.Text = 250
        If Int(Int_TxtBx.Text) < 10 Then Int_TxtBx.Text = 10

        If Int(Agi_TxtBx.Text) > 250 Then Agi_TxtBx.Text = 250
        If Int(Agi_TxtBx.Text) < 10 Then Agi_TxtBx.Text = 10

        If Int(Str_TxtBx.Text) > 250 Then Str_TxtBx.Text = 250
        If Int(Str_TxtBx.Text) < 10 Then Str_TxtBx.Text = 10

        If Int(Dagger_TxtBx.Text) > 250 Then Dagger_TxtBx.Text = 250
        If Int(Dagger_TxtBx.Text) < 1 Then Dagger_TxtBx.Text = 1

        If Int(H2H_TxtBx.Text) > 250 Then H2H_TxtBx.Text = 250
        If Int(H2H_TxtBx.Text) < 1 Then H2H_TxtBx.Text = 1

        If Int(Staff_TxtBx.Text) > 250 Then Staff_TxtBx.Text = 250
        If Int(Staff_TxtBx.Text) < 1 Then Staff_TxtBx.Text = 1

        If Int(Duration_TxtBx.Text) > 250 Then Duration_TxtBx.Text = 250
        If Int(Duration_TxtBx.Text) < 0 Then Duration_TxtBx.Text = 0

        If Int(LDW_TxtBx.Text) > 30 Then LDW_TxtBx.Text = 30
        If Int(LDW_TxtBx.Text) < 0 Then LDW_TxtBx.Text = 0

        If Int(CW_TxtBx.Text) > 30 Then CW_TxtBx.Text = 30
        If Int(CW_TxtBx.Text) < 0 Then CW_TxtBx.Text = 0

        If Int(Ath_TxtBx.Text) > 30 Then Ath_TxtBx.Text = 30
        If Int(Ath_TxtBx.Text) < 0 Then Ath_TxtBx.Text = 0

        If Int(Thief_TxtBx.Text) > 30 Then Thief_TxtBx.Text = 30
        If Int(Thief_TxtBx.Text) < 0 Then Thief_TxtBx.Text = 0

        If Int(Immunity_TxtBx.Text) > 250 Then Immunity_TxtBx.Text = 250
        If Int(Immunity_TxtBx.Text) < 1 Then Immunity_TxtBx.Text = 1

        If Int(Bless_TxtBx.Text) > 250 Then Bless_TxtBx.Text = 250
        If Int(Bless_TxtBx.Text) < 1 Then Bless_TxtBx.Text = 1

        If Int(Heal_TxtBx.Text) > 250 Then Heal_TxtBx.Text = 250
        If Int(Heal_TxtBx.Text) < 1 Then Heal_TxtBx.Text = 1

        If Int(Freeze_TxtBx.Text) > 250 Then Freeze_TxtBx.Text = 250
        If Int(Freeze_TxtBx.Text) < 1 Then Freeze_TxtBx.Text = 1

        If Int(MagicShield_TxtBx.Text) > 250 Then MagicShield_TxtBx.Text = 250
        If Int(MagicShield_TxtBx.Text) < 1 Then MagicShield_TxtBx.Text = 1

        If Int(Lightning_TxtBx.Text) > 250 Then Lightning_TxtBx.Text = 250
        If Int(Lightning_TxtBx.Text) < 1 Then Lightning_TxtBx.Text = 1

        If Int(Fire_TxtBx.Text) > 250 Then Fire_TxtBx.Text = 250
        If Int(Fire_TxtBx.Text) < 1 Then Fire_TxtBx.Text = 1

        If Int(Pulse_TxtBx.Text) > 250 Then Pulse_TxtBx.Text = 250
        If Int(Pulse_TxtBx.Text) < 1 Then Pulse_TxtBx.Text = 1

        If Int(Barter_TxtBx.Text) > 250 Then Barter_TxtBx.Text = 250
        If Int(Barter_TxtBx.Text) < 1 Then Barter_TxtBx.Text = 1

        If Int(Perc_TxtBx.Text) > 250 Then Perc_TxtBx.Text = 250
        If Int(Perc_TxtBx.Text) < 1 Then Perc_TxtBx.Text = 1

        If Int(Stealth_TxtBx.Text) > 250 Then Stealth_TxtBx.Text = 250
        If Int(Stealth_TxtBx.Text) < 1 Then Stealth_TxtBx.Text = 1

        If Int(Meditate_TxtBx.Text) > 250 Then Meditate_TxtBx.Text = 250
        If Int(Meditate_TxtBx.Text) < 1 Then Meditate_TxtBx.Text = 1

        If Int(Prof_TxtBx.Text) > 100 Then Prof_TxtBx.Text = 100
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
        If Int(Pulse_Mod_TxtBx.Text) > Math.Floor(Int(Pulse_TxtBx.Text) / 2) Then Pulse_Mod_TxtBx.ForeColor = Color.Red Else Pulse_Mod_TxtBx.ForeColor = Color.Black
        If Int(Barter_Mod_TxtBx.Text) > Math.Floor(Int(Barter_TxtBx.Text) / 2) Then Barter_Mod_TxtBx.ForeColor = Color.Red Else Barter_Mod_TxtBx.ForeColor = Color.Black
        If Int(Perc_Mod_TxtBx.Text) > Math.Floor(Int(Perc_TxtBx.Text) / 2) Then Perc_Mod_TxtBx.ForeColor = Color.Red Else Perc_Mod_TxtBx.ForeColor = Color.Black
        If Int(Stealth_Mod_TxtBx.Text) > Math.Floor(Int(Stealth_TxtBx.Text) / 2) Then Stealth_Mod_TxtBx.ForeColor = Color.Red Else Stealth_Mod_TxtBx.ForeColor = Color.Black
        If Int(Meditate_Mod_TxtBx.Text) > Math.Floor(Int(Meditate_TxtBx.Text) / 2) Then Meditate_Mod_TxtBx.ForeColor = Color.Red Else Meditate_Mod_TxtBx.ForeColor = Color.Black

        'Whitestar Mod
        If gLevel > 200 Then
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

    Private Sub LDW_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LDW_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub CW_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CW_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Ath_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ath_TxtBx.Leave
        PreValidateDate()
    End Sub

    Private Sub Thief_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Thief_TxtBx.Leave
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

    Private Sub Pulse_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pulse_TxtBx.Leave
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

    Private Sub Pulse_Mod_TxtBx_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pulse_Mod_TxtBx.Leave
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

    Private Sub ShowBless_Chk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowBless_Chk.CheckedChanged
        PreValidateDate()
    End Sub

    Private Sub Max_Bless_Chk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Max_Bless_Chk.CheckedChanged
        PreValidateDate()
    End Sub

    Private Sub Hardcore_Chk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Hardcore_Chk.CheckedChanged
        PreValidateDate()
    End Sub

    Private Sub HP_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles HP_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(HP_TxtBx.Text) >= MaxBase Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                HP_TxtBx.Text = Int(HP_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(HP_TxtBx.Text) = 10 Then
                    MsgBox("Can not reduce below 10")
                    Exit Sub
                End If

                HP_TxtBx.Text = Int(HP_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub End_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles End_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(End_TxtBx.Text) >= MaxBase Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                End_TxtBx.Text = Int(End_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(End_TxtBx.Text) = 10 Then
                    MsgBox("Can not reduce below 10")
                    Exit Sub
                End If

                End_TxtBx.Text = Int(End_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub Mana_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Mana_Btn.MouseDown

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Mana_TxtBx.Text) >= MaxBase Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                Mana_TxtBx.Text = Int(Mana_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Mana_TxtBx.Text) = 10 Then
                    MsgBox("Can not reduce below 10")
                    Exit Sub
                End If

                Mana_TxtBx.Text = Int(Mana_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub Wis_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Wis_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Wis_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Int_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Agi_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Str_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Dagger_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(H2H_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Staff_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Duration_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                Duration_TxtBx.Text = Int(Duration_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Duration_TxtBx.Text) = 0 Then
                    MsgBox("Can not reduce below 0")
                    Exit Sub
                End If

                Duration_TxtBx.Text = Int(Duration_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub LDW_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LDW_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(LDW_TxtBx.Text) >= 30 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                If LDW_TxtBx.Text < 6 Then LDW_TxtBx.Text = 6 Else LDW_TxtBx.Text = (Int(LDW_TxtBx.Text / 3) + 1) * 3

            Case Windows.Forms.MouseButtons.Right
                If Int(LDW_TxtBx.Text) = 0 Then
                    MsgBox("Can not reduce below 0")
                    Exit Sub
                End If

                If LDW_TxtBx.Text <= 6 Then LDW_TxtBx.Text = 0 Else LDW_TxtBx.Text = (Math.Ceiling(LDW_TxtBx.Text / 3) - 1) * 3

        End Select

        SetExpValues()
    End Sub

    Private Sub CW_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CW_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(CW_TxtBx.Text) >= 30 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                If CW_TxtBx.Text < 6 Then CW_TxtBx.Text = 6 Else CW_TxtBx.Text = (Int(CW_TxtBx.Text / 3) + 1) * 3

            Case Windows.Forms.MouseButtons.Right
                If Int(CW_TxtBx.Text) = 0 Then
                    MsgBox("Can not reduce below 0")
                    Exit Sub
                End If

                If CW_TxtBx.Text <= 6 Then CW_TxtBx.Text = 0 Else CW_TxtBx.Text = (Math.Ceiling(CW_TxtBx.Text / 3) - 1) * 3

        End Select

        SetExpValues()
    End Sub


    Private Sub Ath_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Ath_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Ath_TxtBx.Text) >= 30 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                If Ath_TxtBx.Text < 6 Then Ath_TxtBx.Text = 6 Else Ath_TxtBx.Text = (Int(Ath_TxtBx.Text / 3) + 1) * 3

            Case Windows.Forms.MouseButtons.Right
                If Int(Ath_TxtBx.Text) = 0 Then
                    MsgBox("Can not reduce below 0")
                    Exit Sub
                End If

                If Ath_TxtBx.Text <= 6 Then Ath_TxtBx.Text = 0 Else Ath_TxtBx.Text = (Math.Ceiling(Ath_TxtBx.Text / 3) - 1) * 3

        End Select

        SetExpValues()
    End Sub

    Private Sub Thief_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Thief_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Thief_TxtBx.Text) >= 30 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                If Thief_TxtBx.Text < 6 Then Thief_TxtBx.Text = 6 Else Thief_TxtBx.Text = (Int(Thief_TxtBx.Text / 3) + 1) * 3

            Case Windows.Forms.MouseButtons.Right
                If Int(Thief_TxtBx.Text) = 0 Then
                    MsgBox("Can not reduce below 0")
                    Exit Sub
                End If

                If Thief_TxtBx.Text <= 6 Then Thief_TxtBx.Text = 0 Else Thief_TxtBx.Text = (Math.Ceiling(Thief_TxtBx.Text / 3) - 1) * 3

        End Select

        SetExpValues()
    End Sub

    Private Sub Immunity_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Immunity_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Immunity_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Bless_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Heal_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Freeze_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(MagicShield_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Lightning_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Fire_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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

    Private Sub Pulse_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Pulse_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Pulse_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
                    Exit Sub
                End If

                Pulse_TxtBx.Text = Int(Pulse_TxtBx.Text) + 1

            Case Windows.Forms.MouseButtons.Right
                If Int(Pulse_TxtBx.Text) = 1 Then
                    MsgBox("Can not reduce below 1")
                    Exit Sub
                End If

                Pulse_TxtBx.Text = Int(Pulse_TxtBx.Text) - 1

        End Select

        SetExpValues()
    End Sub

    Private Sub Barter_Btn_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Barter_Btn.MouseDown
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If Int(Barter_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Perc_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Stealth_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Meditate_TxtBx.Text) >= 200 Then
                    MsgBox("Cannot raise stats beyond MaxBase")
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
                If Int(Prof_TxtBx.Text) = 100 Then
                    MsgBox("Can not raise beyond 100")
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

        sFD.Filter = "Mage File|*.mag"
        sFD.FileName = FileName
        sFD.ShowDialog()

        If sFD.FileName = "" Then MsgBox("File not saved, Invalid file name.") : Exit Sub

        FileName = System.IO.Path.GetFileName(sFD.FileName)

        Using sw As StreamWriter = File.CreateText(sFD.FileName)
            sw.WriteLine(HP_TxtBx.Text)
            sw.WriteLine(End_TxtBx.Text)
            sw.WriteLine(Mana_TxtBx.Text)
            sw.WriteLine(Wis_TxtBx.Text)
            sw.WriteLine(Int_TxtBx.Text)
            sw.WriteLine(Agi_TxtBx.Text)
            sw.WriteLine(Str_TxtBx.Text)
            sw.WriteLine(Dagger_TxtBx.Text)
            sw.WriteLine(H2H_TxtBx.Text)
            sw.WriteLine(Staff_TxtBx.Text)
            sw.WriteLine(Bless_TxtBx.Text)
            sw.WriteLine(Heal_TxtBx.Text)
            sw.WriteLine(Freeze_TxtBx.Text)
            sw.WriteLine(MagicShield_TxtBx.Text)
            sw.WriteLine(Lightning_TxtBx.Text)
            sw.WriteLine(Fire_TxtBx.Text)
            sw.WriteLine(Pulse_TxtBx.Text)
            sw.WriteLine(Duration_TxtBx.Text)
            sw.WriteLine(Barter_TxtBx.Text)
            sw.WriteLine(Perc_TxtBx.Text)
            sw.WriteLine(Stealth_TxtBx.Text)
            sw.WriteLine(Meditate_TxtBx.Text)
            sw.WriteLine(Immunity_TxtBx.Text)
            sw.WriteLine(Prof_TxtBx.Text)
            sw.WriteLine(LDW_TxtBx.Text)
            sw.WriteLine(CW_TxtBx.Text)
            sw.WriteLine(Ath_TxtBx.Text)
            sw.WriteLine(Thief_TxtBx.Text)
            sw.WriteLine(HP_Mod_TxtBx.Text)
            sw.WriteLine(End_Mod_TxtBx.Text)
            sw.WriteLine(Mana_Mod_TxtBx.Text)
            sw.WriteLine(Wis_Mod_TxtBx.Text)
            sw.WriteLine(Int_Mod_TxtBx.Text)
            sw.WriteLine(Agi_Mod_TxtBx.Text)
            sw.WriteLine(Str_Mod_TxtBx.Text)
            sw.WriteLine(Dagger_Mod_TxtBx.Text)
            sw.WriteLine(H2H_Mod_TxtBx.Text)
            sw.WriteLine(Staff_Mod_TxtBx.Text)
            sw.WriteLine(Bless_Mod_TxtBx.Text)
            sw.WriteLine(Heal_Mod_TxtBx.Text)
            sw.WriteLine(Freeze_Mod_TxtBx.Text)
            sw.WriteLine(MagicShield_Mod_TxtBx.Text)
            sw.WriteLine(Lightning_Mod_TxtBx.Text)
            sw.WriteLine(Fire_Mod_TxtBx.Text)
            sw.WriteLine(Pulse_Mod_TxtBx.Text)
            sw.WriteLine(Duration_Mod_TxtBx.Text)
            sw.WriteLine(Barter_Mod_TxtBx.Text)
            sw.WriteLine(Perc_Mod_TxtBx.Text)
            sw.WriteLine(Stealth_Mod_TxtBx.Text)
            sw.WriteLine(Meditate_Mod_TxtBx.Text)
            sw.WriteLine(Immunity_Mod_TxtBx.Text)
            sw.WriteLine(ShowBless_Chk.Checked)
            sw.WriteLine(Max_Bless_Chk.Checked)
            sw.WriteLine(Hardcore_Chk.Checked)
            sw.Close()
        End Using
    End Sub

    Private Sub Load_Btn_Click(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Load_Btn.MouseUp
        Dim oFD As New OpenFileDialog

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left

                oFD.Filter = "Mage File|*.mag"
                oFD.ShowDialog()

                If oFD.FileName = "" Then MsgBox("File not Loaded, Invalid File Name") : Exit Sub

                FileName = System.IO.Path.GetFileName(oFD.FileName)
                Loading = 1

                Using sr As StreamReader = File.OpenText(oFD.FileName)
                    HP_TxtBx.Text = sr.ReadLine()
                    End_TxtBx.Text = sr.ReadLine()
                    Mana_TxtBx.Text = sr.ReadLine()
                    Wis_TxtBx.Text = sr.ReadLine()
                    Int_TxtBx.Text = sr.ReadLine()
                    Agi_TxtBx.Text = sr.ReadLine()
                    Str_TxtBx.Text = sr.ReadLine()
                    Dagger_TxtBx.Text = sr.ReadLine()
                    H2H_TxtBx.Text = sr.ReadLine()
                    Staff_TxtBx.Text = sr.ReadLine()
                    Bless_TxtBx.Text = sr.ReadLine()
                    Heal_TxtBx.Text = sr.ReadLine()
                    Freeze_TxtBx.Text = sr.ReadLine()
                    MagicShield_TxtBx.Text = sr.ReadLine()
                    Lightning_TxtBx.Text = sr.ReadLine()
                    Fire_TxtBx.Text = sr.ReadLine()
                    Pulse_TxtBx.Text = sr.ReadLine()
                    Duration_TxtBx.Text = sr.ReadLine()
                    Barter_TxtBx.Text = sr.ReadLine()
                    Perc_TxtBx.Text = sr.ReadLine()
                    Stealth_TxtBx.Text = sr.ReadLine()
                    Meditate_TxtBx.Text = sr.ReadLine()
                    Immunity_TxtBx.Text = sr.ReadLine()
                    Prof_TxtBx.Text = sr.ReadLine()
                    LDW_TxtBx.Text = sr.ReadLine()
                    CW_TxtBx.Text = sr.ReadLine()
                    Ath_TxtBx.Text = sr.ReadLine()
                    Thief_TxtBx.Text = sr.ReadLine()
                    HP_Mod_TxtBx.Text = sr.ReadLine()
                    End_Mod_TxtBx.Text = sr.ReadLine()
                    Mana_Mod_TxtBx.Text = sr.ReadLine()
                    Wis_Mod_TxtBx.Text = sr.ReadLine()
                    Int_Mod_TxtBx.Text = sr.ReadLine()
                    Agi_Mod_TxtBx.Text = sr.ReadLine()
                    Str_Mod_TxtBx.Text = sr.ReadLine()
                    Dagger_Mod_TxtBx.Text = sr.ReadLine()
                    H2H_Mod_TxtBx.Text = sr.ReadLine()
                    Staff_Mod_TxtBx.Text = sr.ReadLine()
                    Bless_Mod_TxtBx.Text = sr.ReadLine()
                    Heal_Mod_TxtBx.Text = sr.ReadLine()
                    Freeze_Mod_TxtBx.Text = sr.ReadLine()
                    MagicShield_Mod_TxtBx.Text = sr.ReadLine()
                    Lightning_Mod_TxtBx.Text = sr.ReadLine()
                    Fire_Mod_TxtBx.Text = sr.ReadLine()
                    Pulse_Mod_TxtBx.Text = sr.ReadLine()
                    Duration_Mod_TxtBx.Text = sr.ReadLine()
                    Barter_Mod_TxtBx.Text = sr.ReadLine()
                    Perc_Mod_TxtBx.Text = sr.ReadLine()
                    Stealth_Mod_TxtBx.Text = sr.ReadLine()
                    Meditate_Mod_TxtBx.Text = sr.ReadLine()
                    Immunity_Mod_TxtBx.Text = sr.ReadLine()
                    ShowBless_Chk.Checked = sr.ReadLine()
                    Max_Bless_Chk.Checked = sr.ReadLine()
                    Hardcore_Chk.Checked = sr.ReadLine()
                    sr.Close()
                End Using
            Case Windows.Forms.MouseButtons.Right
                oFD.Filter = "Widget Stat Calc file|*"
                oFD.ShowDialog()

                If oFD.FileName = "" Then MsgBox("File not Loaded, Invalid File Name") : Exit Sub

                FileName = System.IO.Path.GetFileName(oFD.FileName)
                Loading = 1

                Using sr As StreamReader = File.OpenText(oFD.FileName)
                    sr.ReadLine()
                    Hardcore_Chk.Checked = sr.ReadLine()
                    Max_Bless_Chk.Checked = sr.ReadLine()
                    Ath_TxtBx.Text = Int(sr.ReadLine())
                    HP_TxtBx.Text = sr.ReadLine()
                    HP_Mod_TxtBx.Text = sr.ReadLine()
                    End_TxtBx.Text = sr.ReadLine()
                    End_Mod_TxtBx.Text = sr.ReadLine()
                    Mana_TxtBx.Text = sr.ReadLine()
                    Mana_Mod_TxtBx.Text = sr.ReadLine()
                    Wis_TxtBx.Text = sr.ReadLine()
                    Wis_Mod_TxtBx.Text = sr.ReadLine()
                    Int_TxtBx.Text = sr.ReadLine()
                    Int_Mod_TxtBx.Text = sr.ReadLine()
                    Agi_TxtBx.Text = sr.ReadLine()
                    Agi_Mod_TxtBx.Text = sr.ReadLine()
                    Str_TxtBx.Text = sr.ReadLine()
                    Str_Mod_TxtBx.Text = sr.ReadLine()
                    Dagger_TxtBx.Text = sr.ReadLine()
                    Dagger_Mod_TxtBx.Text = sr.ReadLine()
                    H2H_TxtBx.Text = sr.ReadLine()
                    H2H_Mod_TxtBx.Text = sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    Staff_TxtBx.Text = sr.ReadLine()
                    Staff_Mod_TxtBx.Text = sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    Bless_TxtBx.Text = sr.ReadLine()
                    Bless_Mod_TxtBx.Text = sr.ReadLine()
                    Heal_TxtBx.Text = sr.ReadLine()
                    Heal_Mod_TxtBx.Text = sr.ReadLine()
                    Freeze_TxtBx.Text = sr.ReadLine()
                    Freeze_Mod_TxtBx.Text = sr.ReadLine()
                    MagicShield_TxtBx.Text = sr.ReadLine()
                    MagicShield_Mod_TxtBx.Text = sr.ReadLine()
                    Lightning_TxtBx.Text = sr.ReadLine()
                    Lightning_Mod_TxtBx.Text = sr.ReadLine()
                    Fire_TxtBx.Text = sr.ReadLine()
                    Fire_Mod_TxtBx.Text = sr.ReadLine()
                    Pulse_TxtBx.Text = sr.ReadLine()
                    Pulse_Mod_TxtBx.Text = sr.ReadLine()
                    Duration_TxtBx.Text = sr.ReadLine()
                    Duration_Mod_TxtBx.Text = sr.ReadLine()
                    Barter_TxtBx.Text = sr.ReadLine()
                    Barter_Mod_TxtBx.Text = sr.ReadLine()
                    Perc_TxtBx.Text = sr.ReadLine()
                    Perc_Mod_TxtBx.Text = sr.ReadLine()
                    Stealth_TxtBx.Text = sr.ReadLine()
                    Stealth_Mod_TxtBx.Text = sr.ReadLine()
                    sr.ReadLine()
                    sr.ReadLine()
                    Meditate_TxtBx.Text = sr.ReadLine()
                    Meditate_Mod_TxtBx.Text = sr.ReadLine()
                    Immunity_TxtBx.Text = sr.ReadLine()
                    Immunity_Mod_TxtBx.Text = sr.ReadLine()
                    Prof_TxtBx.Text = sr.ReadLine()
                End Using
        End Select
        Loading = 0

        SetExpValues()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
End Class