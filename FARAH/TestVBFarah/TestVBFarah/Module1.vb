﻿Option Explicit Off
Module Module1


    'Farah CHANGE THE VALUES NEEDED FROM HERE
    Dim SPAN = 35
    Dim NomBayWidth = 20
    Dim ContainerHeight = 8.5
    Dim Purling = 1.2
    Dim ExpCategory = "B"


    'LEAVE THOSE FOR INTERNAL DECLERATION
    Dim Wind_GableI50 = 0
    Dim Wind_GableI51 = 0
    Dim Wind_GableI55 = 0
    Dim Wind_GableI56 = 0
    Dim Wind_GableI60 = 0
    Dim Wind_GableI61 = 0
    Dim Wind_GableI65 = 0
    Dim Wind_GableK81 = 0
    Dim Wind_GableK84 = 0
    Dim Wind_GableK92 = 0
    Dim Wind_GableK95 = 0


    Dim Snow_GableE17 = 0
    Dim Snow_GableF51 = 0
    Dim Snow_GableF52 = 0
    Dim Snow_GableE41 = 0
    Dim Snow_GableE45 = 0
    Dim Snow_GableE46 = 0
    Dim Snow_GableK75 = 0
    Dim Snow_GableK72 = 0
    Dim Snow_GableF56 = 0
    Dim Snow_GableF57 = 0

    Dim Gable_Gambrel_Purlin_DesignE58 = 0
    Dim Gable_Gambrel_Purlin_DesignE60 = 0


    Public Sub Main(args() As String)
        'Your code goes here
        'Console.WriteLine("Hello, world!")


        USWindPressureGable_30()
        USSnowLoadGable_30()
        USFindGambrelRafter()
        USFindGableRafter_30()
    End Sub

    Sub USFindGambrelRafter()

        '1) * Roof Slope,q (°) =
        E52 = 20
        E58 = 6.047 : Gable_Gambrel_Purlin_DesignE58 = E58
        E60 = 1.14 : Gable_Gambrel_Purlin_DesignE60 = E60
        E61 = E60 * Purling

        E65 = 20
        E66 = E65 * Purling


        E71 = Math.Max(Wind_GableI50, Math.Max(Wind_GableI51, Math.Max(Wind_GableI55, Math.Max(Wind_GableI56, Math.Max(Wind_GableI60, Wind_GableI65)))))
        E72 = Purling * E71

        E74 = Math.Min(Wind_GableI50, Math.Min(Wind_GableI51, Math.Min(Wind_GableI55, Math.Min(Wind_GableI56, Math.Min(Wind_GableI60, Wind_GableI65)))))
        E75 = Purling * E74


        If (Snow_GableE17 = 0) Then
            E79 = 0
        Else
            E79 = Math.Max(Snow_GableF51, Math.Max(Snow_GableF52, Math.Max(Snow_GableF56, Snow_GableF57)))
        End If

        E80 = Purling * E79

        E84 = (E58 + E61 + E66) * Math.Cos(E52 * Math.PI / 180)
        E85 = (E58 + E61 + E80) * Math.Cos(E52 * Math.PI / 180)
        E86 = (E58 + E61) * Math.Cos(E52 * Math.PI / 180) + E72
        E87 = (E58 + E61) * Math.Cos(E52 * Math.PI / 180) + 0.75 * E72 + 0.75 * (E66 * Math.Cos(E52 * Math.PI / 180))
        E88 = (E58 + E61) * Math.Cos(E52 * Math.PI / 180) + 0.75 * E72 + 0.75 * (E80 * Math.Cos(E52 * Math.PI / 180))
        E89 = Math.Abs(0.6 * (E58 + E61) * Math.Cos(E52 * Math.PI / 180) + E75)

        E91 = Math.Max(E84, Math.Max(E85, Math.Max(E86, Math.Max(E87, Math.Max(E88, E89)))))

        If (E91 <= 139.61) Then
            E93 = "8.0x4.0C14"
        Else
            If (E91 > 139.61 And E91 <= 245) Then
                E93 = "8.0x4.0C12"
            Else
                E93 = "NA"
            End If
        End If

    End Sub


    Sub USWindPressureGable_30()
        'If OutputCalcs Then Print #1, "USA Gable wind prewgbure calculation"
        'If OutputCalcs Then Print #1, "Basic Wind Speed selected =  " & Region

        wgbE9 = SPAN
        wgbE10 = NomBayWidth
        wgbE11 = ContainerHeight
        wgbE12 = Purling

        wgbE13 = wgbE11 + 0.5 * wgbE9 * Math.Tan(0.5235987756)
        'see "wgbE13=" & wgbE13 & " V heightoa=" & heightOA
        wgbE17 = 110
        'If OutputCalcs Then Print #1, "Exposure Category selected =  " & TerCatDescr(TerrainIdx)
        'If OutputCalcs Then Print #1, "Directionality Factor (Kd) = 0.85"
        wgbE18 = 0.85 'Directionality Factor (Kd)
        wgbE19 = ExpCategory
        'If OutputCalcs Then Print #1, "Risk Factor selected =  " & ImportanceDescr(ImportanceIndex)
        'If OutputCalcs Then Print #1, "Topographic factor (Kzt) =1 "
        wgbE20 = 1
        wgbE23 = (wgbE11 + wgbE13) * 0.6 * 0.5

        Select Case wgbE19
            Case "B"   'exposure B
                wgbU21 = 0.3
                wgbV21 = 320
                wgbW21 = 0.333333
                wgbX21 = 30
                wgbY21 = 7
                wgbZ21 = 1200
            Case "C"   'exposure C
                wgbU21 = 0.2
                wgbV21 = 500
                wgbW21 = 0.2
                wgbX21 = 15
                wgbY21 = 9.5
                wgbZ21 = 900
            Case "D"   'exposure D
                wgbU21 = 0.15
                wgbV21 = 650
                wgbW21 = 0.125
                wgbX21 = 7
                wgbY21 = 11.5
                wgbZ21 = 700
            Case Else
                ' "Terrain Value " & TerrainIdx & " out of range"
                wgbU21 = 0.15
                wgbV21 = 650
                wgbW21 = 0.125
                wgbX21 = 7
                wgbY21 = 11.5
                wgbZ21 = 700
        End Select    'terrain value
        If wgbE23 > wgbX21 Then wgbT24 = wgbE23 Else wgbT24 = wgbX21

        wgbE24 = wgbV21 * ((wgbT24 / 33) ^ (wgbW21))
        wgbE25 = Math.Sqrt(1 / (1 + 0.63 * (((wgbE10 + wgbE11) / wgbE24) ^ 0.63)))
        wgbE26 = wgbU21 * ((33 / wgbE23) ^ (1 / 6))
        wgbE27 = 3.4
        wgbE28 = 3.4
        wgbE30 = 0.925 * (1 + 1.7 * wgbE28 * wgbE26 * wgbE25) / (1 + 1.7 * wgbE27 * wgbE26)

        'Awgbumed
        'Enclosure clawgbification :  Open structure
        'Wind flow   Clear
        'Internal prewgbure coeffecient (CGpi) =  0.00

        wgbE38 = 0.5 * (wgbE11 + wgbE13)
        wgbE39 = 2.01 * ((15 / wgbZ21) ^ (2 / wgbY21))
        wgbE40 = 2.01 * ((wgbE38 / wgbZ21) ^ (2 / wgbY21))
        If wgbE38 < 15 Then
            wgbE41 = wgbE39
        Else
            wgbE41 = wgbE40
        End If

        'Velocity Prewgbure (qh):
        wgbE44 = 0.00256 * wgbE41 * wgbE20 * wgbE18 * (wgbE17 ^ 2)
        '=0.00256    *  E41   *   E20  *  E18     *(E17^2)
        'qh = 0.00256*Kz*Kzt*Kd*(V^2) =
        Wu = wgbE44
        WindPrewgbure = Wu   'Ultimate Dynamic Wind Prewgbure
        ' "Windprewgbure = " & WindPrewgbure

        'Purlin load
        'Roof load case 1 - Wind 0? - Case A
        wgbE33 = "CLEAR"
        wgbG50 = 0
        wgbG51 = 0
        Select Case wgbE33
            Case "CLEAR"
                wgbG50 = 1.3     'Zone 2 for 30deg pitch only !
                wgbG51 = 0.3  'Zone 3 for 30deg pitch only !
            Case "OBSTRUCTED"
                wgbG50 = -0.7  'Zone 2 for 30deg pitch only !
                wgbG51 = -0.7      'Zone 3 for 30deg pitch only !
        End Select
        wgbF50 = wgbE38
        wgbF51 = wgbE38
        wgbH50 = wgbE44
        wgbH51 = wgbE44
        wgbI50 = wgbG50 * wgbH50 * wgbE30 : Wind_GableI50 = wgbI50
        '=$G$50*$H$50*$E$30
        wgbI51 = wgbG51 * wgbH51 * wgbE30 : Wind_GableI51 = wgbI51
        wgbJ50 = wgbE12 * wgbI50
        wgbJ51 = wgbE12 * wgbI51

        ' "Purlin load (lb/ft') = " & wgbJ50
        'multiply this by length of purlin ie baywidth to get equally distributed load per purlin

        '-Roof load case 2 - Wind 0? - Case B
        wgbG55 = 0
        wgbG56 = 0
        Select Case wgbE33
            Case "CLEAR"
                wgbG55 = -0.1   'Zone 2 for 10deg pitch only !
                wgbG56 = -0.9  'Zone 3 for 10deg pitch only !
            Case "OBSTRUCTED"
                wgbG55 = -0.2    'Zone 2 for 10deg pitch only !
                wgbG56 = -1.1  'Zone 3 for 10deg pitch only !
        End Select
        wgbF55 = wgbE38
        wgbF56 = wgbE38
        wgbH55 = wgbE44
        wgbH56 = wgbE44
        wgbI55 = wgbG55 * wgbH55 * wgbE30 : Wind_GableI55 = wgbI55
        wgbI56 = wgbG56 * wgbH56 * wgbE30 : Wind_GableI56 = wgbI56
        wgbJ55 = wgbE12 * wgbI55
        wgbJ56 = wgbE12 * wgbI56
        wgbG60 = 0
        wgbG61 = 0
        '-Roof load case 3 - Wind 90? - Case A
        Select Case wgbE33
            Case "CLEAR"
                wgbG60 = -0.8
                wgbG61 = -0.6
            Case "OBSTRUCTED"
                wgbG60 = -1.2
                wgbG61 = -0.9
        End Select
        wgbF60 = wgbE38
        wgbF61 = wgbE38
        wgbH60 = wgbE44
        wgbH61 = wgbE44
        wgbI60 = wgbG60 * wgbH60 * wgbE30 : Wind_GableI60 = wgbI60
        wgbI61 = wgbG61 * wgbH61 * wgbE30 : Wind_GableI61 = wgbI61
        wgbJ60 = wgbE12 * wgbI60
        wgbJ61 = wgbE12 * wgbI61

        '-Roof load case 4 - Wind 90? - Case B
        wgbG65 = 0
        wgbG66 = 0
        Select Case wgbE33
            Case "CLEAR"
                wgbG65 = 0.8
                wgbG66 = 0.5
            Case "OBSTRUCTED"
                wgbG65 = 0.5
                wgbG66 = 0.5
        End Select
        wgbF65 = wgbE38
        wgbF66 = wgbE38
        wgbH65 = wgbE44
        wgbH66 = wgbE44
        wgbI65 = wgbG65 * wgbH65 * wgbE30 : Wind_GableI65 = wgbI65
        '      =$G$65 * $H$65 * $E$30
        wgbI66 = wgbG66 * wgbH66 * wgbE30
        wgbJ65 = wgbE12 * wgbI65
        wgbJ66 = wgbE12 * wgbI66


        'Uplift for footings
        wgbE75 = 0.5 * wgbE9 * wgbE10 / Math.Cos(0.5235987756) 'Zones 2 & 3 surface area


        wgbU73 = 0.5 * (wgbE13 - wgbE11)
        wgbU75 = 0.25 * (wgbE9)

        '-Roof load case 1 - Wind 0? - Case A
        'zone 2 and 3
        wgbF78 = wgbE75 * wgbI50 / 1000 'Total Net force (Kips) Zone 2
        wgbG78 = wgbE75 * wgbI51 / 1000 'Total Net force (Kips) Zone 3
        wgbF79 = wgbF78 * Math.Cos(30 * Math.PI / 180) * wgbU75    'Moment @ axis A (kips.ft) [Vl. forces] Zone 2
        wgbG79 = wgbG78 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis A (kips.ft) [Vl. forces] Zone 3
        wgbF80 = wgbF78 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 2
        wgbG80 = -wgbG78 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 3
        wgbF81 = wgbF79 + wgbG79 + wgbF80 + wgbG80
        wgbF82 = wgbF78 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis B (kips.ft)        [Vl. forces]
        wgbG82 = wgbG78 * Math.Cos(30 * Math.PI / 180) * wgbU75
        wgbF83 = -wgbF78 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces]
        wgbG83 = wgbG78 * Math.Sin(30 * Math.PI / 180) * wgbU73
        wgbF84 = wgbF82 + wgbF83 + wgbG82 + wgbG83
        wgbK81 = 0.5 * wgbF81 / wgbE9 * 1000 : Wind_GableK81 = wgbK81    'RB1
        wgbK84 = 0.5 * wgbF84 / wgbE9 * 1000 : Wind_GableK84 = wgbK84 'RA1
        ' "Ra1=" & wgbK84

        'Now do all this again for 3 more cases

        '-Roof load case 2 - Wind 0? - Case B
        'zone 2 and 3
        wgbF89 = wgbE75 * wgbI55 / 1000 'Total Net force (Kips) Zone 2
        wgbG89 = wgbE75 * wgbI56 / 1000 'Total Net force (Kips) Zone 3
        wgbF90 = wgbF89 * Math.Cos(30 * Math.PI / 180) * wgbU75    'Moment @ axis A (kips.ft) [Vl. forces] Zone 2
        wgbG90 = wgbG89 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis A (kips.ft) [Vl. forces] Zone 3
        wgbF91 = wgbF89 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 2
        wgbG91 = -wgbG89 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 3
        wgbF92 = wgbF90 + wgbG90 + wgbF91 + wgbG91
        wgbF93 = wgbF89 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis B (kips.ft)        [Vl. forces]
        wgbG93 = wgbG89 * Math.Cos(30 * Math.PI / 180) * wgbU75
        wgbF94 = -wgbF89 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces]
        wgbG94 = wgbG89 * Math.Sin(30 * Math.PI / 180) * wgbU73
        wgbF95 = wgbF93 + wgbF94 + wgbG93 + wgbG94
        wgbK92 = 0.5 * wgbF92 / wgbE9 * 1000 : Wind_GableK92 = wgbK92 'RB2
        wgbK95 = 0.5 * wgbF95 / wgbE9 * 1000 : Wind_GableK95 = wgbK95 'RA2
        ' "Ra2=" & wgbK95

        '-Roof load case 3 - Wind 90? - Case A
        'zone 1-2 and 1-3
        wgbE101 = 0.25 * wgbE9 * (wgbE11 + wgbE13) / Math.Cos(0.5235987756)            'Zones 1-2 & 1-3 surface area =
        wgbE103 = 0.5 * wgbE9 * (wgbE10 - 0.5 * (wgbE11 + wgbE13)) / Math.Cos(0.5235987756)  'Zones 2-2 & 2-3 surface area =

        wgbF106 = wgbE101 * wgbI60 / 1000 'Total Net force (Kips) Zone 1-2
        wgbG106 = wgbE101 * wgbI60 / 1000 'Total Net force (Kips) Zone 1-3
        wgbF107 = wgbF106 * Math.Cos(30 * Math.PI / 180) * wgbU75    'Moment @ axis A (kips.ft) [Vl. forces] Zone 1-2
        wgbG107 = wgbG106 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis A (kips.ft) [Vl. forces] Zone 1-3
        wgbF108 = wgbF106 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 1-2
        wgbG108 = -wgbG106 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 1-3
        wgbF109 = wgbF107 + wgbG107 + wgbF108 + wgbG108

        wgbF111 = wgbE103 * wgbI61 / 1000  'Total Net force (Kips)
        wgbG111 = wgbE103 * wgbI61 / 1000
        wgbF112 = wgbF111 * Math.Cos(30 * Math.PI / 180) * wgbU75  'Moment @ axis A or B (kips.ft)  [Vl. forces]
        wgbG112 = wgbG111 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75)
        wgbF113 = wgbF111 * Math.Sin(30 * Math.PI / 180) * wgbU73    '[Hz. forces]
        wgbG113 = -wgbG111 * Math.Sin(30 * Math.PI / 180) * wgbU73    '[Hz. forces]
        wgbF114 = wgbF112 + wgbF113 + wgbG112 + wgbG113
        wgbK109 = wgbF109 / wgbE9 'R1 =
        wgbK114 = wgbF114 / wgbE9 'R2 =
        wgbK117 = ((wgbK109 * (wgbE10 - 0.5 * wgbE11)) + (wgbK114 * 0.5 * (wgbE10 - wgbE11))) / wgbE10 * 1000 'R1AB =
        wgbK119 = ((wgbK114 * 0.5 * (wgbE10 + wgbE11) + 0.5 * wgbK109 * wgbE11)) / wgbE10 * 1000            'R2AB =
        ' "R2AB case3=" & wgbK119

        '-Roof load case 4 - Wind 90? - Case B
        'zone 1-2 and 1-3
        wgbF123 = wgbE101 * wgbI65 / 1000 'Total Net force (Kips) Zone 1-2
        wgbG123 = wgbE101 * wgbI65 / 1000  'Total Net force (Kips) Zone 1-3    '?? Should use I66 ?
        wgbF124 = wgbF123 * Math.Cos(30 * Math.PI / 180) * wgbU75    'Moment @ axis A (kips.ft) [Vl. forces] Zone 1-2
        wgbG124 = wgbG123 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis A (kips.ft) [Vl. forces] Zone 1-3
        wgbF125 = wgbF123 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 1-2
        wgbG125 = -wgbG123 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 1-3
        wgbF126 = wgbF124 + wgbG124 + wgbF125 + wgbG125

        wgbF128 = wgbE103 * wgbI66 / 1000  'Total Net force (Kips)
        wgbG128 = wgbE103 * wgbI66 / 1000
        wgbF129 = wgbF128 * Math.Cos(30 * Math.PI / 180) * wgbU75  'Moment @ axis A or B (kips.ft)  [Vl. forces]
        wgbG129 = wgbG128 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75)
        wgbF130 = wgbF128 * Math.Sin(30 * Math.PI / 180) * wgbU73    '[Hz. forces]
        wgbG130 = -wgbG128 * Math.Sin(30 * Math.PI / 180) * wgbU73    '[Hz. forces]
        wgbF131 = wgbF129 + wgbF130 + wgbG129 + wgbG130
        wgbK126 = wgbF126 / wgbE9 'R1 =
        wgbK131 = wgbF131 / wgbE9 'R2 =
        wgbK134 = ((wgbK126 * (wgbE10 - 0.5 * wgbE11)) + (wgbK131 * 0.5 * (wgbE10 - wgbE11))) / wgbE10 * 1000 'R1AB =
        wgbK136 = ((wgbK131 * 0.5 * (wgbE10 + wgbE11) + 0.5 * wgbK126 * wgbE11)) / wgbE10 * 1000            'R2AB =
        ' "R2AB case 4=" & wgbK136

        wgbE138 = Math.Min(wgbK136, Math.Min(wgbK134, Math.Min(wgbK119, Math.Min(wgbK117, Math.Min(wgbK95, Math.Min(wgbK92, Math.Min(wgbK84, wgbK81)))))))    'Max. Vl. upthrust on rafter support
        UpliftAtCorner = 1 - wgbE138
        '=MIN(K136     ,K134,   K119,   K117,   K95,   K92,   K84,   K81)
        wgbE139 = Math.Max(wgbK136, Math.Max(wgbK134, Math.Max(wgbK119, Math.Max(wgbK117, Math.Max(wgbK95, Math.Max(wgbK92, Math.Max(wgbK84, wgbK81)))))))    'Max. Vl. down prewgb on rafter support
        Console.WriteLine("wgbE139 --> " + wgbE139.ToString())
        ' "wgbE138=" & wgbE138
        UpLift = wgbE138
        'If OutputCalcs Then Print #1, "UpLift = WindPrewgbure * 0.9 giving " & UpLift
        UpLoad = UpLift - 3    'factor of 3 deducted for weight of roof when used for footings
        'If OutputCalcs Then Print #1, "Upload = UpLift - 3 =" & UpLoad & "  ie factor of 3 deducted for weight of roof when used for footings"
        ColumnRoofTrib = SPAN * BayWidth * 0.25
        'If OutputCalcs Then Print #1, "Column Tributary area =" & ColumnRoofTrib & " as Span x Bay width / 4"
        FootingColumnUpload = UpLoad * ColumnRoofTrib
        'If OutputCalcs Then Print #1, "FootingColumnUpload = Upload * ColumnRoofTrib = " & FootingColumnUpload

    End Sub


    Sub USSnowLoadGable_30()
        'AddNotification "No US Gable snow load yet"

        sgbE9 = SPAN 'B1 =
        sgbE10 = NomBayWidth 'L =
        sgbE11 = ContainerHeight
        sgbE12 = Purling  'Eave (container) height =
        sgbE13 = sgbE11 + 0.5 * SPAN * Math.Tan(0.5235987756)
        'sgbE14 = heightOA 'Apex height =

        sgbE16 = "I" 'Building Risk Category:
        sgbE17 = 20  'Ground Snow load (Pg) in psf
        Snow_GableE17 = sgbE17
        sgbE18 = Math.Min((0.13 * sgbE17 + 14), 30) 'Density of snow (g)
        sgbE19 = ExpCategory 'Exposure Category (Terrain Type):
        sgbE20 = "Fully Exposed" 'Exposure condition [Table 7-2] :
        'Exposure factor (Ce) [Table 7-2] :
        sgbE21 = 0
        Select Case sgbE19
            Case "B"   'exposure B
                Select Case sgbE20
                    Case "Fully Exposed" : sgbE21 = 0.9
                    Case "Partially Exposed" : sgbE21 = 1
                    Case "Sheltered" : sgbE21 = 1.2
                End Select
            Case "C"   'exposure C
                Select Case sgbE20
                    Case "Fully Exposed" : sgbE21 = 0.9
                    Case "Partially Exposed" : sgbE21 = 1
                    Case "Sheltered" : sgbE21 = 1.1
                End Select
            Case "D"   'exposure D
                Select Case sgbE20
                    Case "Fully Exposed" : sgbE21 = 0.8
                    Case "Partially Exposed" : sgbE21 = 0.9
                    Case "Sheltered" : sgbE21 = 1
                End Select
            Case Else
                'see "Terrain Value " & TerrainIdx & " out of range"
                Select Case sgbE20
                    Case "Fully Exposed" : sgbE21 = 0.8
                    Case "Partially Exposed" : sgbE21 = 0.9
                    Case "Sheltered" : sgbE21 = 1
                End Select
        End Select    'terrain value
        sgbE22 = "Unheated Structure" 'Thermal condition [Table 7-3] :
        sgbE23 = 1.2 'Thermal factor (Ct) [Table 7-3] :
        'Importance factor (Is) [Table 1.5-2] :   =VLOOKUP(E17,T24:U27,2)
        Select Case sgbE16
            Case "I" : sgbE24 = 0.8
            Case "II" : sgbE24 = 1
            Case "III" : sgbE24 = 1.1
            Case Else : sgbE24 = 1.2
        End Select

        sgbE30 = sgbE24 * sgbE17 : sgbE30 = 0 'Cancelled   'Min. load for low slope roofs [Sect 7.3.4] (Pfmin) :
        sgbE31 = 0.7 * sgbE21 * sgbE23 * sgbE24 * sgbE17 'Flat roof snow load [Sect 7.3.4] (Pf) :
        'Cold roof Snow factor (Ct>1.0)
        sgbE33 = "Slippery"  'Roof surface type
        sgbE34 = "Ventilated" 'Ventilation
        sgbE35 = 0.73 'Roof slope factor [Fig 7-2c (dashed line)] (Cs):
        sgbE37 = sgbE35 * sgbE31 'Balanced sloped snow load (Ps):
        sgbE38 = 1.732   'Slope of roof = 1/tan(10?) :
        sgbE40 = 0.3 * sgbE37 'Unbalanced Load(ps)(windward):
        sgbE41 = sgbE37 : Snow_GableE41 = sgbE41 'Unbalanced Load(ps)(leeward):
        sgbE42 = SPAN / 2 'Length of eave to ridge for drift height:
        sgbE43 = 0.43 * ((Math.Max(sgbE42, 20)) ^ 0.333) * ((sgbE17 + 10) ^ 0.25) - 1.5 'Drift Height(hdr):
        sgbE45 = sgbE43 * sgbE18 / Math.Sqrt(sgbE38) : Snow_GableE45 = sgbE45  'rectangular surcharge(leeward):
        sgbE46 = Math.Min(2.667 * sgbE43 * Math.Sqrt(sgbE38), 0.5 * SPAN) : Snow_GableE46 = sgbE46 'Length of rectangular surcharge :
        '-Roof load case 1 - Snow @ Wind 0? - Unbalanced
        sgbF51 = sgbE40 : Snow_GableF51 = sgbF51
        sgbF52 = sgbE45 + sgbE37 : Snow_GableF52 = sgbF52
        sgbF56 = sgbE37 : Snow_GableF56 = sgbF56
        sgbF57 = sgbE37 : Snow_GableF57 = sgbF57


        sgbG51 = sgbE12 * sgbF51
        sgbG52 = sgbE12 * sgbF52
        '-Roof load case 2 - Snow @ Wind 90? - Balanced
        sgbG56 = sgbE12 * sgbF56
        sgbG57 = sgbE12 * sgbF57
        '*Snow Vertical presgb on eave supports:

        '-Roof load case 1 - Snow @ Wind 0? - Unbalanced
        sgbE66 = 0.5 * (SPAN) * sgbE10 / Math.Cos(0.5235987756) 'Zones 2 & 3 surface area =
        sgbU67 = 0.25 * SPAN
        'Zone#2
        sgbF69 = sgbE66 * sgbF51 / 1000
        sgbF70 = sgbF69 * sgbU67
        'Zone#3
        sgbG69 = sgbE66 * sgbF52 / 1000
        sgbG70 = sgbG69 * (SPAN - sgbU67)

        sgbF72 = sgbF70 + sgbG70
        sgbK73 = sgbF69 * (SPAN - sgbU67) ' Not Found in Sheet
        sgbK72 = 0.5 * sgbF72 / SPAN * 1000 : Snow_GableK72 = sgbK72

        sgbG73 = sgbG69 * sgbU67
        sgbF73 = sgbF69 * (SPAN - sgbU67)
        sgbF75 = sgbF73 + sgbG73
        sgbK75 = 0.5 * sgbF75 / SPAN * 1000 : Snow_GableK75 = sgbK75

        '-Roof load case 2 - Snow @ Wind 90? - Balanced
        'Zone#2
        sgbF80 = sgbE66 * sgbF56 / 1000
        sgbF81 = sgbF80 * sgbU67
        'Zone#3
        sgbF82 = 0
        sgbG82 = 0
        sgbF85 = 0
        sgbG85 = 0
        sgbG80 = sgbE66 * sgbF57 / 1000
        sgbG81 = sgbG80 * (SPAN - sgbU67)
        sgbF83 = sgbF81 + sgbF82 + sgbG81 + sgbG82
        sgbK83 = 0.5 * sgbF83 / SPAN * 1000
        sgbF84 = sgbF80 * (SPAN - sgbU67)
        sgbG84 = sgbG80 * sgbU67
        sgbF86 = sgbF84 + sgbF85 + sgbG84 + sgbG85
        sgbK86 = 0.5 * sgbF86 / SPAN * 1000
        span1 = Math.Max(sgbK87, Math.Max(sgbK84, Math.Max(sgbK76, sgbK73))) 'NOT USED 'Max. Vl. presgb on rafter support =
        span0 = Math.Max(sgbK86, Math.Max(sgbK83, Math.Max(sgbK75, sgbK72)))
        Console.WriteLine("sgbE90 --> " + span0.ToString())
    End Sub


    Sub USFindGableRafter_30()
        'On Error GoTo 0

        'SetQz
        '* Design straining actions:
        '1) Dead loads:
        'Rafter self weight =    7.475   lb/ft'  (Auming 12.0x4.0C12 rafter)
        E12 = 7.475   'this can be read from price file using  = itmweight(FrameMtrl(RafterVal).Itm)
        E13 = Gable_Gambrel_Purlin_DesignE58 / Purling
        'Purlin Equivalent load =    1.56    psf
        '=Gable_Gambrel_Purlin_Design!E14/Wind_Gable!E12
        'ie fixed values of 6.047/3.88=1.56
        E14 = Gable_Gambrel_Purlin_DesignE60    'PBR Panels unit weight =    1.14    psf   = itmWeight(roofShtItm)
        E16 = E13 + E14   'Total dead load     psf
        E18 = E12 + E16 * (0.5 * NomBayWidth)    'Total rafter dead load =

        I18 = E18 * (SPAN ^ 2) / 55 'Mdead1 =
        I19 = -E18 * (SPAN ^ 2) / 28 'Mdead2 =
        I20 = E18 * (SPAN ^ 2) / 55 'Mdead3 =
        M18 = -E18 * (0.5 * SPAN) * 1.67 'Ndead =

        '2) Live loads:
        'Gable_Gambrel_Purlin_Design!E68=fixed value Distributed live pressure = 20.00   psf
        E24 = 20 * (0.5 * NomBayWidth)
        I24 = E24 * (SPAN ^ 2) / 55 'Mlive1 =
        I25 = -E24 * (SPAN ^ 2) / 28 'Mlive2 =
        I26 = E24 * (SPAN ^ 2) / 55 'Mlive3 =
        M24 = -E24 * (0.5 * SPAN) * 1.67 'Nlive =

        '3) Wind loads:
        '-Roof load case 1 - Wind 0? - Case A
        E34 = Wind_GableI50 * (0.5 * NomBayWidth)
        E35 = Wind_GableI51 * (0.5 * NomBayWidth)
        I33 = -((SPAN ^ 2) * (E34 + E35)) / (64 * (0.866) ^ 2)   'Mw0A2 =
        I32 = 0.5 * I33 + E34 * ((0.5 * SPAN / 0.866) ^ 2) / 8   'Mw0A1  uses I33 from above hence out of order
        I34 = 0.5 * I33 + E35 * ((0.5 * SPAN / 0.866) ^ 2) / 8   'Mw0A3 =
        Q32 = ((0.5 * Wind_GableK84 * SPAN) - (I33) - (0.5 * E34 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2886 * SPAN)
        Q34 = ((0.5 * Wind_GableK81 * SPAN) - (I33) - (0.5 * E35 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2866 * SPAN)
        M32 = -0.866 * Q32 - 0.5774 * Wind_GableK84    'Nw0A1 =
        M34 = -0.866 * Q34 - 0.5774 * Wind_GableK81    'Nw0A3 =
        M33 = Math.Min(M32, M34)   'Nw0A2 =
        '-Roof load case 2 - Wind 0? - Case B
        E42 = Wind_GableI55 * (0.5 * NomBayWidth)
        E43 = Wind_GableI56 * (0.5 * NomBayWidth)
        I41 = -((SPAN ^ 2) * (E42 + E43)) / (64 * (0.866) ^ 2)
        I42 = 0.5 * I41 + E43 * ((0.5 * SPAN / 0.866) ^ 2) / 8
        I40 = 0.5 * I41 + E42 * ((0.5 * SPAN / 0.866) ^ 2) / 8
        Q40 = ((0.5 * Wind_GableK95 * SPAN) - (I41) - (0.5 * E42 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2886 * SPAN)
        Q42 = ((0.5 * Wind_GableK92 * SPAN) - (I41) - (0.5 * E43 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2886 * SPAN)
        M40 = -0.866 * Q40 - 0.5774 * Wind_GableK95
        M42 = -0.866 * Q42 - 0.5774 * Wind_GableK92
        M41 = Math.Min(M40, M42)
        '-Roof load case 3 - Wind 90? - Case A
        E50 = Wind_GableI60 * (0.5 * NomBayWidth)
        E51 = Wind_GableI60 * (0.5 * NomBayWidth)
        I48 = E50 * (SPAN ^ 2) / 55 / Math.Cos(0.52359877)
        I49 = -E50 * (SPAN ^ 2) / 28 / Math.Cos(0.52359877)
        I50 = E50 * (SPAN ^ 2) / 55 / Math.Cos(0.52359877)
        M48 = -E50 * (0.5 * SPAN) * 1.25
        '-Roof load case 4 - Wind 90? - Case B
        E58 = Wind_GableI65 * (0.5 * NomBayWidth)
        E59 = Wind_GableI65 * (0.5 * NomBayWidth)
        I56 = E58 * (SPAN ^ 2) / 55 / Math.Cos(0.52359877)
        I57 = -E58 * (SPAN ^ 2) / 28 / Math.Cos(0.52359877) '    =-E58*(Wind_Gable!$E$9^2)/40
        I58 = E58 * (SPAN ^ 2) / 55 / Math.Cos(0.52359877)
        M56 = -E58 * (0.5 * SPAN) * 1.25
        '4) Snow loads:
        '-Roof load case 1 - Snow @ Wind 0? - Unbalanced
        E67 = Snow_GableF51 * (0.5 * NomBayWidth)
        E68 = (Snow_GableE41 + (Snow_GableE45) * (2 * Snow_GableE46 / SPAN)) * (0.5 * NomBayWidth)

        I66 = -((SPAN ^ 2) * (E67 + E68)) / 64
        I65 = 0.5 * I66 + E67 * ((0.5 * SPAN) ^ 2) / 8
        I67 = 0.5 * I66 + E68 * ((0.5 * SPAN) ^ 2) / 8

        Q65 = ((0.5 * Snow_GableK75 * SPAN) - (I66) - (0.5 * E67 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2887 * SPAN)
        Q67 = ((0.5 * Snow_GableK72 * SPAN) - (I66) - (0.5 * E68 * ((0.5 * SPAN) ^ 2 / 0.886))) / (0.2887 * SPAN)

        M65 = -0.866 * Q65 - 0.5774 * Snow_GableK75
        M67 = -0.866 * Q67 - 0.5774 * Snow_GableK72
        M66 = Math.Min(M65, M67)

        '-Roof load case 2 - Snow @ Wind 90? - Balanced
        E75 = Snow_GableF56 * (0.5 * NomBayWidth)
        E76 = Snow_GableF57 * (0.5 * NomBayWidth)

        I73 = E75 * (SPAN ^ 2) * 0.866 / 55
        I74 = -E75 * (SPAN ^ 2) * 0.866 / 28
        I75 = E75 * (SPAN ^ 2) * 0.866 / 55

        M73 = -E75 * (0.5 * SPAN) * 1.67


        '* Design  Combinations:

        E85 = 0.012 * (I18 + I24)
        F85 = 0.012 * (I18 + I65)
        G85 = 0.012 * (I18 + I32)
        H85 = 0.012 * (I18 + 0.75 * I32 + 0.75 * I24)
        I85 = 0.012 * (I18 + 0.75 * I32 + 0.75 * I65)

        E86 = (M18 + M24) * 0.001
        F86 = 0.001 * (M18 + M65)
        G86 = 0.001 * (M18 + M32)
        H86 = 0.001 * (M18 + 0.75 * M32 + 0.75 * M24)
        I86 = 0.001 * (M18 + 0.75 * M32 + 0.75 * M65)

        E87 = (I19 + I25) * 0.012
        F87 = 0.012 * (I19 + I66)
        G87 = 0.012 * (I19 + I33)
        H87 = 0.012 * (I19 + 0.75 * I33 + 0.75 * I25)
        I87 = 0.012 * (I19 + 0.75 * I33 + 0.75 * I57)

        E88 = (M18 + M24) * 0.001
        F88 = 0.001 * (M18 + M66)
        G88 = 0.001 * (M18 + M33)
        H88 = 0.001 * (M18 + 0.75 * M33 + 0.75 * M24)
        I88 = 0.001 * (M18 + 0.75 * M33 + 0.75 * M66)

        F89 = 0.012 * (I20 + I67)
        G89 = 0.012 * (I20 + I34)
        H89 = 0.012 * (I20 + 0.75 * I34 + 0.75 * I26)
        I89 = 0.012 * (I20 + 0.75 * I34 + 0.75 * I67)

        F90 = 0.001 * (M18 + M67)
        G90 = 0.001 * (M18 + M34)
        H90 = 0.001 * (M18 + 0.75 * M34 + 0.75 * M24)
        I90 = 0.001 * (M18 + 0.75 * M34 + 0.75 * M67)

        F91 = 0.012 * (I18 + I73)
        G91 = 0.012 * (I18 + I40)
        H91 = 0.012 * (I18 + 0.75 * I40 + 0.75 * I24)
        I91 = 0.012 * (I18 + 0.75 * I40 + 0.75 * I65)

        F92 = 0.001 * (M18 + M73)
        G92 = 0.001 * (M18 + M40)
        H92 = 0.001 * (M18 + 0.75 * M40 + 0.75 * M24)
        I92 = 0.001 * (M18 + 0.75 * M40 + 0.75 * M65)

        F93 = 0.012 * (I18 + I74)
        G93 = 0.012 * (I19 + I41)
        H93 = 0.012 * (I19 + 0.75 * I41 + 0.75 * I25)
        I93 = 0.012 * (I19 + 0.75 * I41 + 0.75 * I66)

        F94 = 0.001 * (M18 + M73)
        G94 = 0.001 * (M18 + M41)
        H94 = 0.001 * (M18 + 0.75 * M41 + 0.75 * M24)
        I94 = 0.001 * (M18 + 0.75 * M41 + 0.75 * M66)

        G95 = 0.012 * (I20 + I42)
        H95 = 0.012 * (I20 + 0.75 * I42 + 0.75 * I26)
        I95 = 0.012 * (I20 + 0.75 * I42 + 0.75 * I67)

        G96 = 0.001 * (M18 + M42)
        H96 = 0.001 * (M18 + 0.75 * M42 + 0.75 * M24)
        I96 = 0.001 * (M18 + 0.75 * M42 + 0.75 * M67)

        G97 = 0.012 * (I18 + I48)
        H97 = 0.012 * (I18 + 0.75 * I48 + 0.75 * I24)
        I97 = 0.012 * (I18 + 0.75 * I48 + 0.75 * I73)

        G98 = 0.001 * (M18 + M48)
        H98 = 0.001 * (M18 + 0.75 * M48 + 0.75 * M24)
        I98 = 0.001 * (M18 + 0.75 * M48 + 0.75 * M73)

        G99 = 0.012 * (I19 + I49)
        H99 = 0.012 * (I19 + 0.75 * I49 + 0.75 * I25)
        I99 = 0.012 * (I19 + 0.75 * I49 + 0.75 * I74)

        G100 = 0.001 * (M18 + M48)
        H100 = 0.001 * (M18 + 0.75 * M48 + 0.75 * M24)
        I100 = 0.001 * (M18 + 0.75 * M48 + 0.75 * M73)

        G101 = 0.012 * (I18 + I56) : H101 = 0.012 * (I18 + 0.75 * I56 + 0.75 * I24) : I101 = 0.012 * (I18 + 0.75 * I56 + 0.75 * I73)
        G102 = 0.001 * (M18 + M56) : H102 = 0.001 * (M18 + 0.75 * M56 + 0.75 * M24) : I102 = 0.001 * (M18 + 0.75 * M56 + 0.75 * M73)
        G103 = 0.012 * (I19 + I57) : H103 = 0.012 * (I19 + 0.75 * I57 + 0.75 * I25) : I103 = 0.012 * (I19 + 0.75 * I57 + 0.75 * I74)
        G104 = 0.001 * (M18 + M56) : H104 = 0.001 * (M18 + 0.75 * M56 + 0.75 * M24) : I104 = 0.001 * (M18 + 0.75 * M56 + 0.75 * M73)

        '* Stress Ratios:
        'C110 = "8.0x4.0C14"
        C110 = 9
        maxval = 0
        thisval = 0

        If E86 < 0 Then thisval = 0.1239 * Math.Abs(E86) + 0.0125 * Math.Abs(E85) : E110 = thisval : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.1239 * Math.Abs(F86) + 0.0125 * Math.Abs(F85) : F110 = thisval : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.1239 * Math.Abs(G86) + 0.0125 * Math.Abs(G85) : G110 = thisval : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.1239 * Math.Abs(H86) + 0.0125 * Math.Abs(H85) : H110 = thisval : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.1239 * Math.Abs(I86) + 0.0125 * Math.Abs(I85) : I110 = thisval : If thisval > maxval Then maxval = thisval

        If E86 < 0 Then thisval = 0.0536 * Math.Abs(E86) + 0.0145 * Math.Abs(E85) : E111 = thisval : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0536 * Math.Abs(F86) + 0.0145 * Math.Abs(F85) : F111 = thisval : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0536 * Math.Abs(G86) + 0.0145 * Math.Abs(G85) : G111 = thisval : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0536 * Math.Abs(H86) + 0.0145 * Math.Abs(H85) : H111 = thisval : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0536 * Math.Abs(I86) + 0.0145 * Math.Abs(I85) : I111 = thisval : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.1239 * Math.Abs(E88) + 0.0125 * Math.Abs(E87) : E112 = thisval : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.1239 * Math.Abs(F88) + 0.0125 * Math.Abs(F87) : F112 = thisval : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.1239 * Math.Abs(G88) + 0.0125 * Math.Abs(G87) : G112 = thisval : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.1239 * Math.Abs(H88) + 0.0125 * Math.Abs(H87) : H112 = thisval : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.1239 * Math.Abs(I88) + 0.0125 * Math.Abs(I87) : I112 = thisval : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0536 * Math.Abs(E88) + 0.0145 * Math.Abs(E87) : E113 = thisval : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0536 * Math.Abs(F88) + 0.0145 * Math.Abs(F87) : F113 = thisval : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0536 * Math.Abs(G88) + 0.0145 * Math.Abs(G87) : G113 = thisval : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0536 * Math.Abs(H88) + 0.0145 * Math.Abs(H87) : H113 = thisval : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0536 * Math.Abs(I88) + 0.0145 * Math.Abs(I87) : I113 = thisval : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.1239 * Math.Abs(F90) + 0.0125 * Math.Abs(F89) : F114 = thisval : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.1239 * Math.Abs(G90) + 0.0125 * Math.Abs(G89) : G114 = thisval : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.1239 * Math.Abs(H90) + 0.0125 * Math.Abs(H89) : H114 = thisval : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.1239 * Math.Abs(I90) + 0.0125 * Math.Abs(I89) : I114 = thisval : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0536 * Math.Abs(F90) + 0.0145 * Math.Abs(F89) : F115 = thisval : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0536 * Math.Abs(G90) + 0.0145 * Math.Abs(G89) : G115 = thisval : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0536 * Math.Abs(H90) + 0.0145 * Math.Abs(H89) : H115 = thisval : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0536 * Math.Abs(I90) + 0.0145 * Math.Abs(I89) : I115 = thisval : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.1239 * Math.Abs(F92) + 0.0125 * Math.Abs(F91) : F116 = thisval : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.1239 * Math.Abs(G92) + 0.0125 * Math.Abs(G91) : G116 = thisval : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.1239 * Math.Abs(H92) + 0.0125 * Math.Abs(H91) : H116 = thisval : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.1239 * Math.Abs(I92) + 0.0125 * Math.Abs(I91) : I116 = thisval : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0536 * Math.Abs(F92) + 0.0145 * Math.Abs(F91) : F117 = thisval : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0536 * Math.Abs(G92) + 0.0145 * Math.Abs(G91) : G117 = thisval : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0536 * Math.Abs(H92) + 0.0145 * Math.Abs(H91) : H117 = thisval : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0536 * Math.Abs(I92) + 0.0145 * Math.Abs(I91) : I117 = thisval : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.1239 * Math.Abs(F94) + 0.0125 * Math.Abs(F93) : F118 = thisval : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.1239 * Math.Abs(G94) + 0.0125 * Math.Abs(G93) : G118 = thisval : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.1239 * Math.Abs(H94) + 0.0125 * Math.Abs(H93) : H118 = thisval : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.1239 * Math.Abs(I94) + 0.0125 * Math.Abs(I93) : I118 = thisval : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0536 * Math.Abs(F94) + 0.0145 * Math.Abs(F93) : F119 = thisval : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0536 * Math.Abs(G94) + 0.0145 * Math.Abs(G93) : G119 = thisval : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0536 * Math.Abs(H94) + 0.0145 * Math.Abs(H93) : H119 = thisval : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0536 * Math.Abs(I94) + 0.0145 * Math.Abs(I93) : I119 = thisval : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.1239 * Math.Abs(G96) + 0.0125 * Math.Abs(G95) : G120 = thisval : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.1239 * Math.Abs(H96) + 0.0125 * Math.Abs(H95) : H120 = thisval : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.1239 * Math.Abs(I96) + 0.0125 * Math.Abs(I95) : I120 = thisval : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0536 * Math.Abs(G96) + 0.0145 * Math.Abs(G95) : G121 = thisval : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0536 * Math.Abs(H96) + 0.0145 * Math.Abs(H95) : H121 = thisval : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0536 * Math.Abs(I96) + 0.0145 * Math.Abs(I95) : I121 = thisval : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.1239 * Math.Abs(G98) + 0.0125 * Math.Abs(G97) : G122 = thisval : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.1239 * Math.Abs(H98) + 0.0125 * Math.Abs(H97) : H122 = thisval : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.1239 * Math.Abs(I98) + 0.0125 * Math.Abs(I97) : I122 = thisval : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0536 * Math.Abs(G98) + 0.0145 * Math.Abs(G97) : G123 = thisval : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0536 * Math.Abs(H98) + 0.0145 * Math.Abs(H97) : H123 = thisval : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0536 * Math.Abs(I98) + 0.0145 * Math.Abs(I97) : I123 = thisval : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.1239 * Math.Abs(G100) + 0.0125 * Math.Abs(G99) : G124 = thisval : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.1239 * Math.Abs(H100) + 0.0125 * Math.Abs(H99) : H124 = thisval : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.1239 * Math.Abs(I100) + 0.0125 * Math.Abs(I99) : I124 = thisval : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0536 * Math.Abs(G100) + 0.0145 * Math.Abs(G99) : G125 = thisval : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0536 * Math.Abs(H100) + 0.0145 * Math.Abs(H99) : H125 = thisval : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0536 * Math.Abs(I100) + 0.0145 * Math.Abs(I99) : I125 = thisval : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.1239 * Math.Abs(G102) + 0.0125 * Math.Abs(G101) : G126 = thisval : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.1239 * Math.Abs(H102) + 0.0125 * Math.Abs(H101) : H126 = thisval : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.1239 * Math.Abs(I102) + 0.0125 * Math.Abs(I101) : I126 = thisval : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0536 * Math.Abs(G102) + 0.0145 * Math.Abs(G101) : G127 = thisval : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0536 * Math.Abs(H102) + 0.0145 * Math.Abs(H101) : H127 = thisval : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0536 * Math.Abs(I102) + 0.0145 * Math.Abs(I101) : I127 = thisval : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.1239 * Math.Abs(G104) + 0.0125 * Math.Abs(G103) : G128 = thisval : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.1239 * Math.Abs(H104) + 0.0125 * Math.Abs(H103) : H128 = thisval : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.1239 * Math.Abs(I104) + 0.0125 * Math.Abs(I103) : I128 = thisval : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0536 * Math.Abs(G104) + 0.0145 * Math.Abs(G103) : G129 = thisval : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0536 * Math.Abs(H104) + 0.0145 * Math.Abs(H103) : H129 = thisval : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0536 * Math.Abs(I104) + 0.0145 * Math.Abs(I103) : I129 = thisval : If thisval > maxval Then maxval = thisval

        I130 = maxval
        '"Checked up to here"

        'C132 = "12.0x4.0C14"
        C132 = 14
        maxval = 0
        If E86 < 0 Then thisval = 0.1072 * Math.Abs(E86) + 0.0089 * Math.Abs(E85) : E132 = thisval : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.1072 * Math.Abs(F86) + 0.0089 * Math.Abs(F85) : F132 = thisval : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.1072 * Math.Abs(G86) + 0.0089 * Math.Abs(G85) : G132 = thisval : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.1072 * Math.Abs(H86) + 0.0089 * Math.Abs(H85) : H132 = thisval : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.1072 * Math.Abs(I86) + 0.0089 * Math.Abs(I85) : I132 = thisval : If thisval > maxval Then maxval = thisval

        If E86 < 0 Then thisval = 0.0572 * Math.Abs(E86) + 0.0105 * Math.Abs(E85) : E133 = thisval : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0572 * Math.Abs(F86) + 0.0105 * Math.Abs(F85) : F133 = thisval : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0572 * Math.Abs(G86) + 0.0105 * Math.Abs(G85) : G133 = thisval : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0572 * Math.Abs(H86) + 0.0105 * Math.Abs(H85) : H133 = thisval : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0572 * Math.Abs(I86) + 0.0105 * Math.Abs(I85) : I133 = thisval : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.1072 * Math.Abs(E88) + 0.0089 * Math.Abs(E87) : E134 = thisval : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.1072 * Math.Abs(F88) + 0.0089 * Math.Abs(F87) : F134 = thisval : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.1072 * Math.Abs(G88) + 0.0089 * Math.Abs(G87) : G134 = thisval : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.1072 * Math.Abs(H88) + 0.0089 * Math.Abs(H87) : H134 = thisval : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.1072 * Math.Abs(I88) + 0.0089 * Math.Abs(I87) : I134 = thisval : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0572 * Math.Abs(E88) + 0.0105 * Math.Abs(E87) : E135 = thisval : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0572 * Math.Abs(F88) + 0.0105 * Math.Abs(F87) : F135 = thisval : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0572 * Math.Abs(G88) + 0.0105 * Math.Abs(G87) : G135 = thisval : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0572 * Math.Abs(H88) + 0.0105 * Math.Abs(H87) : H135 = thisval : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0572 * Math.Abs(I88) + 0.0105 * Math.Abs(I87) : I135 = thisval : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.1072 * Math.Abs(F90) + 0.0089 * Math.Abs(F89) : F136 = thisval : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.1072 * Math.Abs(G90) + 0.0089 * Math.Abs(G89) : G136 = thisval : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.1072 * Math.Abs(H90) + 0.0089 * Math.Abs(H89) : H136 = thisval : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.1072 * Math.Abs(I90) + 0.0089 * Math.Abs(I89) : I136 = thisval : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0572 * Math.Abs(F90) + 0.0105 * Math.Abs(F89) : F137 = thisval : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0572 * Math.Abs(G90) + 0.0105 * Math.Abs(G89) : G137 = thisval : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0572 * Math.Abs(H90) + 0.0105 * Math.Abs(H89) : H137 = thisval : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0572 * Math.Abs(I90) + 0.0105 * Math.Abs(I89) : I137 = thisval : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.1072 * Math.Abs(F92) + 0.0089 * Math.Abs(F91) : F138 = thisval : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.1072 * Math.Abs(G92) + 0.0089 * Math.Abs(G91) : G138 = thisval : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.1072 * Math.Abs(H92) + 0.0089 * Math.Abs(H91) : H138 = thisval : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.1072 * Math.Abs(I92) + 0.0089 * Math.Abs(I91) : I138 = thisval : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0572 * Math.Abs(F92) + 0.0105 * Math.Abs(F91) : F139 = thisval : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0572 * Math.Abs(G92) + 0.0105 * Math.Abs(G91) : G139 = thisval : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0572 * Math.Abs(H92) + 0.0105 * Math.Abs(H91) : H139 = thisval : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0572 * Math.Abs(I92) + 0.0105 * Math.Abs(I91) : I139 = thisval : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.1072 * Math.Abs(F94) + 0.0089 * Math.Abs(F93) : F140 = thisval : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.1072 * Math.Abs(G94) + 0.0089 * Math.Abs(G93) : G140 = thisval : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.1072 * Math.Abs(H94) + 0.0089 * Math.Abs(H93) : H140 = thisval : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.1072 * Math.Abs(I94) + 0.0089 * Math.Abs(I93) : I140 = thisval : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0572 * Math.Abs(F94) + 0.0105 * Math.Abs(F93) : F141 = thisval : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0572 * Math.Abs(G94) + 0.0105 * Math.Abs(G93) : G141 = thisval : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0572 * Math.Abs(H94) + 0.0105 * Math.Abs(H93) : H141 = thisval : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0572 * Math.Abs(I94) + 0.0105 * Math.Abs(I93) : I141 = thisval : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.1072 * Math.Abs(G96) + 0.0089 * Math.Abs(G95) : G142 = thisval : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.1072 * Math.Abs(H96) + 0.0089 * Math.Abs(H95) : H142 = thisval : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.1072 * Math.Abs(I96) + 0.0089 * Math.Abs(I95) : I142 = thisval : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0572 * Math.Abs(G96) + 0.0105 * Math.Abs(G95) : G143 = thisval : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0572 * Math.Abs(H96) + 0.0105 * Math.Abs(H95) : H143 = thisval : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0572 * Math.Abs(I96) + 0.0105 * Math.Abs(I95) : I143 = thisval : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.1072 * Math.Abs(G98) + 0.0089 * Math.Abs(G97) : G144 = thisval : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.1072 * Math.Abs(H98) + 0.0089 * Math.Abs(H97) : H144 = thisval : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.1072 * Math.Abs(I98) + 0.0089 * Math.Abs(I97) : I144 = thisval : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0572 * Math.Abs(G98) + 0.0105 * Math.Abs(G97) : G145 = thisval : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0572 * Math.Abs(H98) + 0.0105 * Math.Abs(H97) : H145 = thisval : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0572 * Math.Abs(I98) + 0.0105 * Math.Abs(I97) : I145 = thisval : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.1072 * Math.Abs(G100) + 0.0089 * Math.Abs(G99) : G146 = thisval : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.1072 * Math.Abs(H100) + 0.0089 * Math.Abs(H99) : H146 = thisval : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.1072 * Math.Abs(I100) + 0.0089 * Math.Abs(I99) : I146 = thisval : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0572 * Math.Abs(G100) + 0.0105 * Math.Abs(G99) : G147 = thisval : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0572 * Math.Abs(H100) + 0.0105 * Math.Abs(H99) : H147 = thisval : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0572 * Math.Abs(I100) + 0.0105 * Math.Abs(I99) : I147 = thisval : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.1072 * Math.Abs(G102) + 0.0089 * Math.Abs(G101) : G148 = thisval : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.1072 * Math.Abs(H102) + 0.0089 * Math.Abs(H101) : H148 = thisval : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.1072 * Math.Abs(I102) + 0.0089 * Math.Abs(I101) : I148 = thisval : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0572 * Math.Abs(G102) + 0.0105 * Math.Abs(G101) : G149 = thisval : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0572 * Math.Abs(H102) + 0.0105 * Math.Abs(H101) : H149 = thisval : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0572 * Math.Abs(I102) + 0.0105 * Math.Abs(I101) : I149 = thisval : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.1072 * Math.Abs(G104) + 0.0089 * Math.Abs(G103) : G150 = thisval : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.1072 * Math.Abs(H104) + 0.0089 * Math.Abs(H103) : H150 = thisval : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.1072 * Math.Abs(I104) + 0.0089 * Math.Abs(I103) : I150 = thisval : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0572 * Math.Abs(G104) + 0.0105 * Math.Abs(G103) : G151 = thisval : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0572 * Math.Abs(H104) + 0.0105 * Math.Abs(H103) : H151 = thisval : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0572 * Math.Abs(I104) + 0.0105 * Math.Abs(I103) : I151 = thisval : If thisval > maxval Then maxval = thisval
        I152 = maxval

        'C154 = "8.0x4.0C12"
        C154 = 10
        maxval = 0
        If E86 < 0 Then thisval = 0.0597 * Math.Abs(E86) + 0.0071 * Math.Abs(E85) : E154 = thisval : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0597 * Math.Abs(F86) + 0.0071 * Math.Abs(F85) : F154 = thisval : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0597 * Math.Abs(G86) + 0.0071 * Math.Abs(G85) : G154 = thisval : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0597 * Math.Abs(H86) + 0.0071 * Math.Abs(H85) : H154 = thisval : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0597 * Math.Abs(I86) + 0.0071 * Math.Abs(I85) : I154 = thisval : If thisval > maxval Then maxval = thisval

        If E86 < 0 Then thisval = 0.0265 * Math.Abs(E86) + 0.0083 * Math.Abs(E85) : E155 = thisval : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0265 * Math.Abs(F86) + 0.0083 * Math.Abs(F85) : F155 = thisval : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0265 * Math.Abs(G86) + 0.0083 * Math.Abs(G85) : G155 = thisval : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0265 * Math.Abs(H86) + 0.0083 * Math.Abs(H85) : H155 = thisval : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0265 * Math.Abs(I86) + 0.0083 * Math.Abs(I85) : I155 = thisval : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0597 * Math.Abs(E88) + 0.0071 * Math.Abs(E87) : E156 = thisval : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0597 * Math.Abs(F88) + 0.0071 * Math.Abs(F87) : F156 = thisval : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0597 * Math.Abs(G88) + 0.0071 * Math.Abs(G87) : G156 = thisval : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0597 * Math.Abs(H88) + 0.0071 * Math.Abs(H87) : H156 = thisval : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0597 * Math.Abs(I88) + 0.0071 * Math.Abs(I87) : I156 = thisval : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0265 * Math.Abs(E88) + 0.0083 * Math.Abs(E87) : E157 = thisval : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0265 * Math.Abs(F88) + 0.0083 * Math.Abs(F87) : F157 = thisval : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0265 * Math.Abs(G88) + 0.0083 * Math.Abs(G87) : G157 = thisval : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0265 * Math.Abs(H88) + 0.0083 * Math.Abs(H87) : H157 = thisval : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0265 * Math.Abs(I88) + 0.0083 * Math.Abs(I87) : I157 = thisval : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0597 * Math.Abs(F90) + 0.0071 * Math.Abs(F89) : F158 = thisval : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0597 * Math.Abs(G90) + 0.0071 * Math.Abs(G89) : G158 = thisval : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0597 * Math.Abs(H90) + 0.0071 * Math.Abs(H89) : H158 = thisval : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0597 * Math.Abs(I90) + 0.0071 * Math.Abs(I89) : I158 = thisval : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0265 * Math.Abs(F90) + 0.0083 * Math.Abs(F89) : F159 = thisval : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0265 * Math.Abs(G90) + 0.0083 * Math.Abs(G89) : G159 = thisval : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0265 * Math.Abs(H90) + 0.0083 * Math.Abs(H89) : H159 = thisval : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0265 * Math.Abs(I90) + 0.0083 * Math.Abs(I89) : I159 = thisval : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0597 * Math.Abs(F92) + 0.0071 * Math.Abs(F91) : F160 = thisval : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0597 * Math.Abs(G92) + 0.0071 * Math.Abs(G91) : G160 = thisval : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0597 * Math.Abs(H92) + 0.0071 * Math.Abs(H91) : H160 = thisval : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0597 * Math.Abs(I92) + 0.0071 * Math.Abs(I91) : I160 = thisval : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0265 * Math.Abs(F92) + 0.0083 * Math.Abs(F91) : F161 = thisval : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0265 * Math.Abs(G92) + 0.0083 * Math.Abs(G91) : G161 = thisval : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0265 * Math.Abs(H92) + 0.0083 * Math.Abs(H91) : H161 = thisval : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0265 * Math.Abs(I92) + 0.0083 * Math.Abs(I91) : I161 = thisval : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0597 * Math.Abs(F94) + 0.0071 * Math.Abs(F93) : F162 = thisval : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0597 * Math.Abs(G94) + 0.0071 * Math.Abs(G93) : G162 = thisval : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0597 * Math.Abs(H94) + 0.0071 * Math.Abs(H93) : H162 = thisval : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0597 * Math.Abs(I94) + 0.0071 * Math.Abs(I93) : I162 = thisval : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0265 * Math.Abs(F94) + 0.0083 * Math.Abs(F93) : F163 = thisval : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0265 * Math.Abs(G94) + 0.0083 * Math.Abs(G93) : G163 = thisval : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0265 * Math.Abs(H94) + 0.0083 * Math.Abs(H93) : H163 = thisval : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0265 * Math.Abs(I94) + 0.0083 * Math.Abs(I93) : I163 = thisval : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0597 * Math.Abs(G96) + 0.0071 * Math.Abs(G95) : G164 = thisval : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0597 * Math.Abs(H96) + 0.0071 * Math.Abs(H95) : H164 = thisval : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0597 * Math.Abs(I96) + 0.0071 * Math.Abs(I95) : I164 = thisval : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0265 * Math.Abs(G96) + 0.0083 * Math.Abs(G95) : G165 = thisval : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0265 * Math.Abs(H96) + 0.0083 * Math.Abs(H95) : H165 = thisval : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0265 * Math.Abs(I96) + 0.0083 * Math.Abs(I95) : I165 = thisval : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0597 * Math.Abs(G98) + 0.0071 * Math.Abs(G97) : G166 = thisval : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0597 * Math.Abs(H98) + 0.0071 * Math.Abs(H97) : H166 = thisval : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0597 * Math.Abs(I98) + 0.0071 * Math.Abs(I97) : I166 = thisval : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0265 * Math.Abs(G98) + 0.0083 * Math.Abs(G97) : G167 = thisval : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0265 * Math.Abs(H98) + 0.0083 * Math.Abs(H97) : H167 = thisval : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0265 * Math.Abs(I98) + 0.0083 * Math.Abs(I97) : I167 = thisval : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0597 * Math.Abs(G100) + 0.0071 * Math.Abs(G99) : G168 = thisval : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0597 * Math.Abs(H100) + 0.0071 * Math.Abs(H99) : H168 = thisval : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0597 * Math.Abs(I100) + 0.0071 * Math.Abs(I99) : I168 = thisval : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0265 * Math.Abs(G100) + 0.0083 * Math.Abs(G99) : G169 = thisval : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0265 * Math.Abs(H100) + 0.0083 * Math.Abs(H99) : H169 = thisval : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0265 * Math.Abs(I100) + 0.0083 * Math.Abs(I99) : I169 = thisval : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0597 * Math.Abs(G102) + 0.0071 * Math.Abs(G101) : G170 = thisval : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0597 * Math.Abs(H102) + 0.0071 * Math.Abs(H101) : H170 = thisval : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0597 * Math.Abs(I102) + 0.0071 * Math.Abs(I101) : I170 = thisval : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0265 * Math.Abs(G102) + 0.0083 * Math.Abs(G101) : G171 = thisval : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0265 * Math.Abs(H102) + 0.0083 * Math.Abs(H101) : H171 = thisval : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0265 * Math.Abs(I102) + 0.0083 * Math.Abs(I101) : I171 = thisval : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0597 * Math.Abs(G104) + 0.0071 * Math.Abs(G103) : G172 = thisval : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0597 * Math.Abs(H104) + 0.0071 * Math.Abs(H103) : H172 = thisval : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0597 * Math.Abs(I104) + 0.0071 * Math.Abs(I103) : I172 = thisval : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0265 * Math.Abs(G104) + 0.0083 * Math.Abs(G103) : G173 = thisval : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0265 * Math.Abs(H104) + 0.0083 * Math.Abs(H103) : H173 = thisval : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0265 * Math.Abs(I104) + 0.0083 * Math.Abs(I103) : I173 = thisval : If thisval > maxval Then maxval = thisval
        I174 = maxval

        'C176 = "12.0x4.0C12"
        C176 = 15
        maxval = 0

        If E86 < 0 Then thisval = 0.0515 * Math.Abs(E86) + 0.0046 * Math.Abs(E85) : E176 = thisval : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0515 * Math.Abs(F86) + 0.0046 * Math.Abs(F85) : F176 = thisval : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0515 * Math.Abs(G86) + 0.0046 * Math.Abs(G85) : G176 = thisval : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0515 * Math.Abs(H86) + 0.0046 * Math.Abs(H85) : H176 = thisval : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0515 * Math.Abs(I86) + 0.0046 * Math.Abs(I85) : I176 = thisval : If thisval > maxval Then maxval = thisval

        If E86 < 0 Then thisval = 0.0281 * Math.Abs(E86) + 0.0054 * Math.Abs(E85) : E177 = thisval : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0281 * Math.Abs(F86) + 0.0054 * Math.Abs(F85) : F177 = thisval : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0281 * Math.Abs(G86) + 0.0054 * Math.Abs(G85) : G177 = thisval : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0281 * Math.Abs(H86) + 0.0054 * Math.Abs(H85) : H177 = thisval : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0281 * Math.Abs(I86) + 0.0054 * Math.Abs(I85) : I177 = thisval : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0515 * Math.Abs(E88) + 0.0046 * Math.Abs(E87) : E178 = thisval : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0515 * Math.Abs(F88) + 0.0046 * Math.Abs(F87) : F178 = thisval : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0515 * Math.Abs(G88) + 0.0046 * Math.Abs(G87) : G178 = thisval : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0515 * Math.Abs(H88) + 0.0046 * Math.Abs(H87) : H178 = thisval : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0515 * Math.Abs(I88) + 0.0046 * Math.Abs(I87) : I178 = thisval : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0281 * Math.Abs(E88) + 0.0054 * Math.Abs(E87) : E179 = thisval : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0281 * Math.Abs(F88) + 0.0054 * Math.Abs(F87) : F179 = thisval : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0281 * Math.Abs(G88) + 0.0054 * Math.Abs(G87) : G179 = thisval : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0281 * Math.Abs(H88) + 0.0054 * Math.Abs(H87) : H179 = thisval : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0281 * Math.Abs(I88) + 0.0054 * Math.Abs(I87) : I179 = thisval : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0515 * Math.Abs(F90) + 0.0046 * Math.Abs(F89) : F180 = thisval : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0515 * Math.Abs(G90) + 0.0046 * Math.Abs(G89) : G180 = thisval : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0515 * Math.Abs(H90) + 0.0046 * Math.Abs(H89) : H180 = thisval : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0515 * Math.Abs(I90) + 0.0046 * Math.Abs(I89) : I180 = thisval : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0281 * Math.Abs(F90) + 0.0054 * Math.Abs(F89) : F181 = thisval : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0281 * Math.Abs(G90) + 0.0054 * Math.Abs(G89) : G181 = thisval : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0281 * Math.Abs(H90) + 0.0054 * Math.Abs(H89) : H181 = thisval : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0281 * Math.Abs(I90) + 0.0054 * Math.Abs(I89) : I181 = thisval : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0515 * Math.Abs(F92) + 0.0046 * Math.Abs(F91) : F182 = thisval : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0515 * Math.Abs(G92) + 0.0046 * Math.Abs(G91) : G182 = thisval : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0515 * Math.Abs(H92) + 0.0046 * Math.Abs(H91) : H182 = thisval : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0515 * Math.Abs(I92) + 0.0046 * Math.Abs(I91) : I182 = thisval : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0281 * Math.Abs(F92) + 0.0054 * Math.Abs(F91) : F183 = thisval : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0281 * Math.Abs(G92) + 0.0054 * Math.Abs(G91) : G183 = thisval : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0281 * Math.Abs(H92) + 0.0054 * Math.Abs(H91) : H183 = thisval : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0281 * Math.Abs(I92) + 0.0054 * Math.Abs(I91) : I183 = thisval : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0515 * Math.Abs(F94) + 0.0046 * Math.Abs(F93) : F184 = thisval : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0515 * Math.Abs(G94) + 0.0046 * Math.Abs(G93) : G184 = thisval : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0515 * Math.Abs(H94) + 0.0046 * Math.Abs(H93) : H184 = thisval : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0515 * Math.Abs(I94) + 0.0046 * Math.Abs(I93) : I184 = thisval : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0281 * Math.Abs(F94) + 0.0054 * Math.Abs(F93) : F185 = thisval : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0281 * Math.Abs(G94) + 0.0054 * Math.Abs(G93) : G185 = thisval : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0281 * Math.Abs(H94) + 0.0054 * Math.Abs(H93) : H185 = thisval : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0281 * Math.Abs(I94) + 0.0054 * Math.Abs(I93) : I185 = thisval : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0515 * Math.Abs(G96) + 0.0046 * Math.Abs(G95) : G186 = thisval : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0515 * Math.Abs(H96) + 0.0046 * Math.Abs(H95) : H186 = thisval : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0515 * Math.Abs(I96) + 0.0046 * Math.Abs(I95) : I186 = thisval : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0281 * Math.Abs(G96) + 0.0054 * Math.Abs(G95) : G187 = thisval : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0281 * Math.Abs(H96) + 0.0054 * Math.Abs(H95) : H187 = thisval : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0281 * Math.Abs(I96) + 0.0054 * Math.Abs(I95) : I187 = thisval : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0515 * Math.Abs(G98) + 0.0046 * Math.Abs(G97) : G188 = thisval : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0515 * Math.Abs(H98) + 0.0046 * Math.Abs(H97) : H188 = thisval : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0515 * Math.Abs(I98) + 0.0046 * Math.Abs(I97) : I188 = thisval : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0281 * Math.Abs(G98) + 0.0054 * Math.Abs(G97) : G189 = thisval : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0281 * Math.Abs(H98) + 0.0054 * Math.Abs(H97) : H189 = thisval : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0281 * Math.Abs(I98) + 0.0054 * Math.Abs(I97) : I189 = thisval : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0515 * Math.Abs(G100) + 0.0046 * Math.Abs(G99) : G190 = thisval : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0515 * Math.Abs(H100) + 0.0046 * Math.Abs(H99) : H190 = thisval : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0515 * Math.Abs(I100) + 0.0046 * Math.Abs(I99) : I190 = thisval : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0281 * Math.Abs(G100) + 0.0054 * Math.Abs(G99) : G191 = thisval : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0281 * Math.Abs(H100) + 0.0054 * Math.Abs(H99) : H191 = thisval : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0281 * Math.Abs(I100) + 0.0054 * Math.Abs(I99) : I191 = thisval : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0515 * Math.Abs(G102) + 0.0046 * Math.Abs(G101) : G192 = thisval : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0515 * Math.Abs(H102) + 0.0046 * Math.Abs(H101) : H192 = thisval : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0515 * Math.Abs(I102) + 0.0046 * Math.Abs(I101) : I192 = thisval : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0281 * Math.Abs(G102) + 0.0054 * Math.Abs(G101) : G193 = thisval : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0281 * Math.Abs(H102) + 0.0054 * Math.Abs(H101) : H193 = thisval : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0281 * Math.Abs(I102) + 0.0054 * Math.Abs(I101) : I193 = thisval : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0515 * Math.Abs(G104) + 0.0046 * Math.Abs(G103) : G194 = thisval : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0515 * Math.Abs(H104) + 0.0046 * Math.Abs(H103) : H194 = thisval : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0515 * Math.Abs(I104) + 0.0046 * Math.Abs(I103) : I194 = thisval : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0281 * Math.Abs(G104) + 0.0054 * Math.Abs(G103) : G195 = thisval : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0281 * Math.Abs(H104) + 0.0054 * Math.Abs(H103) : H195 = thisval : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0281 * Math.Abs(I104) + 0.0054 * Math.Abs(I103) : I195 = thisval : If thisval > maxval Then maxval = thisval
        I196 = maxval

        'If OutputCalcs Then Print #1, "xxxxX"
        '* Preliminary Rafter Profile selection :
        '=IF(I130<=1 then E220=C110,IF(I152<=1,C132,IF(I174<=1,C154,IF(I196<=1.02,C176,"Check braced apex rafter solution"))))
        If I130 <= 1 Then
            E202 = C110.ToString()
        ElseIf I152 <= 1 Then
            E202 = C132.ToString()
        ElseIf I174 <= 1 Then
            E202 = C154.ToString()
        ElseIf I196 <= 1.02 Then
            E202 = C176.ToString()
        Else
            E202 = "Check braced apex rafter solution"
        End If
        'see E202

        '* Check of 12.0x4.0C12 apex brace  solution (Max. brace length = 13 ft):
        '1) check apex brace saftey:
        mv = 100000
        tv = 0
        tv = 0.8 * E86 : E209 = tv : If tv < mv Then mv = tv
        tv = 0.8 * F86 : F209 = tv : If tv < mv Then mv = tv
        tv = 0.8 * G86 : G209 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H86 : H209 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I86 : I209 = tv : If tv < mv Then mv = tv

        tv = 0.8 * E88 : E210 = tv : If tv < mv Then mv = tv
        tv = 0.8 * F88 : F210 = tv : If tv < mv Then mv = tv
        tv = 0.8 * G88 : G210 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H88 : H210 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I88 : I210 = tv : If tv < mv Then mv = tv

        tv = 0.8 * F90 : F211 = tv : If tv < mv Then mv = tv
        tv = 0.8 * G90 : G211 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H90 : H211 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I90 : I211 = tv : If tv < mv Then mv = tv

        tv = 0.8 * F92 : F212 = tv : If tv < mv Then mv = tv
        tv = 0.8 * G92 : G212 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H92 : H212 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I92 : I212 = tv : If tv < mv Then mv = tv

        tv = 0.8 * F94 : F213 = tv : If tv < mv Then mv = tv
        tv = 0.8 * G94 : G213 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H94 : H213 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I94 : I213 = tv : If tv < mv Then mv = tv

        tv = 0.8 * G96 : G214 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H96 : H214 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I96 : I214 = tv : If tv < mv Then mv = tv

        tv = 0.8 * G98 : G215 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H98 : H215 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I98 : I215 = tv : If tv < mv Then mv = tv

        tv = 0.8 * G100 : G216 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H100 : H216 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I100 : I216 = tv : If tv < mv Then mv = tv

        tv = 0.8 * G102 : G217 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H102 : H217 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I102 : I217 = tv : If tv < mv Then mv = tv

        tv = 0.8 * G104 : G218 = tv : If tv < mv Then mv = tv
        tv = 0.8 * H104 : H218 = tv : If tv < mv Then mv = tv
        tv = 0.8 * I104 : I218 = tv : If tv < mv Then mv = tv
        F220 = Math.Abs(mv)
        If F220 <= 16.86 Then F222 = "Apex brace is safe" Else F222 = "Apex brace not available"
        'see F220

        '2) check sagging moment sections:
        E229 = 0.6 * E85 : G229 = 0.012 * (0.6 * I18 + 0.75 * I32) : H229 = 0.012 * (0.6 * I18 + 0.75 * 0.75 * I32 + 0.6 * 0.75 * I24)

        E230 = 1.1 * E86 : G230 = 0.001 * (1.1 * M18 + 1.1 * M32) : H230 = 0.001 * (1.1 * M18 + 1.1 * 0.75 * M32 + 1.1 * 0.75 * M24)

        F231 = 0.012 * (0.6 * I20 + 0.75 * I67) : G231 = 0.012 * (0.6 * I20 + 6.87 * I34) : H231 = 0.012 * (0.6 * I20 + 6.87 * 0.75 * I34 + 0.6 * 0.75 * I26) : I231 = 0.012 * (0.6 * I20 + 0.75 * 6.87 * I34 + 0.73 * 0.75 * I67)

        F232 = 0.001 * (1.1 * M18 + 0.37 * M67) : G232 = 0.001 * (1.1 * M18 + 1.1 * M34) : H232 = 0.001 * (1.1 * M18 + 1.1 * 0.75 * M34 + 1.1 * 0.75 * M24) : I232 = 0.001 * (1.1 * M18 + 1.1 * 0.75 * M34 + 0.37 * 0.75 * M67)

        F233 = 0.6 * F91 : G233 = 0.012 * (0.6 * I18 + 2.5 * I40) : H233 = 0.012 * (0.6 * I18 + 2.5 * 0.75 * I40 + 0.6 * 0.75 * I24)

        F234 = 1.1 * F92 : G234 = 0.001 * (1.1 * M18 + 1.1 * M40) : H234 = 0.001 * (1.1 * M18 + 1.1 * 0.75 * M40 + 1.1 * 0.75 * M24)

        G235 = 0.012 * (0.6 * I20 + 0.8 * I42) : H235 = 0.012 * (0.6 * I20 + 0.8 * 0.75 * I42 + 0.6 * 0.75 * I26) : I235 = 0.012 * (0.6 * I20 + 0.8 * 0.75 * I42 + 0.73 * 0.75 * I67)

        G236 = 0.001 * (1.1 * M18 + 1.1 * M42) : H236 = 0.001 * (1.1 * M18 + 1.1 * 0.75 * M42 + 1.1 * 0.75 * M24) : I236 = 0.001 * (1.1 * M18 + 1.1 * 0.75 * M42 + 0.37 * 0.75 * M67)

        G237 = 0.6 * G97 : H237 = 0.6 * H97 : I237 = 0.6 * I97

        G238 = 1.1 * G98 : H238 = 1.1 * H98 : I238 = 1.1 * I98

        G239 = 0.6 * G101 : H239 = 0.6 * H101 : I239 = 0.6 * I101

        G240 = 1.1 * G102 : H240 = 1.1 * H102 : I240 = 1.1 * I102


        '12.0x4.0C12 S.R.
        mv = 0
        If E230 < 0 Then tv = 0.038 * Math.Abs(E230) + 0.0046 * Math.Abs(E229) : E244 = tv : If tv > mv Then mv = tv
        If G230 < 0 Then tv = 0.038 * Math.Abs(G230) + 0.0046 * Math.Abs(G229) : G244 = tv : If tv > mv Then mv = tv
        If H230 < 0 Then tv = 0.038 * Math.Abs(H230) + 0.0046 * Math.Abs(H229) : H244 = tv : If tv > mv Then mv = tv

        If E230 < 0 Then tv = 0.0281 * Math.Abs(E230) + 0.0054 * Math.Abs(E229) : E245 = tv : If tv > mv Then mv = tv
        If G230 < 0 Then tv = 0.0281 * Math.Abs(G230) + 0.0054 * Math.Abs(G229) : G245 = tv : If tv > mv Then mv = tv
        If H230 < 0 Then tv = 0.0281 * Math.Abs(H230) + 0.0054 * Math.Abs(H229) : H245 = tv : If tv > mv Then mv = tv

        If F232 < 0 Then tv = 0.038 * Math.Abs(F232) + 0.0046 * Math.Abs(F231) : F246 = tv : If tv > mv Then mv = tv
        If G232 < 0 Then tv = 0.038 * Math.Abs(G232) + 0.0046 * Math.Abs(G231) : G246 = tv : If tv > mv Then mv = tv
        If H232 < 0 Then tv = 0.038 * Math.Abs(H232) + 0.0046 * Math.Abs(H231) : H246 = tv : If tv > mv Then mv = tv
        If I232 < 0 Then tv = 0.038 * Math.Abs(I232) + 0.0046 * Math.Abs(I231) : I246 = tv : If tv > mv Then mv = tv

        If F232 < 0 Then tv = 0.0281 * Math.Abs(F232) + 0.0054 * Math.Abs(F231) : F247 = tv : If tv > mv Then mv = tv
        If G232 < 0 Then tv = 0.0281 * Math.Abs(G232) + 0.0054 * Math.Abs(G231) : G247 = tv : If tv > mv Then mv = tv
        If H232 < 0 Then tv = 0.0281 * Math.Abs(H232) + 0.0054 * Math.Abs(H231) : H247 = tv : If tv > mv Then mv = tv
        If I232 < 0 Then tv = 0.0281 * Math.Abs(I232) + 0.0054 * Math.Abs(I231) : I247 = tv : If tv > mv Then mv = tv

        If F234 < 0 Then tv = 0.038 * Math.Abs(F234) + 0.0046 * Math.Abs(F233) : F248 = tv : If tv > mv Then mv = tv
        If G234 < 0 Then tv = 0.038 * Math.Abs(G234) + 0.0046 * Math.Abs(G233) : G248 = tv : If tv > mv Then mv = tv
        If H234 < 0 Then tv = 0.038 * Math.Abs(H234) + 0.0046 * Math.Abs(H233) : H248 = tv : If tv > mv Then mv = tv

        If F234 < 0 Then tv = 0.0281 * Math.Abs(F234) + 0.0054 * Math.Abs(F233) : F249 = tv : If tv > mv Then mv = tv
        If G234 < 0 Then tv = 0.0281 * Math.Abs(G234) + 0.0054 * Math.Abs(G233) : G249 = tv : If tv > mv Then mv = tv
        If H234 < 0 Then tv = 0.0281 * Math.Abs(H234) + 0.0054 * Math.Abs(H233) : H249 = tv : If tv > mv Then mv = tv

        If G236 < 0 Then tv = 0.038 * Math.Abs(G236) + 0.0046 * Math.Abs(G235) : G250 = tv : If tv > mv Then mv = tv
        If H236 < 0 Then tv = 0.038 * Math.Abs(H236) + 0.0046 * Math.Abs(H235) : H250 = tv : If tv > mv Then mv = tv
        If I236 < 0 Then tv = 0.038 * Math.Abs(I236) + 0.0046 * Math.Abs(I235) : I250 = tv : If tv > mv Then mv = tv

        If G236 < 0 Then tv = 0.0281 * Math.Abs(G236) + 0.0054 * Math.Abs(G235) : G251 = tv : If tv > mv Then mv = tv
        If H236 < 0 Then tv = 0.0281 * Math.Abs(H236) + 0.0054 * Math.Abs(H235) : H251 = tv : If tv > mv Then mv = tv
        If I236 < 0 Then tv = 0.0281 * Math.Abs(I236) + 0.0054 * Math.Abs(I235) : I251 = tv : If tv > mv Then mv = tv

        If G238 < 0 Then tv = 0.038 * Math.Abs(G238) + 0.0046 * Math.Abs(G237) : G252 = tv : If tv > mv Then mv = tv
        If H238 < 0 Then tv = 0.038 * Math.Abs(H238) + 0.0046 * Math.Abs(H237) : H252 = tv : If tv > mv Then mv = tv
        If I238 < 0 Then tv = 0.038 * Math.Abs(I238) + 0.0046 * Math.Abs(I237) : I252 = tv : If tv > mv Then mv = tv

        If G238 < 0 Then tv = 0.0281 * Math.Abs(G238) + 0.0054 * Math.Abs(G237) : G253 = tv : If tv > mv Then mv = tv
        If H238 < 0 Then tv = 0.0281 * Math.Abs(H238) + 0.0054 * Math.Abs(H237) : H253 = tv : If tv > mv Then mv = tv
        If I238 < 0 Then tv = 0.0281 * Math.Abs(I238) + 0.0054 * Math.Abs(I237) : I253 = tv : If tv > mv Then mv = tv

        If G240 < 0 Then tv = 0.038 * Math.Abs(G240) + 0.0046 * Math.Abs(G239) : G254 = tv : If tv > mv Then mv = tv
        If H240 < 0 Then tv = 0.038 * Math.Abs(H240) + 0.0046 * Math.Abs(H239) : H254 = tv : If tv > mv Then mv = tv
        If I240 < 0 Then tv = 0.038 * Math.Abs(I240) + 0.0046 * Math.Abs(I239) : I254 = tv : If tv > mv Then mv = tv

        If G240 < 0 Then tv = 0.0281 * Math.Abs(G240) + 0.0054 * Math.Abs(G239) : G255 = tv : If tv > mv Then mv = tv
        If H240 < 0 Then tv = 0.0281 * Math.Abs(H240) + 0.0054 * Math.Abs(H239) : H255 = tv : If tv > mv Then mv = tv
        If I240 < 0 Then tv = 0.0281 * Math.Abs(I240) + 0.0054 * Math.Abs(I239) : I255 = tv : If tv > mv Then mv = tv
        I256 = mv
        If I256 <= 1.04 Then f258 = "Rafter Safe" Else f258 = "Rafter Unsafe"
        'RAFTER Profile
        E262 = "No rafter found"
        ApexBrace = False
        'ValidS = False"INVALID"
        ProfileFound = False : CertifiedProfile = False
        If E202 = "Check braced apex rafter solution" Then
            If F222 = "Apex brace is safe" And f258 = "Rafter Safe" Then
                E262 = "OK"   '"12.0x4.0C12 with apex braces of length 1/3 span"
                ApexBrace = True
                ApexBraceInVal = 6 : ApexBracesInt = sideBays * 2
                BApexBrace = 6
                RafterInVal = 6
                'ValidS = True
                ProfileFound = True
                CertifiedProfile = True
            Else
                E262 = "Failed"    '"Kit unavailable. Rafter failed"
                RafterInVal = 6
                ApexBraceInVal = 6 : ApexBracesInt = sideBays * 2
                BApexBrace = 6 : ApexBrace = True
                'ValidS = False
            End If
        Else
            E262 = "OK"
            RafterInVal = E202
            'ValidS = True
            ApexBrace = False
            ApexBraceInVal = 0 : ApexBracesInt = 0
            BApexBrace = 15
            ProfileFound = True
            CertifiedProfile = True
        End If

        ' "Rafter is " & E262 & nl & " ie " & FrameMtrl(RafterInVal).Mtrl & nl & "ApexBrace=" & ApexBrace
        Console.WriteLine("I130 : " + I130.ToString())
        Console.WriteLine("I152 : " + I152.ToString())
        Console.WriteLine("I174 : " + I174.ToString())
        Console.WriteLine("I196 : " + I196.ToString())
        Console.WriteLine("E202 : " + E202.ToString())
        Console.WriteLine("F220 : " + F220.ToString())
        Console.WriteLine("I256 : " + I256.ToString())
        Console.WriteLine("E262 : " + E262.ToString())
        'Close #1


    End Sub





End Module
