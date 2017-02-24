Module Module1
    Sub Main()

        Console.WriteLine("test")


    End Sub
    Sub USFindGableRafter_30()
        On Error GoTo 0

        'SetQz
        Console.WriteLine("Enter SPAN:")
        Dim SPAN = Console.ReadLine()
        Console.WriteLine("EnterNomBayWidth:")
        Dim NomBayWidth = Console.ReadLine()
        '* Design straining actions:
        '1) Dead loads:
        'Rafter self weight =    7.475   lb/ft'  (Auming 12.0x4.0C12 rafter)
        Dim E12 = 7.475   'this can be read from price file using  = itmweight(FrameMtrl(RafterVal).Itm)
        Dim E13 = 1.56    'Purlin Equivalent load =    1.56    psf
        '=Gable_Gambrel_Purlin_Design!E14/Wind_Gable!E12
        'ie fixed values of 6.047/3.88=1.56
        Dim E14 = 1.14    'PBR Panels unit weight =    1.14    psf   = itmWeight(roofShtItm)
        Dim E16 = E13 + E14   'Total dead load     psf
        Dim E18 = E12 + E16 * (0.5 * NomBayWidth)    'Total rafter dead load =

        Dim I18 = E18 * (SPAN ^ 2) / 55 'Mdead1 =
        Dim I19 = -E18 * (SPAN ^ 2) / 28 'Mdead2 =
        Dim I20 = E18 * (SPAN ^ 2) / 55 'Mdead3 =
        Dim M18 = -E18 * (0.5 * SPAN) * 1.67 'Ndead =

        '2) Live loads:
        'Gable_Gambrel_Purlin_Design!E68=fixed value Distributed live pressure = 20.00   psf
        Dim E24 = 20 * (0.5 * NomBayWidth)
        Dim I24 = E24 * (SPAN ^ 2) / 55 'Mlive1 =
        Dim I25 = -E24 * (SPAN ^ 2) / 28 'Mlive2 =
        Dim I26 = E24 * (SPAN ^ 2) / 55 'Mlive3 =
        Dim M24 = -E24 * (0.5 * SPAN) * 1.67 'Nlive =

        '3) Wind loads:
        '-Roof load case 1 - Wind 0? - Case A
        Dim E34 = Wind_GableI50 * (0.5 * NomBayWidth)
        Dim E35 = Wind_GableI51 * (0.5 * NomBayWidth)
        Dim I33 = -((SPAN ^ 2) * (E34 + E35)) / (64 * (0.866) ^ 2)   'Mw0A2 =
        Dim I32 = 0.5 * I33 + E34 * ((0.5 * SPAN / 0.866) ^ 2) / 8   'Mw0A1  uses I33 from above hence out of order
        Dim I34 = 0.5 * I33 + E35 * ((0.5 * SPAN / 0.866) ^ 2) / 8   'Mw0A3 =
        Dim Q32 = ((0.5 * Wind_GableK84 * SPAN) - (I33) - (0.5 * E34 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2886 * SPAN)
        Dim Q34 = ((0.5 * Wind_GableK81 * SPAN) - (I33) - (0.5 * E35 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2886 * SPAN)
        Dim M32 = -0.866 * Q32 - 0.5774 * Wind_GableK84    'Nw0A1 =
        Dim M34 = -0.866 * Q34 - 0.5774 * Wind_GableK81    'Nw0A3 =
        Dim M33 = Min(M32, M34)   'Nw0A2 =
        '-Roof load case 2 - Wind 0? - Case B
        Dim E42 = Wind_GableI55 * (0.5 * NomBayWidth)
        Dim E43 = Wind_GableI56 * (0.5 * NomBayWidth)
        Dim I41 = -((SPAN ^ 2) * (E42 + E43)) / (64 * (0.866) ^ 2)
        Dim I42 = 0.5 * I41 + E43 * ((0.5 * SPAN / 0.866) ^ 2) / 8
        Dim I40 = 0.5 * I41 + E42 * ((0.5 * SPAN / 0.866) ^ 2) / 8
        Dim Q40 = ((0.5 * Wind_GableK95 * SPAN) - (I41) - (0.5 * E42 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2886 * SPAN)
        Dim Q42 = ((0.5 * Wind_GableK92 * SPAN) - (I41) - (0.5 * E43 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2886 * SPAN)
        Dim M40 = -0.866 * Q40 - 0.5774 * Wind_GableK95
        Dim M42 = -0.866 * Q42 - 0.5774 * Wind_GableK92
        Dim M41 = Math.Min(M40, M42)
        '-Roof load case 3 - Wind 90? - Case A
        Dim E50 = Wind_GableI60 * (0.5 * NomBayWidth)
        Dim E51 = Wind_GableI60 * (0.5 * NomBayWidth)
        Dim I48 = E50 * (SPAN ^ 2) / 55 / COS(0.52359877)
        Dim I49 = -E50 * (SPAN ^ 2) / 28 / COS(0.52359877)
        Dim I50 = E50 * (SPAN ^ 2) / 55 / COS(0.52359877)
        Dim M48 = -E50 * (0.5 * SPAN) * 1.25
        '-Roof load case 4 - Wind 90? - Case B
        Dim E58 = Wind_GableI65 * (0.5 * NomBayWidth)
        Dim E59 = Wind_GableI65 * (0.5 * NomBayWidth)
        Dim I56 = E58 * (SPAN ^ 2) / 55 / Math.Cos(0.52359877)
        Dim I57 = -E58 * (SPAN ^ 2) / 28 / Math.Cos(0.52359877) '    =-E58*(Wind_Gable!$E$9^2)/40
        Dim I58 = E58 * (SPAN ^ 2) / 55 / Math.Cos(0.52359877)
        Dim M56 = -E58 * (0.5 * SPAN) * 1.25
        '4) Snow loads:
        '-Roof load case 1 - Snow @ Wind 0? - Unbalanced
        Dim E67 = Snow_GableF51 * (0.5 * NomBayWidth)
        Dim E68 = (Snow_GableE41 + (Snow_GableE45) * (2 * Snow_GableE46 / NomBayWidth)) * (0.5 * SPAN)

        Dim I66 = -((SPAN ^ 2) * (E67 + E68)) / 64
        Dim I65 = 0.5 * I66 + E67 * ((0.5 * SPAN) ^ 2) / 8
        Dim I67 = 0.5 * I66 + E68 * ((0.5 * SPAN) ^ 2) / 8

        Dim Q65 = ((0.5 * Snow_GableK75 * SPAN) - (I66) - (0.5 * E67 * ((0.5 * SPAN / 0.866) ^ 2))) / (0.2887 * SPAN)
        Dim Q67 = ((0.5 * Snow_GableK72 * SPAN) - (I66) - (0.5 * E68 * ((0.5 * SPAN) ^ 2 / 0.866))) / (0.2887 * SPAN)

        Dim M65 = -0.866 * Q65 - 0.5774 * Snow_GableK75
        Dim M67 = -0.866 * Q67 - 0.5774 * Snow_GableK72
        Dim M66 = Math.Min(M65, M67)

        '-Roof load case 2 - Snow @ Wind 90? - Balanced
        Dim E75 = Snow_GableF56 * (0.5 * NomBayWidth)
        Dim E76 = Snow_GableF57 * (0.5 * NomBayWidth)

        Dim I73 = E75 * (SPAN ^ 2) * 0.866 / 55
        Dim I74 = -E75 * (SPAN ^ 2) * 0.866 / 28
        Dim I75 = E75 * (SPAN ^ 2) * 0.866 / 55

        Dim M73 = -E75 * (0.5 * SPAN) * 1.67


        '* Design  Combinations:

        Dim E85 = 0.012 * (I18 + I24)
        Dim F85 = 0.012 * (I18 + I65)
        Dim G85 = 0.012 * (I18 + I32)
        Dim H85 = 0.012 * (I18 + 0.75 * I32 + 0.75 * I24)
        Dim I85 = 0.012 * (I18 + 0.75 * I32 + 0.75 * I65)

        Dim E86 = (M18 + M24) * 0.001
        Dim F86 = 0.001 * (M18 + M65)
        Dim G86 = 0.001 * (M18 + M32)
        Dim H86 = 0.001 * (M18 + 0.75 * M32 + 0.75 * M24)
        Dim I86 = 0.001 * (M18 + 0.75 * M32 + 0.75 * M65)

        Dim E87 = (I19 + I25) * 0.012
        Dim F87 = 0.012 * (I19 + I66)
        Dim G87 = 0.012 * (I19 + I33)
        Dim H87 = 0.012 * (I19 + 0.75 * I33 + 0.75 * I25)
        Dim I87 = 0.012 * (I19 + 0.75 * I33 + 0.75 * I57)

        Dim E88 = (M18 + M24) * 0.001
        Dim F88 = 0.001 * (M18 + M66)
        Dim G88 = 0.001 * (M18 + M33)
        Dim H88 = 0.001 * (M18 + 0.75 * M33 + 0.75 * M24)
        Dim I88 = 0.001 * (M18 + 0.75 * M33 + 0.75 * M66)

        Dim F89 = 0.012 * (I20 + I67)
        Dim G89 = 0.012 * (I20 + I34)
        Dim H89 = 0.012 * (I20 + 0.75 * I34 + 0.75 * I26)
        Dim I89 = 0.012 * (I20 + 0.75 * I34 + 0.75 * I67)

        Dim F90 = 0.001 * (M18 + M67)
        Dim G90 = 0.001 * (M18 + M34)
        Dim H90 = 0.001 * (M18 + 0.75 * M34 + 0.75 * M24)
        Dim I90 = 0.001 * (M18 + 0.75 * M34 + 0.75 * M67)

        Dim F91 = 0.012 * (I18 + I73)
        Dim G91 = 0.012 * (I18 + I40)
        Dim H91 = 0.012 * (I18 + 0.75 * I40 + 0.75 * I24)
        Dim I91 = 0.012 * (I18 + 0.75 * I40 + 0.75 * I65)

        Dim F92 = 0.001 * (M18 + M73)
        Dim G92 = 0.001 * (M18 + M40)
        Dim H92 = 0.001 * (M18 + 0.75 * M40 + 0.75 * M24)
        Dim I92 = 0.001 * (M18 + 0.75 * M40 + 0.75 * M65)

        Dim F93 = 0.012 * (I18 + I74)
        Dim G93 = 0.012 * (I19 + I41)
        Dim H93 = 0.012 * (I19 + 0.75 * I41 + 0.75 * I25)
        Dim I93 = 0.012 * (I19 + 0.75 * I41 + 0.75 * I66)

        Dim F94 = 0.001 * (M18 + M73)
        Dim G94 = 0.001 * (M18 + M41)
        Dim H94 = 0.001 * (M18 + 0.75 * M41 + 0.75 * M24)
        Dim I94 = 0.001 * (M18 + 0.75 * M41 + 0.75 * M66)

        Dim G95 = 0.012 * (I20 + I42)
        Dim H95 = 0.012 * (I20 + 0.75 * I42 + 0.75 * I26)
        Dim I95 = 0.012 * (I20 + 0.75 * I42 + 0.75 * I67)

        Dim G96 = 0.001 * (M18 + M42)
        Dim H96 = 0.001 * (M18 + 0.75 * M42 + 0.75 * M24)
        Dim I96 = 0.001 * (M18 + 0.75 * M42 + 0.75 * M67)

        Dim G97 = 0.012 * (I18 + I48)
        Dim H97 = 0.012 * (I18 + 0.75 * I48 + 0.75 * I24)
        Dim I97 = 0.012 * (I18 + 0.75 * I48 + 0.75 * I73)

        Dim G98 = 0.001 * (M18 + M48)
        Dim H98 = 0.001 * (M18 + 0.75 * M48 + 0.75 * M24)
        Dim I98 = 0.001 * (M18 + 0.75 * M48 + 0.75 * M73)

        Dim G99 = 0.012 * (I19 + I49)
        Dim H99 = 0.012 * (I19 + 0.75 * I49 + 0.75 * I25)
        Dim I99 = 0.012 * (I19 + 0.75 * I49 + 0.75 * I74)

        Dim G100 = 0.001 * (M18 + M48)
        Dim H100 = 0.001 * (M18 + 0.75 * M48 + 0.75 * M24)
        Dim I100 = 0.001 * (M18 + 0.75 * M48 + 0.75 * M73)

        Dim G101 = 0.012 * (I18 + I56) : Dim H101 = 0.012 * (I18 + 0.75 * I56 + 0.75 * I24) : Dim I101 = 0.012 * (I18 + 0.75 * I56 + 0.75 * I73)
        Dim G102 = 0.001 * (M18 + M56) : Dim H102 = 0.001 * (M18 + 0.75 * M56 + 0.75 * M24) : Dim I102 = 0.001 * (M18 + 0.75 * M56 + 0.75 * M73)
        Dim G103 = 0.012 * (I19 + I57) : Dim H103 = 0.012 * (I19 + 0.75 * I57 + 0.75 * I25) : Dim I103 = 0.012 * (I19 + 0.75 * I57 + 0.75 * I74)
        Dim G104 = 0.001 * (M18 + M56) : Dim H104 = 0.001 * (M18 + 0.75 * M56 + 0.75 * M24) : Dim I104 = 0.001 * (M18 + 0.75 * M56 + 0.75 * M73)

        '* Stress Ratios:
        'C110 = "8.0x4.0C14"
        Dim C110 = 9
        Dim maxval = 0
        Dim thisval = 0

        If E86 < 0 Then thisval = 0.1239 * Math.Abs(E86) + 0.0125 * Math.Abs(E85) : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.1239 * Math.Abs(F86) + 0.0125 * Math.Abs(F85) : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.1239 * Math.Abs(G86) + 0.0125 * Math.Abs(G85) : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.1239 * Math.Abs(H86) + 0.0125 * Math.Abs(H85) : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.1239 * Math.Abs(I86) + 0.0125 * Math.Abs(I85) : If thisval > maxval Then maxval = thisval

        If E86 < 0 Then thisval = 0.0536 * Math.Abs(E86) + 0.0145 * Math.Abs(E85) : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0536 * Math.Abs(F86) + 0.0145 * Math.Abs(F85) : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0536 * Math.Abs(G86) + 0.0145 * Math.Abs(G85) : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0536 * Math.Abs(H86) + 0.0145 * Math.Abs(H85) : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0536 * Math.Abs(I86) + 0.0145 * Math.Abs(I85) : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.1239 * Math.Abs(E88) + 0.0125 * Math.Abs(E87) : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.1239 * Math.Abs(F88) + 0.0125 * Math.Abs(F87) : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.1239 * Math.Abs(G88) + 0.0125 * Math.Abs(G87) : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.1239 * Math.Abs(H88) + 0.0125 * Math.Abs(H87) : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.1239 * Math.Abs(I88) + 0.0125 * Math.Abs(I87) : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0536 * Math.Abs(E88) + 0.0145 * Math.Abs(E87) : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0536 * Math.Abs(F88) + 0.0145 * Math.Abs(F87) : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0536 * Math.Abs(G88) + 0.0145 * Math.Abs(G87) : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0536 * Math.Abs(H88) + 0.0145 * Math.Abs(H87) : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0536 * Math.Abs(I88) + 0.0145 * Math.Abs(I87) : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.1239 * Math.Abs(F90) + 0.0125 * Math.Abs(F89) : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.1239 * Math.Abs(G90) + 0.0125 * Math.Abs(G89) : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.1239 * Math.Abs(H90) + 0.0125 * Math.Abs(H89) : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.1239 * Math.Abs(I90) + 0.0125 * Math.Abs(I89) : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0536 * Math.Abs(F90) + 0.0145 * Math.Abs(F89) : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0536 * Math.Abs(G90) + 0.0145 * Math.Abs(G89) : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0536 * Math.Abs(H90) + 0.0145 * Math.Abs(H89) : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0536 * Math.Abs(I90) + 0.0145 * Math.Abs(I89) : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.1239 * Math.Abs(F92) + 0.0125 * Math.Abs(F91) : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.1239 * Math.Abs(G92) + 0.0125 * Math.Abs(G91) : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.1239 * Math.Abs(H92) + 0.0125 * Math.Abs(H91) : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.1239 * Math.Abs(I92) + 0.0125 * Math.Abs(I91) : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0536 * Math.Abs(F92) + 0.0145 * Math.Abs(F91) : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0536 * Math.Abs(G92) + 0.0145 * Math.Abs(G91) : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0536 * Math.Abs(H92) + 0.0145 * Math.Abs(H91) : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0536 * Math.Abs(I92) + 0.0145 * Math.Abs(I91) : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.1239 * Math.Abs(F94) + 0.0125 * Math.Abs(F93) : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.1239 * Math.Abs(G94) + 0.0125 * Math.Abs(G93) : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.1239 * Math.Abs(H94) + 0.0125 * Math.Abs(H93) : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.1239 * Math.Abs(I94) + 0.0125 * Math.Abs(I93) : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0536 * Math.Abs(F94) + 0.0145 * Math.Abs(F93) : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0536 * Math.Abs(G94) + 0.0145 * Math.Abs(G93) : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0536 * Math.Abs(H94) + 0.0145 * Math.Abs(H93) : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0536 * Math.Abs(I94) + 0.0145 * Math.Abs(I93) : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.1239 * Math.Abs(G96) + 0.0125 * Math.Abs(G95) : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.1239 * Math.Abs(H96) + 0.0125 * Math.Abs(H95) : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.1239 * Math.Abs(I96) + 0.0125 * Math.Abs(I95) : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0536 * Math.Abs(G96) + 0.0145 * Math.Abs(G95) : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0536 * Math.Abs(H96) + 0.0145 * Math.Abs(H95) : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0536 * Math.Abs(I96) + 0.0145 * Math.Abs(I95) : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.1239 * Math.Abs(G98) + 0.0125 * Math.Abs(G97) : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.1239 * Math.Abs(H98) + 0.0125 * Math.Abs(H97) : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.1239 * Math.Abs(I98) + 0.0125 * Math.Abs(I97) : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0536 * Math.Abs(G98) + 0.0145 * Math.Abs(G97) : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0536 * Math.Abs(H98) + 0.0145 * Math.Abs(H97) : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0536 * Math.Abs(I98) + 0.0145 * Math.Abs(I97) : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.1239 * Math.Abs(G100) + 0.0125 * Math.Abs(G99) : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.1239 * Math.Abs(H100) + 0.0125 * Math.Abs(H99) : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.1239 * Math.Abs(I100) + 0.0125 * Math.Abs(I99) : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0536 * Math.Abs(G100) + 0.0145 * Math.Abs(G99) : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0536 * Math.Abs(H100) + 0.0145 * Math.Abs(H99) : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0536 * Math.Abs(I100) + 0.0145 * Math.Abs(I99) : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.1239 * Math.Abs(G102) + 0.0125 * Math.Abs(G101) : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.1239 * Math.Abs(H102) + 0.0125 * Math.Abs(H101) : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.1239 * Math.Abs(I102) + 0.0125 * Math.Abs(I101) : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0536 * Math.Abs(G102) + 0.0145 * Math.Abs(G101) : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0536 * Math.Abs(H102) + 0.0145 * Math.Abs(H101) : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0536 * Math.Abs(I102) + 0.0145 * Math.Abs(I101) : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.1239 * Math.Abs(G104) + 0.0125 * Math.Abs(G103) : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.1239 * Math.Abs(H104) + 0.0125 * Math.Abs(H103) : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.1239 * Math.Abs(I104) + 0.0125 * Math.Abs(I103) : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0536 * Math.Abs(G104) + 0.0145 * Math.Abs(G103) : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0536 * Math.Abs(H104) + 0.0145 * Math.Abs(H103) : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0536 * Math.Abs(I104) + 0.0145 * Math.Abs(I103) : If thisval > maxval Then maxval = thisval

        Dim I130 = maxval
        '"Checked up to here"

        'C132 = "12.0x4.0C14"
        Dim C132 = 14
        maxval = 0
        If E86 < 0 Then thisval = 0.1072 * Math.Abs(E86) + 0.0089 * Math.Abs(E85) : If thisval > maxval Then maxval = thisval
        '=IF(E86<0,               0.1072 * ABS(E86) + 0.0089  *ABS(E85),0)
        If F86 < 0 Then thisval = 0.1072 * Math.Abs(F86) + 0.0089 * Math.Abs(F85) : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.1072 * Math.Abs(G86) + 0.0089 * Math.Abs(G85) : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.1072 * Math.Abs(H86) + 0.0089 * Math.Abs(H85) : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.1072 * Math.Abs(I86) + 0.0089 * Math.Abs(I85) : If thisval > maxval Then maxval = thisval

        If E86 < 0 Then thisval = 0.0572 * Math.Abs(E86) + 0.0105 * Math.Abs(E85) : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0572 * Math.Abs(F86) + 0.0105 * Math.Abs(F85) : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0572 * Math.Abs(G86) + 0.0105 * Math.Abs(G85) : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0572 * Math.Abs(H86) + 0.0105 * Math.Abs(H85) : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0572 * Math.Abs(I86) + 0.0105 * Math.Abs(I85) : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.1072 * Math.Abs(E88) + 0.0089 * Math.Abs(E87) : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.1072 * Math.Abs(F88) + 0.0089 * Math.Abs(F87) : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.1072 * Math.Abs(G88) + 0.0089 * Math.Abs(G87) : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.1072 * Math.Abs(H88) + 0.0089 * Math.Abs(H87) : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.1072 * Math.Abs(I88) + 0.0089 * Math.Abs(I87) : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0572 * Math.Abs(E88) + 0.0105 * Math.Abs(E87) : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0572 * Math.Abs(F88) + 0.0105 * Math.Abs(F87) : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0572 * Math.Abs(G88) + 0.0105 * Math.Abs(G87) : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0572 * Math.Abs(H88) + 0.0105 * Math.Abs(H87) : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0572 * Math.Abs(I88) + 0.0105 * Math.Abs(I87) : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.1072 * Math.Abs(F90) + 0.0089 * Math.Abs(F89) : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.1072 * Math.Abs(G90) + 0.0089 * Math.Abs(G89) : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.1072 * Math.Abs(H90) + 0.0089 * Math.Abs(H89) : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.1072 * Math.Abs(I90) + 0.0089 * Math.Abs(I89) : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0572 * Math.Abs(F90) + 0.0105 * Math.Abs(F89) : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0572 * Math.Abs(G90) + 0.0105 * Math.Abs(G89) : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0572 * Math.Abs(H90) + 0.0105 * Math.Abs(H89) : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0572 * Math.Abs(I90) + 0.0105 * Math.Abs(I89) : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.1072 * Math.Abs(F92) + 0.0089 * Math.Abs(F91) : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.1072 * Math.Abs(G92) + 0.0089 * Math.Abs(G91) : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.1072 * Math.Abs(H92) + 0.0089 * Math.Abs(H91) : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.1072 * Math.Abs(I92) + 0.0089 * Math.Abs(I91) : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0572 * Math.Abs(F92) + 0.0105 * Math.Abs(F91) : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0572 * Math.Abs(G92) + 0.0105 * Math.Abs(G91) : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0572 * Math.Abs(H92) + 0.0105 * Math.Abs(H91) : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0572 * Math.Abs(I92) + 0.0105 * Math.Abs(I91) : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.1072 * Math.Abs(F94) + 0.0089 * Math.Abs(F93) : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.1072 * Math.Abs(G94) + 0.0089 * Math.Abs(G93) : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.1072 * Math.Abs(H94) + 0.0089 * Math.Abs(H93) : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.1072 * Math.Abs(I94) + 0.0089 * Math.Abs(I93) : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0572 * Math.Abs(F94) + 0.0105 * Math.Abs(F93) : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0572 * Math.Abs(G94) + 0.0105 * Math.Abs(G93) : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0572 * Math.Abs(H94) + 0.0105 * Math.Abs(H93) : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0572 * Math.Abs(I94) + 0.0105 * Math.Abs(I93) : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.1072 * Math.Abs(G96) + 0.0089 * Math.Abs(G95) : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.1072 * Math.Abs(H96) + 0.0089 * Math.Abs(H95) : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.1072 * Math.Abs(I96) + 0.0089 * Math.Abs(I95) : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0572 * Math.Abs(G96) + 0.0105 * Math.Abs(G95) : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0572 * Math.Abs(H96) + 0.0105 * Math.Abs(H95) : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0572 * Math.Abs(I96) + 0.0105 * Math.Abs(I95) : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.1072 * Math.Abs(G98) + 0.0089 * Math.Abs(G97) : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.1072 * Math.Abs(H98) + 0.0089 * Math.Abs(H97) : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.1072 * Math.Abs(I98) + 0.0089 * Math.Abs(I97) : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0572 * Math.Abs(G98) + 0.0105 * Math.Abs(G97) : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0572 * Math.Abs(H98) + 0.0105 * Math.Abs(H97) : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0572 * Math.Abs(I98) + 0.0105 * Math.Abs(I97) : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.1072 * Math.Abs(G100) + 0.0089 * Math.Abs(G99) : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.1072 * Math.Abs(H100) + 0.0089 * Math.Abs(H99) : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.1072 * Math.Abs(I100) + 0.0089 * Math.Abs(I99) : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0572 * Math.Abs(G100) + 0.0105 * Math.Abs(G99) : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0572 * Math.Abs(H100) + 0.0105 * Math.Abs(H99) : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0572 * Math.Abs(I100) + 0.0105 * Math.Abs(I99) : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.1072 * Math.Abs(G102) + 0.0089 * Math.Abs(G101) : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.1072 * Math.Abs(H102) + 0.0089 * Math.Abs(H101) : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.1072 * Math.Abs(I102) + 0.0089 * Math.Abs(I101) : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0572 * Math.Abs(G102) + 0.0105 * Math.Abs(G101) : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0572 * Math.Abs(H102) + 0.0105 * Math.Abs(H101) : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0572 * Math.Abs(I102) + 0.0105 * Math.Abs(I101) : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.1072 * Math.Abs(G104) + 0.0089 * Math.Abs(G103) : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.1072 * Math.Abs(H104) + 0.0089 * Math.Abs(H103) : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.1072 * Math.Abs(I104) + 0.0089 * Math.Abs(I103) : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0572 * Math.Abs(G104) + 0.0105 * Math.Abs(G103) : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0572 * Math.Abs(H104) + 0.0105 * Math.Abs(H103) : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0572 * Math.Abs(I104) + 0.0105 * Math.Abs(I103) : If thisval > maxval Then maxval = thisval
        Dim I152 = maxval

        'C154 = "8.0x4.0C12"
        Dim C154 = 10
        maxval = 0
        If E86 < 0 Then thisval = 0.0597 * Math.Abs(E86) + 0.0071 * Math.Abs(E85) : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0597 * Math.Abs(F86) + 0.0071 * Math.Abs(F85) : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0597 * Math.Abs(G86) + 0.0071 * Math.Abs(G85) : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0597 * Math.Abs(H86) + 0.0071 * Math.Abs(H85) : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0597 * Math.Abs(I86) + 0.0071 * Math.Abs(I85) : If thisval > maxval Then maxval = thisval

        If E86 < 0 Then thisval = 0.0265 * Math.Abs(E86) + 0.0083 * Math.Abs(E85) : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0265 * Math.Abs(F86) + 0.0083 * Math.Abs(F85) : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0265 * Math.Abs(G86) + 0.0083 * Math.Abs(G85) : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0265 * Math.Abs(H86) + 0.0083 * Math.Abs(H85) : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0265 * Math.Abs(I86) + 0.0083 * Math.Abs(I85) : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0597 * Math.Abs(E88) + 0.0071 * Math.Abs(E87) : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0597 * Math.Abs(F88) + 0.0071 * Math.Abs(F87) : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0597 * Math.Abs(G88) + 0.0071 * Math.Abs(G87) : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0597 * Math.Abs(H88) + 0.0071 * Math.Abs(H87) : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0597 * Math.Abs(I88) + 0.0071 * Math.Abs(I87) : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0265 * Math.Abs(E88) + 0.0083 * Math.Abs(E87) : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0265 * Math.Abs(F88) + 0.0083 * Math.Abs(F87) : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0265 * Math.Abs(G88) + 0.0083 * Math.Abs(G87) : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0265 * Math.Abs(H88) + 0.0083 * Math.Abs(H87) : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0265 * Math.Abs(I88) + 0.0083 * Math.Abs(I87) : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0597 * Math.Abs(F90) + 0.0071 * Math.Abs(F89) : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0597 * Math.Abs(G90) + 0.0071 * Math.Abs(G89) : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0597 * Math.Abs(H90) + 0.0071 * Math.Abs(H89) : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0597 * Math.Abs(I90) + 0.0071 * Math.Abs(I89) : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0265 * Math.Abs(F90) + 0.0083 * Math.Abs(F89) : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0265 * Math.Abs(G90) + 0.0083 * Math.Abs(G89) : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0265 * Math.Abs(H90) + 0.0083 * Math.Abs(H89) : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0265 * Math.Abs(I90) + 0.0083 * Math.Abs(I89) : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0597 * Math.Abs(F92) + 0.0071 * Math.Abs(F91) : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0597 * Math.Abs(G92) + 0.0071 * Math.Abs(G91) : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0597 * Math.Abs(H92) + 0.0071 * Math.Abs(H91) : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0597 * Math.Abs(I92) + 0.0071 * Math.Abs(I91) : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0265 * Math.Abs(F92) + 0.0083 * Math.Abs(F91) : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0265 * Math.Abs(G92) + 0.0083 * Math.Abs(G91) : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0265 * Math.Abs(H92) + 0.0083 * Math.Abs(H91) : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0265 * Math.Abs(I92) + 0.0083 * Math.Abs(I91) : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0597 * Math.Abs(F94) + 0.0071 * Math.Abs(F93) : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0597 * Math.Abs(G94) + 0.0071 * Math.Abs(G93) : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0597 * Math.Abs(H94) + 0.0071 * Math.Abs(H93) : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0597 * Math.Abs(I94) + 0.0071 * Math.Abs(I93) : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0265 * Math.Abs(F94) + 0.0083 * Math.Abs(F93) : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0265 * Math.Abs(G94) + 0.0083 * Math.Abs(G93) : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0265 * Math.Abs(H94) + 0.0083 * Math.Abs(H93) : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0265 * Math.Abs(I94) + 0.0083 * Math.Abs(I93) : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0597 * Math.Abs(G96) + 0.0071 * Math.Abs(G95) : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0597 * Math.Abs(H96) + 0.0071 * Math.Abs(H95) : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0597 * Math.Abs(I96) + 0.0071 * Math.Abs(I95) : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0265 * Math.Abs(G96) + 0.0083 * Math.Abs(G95) : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0265 * Math.Abs(H96) + 0.0083 * Math.Abs(H95) : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0265 * Math.Abs(I96) + 0.0083 * Math.Abs(I95) : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0597 * Math.Abs(G98) + 0.0071 * Math.Abs(G97) : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0597 * Math.Abs(H98) + 0.0071 * Math.Abs(H97) : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0597 * Math.Abs(I98) + 0.0071 * Math.Abs(I97) : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0265 * Math.Abs(G98) + 0.0083 * Math.Abs(G97) : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0265 * Math.Abs(H98) + 0.0083 * Math.Abs(H97) : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0265 * Math.Abs(I98) + 0.0083 * Math.Abs(I97) : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0597 * Math.Abs(G100) + 0.0071 * Math.Abs(G99) : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0597 * Math.Abs(H100) + 0.0071 * Math.Abs(H99) : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0597 * Math.Abs(I100) + 0.0071 * Math.Abs(I99) : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0265 * Math.Abs(G100) + 0.0083 * Math.Abs(G99) : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0265 * Math.Abs(H100) + 0.0083 * Math.Abs(H99) : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0265 * Math.Abs(I100) + 0.0083 * Math.Abs(I99) : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0597 * Math.Abs(G102) + 0.0071 * Math.Abs(G101) : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0597 * Math.Abs(H102) + 0.0071 * Math.Abs(H101) : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0597 * Math.Abs(I102) + 0.0071 * Math.Abs(I101) : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0265 * Math.Abs(G102) + 0.0083 * Math.Abs(G101) : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0265 * Math.Abs(H102) + 0.0083 * Math.Abs(H101) : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0265 * Math.Abs(I102) + 0.0083 * Math.Abs(I101) : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0597 * Math.Abs(G104) + 0.0071 * Math.Abs(G103) : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0597 * Math.Abs(H104) + 0.0071 * Math.Abs(H103) : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0597 * Math.Abs(I104) + 0.0071 * Math.Abs(I103) : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0265 * Math.Abs(G104) + 0.0083 * Math.Abs(G103) : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0265 * Math.Abs(H104) + 0.0083 * Math.Abs(H103) : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0265 * Math.Abs(I104) + 0.0083 * Math.Abs(I103) : If thisval > maxval Then maxval = thisval
        Dim I174 = maxval

        'C176 = "12.0x4.0C12"
        Dim C176 = 15
        maxval = 0

        If E86 < 0 Then thisval = 0.0515 * Math.Abs(E86) + 0.0046 * Math.Abs(E85) : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0515 * Math.Abs(F86) + 0.0046 * Math.Abs(F85) : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0515 * Math.Abs(G86) + 0.0046 * Math.Abs(G85) : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0515 * Math.Abs(H86) + 0.0046 * Math.Abs(H85) : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0515 * Math.Abs(I86) + 0.0046 * Math.Abs(I85) : If thisval > maxval Then maxval = thisval

        If E86 < 0 Then thisval = 0.0281 * Math.Abs(E86) + 0.0054 * Math.Abs(E85) : If thisval > maxval Then maxval = thisval
        If F86 < 0 Then thisval = 0.0281 * Math.Abs(F86) + 0.0054 * Math.Abs(F85) : If thisval > maxval Then maxval = thisval
        If G86 < 0 Then thisval = 0.0281 * Math.Abs(G86) + 0.0054 * Math.Abs(G85) : If thisval > maxval Then maxval = thisval
        If H86 < 0 Then thisval = 0.0281 * Math.Abs(H86) + 0.0054 * Math.Abs(H85) : If thisval > maxval Then maxval = thisval
        If I86 < 0 Then thisval = 0.0281 * Math.Abs(I86) + 0.0054 * Math.Abs(I85) : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0515 * Math.Abs(E88) + 0.0046 * Math.Abs(E87) : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0515 * Math.Abs(F88) + 0.0046 * Math.Abs(F87) : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0515 * Math.Abs(G88) + 0.0046 * Math.Abs(G87) : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0515 * Math.Abs(H88) + 0.0046 * Math.Abs(H87) : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0515 * Math.Abs(I88) + 0.0046 * Math.Abs(I87) : If thisval > maxval Then maxval = thisval

        If E88 < 0 Then thisval = 0.0281 * Math.Abs(E88) + 0.0054 * Math.Abs(E87) : If thisval > maxval Then maxval = thisval
        If F88 < 0 Then thisval = 0.0281 * Math.Abs(F88) + 0.0054 * Math.Abs(F87) : If thisval > maxval Then maxval = thisval
        If G88 < 0 Then thisval = 0.0281 * Math.Abs(G88) + 0.0054 * Math.Abs(G87) : If thisval > maxval Then maxval = thisval
        If H88 < 0 Then thisval = 0.0281 * Math.Abs(H88) + 0.0054 * Math.Abs(H87) : If thisval > maxval Then maxval = thisval
        If I88 < 0 Then thisval = 0.0281 * Math.Abs(I88) + 0.0054 * Math.Abs(I87) : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0515 * Math.Abs(F90) + 0.0046 * Math.Abs(F89) : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0515 * Math.Abs(G90) + 0.0046 * Math.Abs(G89) : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0515 * Math.Abs(H90) + 0.0046 * Math.Abs(H89) : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0515 * Math.Abs(I90) + 0.0046 * Math.Abs(I89) : If thisval > maxval Then maxval = thisval

        If F90 < 0 Then thisval = 0.0281 * Math.Abs(F90) + 0.0054 * Math.Abs(F89) : If thisval > maxval Then maxval = thisval
        If G90 < 0 Then thisval = 0.0281 * Math.Abs(G90) + 0.0054 * Math.Abs(G89) : If thisval > maxval Then maxval = thisval
        If H90 < 0 Then thisval = 0.0281 * Math.Abs(H90) + 0.0054 * Math.Abs(H89) : If thisval > maxval Then maxval = thisval
        If I90 < 0 Then thisval = 0.0281 * Math.Abs(I90) + 0.0054 * Math.Abs(I89) : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0515 * Math.Abs(F92) + 0.0046 * Math.Abs(F91) : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0515 * Math.Abs(G92) + 0.0046 * Math.Abs(G91) : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0515 * Math.Abs(H92) + 0.0046 * Math.Abs(H91) : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0515 * Math.Abs(I92) + 0.0046 * Math.Abs(I91) : If thisval > maxval Then maxval = thisval

        If F92 < 0 Then thisval = 0.0281 * Math.Abs(F92) + 0.0054 * Math.Abs(F91) : If thisval > maxval Then maxval = thisval
        If G92 < 0 Then thisval = 0.0281 * Math.Abs(G92) + 0.0054 * Math.Abs(G91) : If thisval > maxval Then maxval = thisval
        If H92 < 0 Then thisval = 0.0281 * Math.Abs(H92) + 0.0054 * Math.Abs(H91) : If thisval > maxval Then maxval = thisval
        If I92 < 0 Then thisval = 0.0281 * Math.Abs(I92) + 0.0054 * Math.Abs(I91) : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0515 * Math.Abs(F94) + 0.0046 * Math.Abs(F93) : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0515 * Math.Abs(G94) + 0.0046 * Math.Abs(G93) : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0515 * Math.Abs(H94) + 0.0046 * Math.Abs(H93) : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0515 * Math.Abs(I94) + 0.0046 * Math.Abs(I93) : If thisval > maxval Then maxval = thisval

        If F94 < 0 Then thisval = 0.0281 * Math.Abs(F94) + 0.0054 * Math.Abs(F93) : If thisval > maxval Then maxval = thisval
        If G94 < 0 Then thisval = 0.0281 * Math.Abs(G94) + 0.0054 * Math.Abs(G93) : If thisval > maxval Then maxval = thisval
        If H94 < 0 Then thisval = 0.0281 * Math.Abs(H94) + 0.0054 * Math.Abs(H93) : If thisval > maxval Then maxval = thisval
        If I94 < 0 Then thisval = 0.0281 * Math.Abs(I94) + 0.0054 * Math.Abs(I93) : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0515 * Math.Abs(G96) + 0.0046 * Math.Abs(G95) : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0515 * Math.Abs(H96) + 0.0046 * Math.Abs(H95) : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0515 * Math.Abs(I96) + 0.0046 * Math.Abs(I95) : If thisval > maxval Then maxval = thisval

        If G96 < 0 Then thisval = 0.0281 * Math.Abs(G96) + 0.0054 * Math.Abs(G95) : If thisval > maxval Then maxval = thisval
        If H96 < 0 Then thisval = 0.0281 * Math.Abs(H96) + 0.0054 * Math.Abs(H95) : If thisval > maxval Then maxval = thisval
        If I96 < 0 Then thisval = 0.0281 * Math.Abs(I96) + 0.0054 * Math.Abs(I95) : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0515 * Math.Abs(G98) + 0.0046 * Math.Abs(G97) : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0515 * Math.Abs(H98) + 0.0046 * Math.Abs(H97) : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0515 * Math.Abs(I98) + 0.0046 * Math.Abs(I97) : If thisval > maxval Then maxval = thisval

        If G98 < 0 Then thisval = 0.0281 * Math.Abs(G98) + 0.0054 * Math.Abs(G97) : If thisval > maxval Then maxval = thisval
        If H98 < 0 Then thisval = 0.0281 * Math.Abs(H98) + 0.0054 * Math.Abs(H97) : If thisval > maxval Then maxval = thisval
        If I98 < 0 Then thisval = 0.0281 * Math.Abs(I98) + 0.0054 * Math.Abs(I97) : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0515 * Math.Abs(G100) + 0.0046 * Math.Abs(G99) : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0515 * Math.Abs(H100) + 0.0046 * Math.Abs(H99) : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0515 * Math.Abs(I100) + 0.0046 * Math.Abs(I99) : If thisval > maxval Then maxval = thisval

        If G100 < 0 Then thisval = 0.0281 * Math.Abs(G100) + 0.0054 * Math.Abs(G99) : If thisval > maxval Then maxval = thisval
        If H100 < 0 Then thisval = 0.0281 * Math.Abs(H100) + 0.0054 * Math.Abs(H99) : If thisval > maxval Then maxval = thisval
        If I100 < 0 Then thisval = 0.0281 * Math.Abs(I100) + 0.0054 * Math.Abs(I99) : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0515 * Math.Abs(G102) + 0.0046 * Math.Abs(G101) : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0515 * Math.Abs(H102) + 0.0046 * Math.Abs(H101) : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0515 * Math.Abs(I102) + 0.0046 * Math.Abs(I101) : If thisval > maxval Then maxval = thisval

        If G102 < 0 Then thisval = 0.0281 * Math.Abs(G102) + 0.0054 * Math.Abs(G101) : If thisval > maxval Then maxval = thisval
        If H102 < 0 Then thisval = 0.0281 * Math.Abs(H102) + 0.0054 * Math.Abs(H101) : If thisval > maxval Then maxval = thisval
        If I102 < 0 Then thisval = 0.0281 * Math.Abs(I102) + 0.0054 * Math.Abs(I101) : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0515 * Math.Abs(G104) + 0.0046 * Math.Abs(G103) : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0515 * Math.Abs(H104) + 0.0046 * Math.Abs(H103) : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0515 * Math.Abs(I104) + 0.0046 * Math.Abs(I103) : If thisval > maxval Then maxval = thisval

        If G104 < 0 Then thisval = 0.0281 * Math.Abs(G104) + 0.0054 * Math.Abs(G103) : If thisval > maxval Then maxval = thisval
        If H104 < 0 Then thisval = 0.0281 * Math.Abs(H104) + 0.0054 * Math.Abs(H103) : If thisval > maxval Then maxval = thisval
        If I104 < 0 Then thisval = 0.0281 * Math.Abs(I104) + 0.0054 * Math.Abs(I103) : If thisval > maxval Then maxval = thisval
        Dim I196 = maxval

        'If OutputCalcs Then Print #1, "xxxxX"
        '* Preliminary Rafter Profile selection :
        '=IF(I130<=1 then E220=C110,IF(I152<=1,C132,IF(I174<=1,C154,IF(I196<=1.02,C176,"Check braced apex rafter solution"))))
        Dim E202
        If I130 <= 1 Then
            E202 = C110
        ElseIf I152 <= 1 Then
            E202 = C132
        ElseIf I174 <= 1 Then
            E202 = C154
        ElseIf I196 <= 1.02 Then
            E202 = C176
        Else
            E202 = "Check braced apex rafter solution"
        End If
        'see E202

        '* Check of 12.0x4.0C12 apex brace  solution (Max. brace length = 13 ft):
        '1) check apex brace saftey:
        Dim mv = 100000
        Dim tv = 0.8 * E86 : If tv < mv Then mv = tv
        tv = 0.8 * F86 : If tv < mv Then mv = tv
        tv = 0.8 * G86 : If tv < mv Then mv = tv
        tv = 0.8 * H86 : If tv < mv Then mv = tv
        tv = 0.8 * I86 : If tv < mv Then mv = tv
        tv = 0.8 * E88 : If tv < mv Then mv = tv
        tv = 0.8 * F88 : If tv < mv Then mv = tv
        tv = 0.8 * G88 : If tv < mv Then mv = tv
        tv = 0.8 * H88 : If tv < mv Then mv = tv
        tv = 0.8 * I88 : If tv < mv Then mv = tv
        tv = 0.8 * F90 : If tv < mv Then mv = tv
        tv = 0.8 * G90 : If tv < mv Then mv = tv
        tv = 0.8 * H90 : If tv < mv Then mv = tv
        tv = 0.8 * I90 : If tv < mv Then mv = tv
        tv = 0.8 * F92 : If tv < mv Then mv = tv
        tv = 0.8 * G92 : If tv < mv Then mv = tv
        tv = 0.8 * H92 : If tv < mv Then mv = tv
        tv = 0.8 * I92 : If tv < mv Then mv = tv
        tv = 0.8 * F94 : If tv < mv Then mv = tv
        tv = 0.8 * G94 : If tv < mv Then mv = tv
        tv = 0.8 * H94 : If tv < mv Then mv = tv
        tv = 0.8 * I94 : If tv < mv Then mv = tv
        tv = 0.8 * G96 : If tv < mv Then mv = tv
        tv = 0.8 * H96 : If tv < mv Then mv = tv
        tv = 0.8 * I96 : If tv < mv Then mv = tv
        tv = 0.8 * G98 : If tv < mv Then mv = tv
        tv = 0.8 * H98 : If tv < mv Then mv = tv
        tv = 0.8 * I98 : If tv < mv Then mv = tv
        tv = 0.8 * G100 : If tv < mv Then mv = tv
        tv = 0.8 * H100 : If tv < mv Then mv = tv
        tv = 0.8 * I100 : If tv < mv Then mv = tv
        tv = 0.8 * G102 : If tv < mv Then mv = tv
        tv = 0.8 * H102 : If tv < mv Then mv = tv
        tv = 0.8 * I102 : If tv < mv Then mv = tv
        tv = 0.8 * G104 : If tv < mv Then mv = tv
        tv = 0.8 * H104 : If tv < mv Then mv = tv
        tv = 0.8 * I104 : If tv < mv Then mv = tv
        Dim F220 = Math.Abs(mv)
        Dim F222
        If F220 <= 16.86 Then F222 = "Apex brace is safe" Else F222 = "Apex brace not available"
        'see F220

        '2) check sagging moment sections:
        Dim E229 = 0.43 * E85 : Dim G229 = 0.012 * (0.43 * I18 + 0.81 * I32) : Dim H229 = 0.012 * (0.43 * I18 + 0.81 * 0.75 * I32 + 0.43 * 0.75 * I24)
        Dim E230 = 1.22 * E86 : Dim G230 = 0.001 * (1.22 * M18 + 1.2 * M32) : Dim H230 = 0.001 * (1.22 * M18 + 1.2 * 0.75 * M32 + 1.22 * 0.75 * M24)
        Dim F231 = 0.012 * (0.43 * I20 + 0.73 * I67) : Dim G231 = 0.012 * (0.43 * I20 + 1.46 * I34) : Dim H231 = 0.012 * (0.43 * I20 + 1.46 * 0.75 * I34 + 0.43 * 0.75 * I26) : Dim I231 = 0.012 * (0.43 * I20 + 0.75 * 1.46 * I34 + 0.73 * 0.75 * I67)
        Dim F232 = 0.001 * (1.22 * M18 + 0.87 * M67) : Dim G232 = 0.001 * (1.22 * M18 + 1.2 * M34) : Dim H232 = 0.001 * (1.22 * M18 + 1.2 * 0.75 * M34 + 1.22 * 0.75 * M24) : Dim I232 = 0.001 * (1.22 * M18 + 1.2 * 0.75 * M34 + 0.87 * 0.75 * M67)
        Dim F233 = 0.43 * F91 : Dim G233 = 0.012 * (0.43 * I18 + 1.8 * I40) : Dim H233 = 0.012 * (0.43 * I18 + 1.8 * 0.75 * I40 + 0.43 * 0.75 * I24)
        Dim F234 = 1.22 * F92 : Dim G234 = 0.001 * (1.22 * M18 + 1.2 * M40) : Dim H234 = 0.001 * (1.22 * M18 + 1.2 * 0.75 * M40 + 1.22 * 0.75 * M24)
        Dim G235 = 0.012 * (0.43 * I20 + 0.8 * I42) : Dim H235 = 0.012 * (0.43 * I20 + 0.8 * 0.75 * I42 + 0.43 * 0.75 * I26) : Dim I235 = 0.012 * (0.43 * I20 + 0.8 * 0.75 * I42 + 0.73 * 0.75 * I67)
        Dim G236 = 0.001 * (1.22 * M18 + 1.2 * M42) : Dim H236 = 0.001 * (1.22 * M18 + 1.2 * 0.75 * M42 + 1.22 * 0.75 * M24) : Dim I236 = 0.001 * (1.22 * M18 + 1.2 * 0.75 * M42 + 0.87 * 0.75 * M67)
        Dim G237 = 0.43 * G97 : Dim H237 = 0.43 * H97 : Dim I237 = 0.43 * I97
        Dim G238 = 1.22 * G98 : Dim H238 = 1.22 * H98 : Dim I238 = 1.22 * I98
        Dim G239 = 0.43 * G101 : Dim H239 = 0.43 * H101 : Dim I239 = 0.43 * I101
        Dim G240 = 1.22 * G102 : Dim H240 = 1.22 * H102 : Dim I240 = 1.22 * I102


        '12.0x4.0C12 S.R.
        mv = 0
        If E230 < 0 Then tv = 0.038 * Math.Abs(E230) + 0.0046 * Math.Abs(E229) : If tv > mv Then mv = tv
        If G230 < 0 Then tv = 0.038 * Math.Abs(G230) + 0.0046 * Math.Abs(G229) : If tv > mv Then mv = tv
        If H230 < 0 Then tv = 0.038 * Math.Abs(H230) + 0.0046 * Math.Abs(H229) : If tv > mv Then mv = tv

        If E230 < 0 Then tv = 0.0281 * Math.Abs(E230) + 0.0054 * Math.Abs(E229) : If tv > mv Then mv = tv
        If G230 < 0 Then tv = 0.0281 * Math.Abs(G230) + 0.0054 * Math.Abs(G229) : If tv > mv Then mv = tv
        If H230 < 0 Then tv = 0.0281 * Math.Abs(H230) + 0.0054 * Math.Abs(H229) : If tv > mv Then mv = tv

        If F232 < 0 Then tv = 0.038 * Math.Abs(F232) + 0.0046 * Math.Abs(F231) : If tv > mv Then mv = tv
        If G232 < 0 Then tv = 0.038 * Math.Abs(G232) + 0.0046 * Math.Abs(G231) : If tv > mv Then mv = tv
        If H232 < 0 Then tv = 0.038 * Math.Abs(H232) + 0.0046 * Math.Abs(H231) : If tv > mv Then mv = tv
        If I232 < 0 Then tv = 0.038 * Math.Abs(I232) + 0.0046 * Math.Abs(I231) : If tv > mv Then mv = tv

        If F232 < 0 Then tv = 0.0281 * Math.Abs(F232) + 0.0054 * Math.Abs(F231) : If tv > mv Then mv = tv
        If G232 < 0 Then tv = 0.0281 * Math.Abs(G232) + 0.0054 * Math.Abs(G231) : If tv > mv Then mv = tv
        If H232 < 0 Then tv = 0.0281 * Math.Abs(H232) + 0.0054 * Math.Abs(H231) : If tv > mv Then mv = tv
        If I232 < 0 Then tv = 0.0281 * Math.Abs(I232) + 0.0054 * Math.Abs(I231) : If tv > mv Then mv = tv

        If F234 < 0 Then tv = 0.038 * Math.Abs(F234) + 0.0046 * Math.Abs(F233) : If tv > mv Then mv = tv
        If G234 < 0 Then tv = 0.038 * Math.Abs(G234) + 0.0046 * Math.Abs(G233) : If tv > mv Then mv = tv
        If H234 < 0 Then tv = 0.038 * Math.Abs(H234) + 0.0046 * Math.Abs(H233) : If tv > mv Then mv = tv

        If F234 < 0 Then tv = 0.0281 * Math.Abs(F234) + 0.0054 * Math.Abs(F233) : If tv > mv Then mv = tv
        If G234 < 0 Then tv = 0.0281 * Math.Abs(G234) + 0.0054 * Math.Abs(G233) : If tv > mv Then mv = tv
        If H234 < 0 Then tv = 0.0281 * Math.Abs(H234) + 0.0054 * Math.Abs(H233) : If tv > mv Then mv = tv

        If G236 < 0 Then tv = 0.038 * Math.Abs(G236) + 0.0046 * Math.Abs(G235) : If tv > mv Then mv = tv
        If H236 < 0 Then tv = 0.038 * Math.Abs(H236) + 0.0046 * Math.Abs(H235) : If tv > mv Then mv = tv
        If I236 < 0 Then tv = 0.038 * Math.Abs(I236) + 0.0046 * Math.Abs(I235) : If tv > mv Then mv = tv

        If G236 < 0 Then tv = 0.0281 * Math.Abs(G236) + 0.0054 * Math.Abs(G235) : If tv > mv Then mv = tv
        If H236 < 0 Then tv = 0.0281 * Math.Abs(H236) + 0.0054 * Math.Abs(H235) : If tv > mv Then mv = tv
        If I236 < 0 Then tv = 0.0281 * Math.Abs(I236) + 0.0054 * Math.Abs(I235) : If tv > mv Then mv = tv

        If G238 < 0 Then tv = 0.038 * Math.Abs(G238) + 0.0046 * Math.Abs(G237) : If tv > mv Then mv = tv
        If H238 < 0 Then tv = 0.038 * Math.Abs(H238) + 0.0046 * Math.Abs(H237) : If tv > mv Then mv = tv
        If I238 < 0 Then tv = 0.038 * Math.Abs(I238) + 0.0046 * Math.Abs(I237) : If tv > mv Then mv = tv

        If G238 < 0 Then tv = 0.0281 * Math.Abs(G238) + 0.0054 * Math.Abs(G237) : If tv > mv Then mv = tv
        If H238 < 0 Then tv = 0.0281 * Math.Abs(H238) + 0.0054 * Math.Abs(H237) : If tv > mv Then mv = tv
        If I238 < 0 Then tv = 0.0281 * Math.Abs(I238) + 0.0054 * Math.Abs(I237) : If tv > mv Then mv = tv

        If G240 < 0 Then tv = 0.038 * Math.Abs(G240) + 0.0046 * Math.Abs(G239) : If tv > mv Then mv = tv
        If H240 < 0 Then tv = 0.038 * Math.Abs(H240) + 0.0046 * Math.Abs(H239) : If tv > mv Then mv = tv
        If I240 < 0 Then tv = 0.038 * Math.Abs(I240) + 0.0046 * Math.Abs(I239) : If tv > mv Then mv = tv

        If G240 < 0 Then tv = 0.0281 * Math.Abs(G240) + 0.0054 * Math.Abs(G239) : If tv > mv Then mv = tv
        If H240 < 0 Then tv = 0.0281 * Math.Abs(H240) + 0.0054 * Math.Abs(H239) : If tv > mv Then mv = tv
        If I240 < 0 Then tv = 0.0281 * Math.Abs(I240) + 0.0054 * Math.Abs(I239) : If tv > mv Then mv = tv
        Dim I256 = mv
        Dim f258
        If I256 <= 1.04 Then F258 = "Rafter Safe" Else F258 = "Rafter Unsafe"
        'RAFTER Profile
        Dim E262 = "No rafter found"
        Dim ApexBrace = False
        Dim sideBays
        'ValidS = False"INVALID"
        Dim ApexBraceInVal
        Dim BApexBrace
        Dim RafterInVal
        Dim ApexBracesInt
        Dim ProfileFound = False : Dim CertifiedProfile = False
        If E202 = "Check braced apex rafter solution" Then
            If F222 = "Apex brace is safe" And F258 = "Rafter Safe" Then
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


        'Close #1
        Dim Results = ""
        Dim FoundPitch
        Dim FoundSpan
        Dim FoundStubs
        Dim spanVar
        Dim FoundCap
        Dim capVar
        If ProfileFound Then
            FoundPitch = 1
            FoundSpan = SPAN
            FoundStubs = False
            If FoundSpan - SPAN <= 0 Then
                spanVar = 0
            Else
                spanVar = 100 / (SPAN / (FoundSpan - SPAN))
            End If
            FoundCap = BCapability
            If FoundCap < WindSpeedReq And FoundCap >= WindSpeedReqLe5 Then 'using 5% overload
                capVar = 1 - (WindSpeedReqLe5 / WindSpeedReq)
                capVar = 100 / (WindSpeedReq / (FoundCap - WindSpeedReq))
            ElseIf FoundCap - WindSpeedReq <= 0 Then
                capVar = 0
            Else
                capVar = 100 / (WindSpeedReq / (FoundCap - WindSpeedReq))
            End If
            Con.DebugOutput Format$(capVar, "000#.#00")

    HoistedLoadFound = 0 'BHoistedLoad

            'Results = Results & "Span =     " & Tabb & Format$(SPAN, "##0.0#") & Tabb & " >= " & Tabb & BSpan & Tabb & Format$(spanVar, "##0") & "%" & vbCrLf
            'Results = Results & Tabb & Tabb & "Sought" & Tabb & "  v" & Tabb & "Found" & Tabb & "Variation" & vbCrLf
            'Results = Results & "WallHeight= " & Tabb & Format$(WallHeight, "##0.0#") & Tabb & " >= " & Tabb & BHeight & Tabb & Format$(heightVar, "##0") & "%" & vbCrLf
            'Results = Results & "Roof " & Pitchs(RoofStyleIdx).Descr & Tabb & RoofStyleIdx & Tabb & "  = " & Tabb & BPitch & vbCrLf
            'Results = Results & "Profile Found. " ' & wc & vbCrLf & BldRs.Fields("BDescription").Value
            'FoundDescription = BDescription
            'set these values because profile is Found
            ' "Setting Mullions per end MullionInVal =" & MullionInVal
            MullionsPerEndIntIn = 0 'BMullions
            'If MullionsPerEndIntIn < 0 Then MullionsPerEndIntIn = olceil(SPAN / NomBayWidth, 0, "InitMulions")
            'If MullionsPerEndIntIn > 4 Then MullionsPerEndIntIn = 4
            '    If OutputCalcs Then Print #1, MullionsPerEndInt & " End-wall mullions of " & MullionInS '& vbCrLf & "Put an overide here !", 16

            If RafterInVal <= 0 Or RafterInVal > MaxFrameMtrls Then AddNotification "Rafter material is invalid": RafterInVal = 6 ': ValidS = "INVALID": Exit Sub
            If OutputCalcs Then Print #1, "Rafter = " & FrameMtrl(RafterInVal).Mtrl & " on line " & rLine

    ColumnInVal = 7 'BColumn
            MullionInVal = 7 'BMullion
            KneeBraceInVal = 0 'BKneeBrace
        Else    'profile not found
            CertifiedProfile = False
            FoundSpan = SPAN : spanVar = 100
            FoundHeight = WallHeight : heightVar = 100
            FoundStubs = False
            FoundNomBayWidth = SearchNomBayWidth ': NomBayWidthVar = 0
            spanVar = 0 : heightVar = 0 : NomBayWidthVar = 0
            FoundCap = 0 : capVar = 100
            FoundDescription = "Profile NOT Found * "
            'set these values to DEFAULTS because profile is NOT Found
            RafterInVal = 6 : If OutputCalcs Then Print #1, "Rafter = 2 x " & FrameMtrl(RafterInVal).Mtrl
    MullionInVal = 4 : If OutputCalcs Then Print #1, "Mullion = 1 x " & FrameMtrl(ColumnInVal).Mtrl
    ApexBraceInVal = 6 : If OutputCalcs Then Print #1, "Apex Brace = 2 x " & FrameMtrl(ApexBraceInVal).Mtrl
    ApexBracesInt = sideBays * 2
            HoistedLoadFound = 0
            ValidS = "INVALID"
        End If

        CheckApexBraceOveride

        If CertifiedProfile Then
            If OutputCalcs Then Print #1, "Certified profile found"
Else
            If OutputCalcs Then Print #1, "Profile found is Un Certified"
End If

        ApexBraceVal = ApexBraceInVal
        ApexBraceItm = FrameMtrl(ApexBraceVal).Itm
        ApexBraceLth = SPAN / 3
        ApexBraceDescS = itmdescr(ApexBraceItm)
        ApexBraceS = ""
        ApexBrace = itmprice(ApexBraceItm)
        ApexBraceC = 0
        ApexBraceHeight = UnderApexHeight - ((ApexBraceLth / 2) * vertRiseM)

        Results = Results & FoundDescription & vbCrLf
        Results = Results & Tabb & Tabb & "Sought" & Tabb & "  v" & Tabb & "Found" & Tabb & "Variation" & vbCrLf & vbCrLf
        Results = Results & "Span       " & Tabb & Format$(SPAN, "##0.0#") & Tabb & " >= " & Tabb & BSpan & Tabb & Format$(spanVar, "##0") & "%" & vbCrLf
        Results = Results & "Pitch " & Pitchs(RoofStyleIdx).Descr & Tabb & RoofStyleIdx & Tabb & "  = " & Tabb & BPitch & vbCrLf
        Results = Results & "Using Stubs " & FoundStubs & vbCrLf
        Results = Results & "Capability " & Tabb & Format$(WindSpeedReq, "##0.0#") & Tabb & " >= " & Tabb & FoundCap & Tabb & Format$(capVar, "##0") & "%" & vbCrLf
        ' Results
        If OutputCalcs Then Print #1,
If OutputCalcs Then Print #1, "Profile search completed"
If OutputCalcs Then Print #1,
If OutputCalcs Then Print #1, Results

variations:
        heightVar = 0
        allVar = spanVar + capVar
        Msg = "Span " & SPAN & Tabb & FoundSpan & Tabb & spanVar & nl
        'Msg = Msg & "Height " & WallHeight & Tabb & FoundHeight & Tabb & heightVar & NL
        Msg = Msg & "Cap " & WindSpeedReq & Tabb & FoundCap & Tabb & capVar & nl
        Msg = Msg & "All " & allVar
        If spanVar > 50 Then spanVarColour = RGB(0, 200, 0) Else spanVarColour = RGB(0, spanVar * 5, 0)
        If heightVar > 50 Then heightVarColour = RGB(0, 200, 0) Else heightVarColour = RGB(0, heightVar * 4, 0)
        If capVar > 50 Then
            capVarColour = RGB(200, 0, 0)
        ElseIf capVar < 0 Then
            capVarColour = RGB(-capVar * 50, 0, 0)
        Else
            capVarColour = RGB(0, capVar * 5, 0)
        End If
        If allVar > 100 Then allVarColour = RGB(250, 0, 0) Else allVarColour = RGB(Abs(allVar) * 2, 0, 0)
        ' Msg
        '"SpanVar=" & spanVar & vbCrLf & "FoundSpan" & FoundSpan & vbCrLf & "SearchSpan=" & SearchSpan

        'SpanVars set here
        If OutputCalcs Then Print #1,
If OutputCalcs Then Print #1, "End of load analysis"
Close #1
' "See ProfileLookup.txt"

'needs USA version of this
GetFootingDetails FootingType   'this came to Wintrang from USA

EOFUSFindGableRafter:
        Exit Sub

    End Sub


    Sub USWindPressureGable()
        'If OutputCalcs Then Print #1, "USA Gable wind prewgbure calculation"
        'If OutputCalcs Then Print #1, "Basic Wind Speed selected =  " & Region
        Console.WriteLine("Enter SPAN:")
        Dim SPAN = Console.ReadLine()
        Console.WriteLine("EnterNomBayWidth:")
        Dim NomBayWidth = Console.ReadLine()

        Dim wgbE9 = SPAN
        Dim wgbE10 = NomBayWidth
        Dim wgbE11 = 8.5
        Dim wgbE12 = 3.88
        Dim wgbE13 = wgbE11 + 0.5 * wgbE9 * Math.Tan(0.5235987756)
        'see "wgbE13=" & wgbE13 & " V heightoa=" & heightOA
        Dim wgbE17 = 110
        'If OutputCalcs Then Print #1, "Exposure Category selected =  " & TerCatDescr(TerrainIdx)
        'If OutputCalcs Then Print #1, "Directionality Factor (Kd) = 0.85"
        Dim wgbE18 = 0.85 'Directionality Factor (Kd)
        Dim wgbE19 = "B"
        'If OutputCalcs Then Print #1, "Risk Factor selected =  " & ImportanceDescr(ImportanceIndex)
        'If OutputCalcs Then Print #1, "Topographic factor (Kzt) =1 "
        Dim wgbE20 = 1
        Dim wgbE23 = (wgbE11 + wgbE13) * 0.6 * 0.5
        Dim wgbU21, wgbV21, wgbW21, wgbX21, wgbY21, wgbZ21

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
        Dim wgbT24
        If wgbE23 > wgbX21 Then wgbT24 = wgbE23 Else wgbT24 = wgbX21

        Dim wgbE24 = wgbV21 * ((wgbT24 / 33) ^ (wgbW21))
        Dim wgbE25 = Math.Sqrt(1 / (1 + 0.63 * (((wgbE10 + wgbE11) / wgbE24) ^ 0.63)))
        Dim wgbE26 = wgbU21 * ((33 / wgbE23) ^ (1 / 6))
        Dim wgbE27 = 3.4
        Dim wgbE28 = 3.4
        Dim wgbE30 = 0.925 * (1 + 1.7 * wgbE28 * wgbE26 * wgbE25) / (1 + 1.7 * wgbE27 * wgbE26)

        'Awgbumed
        'Enclosure clawgbification :  Open structure
        'Wind flow   Clear
        'Internal prewgbure coeffecient (CGpi) =  0.00

        Dim wgbE38 = 0.5 * (wgbE11 + wgbE13)
        Dim wgbE39 = 2.01 * ((15 / wgbZ21) ^ (2 / wgbY21))
        Dim wgbE40 = 2.01 * ((wgbE38 / wgbZ21) ^ (2 / wgbY21))
        Dim wgbE41
        If wgbE38 < 15 Then
            wgbE41 = wgbE39
        Else
            wgbE41 = wgbE40
        End If

        'Velocity Prewgbure (qh):
        Dim wgbE44 = 0.00256 * wgbE41 * wgbE20 * wgbE18 * (wgbE17 ^ 2)
        '=0.00256    *  E41   *   E20  *  E18     *(E17^2)
        'qh = 0.00256*Kz*Kzt*Kd*(V^2) =
        Dim Wu = wgbE44
        WindPrewgbure = Wu   'Ultimate Dynamic Wind Prewgbure
        ' "Windprewgbure = " & WindPrewgbure

        'Purlin load
        'Roof load case 1 - Wind 0? - Case A
        Dim wgbE33 = "CLEAR"
        Dim wgbG50, wgbG51
        Select Case wgbE33
            Case "CLEAR"
                wgbG50 = 1.1     'Zone 2 for 30deg pitch only !
                wgbG51 = -0.333  'Zone 3 for 30deg pitch only !
            Case "OBSTRUCTED"
                wgbG50 = -1.467  'Zone 2 for 30deg pitch only !
                wgbG51 = -1      'Zone 3 for 30deg pitch only !
        End Select
        Dim wgbF50 = wgbE38
        Dim wgbF51 = wgbE38
        Dim wgbH50 = wgbE44
        Dim wgbH51 = wgbE44
        Dim wgbI50 = wgbG50 * wgbH50 * wgbE30 : Dim Wind_GableI50 = wgbI50
        '=$G$50*$H$50*$E$30
        Dim wgbI51 = wgbG51 * wgbH51 * wgbE30 : Dim Wind_GableI51 = wgbI51
        Dim wgbJ50 = wgbE12 * wgbI50
        Dim wgbJ51 = wgbE12 * wgbI51

        ' "Purlin load (lb/ft') = " & wgbJ50
        'multiply this by length of purlin ie baywidth to get equally distributed load per purlin

        '-Roof load case 2 - Wind 0? - Case B
        Dim wgbG55, wgbG56
        Select Case wgbE33
            Case "CLEAR"
                wgbG55 = 0.167   'Zone 2 for 10deg pitch only !
                wgbG56 = -1.167  'Zone 3 for 10deg pitch only !
            Case "OBSTRUCTED"
                wgbG55 = -0.8    'Zone 2 for 10deg pitch only !
                wgbG56 = -1.667  'Zone 3 for 10deg pitch only !
        End Select
        Dim wgbF55 = wgbE38
        Dim wgbF56 = wgbE38
        Dim wgbH55 = wgbE44
        Dim wgbH56 = wgbE44
        Dim wgbI55 = wgbG55 * wgbH55 * wgbE30 : Dim Wind_GableI55 = wgbI55
        Dim wgbI56 = wgbG56 * wgbH56 * wgbE30 : Dim Wind_GableI56 = wgbI56
        Dim wgbJ55 = wgbE12 * wgbI55
        Dim wgbJ56 = wgbE12 * wgbI56

        '-Roof load case 3 - Wind 90? - Case A
        Dim wgbG60, wgbG61
        Select Case wgbE33
            Case "CLEAR"
                wgbG60 = -0.8
                wgbG61 = -0.6
            Case "OBSTRUCTED"
                wgbG60 = -1.2
                wgbG61 = -0.9
        End Select
        Dim wgbF60 = wgbE38
        Dim wgbF61 = wgbE38
        Dim wgbH60 = wgbE44
        Dim wgbH61 = wgbE44
        Dim wgbI60 = wgbG60 * wgbH60 * wgbE30 : Dim Wind_GableI60 = wgbI60
        Dim wgbI61 = wgbG61 * wgbH61 * wgbE30 : Dim Wind_GableI61 = wgbI61
        Dim wgbJ60 = wgbE12 * wgbI60
        Dim wgbJ61 = wgbE12 * wgbI61

        '-Roof load case 4 - Wind 90? - Case B
        Dim wgbG65, wgbG66
        Select Case wgbE33
            Case "CLEAR"
                wgbG65 = 0.8
                wgbG66 = 0.5
            Case "OBSTRUCTED"
                wgbG65 = 0.5
                wgbG66 = 0.5
        End Select
        Dim wgbF65 = wgbE38
        Dim wgbF66 = wgbE38
        Dim wgbH65 = wgbE44
        Dim wgbH66 = wgbE44
        Dim wgbI65 = wgbG65 * wgbH65 * wgbE30 : Dim Wind_GableI65 = wgbI65
        '      =$G$65 * $H$65 * $E$30
        Dim wgbI66 = wgbG66 * wgbH66 * wgbE30
        Dim wgbJ65 = wgbE12 * wgbI65
        Dim wgbJ66 = wgbE12 * wgbI66


        'Uplift for footings
        Dim wgbE75 = 0.5 * wgbE9 * wgbE10 / Math.Cos(0.5235987756) 'Zones 2 & 3 surface area


        Dim wgbU73 = 0.5 * (wgbE13 - wgbE11)
        Dim wgbU75 = 0.25 * (wgbE9)

        '-Roof load case 1 - Wind 0? - Case A
        'zone 2 and 3
        Dim wgbF78 = wgbE75 * wgbI50 / 1000 'Total Net force (Kips) Zone 2
        Dim wgbG78 = wgbE75 * wgbI51 / 1000 'Total Net force (Kips) Zone 3
        Dim wgbF79 = wgbF78 * Math.Cos(30 * Math.PI / 180) * wgbU75    'Moment @ axis A (kips.ft) [Vl. forces] Zone 2
        Dim wgbG79 = wgbG78 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis A (kips.ft) [Vl. forces] Zone 3
        Dim wgbF80 = wgbF78 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 2
        Dim wgbG80 = -wgbG78 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 3
        Dim wgbF81 = wgbF79 + wgbG79 + wgbF80 + wgbG80
        Dim wgbF82 = wgbF78 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis B (kips.ft)        [Vl. forces]
        Dim wgbG82 = wgbG78 * Math.Cos(30 * Math.PI / 180) * wgbU75
        Dim wgbF83 = -wgbF78 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces]
        Dim wgbG83 = wgbG78 * Math.Sin(30 * Math.PI / 180) * wgbU73
        Dim wgbF84 = wgbF82 + wgbF83 + wgbG82 + wgbG83
        Dim wgbK81 = 0.5 * wgbF81 / wgbE9 * 1000 : Dim Wind_GableK81 = wgbK81    'RB1
        Dim wgbK84 = 0.5 * wgbF84 / wgbE9 * 1000 : Dim Wind_GableK84 = wgbK84 'RA1
        ' "Ra1=" & wgbK84

        'Now do all this again for 3 more cases

        '-Roof load case 2 - Wind 0? - Case B
        'zone 2 and 3
        Dim wgbF89 = wgbE75 * wgbI55 / 1000 'Total Net force (Kips) Zone 2
        Dim wgbG89 = wgbE75 * wgbI56 / 1000 'Total Net force (Kips) Zone 3
        Dim wgbF90 = wgbF89 * Math.Cos(30 * Math.PI / 180) * wgbU75    'Moment @ axis A (kips.ft) [Vl. forces] Zone 2
        Dim wgbG90 = wgbG89 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis A (kips.ft) [Vl. forces] Zone 3
        Dim wgbF91 = wgbF89 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 2
        Dim wgbG91 = -wgbG89 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 3
        Dim wgbF92 = wgbF90 + wgbG90 + wgbF91 + wgbG91
        Dim wgbF93 = wgbF89 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis B (kips.ft)        [Vl. forces]
        Dim wgbG93 = wgbG89 * Math.Cos(30 * Math.PI / 180) * wgbU75
        Dim wgbF94 = -wgbF89 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces]
        Dim wgbG94 = wgbG89 * Math.Sin(30 * Math.PI / 180) * wgbU73
        Dim wgbF95 = wgbF93 + wgbF94 + wgbG93 + wgbG94
        Dim wgbK92 = 0.5 * wgbF92 / wgbE9 * 1000 : Dim Wind_GableK92 = wgbK92 'RB2
        Dim wgbK95 = 0.5 * wgbF95 / wgbE9 * 1000 : Dim Wind_GableK95 = wgbK95 'RA2
        ' "Ra2=" & wgbK95

        '-Roof load case 3 - Wind 90? - Case A
        'zone 1-2 and 1-3
        Dim wgbE101 = 0.25 * wgbE9 * (wgbE11 + wgbE13) / Math.Cos(0.5235987756)            'Zones 1-2 & 1-3 surface area =
        Dim wgbE103 = 0.5 * wgbE9 * (wgbE10 - 0.5 * (wgbE11 + wgbE13)) / Math.Cos(0.5235987756)  'Zones 2-2 & 2-3 surface area =

        Dim wgbF106 = wgbE101 * wgbI60 / 1000 'Total Net force (Kips) Zone 1-2
        Dim wgbG106 = wgbE101 * wgbI60 / 1000 'Total Net force (Kips) Zone 1-3
        Dim wgbF107 = wgbF106 * Math.Cos(30 * Math.PI / 180) * wgbU75    'Moment @ axis A (kips.ft) [Vl. forces] Zone 1-2
        Dim wgbG107 = wgbG106 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis A (kips.ft) [Vl. forces] Zone 1-3
        Dim wgbF108 = wgbF106 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 1-2
        Dim wgbG108 = -wgbG106 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 1-3
        Dim wgbF109 = wgbF107 + wgbG107 + wgbF108 + wgbG108

        Dim wgbF111 = wgbE103 * wgbI61 / 1000  'Total Net force (Kips)
        Dim wgbG111 = wgbE103 * wgbI61 / 1000
        Dim wgbF112 = wgbF111 * Math.Cos(30 * Math.PI / 180) * wgbU75  'Moment @ axis A or B (kips.ft)  [Vl. forces]
        Dim wgbG112 = wgbG111 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75)
        Dim wgbF113 = wgbF111 * Math.Sin(30 * Math.PI / 180) * wgbU73    '[Hz. forces]
        Dim wgbG113 = -wgbG111 * Math.Sin(30 * Math.PI / 180) * wgbU73    '[Hz. forces]
        Dim wgbF114 = wgbF112 + wgbF113 + wgbG112 + wgbG113
        Dim wgbK109 = wgbF109 / wgbE9 'R1 =
        Dim wgbK114 = wgbF114 / wgbE9 'R2 =
        Dim wgbK117 = ((wgbK109 * (wgbE10 - 0.5 * wgbE11)) + (wgbK114 * 0.5 * (wgbE10 - wgbE11))) / wgbE10 * 1000 'R1AB =
        Dim wgbK119 = ((wgbK114 * 0.5 * (wgbE10 + wgbE11) + 0.5 * wgbK109 * wgbE11)) / wgbE10 * 1000            'R2AB =
        ' "R2AB case3=" & wgbK119

        '-Roof load case 4 - Wind 90? - Case B
        'zone 1-2 and 1-3
        Dim wgbF123 = wgbE101 * wgbI65 / 1000 'Total Net force (Kips) Zone 1-2
        Dim wgbG123 = wgbE101 * wgbI65 / 1000  'Total Net force (Kips) Zone 1-3    '?? Should use I66 ?
        Dim wgbF124 = wgbF123 * Math.Cos(30 * Math.PI / 180) * wgbU75    'Moment @ axis A (kips.ft) [Vl. forces] Zone 1-2
        Dim wgbG124 = wgbG123 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75) 'Moment @ axis A (kips.ft) [Vl. forces] Zone 1-3
        Dim wgbF125 = wgbF123 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 1-2
        Dim wgbG125 = -wgbG123 * Math.Sin(30 * Math.PI / 180) * wgbU73 '[Hz. forces] Zone 1-3
        Dim wgbF126 = wgbF124 + wgbG124 + wgbF125 + wgbG125

        Dim wgbF128 = wgbE103 * wgbI61 / 1000  'Total Net force (Kips)
        Dim wgbG128 = wgbE103 * wgbI61 / 1000
        Dim wgbF129 = wgbF128 * Math.Cos(30 * Math.PI / 180) * wgbU75  'Moment @ axis A or B (kips.ft)  [Vl. forces]
        Dim wgbG129 = wgbG128 * Math.Cos(30 * Math.PI / 180) * (wgbE9 - wgbU75)
        Dim wgbF130 = wgbF128 * Math.Sin(30 * Math.PI / 180) * wgbU73    '[Hz. forces]
        Dim wgbG130 = wgbG128 * Math.Sin(30 * Math.PI / 180) * wgbU73    '[Hz. forces]
        Dim wgbF131 = wgbF129 + wgbF130 + wgbG129 + wgbG130
        Dim wgbK126 = wgbF126 / wgbE9 'R1 =
        Dim wgbK131 = wgbF131 / wgbE9 'R2 =
        Dim wgbK134 = ((wgbK126 * (wgbE10 - 0.5 * wgbE11)) + (wgbK114 * 0.5 * (wgbE10 - wgbE11))) / wgbE10 * 1000 'R1AB =
        Dim wgbK136 = ((wgbK131 * 0.5 * (wgbE10 + wgbE11) + 0.5 * wgbK126 * wgbE11)) / wgbE10 * 1000            'R2AB =
        ' "R2AB case 4=" & wgbK136

        Dim wgbE138 = Math.Min(wgbK136, wgbK134, wgbK119, wgbK117, wgbK95, wgbK92, wgbK84, wgbK81)    'Max. Vl. upthrust on rafter support
        Dim UpliftAtCorner = 1 - wgbE138
        '=MIN(K136     ,K134,   K119,   K117,   K95,   K92,   K84,   K81)
        Dim wgbE139 = Math.Max(wgbK136, wgbK134, wgbK119, wgbK117, wgbK95, wgbK92, wgbK84, wgbK81)    'Max. Vl. down prewgb on rafter support
        ' "wgbE138=" & wgbE138
        Dim UpLift = wgbE138
        If OutputCalcs Then Print #1, "UpLift = WindPrewgbure * 0.9 giving " & UpLift
        Dim UpLoad = UpLift - 3    'factor of 3 deducted for weight of roof when used for footings
        If OutputCalcs Then Print #1, "Upload = UpLift - 3 =" & UpLoad & "  ie factor of 3 deducted for weight of roof when used for footings"
        Dim ColumnRoofTrib = SPAN * BayWidth * 0.25
        If OutputCalcs Then Print #1, "Column Tributary area =" & ColumnRoofTrib & " as Span x Bay width / 4"
        Dim FootingColumnUpload = UpLoad * ColumnRoofTrib
        If OutputCalcs Then Print #1, "FootingColumnUpload = Upload * ColumnRoofTrib = " & FootingColumnUpload

End Sub


    Sub USSnowLoadGable()
        'AddNotification "No US Gable snow load yet"
        Console.WriteLine("Enter SPAN:")
        Dim SPAN = Console.ReadLine()
        Console.WriteLine("EnterNomBayWidth:")
        Dim NomBayWidth = Console.ReadLine()

        Dim sgbE10 = SPAN 'B1 =
        Dim sgbE11 = NomBayWidth 'L =
        Dim sgbE12 = 3.88  'Eave (container) height =
        Dim sgbE13 = sgbE11 + 0.5 * SPAN * Math.Tan(0.5235987756)
        'Dim sgbE14 = heightOA 'Apex height =

        Dim sgbE16 = "I" 'Building Risk Category:
        Dim sgbE17 = 20  'Ground Snow load (Pg) in psf
        Dim sgbE18 = Math.Min((0.13 * sgbE17 + 14), 30) 'Density of snow (g)
        Dim sgbE19 = "B" 'Exposure Category (Terrain Type):
        Dim sgbE20 = "Fully Exposed" 'Exposure condition [Table 7-2] :
        'Exposure factor (Ce) [Table 7-2] :
        Dim sgbE21

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
        Dim sgbE22 = "Unheated Structure" 'Thermal condition [Table 7-3] :
        Dim sgbE23 = 1.2 'Thermal factor (Ct) [Table 7-3] :
        'Importance factor (Is) [Table 1.5-2] :   =VLOOKUP(E17,T24:U27,2)
        Dim sgbE24
        Select Case sgbE16
            Case "I" : sgbE24 = 0.8
            Case "II" : sgbE24 = 1
            Case "III" : sgbE24 = 1.1
            Case Else : sgbE24 = 1.2
        End Select

        Dim sgbE30 = sgbE24 * sgbE17   'Min. load for low slope roofs [Sect 7.3.4] (Pfmin) :
        Dim sgbE31 = 0.7 * sgbE21 * sgbE23 * sgbE24 * sgbE17 'Flat roof snow load [Sect 7.3.4] (Pf) :
        'Cold roof Snow factor (Ct>1.0)
        Dim sgbE33 = "Slippery"  'Roof surface type
        Dim sgbE34 = "Ventilated" 'Ventilation
        Dim sgbE35 = 0.73 'Roof slope factor [Fig 7-2c (dashed line)] (Cs):
        Dim sgbE37 = sgbE35 * sgbE31 'Balanced sloped snow load (Ps):
        Dim sgbE38 = 1.732   'Slope of roof = 1/tan(10?) :
        Dim sgbE40 = 0.3 * sgbE37 'Unbalanced Load(ps)(windward):
        Dim sgbE41 = sgbE37 'Unbalanced Load(ps)(leeward):
        Dim sgbE42 = SPAN / 2 'Length of eave to ridge for drift height:
        Dim sgbE43 = 0.43 * ((Math.Max(sgbE42, 20)) ^ 0.333) * ((sgbE17 + 10) ^ 0.25) - 1.5 'Drift Height(hdr):
        Dim sgbE45 = sgbE43 * sgbE18 / Math.Sqrt(sgbE38) 'rectangular surcharge(leeward):
        Dim sgbE46 = Math.Min(2.667 * sgbE43 * Math.Sqrt(sgbE38), 0.5 * SPAN) 'Length of rectangular surcharge :
        '-Roof load case 1 - Snow @ Wind 0? - Unbalanced
        Dim sgbF51 = sgbE40
        Dim sgbF52 = sgbE45 + sgbE37
        Dim sgbF56 = sgbE37
        Dim sgbF57 = sgbE37

        Dim sgbG51 = sgbE12 * sgbF51
        Dim sgbG52 = sgbE13 * sgbF52
        '-Roof load case 2 - Snow @ Wind 90? - Balanced
        Dim sgbG56 = sgbE12 * sgbF56
        Dim sgbG57 = sgbE12 * sgbF57
        '*Snow Vertical presgb on eave supports:

        '-Roof load case 1 - Snow @ Wind 0? - Unbalanced
        Dim sgbE66 = 0.5 * (SPAN) * sgbE10 / Math.Cos(0.5235987756) 'Zones 2 & 3 surface area =
        Dim sgbU67 = 0.25 * SPAN
        'Zone#2
        Dim sgbF69 = sgbE66 * sgbF51 / 1000
        Dim sgbF70 = sgbF69 * sgbU67
        'Zone#3
        Dim sgbG69 = sgbE66 * sgbF52 / 1000
        Dim sgbG70 = sgbG69 * (SPAN - sgbU67)

        Dim sgbF72 = sgbE70 + sgbH71
        Dim sgbK73 = sgbF69 * (SPAN - sgbU67)
        Dim sgbK72 = 0.5 * sgbF72 / SPAN * 1000

        sgbG73 = sgbG69 * sgbU67
        sgbF75 = sgbE73 + sgbH74
        sgbK75 = 0.5 * sgbF75 / SPAN * 1000

        '-Roof load case 2 - Snow @ Wind 90? - Balanced
        'Zone#2
        sgbF80 = sgbE66 * sgbF56 / 1000
        sgbF81 = sgbF80 * sgbU67
        'Zone#3
        sgbG80 = sgbE66 * sgbF57 / 1000
        sgbG81 = G80 * (SPAN - sgbU67)
        sgbF83 = sgbF81 + sgbF82 + sgbG81 + sgbG82
        sgbK83 = 0.5 * sgbF83 / SPAN * 1000
        sgbF84 = sgbF80 * (SPAN - sgbU67)
        sgbG84 = sgbG80 * sgbU67
        sgbF86 = sgbF84 + sgbF85 + sgbG84 + sgbG85
        sgbK86 = 0.5 * sgbF86 / SPAN * 1000
        span1 = Max(sgbK87, sgbK84, sgbK76, sgbK73) 'Max. Vl. presgb on rafter support =
        span0 = Max(sgbK86, sgbK83, sgbK75, sgbK72)

    End Sub


End Module
