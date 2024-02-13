

Sub ll()

Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim rutall As String, rutamp As String, rutator As String, rutampacc As String, rutass As String, rutatensor As String
Dim punto1 As Variant, punto2 As Variant, PI As Variant
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim Lola_550 As String, Lola_1100 As String, Lola_2200 As String, Lola_Bastidor As String, Lola_Cono As String
Dim l550 As Double, l1100 As Double, l2200 As Double, lgatominss As Double, l180 As Double, l360 As Double, l540 As Double, lbastidor As Double, repite As Double, lgatominmp As Double, lcono As Double, lespada As Double, ladaptador As Double, l90 As Double, l270 As Double, l450 As Double, lldesfase As Double
Dim MP_Husillo As String, zMP_Base As String, MP_Giro As String, MP_Fusible As String, mp_180 As String, mp_270 As String, mp_450 As String, ss_180 As String, mp_90 As String, ss_360 As String, ss_540 As String, ll_550 As String, ll_1100 As String, ll_2200 As String, ll_bastidor As String, ss_espada As String, ss_husillo As String, ss_gatoizq As String, ss_gatodrc As String, ss_adaptador As String, ss_llave As String
Dim Distancia As Double, lfija As Double, lfija1 As Double, lfija2 As Double
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double, Punto_aux3(0 To 2) As Double, lhueco As Double, nbastidor As Integer
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String, extremo1 As String, disposicion As String, terminacion As String, vista As String, vistamp As String
Dim Gcapa As Object
Dim n2200 As Integer, n1100 As Integer, n550 As Integer, n450 As Integer, n270 As Integer, n180mp As Integer, n90 As Integer, n540 As Integer, n360 As Integer, n180 As Integer, nespada As Integer, divhueco As Integer
Dim M20x90_4 As String, M20x50_4 As String, M20x60_4 As String, M16x40_4 As String, M16x40_8 As String, M16x40_16 As String, M24x110 As String

Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

Ncapa = "Mega"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Lolashor"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Slims"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30

On Error GoTo terminar

rutall = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Lola\"
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
rutampacc = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"
rutass = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\"
rutatensor = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensores\"


'Valores fijos
PI = 4 * Atn(1)
repite = 1
l90 = 90
l180 = 180
l270 = 270
l450 = 450
l360 = 360
l540 = 540
l550 = 550
l1100 = 1100
l2200 = 2200
lcono = 275.17
lbastidor = 1200
lgatominmp = 435
lgatominss = 420 'Se abre +/- 45mm. ya que hay dos gatos
lespada = 358
ladaptador = 117.25

On Error GoTo terminar

kwordList = "A B C D"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
disposicion = ThisDrawing.Utility.GetKeyword(vbLf & "Viga y disposición: A (Viga en planta) B (Viga en alzado) C (Bastidor en planta) D (Bastidor en alzado): [A/B/C/D]")

If disposicion = "" Or disposicion = "A" Or disposicion = "C" Then
vista = "PL"
vistamp = "PLA"
ElseIf disposicion = "B" Or disposicion = "D" Then
vista = ""
vistamp = "ALZ"
Else
GoTo terminar
End If

If disposicion = "" Or disposicion = "A" Or disposicion = "B" Then
    kwordList = "SS MP"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    terminacion = ThisDrawing.Utility.GetKeyword(vbLf & "Terminación en los extremos: SS o MP: [SS/MP]")

    'If terminacion = "" Or terminacion = "SS" Then
     '   kwordList = "A B"
     '   ThisDrawing.Utility.InitializeUserInput 0, kwordList
     '   extremo1 = ThisDrawing.Utility.GetKeyword(vbLf & "Terminación extremo 1: A (Gato SS) B (Espada SS): [A/B]")
     '   If extremo1 = "" Or extremo1 = "A" Then
     '       nespada = 0
     '   ElseIf extremo1 = "B" Then
     '       nespada = 1
     '   End If
    'End If

End If

If disposicion = "" Or disposicion = "A" Or disposicion = "B" Then
    If terminacion = "" Or terminacion = "SS" Then
        lfija = 2 * lcono + lgatominss + ladaptador + 415
    ElseIf terminacion = "MP" Then
        lfija = 2 * lcono + lgatominmp + 2 * l90
    End If
End If

M24x110 = rutator & "1-M24X110.dwg"
M16x40_4 = rutator & "4-M16X40.dwg"
M20x60_4 = rutator & "4-M20X60.dwg"
M16x40_8 = rutator & "8-M16X40.dwg"
M16x40_16 = rutator & "16-M16X40.dwg"

Do While repite = 1

    'Geometría:
    punto1 = gcadUtil.GetPoint(, "1º Punto: ")
    punto2 = gcadUtil.GetPoint(punto1, "2º Punto: ")
    P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
    P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)

    Set Eje1 = gcadModel.AddLine(P1, P2)
    ANG = gcadUtil.AngleFromXAxis(P1, P2)

    x = P2(0) - P1(0)
    y = P2(1) - P1(1)
    Xs = 1
    Ys = 1
    Zs = 1
    Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

    Punto_inicial(0) = P1(0): Punto_inicial(1) = P1(1): Punto_inicial(2) = P1(2)
    Punto_final(0) = Punto_inicial(0): Punto_final(1) = Punto_inicial(1): Punto_final(2) = Punto_inicial(2)


    If disposicion = "" Or disposicion = "A" Or disposicion = "B" Then

        If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."
        GoTo terminar
        End If

        lpuntal = Distancia - lfija
        n2200 = Fix(lpuntal / l2200)
        lpuntal = lpuntal - n2200 * l2200
        n1100 = Fix(lpuntal / l1100)
        lpuntal = lpuntal - n1100 * l1100
        n550 = Fix(lpuntal / l550)
        lpuntal = lpuntal - n550 * l550
 
        If terminacion = "" Or terminacion = "SS" Then
            n540 = Fix(lpuntal / l540)
            lpuntal = lpuntal - n540 * l540
            n360 = Fix(lpuntal / l360)
            lpuntal = lpuntal - n360 * l360
            n180 = Fix(lpuntal / l180)
            lpuntal = lpuntal - n180 * l180
            lregulacion = (Distancia - n2200 * l2200 - n1100 * l1100 - n550 * l550 - n540 * l540 - n360 * l360 - n180 * l180 + lgatominss - lfija)
        
            'If extremo1 = "B" Then
            ss_espada = rutass & "SS" & vista & "Espada.dwg"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_espada, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Slims"
            'Punto_inicial(0) = Punto_inicial(0) + lespada * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + lespada * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            'End If
    
            ss_husillo = rutass & "zSSHusillo.dwg"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_husillo, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Slims"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, M24x110, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Nonplot"
            
            Dim ss_tubop As String
            ss_tubop = rutator & "TuboPivote" & vista & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_tubop, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Slims"
            
            Punto_inicial(0) = Punto_inicial(0) + 415 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + 415 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            Punto_final(0) = Punto_inicial(0): Punto_final(1) = Punto_inicial(1): Punto_final(2) = Punto_inicial(2)
        
            ss_gatoizq = rutass & "SS" & vista & "Gatorefizq.dwg"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_gatoizq, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Slims"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_4, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Nonplot"
            
            ' aquí va la ESPADA ANTIGIRO para que no rote el sistema
            
            Dim ss_antigiro As String
            ss_antigiro = rutatensor & "ESPantigiro" & vista & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_antigiro, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Lolashor"
            
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete

        
            If n540 > 0 Then
                ss_540 = rutass & "SS" & vista & "0540.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_540, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Slims"
                Punto_final(0) = Punto_inicial(0) + l540 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l540 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            End If
            
        ElseIf terminacion = "MP" Then
            n450 = Fix(lpuntal / l450)
            lpuntal = lpuntal - n450 * l450
            n270 = Fix(lpuntal / l270)
            lpuntal = lpuntal - n270 * l270
            n180 = Fix(lpuntal / l180)
            lpuntal = lpuntal - n180 * l180
            n90 = Fix(lpuntal / l90)
            lpuntal = lpuntal - n90 * l90
            lregulacion = (Distancia - n2200 * l2200 - n1100 * l1100 - n550 * l550 - n450 * l450 - n270 * l270 - n180 * l180 - n90 * l90 + lgatominmp - lfija) / 2
            
            MP_Giro = rutampacc & "MG_AnguloGiro" & vistamp & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, MP_Giro, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"

            Punto_inicial(0) = Punto_inicial(0) + l90 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + l90 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            Punto_final(0) = Punto_inicial(0): Punto_final(1) = Punto_inicial(1): Punto_final(2) = Punto_inicial(2)
            
            If n450 > 0 Then
                mp_450 = rutamp & "Mshor450" & vistamp & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, mp_450, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M20x60_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If

            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x60_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            
        End If
            
        Lola_Cono = rutall & "Lola_Cono.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, Lola_Cono, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Lolashor"
        Punto_final(0) = Punto_final(0) + lcono * Cos(ANG): Punto_final(1) = Punto_final(1) + lcono * Sin(ANG): Punto_final(2) = Punto_final(2)

        
        If disposicion = "" Or disposicion = "A" Then
            Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
            Punto_aux1(0) = Punto_inicial(0) + 250 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inicial(1) + 250 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inicial(2)
            Punto_aux2(0) = Punto_inicial(0) + 250 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inicial(1) + 250 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inicial(2)
            
            If n1100 > 0 Then
                Lola_1100 = rutall & "Lola_1100" & vista & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, Lola_1100, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, Lola_1100, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, M16x40_8, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, M16x40_8, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l1100 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l1100 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If

            If n2200 > 0 Then
                i = 0
                Lola_2200 = rutall & "Lola_2200" & vista & ".dwg"
                Do While i < n2200
                    Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inicial(0) + 250 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inicial(1) + 250 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inicial(2)
                    Punto_aux2(0) = Punto_inicial(0) + 250 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inicial(1) + 250 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inicial(2)
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, Lola_2200, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Lolashor"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, Lola_2200, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Lolashor"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M16x40_8, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M16x40_8, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inicial(0) + l2200 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l2200 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
                    i = i + 1
                Loop
            End If

            If n550 > 0 Then
                Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                Punto_aux1(0) = Punto_inicial(0) + 250 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inicial(1) + 250 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inicial(2)
                Punto_aux2(0) = Punto_inicial(0) + 250 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inicial(1) + 250 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inicial(2)
                Lola_550 = rutall & "Lola_550" & vista & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, Lola_550, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, Lola_550, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, M16x40_8, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, M16x40_8, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l550 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l550 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If
            
            Punto_aux1(0) = Punto_final(0) + 250 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final(1) + 250 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final(2)
            Punto_aux2(0) = Punto_final(0) + 250 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final(1) + 250 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final(2)
            Set blockRef = gcadModel.InsertBlock(Punto_aux1, M16x40_8, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(Punto_aux2, M16x40_8, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
        End If
    
    
        If disposicion = "B" Then
            Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
      
            If n1100 > 0 Then
                Lola_1100 = rutall & "Lola_1100" & vista & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_1100, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_1100, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l1100 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l1100 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If

            If n2200 > 0 Then
                i = 0
                Lola_2200 = rutall & "Lola_2200" & vista & ".dwg"
                Do While i < n2200
                    Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                    Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_2200, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Lolashor"
                    Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_2200, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Lolashor"
                    Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_16, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inicial(0) + l2200 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l2200 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
                    i = i + 1
                Loop
            End If

            If n550 > 0 Then
                Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                Lola_550 = rutall & "Lola_550" & vista & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_550, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_550, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l550 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l550 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If
            Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
        End If

        Punto_inicial(0) = Punto_final(0) + lcono * Cos(ANG): Punto_inicial(1) = Punto_final(1) + lcono * Sin(ANG): Punto_inicial(2) = Punto_final(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_Cono, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Lolashor"
        Punto_final(0) = Punto_inicial(0): Punto_final(1) = Punto_inicial(1): Punto_final(2) = Punto_inicial(2)


        If terminacion = "" Or terminacion = "SS" Then
        
            If n360 > 0 Then
                Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                ss_360 = rutass & "SS" & vista & "0360.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_360, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Slims"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l360 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l360 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If
            
            If n180 > 0 Then
                Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                ss_180 = rutass & "SS" & vista & "0180.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_180, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Slims"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If
            
            Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
            ss_gatodrc = rutass & "SS" & vista & "Gatorefdrc.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_gatodrc, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Slims"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete

            Punto_inicial(0) = Punto_inicial(0) + lregulacion * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + lregulacion * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_husillo, Xs, Ys, Zs, ANG + PI)
            'blockRef.Layer = "Slims"
            
            'ss_llave = rutass & "SSLlave.dwg"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_llave, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Slims"
            
            'Punto_inicial(0) = Punto_inicial(0) + ladaptador * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + ladaptador * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            
            'ss_adaptador = rutass & "SSAdaptador" & vista & ".dwg"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_adaptador, Xs, Ys, Zs, ANG + PI)
            'blockRef.Layer = "Slims"
            
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, M24x110, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Nonplot"
            
            ss_adaptador = rutass & "SSAdaptador" & vista & ".dwg"
            Set blockRef = gcadModel.InsertBlock(P2, ss_adaptador, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Slims"
            
            Dim Bulond23 As String
            Bulond23 = rutator & "1M23_BULOND23.dwg"
            Set blockRef = gcadModel.InsertBlock(P2, Bulond23, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            
            P2(0) = P2(0) - ladaptador * Cos(ANG): P2(1) = P2(1) - ladaptador * Sin(ANG): P2(2) = P2(2)
            
            Set blockRef = gcadModel.InsertBlock(P2, ss_husillo, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Slims"
            
            ss_llave = rutass & "SSLlave.dwg"
            Set blockRef = gcadModel.InsertBlock(P2, ss_llave, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Slims"
            
            
        ElseIf terminacion = "MP" Then
        
            If n270 > 0 Then
                Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                mp_270 = rutamp & "Mshor270" & vistamp & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, mp_270, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M20x60_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If

            If n180 > 0 Then
                Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                mp_180 = rutamp & "Mshor180" & vistamp & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, mp_180, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M20x60_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If
     
            If n90 > 0 Then
                Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
                mp_90 = rutamp & "Mshor90" & vistamp & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, mp_90, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, M20x60_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inicial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inicial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inicial(2)
            End If
     
     
            Punto_inicial(0) = Punto_final(0): Punto_inicial(1) = Punto_final(1): Punto_inicial(2) = Punto_final(2)
            
            zMP_Base = rutampacc & "zMGBaseGato_azul.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, zMP_Base, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, M20x60_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            
            Punto_inicial(0) = Punto_inicial(0) + lregulacion * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + lregulacion * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            
            MP_Husillo = rutampacc & "MGHusilloGato.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, MP_Husillo, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"

            Punto_inicial(0) = Punto_inicial(0) + lregulacion * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + lregulacion * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
    
            zMP_Base = rutampacc & "zMGBaseGato_naranja.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, zMP_Base, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Mega"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, M20x60_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            
            Punto_inicial(0) = Punto_inicial(0) + l90 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + l90 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, MP_Giro, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Mega"
    
        End If
        


    ElseIf disposicion = "C" Or disposicion = "D" Then
        
        lldesfase = 0
        If disposicion = "D" Then lldesfase = 280.1

        nbastidor = InputBox("Introduce el número de arriostramientos necesarios:", "Número de arriostramientos", 1)
        If nbastidor = 0 Then Exit Sub
        lhueco = (Distancia - nbastidor * 1200) / nbastidor
        
        If lhueco < 0 Then
        MsgBox "Distancia entre bastidores negativa " & lhueco & "mm, reducir el número de bastidores."
        GoTo terminar
        End If
        
        Lola_Bastidor = rutall & "Lola_Bastidor" & vista & ".dwg"

        If nbastidor > 0 Then
            i = 0
            Do While i < nbastidor
                If i = 0 Then
                divhueco = 2
                Else
                divhueco = 1
                End If
                Punto_inicial(0) = Punto_inicial(0) + lhueco / divhueco * Cos(ANG) + lldesfase * Cos(ANG + PI / 2): Punto_inicial(1) = Punto_inicial(1) + lhueco / divhueco * Sin(ANG) + lldesfase * Sin(ANG + PI / 2): Punto_inicial(2) = Punto_inicial(2)
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_Bastidor, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Lolashor"
                Punto_inicial(0) = Punto_inicial(0) + lbastidor * Cos(ANG) - 2 * lldesfase * Cos(ANG + PI / 2): Punto_inicial(1) = Punto_inicial(1) + lbastidor * Sin(ANG) - 2 * lldesfase * Sin(ANG + PI / 2): Punto_inicial(2) = Punto_inicial(2)
                Set blockRef = gcadModel.InsertBlock(Punto_inicial, Lola_Bastidor, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Lolashor"
                Punto_inicial(0) = Punto_inicial(0) + lldesfase * Cos(ANG + PI / 2): Punto_inicial(1) = Punto_inicial(1) + lldesfase * Sin(ANG + PI / 2): Punto_inicial(2) = Punto_inicial(2)
                i = i + 1
                Loop
            End If
        
    
    End If
    
    
Eje1.Layer = "Nonplot"
Loop
terminar:
End Sub






