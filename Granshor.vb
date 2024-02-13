Sub gs()

Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim rutags As String, rutamp As String, rutator As String, rutampacc As String, rutass As String
Dim punto1 As Variant, punto2 As Variant, PI As Variant
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double
Dim GS_Bulon120mm As String, GS_Giro As String, GS_Fusible As String, GS_750 As String, GS_1500 As String, GS_3000 As String, GS_4500 As String, GS_6000 As String, GS_Giro80 As String, GS_Bulon80 As String, zGS_Husillo As String, GS_Triangulo_fijo As String, GS_Triangulo_gato As String, GS_Poste As String, GS_transion As String
Dim lgiro80 As Double, lgirogs As Double, lfusible As Double, l750 As Double, l1500 As Double, l3000 As Double, l4500 As Double, l6000 As Double, lposte As Double, lgatomin As Double, ltriangulofijo As Double, ltrasicion As Double
Dim MP_Jack As String, MP_Husillo As String, zMP_Base As String, MP_Giro As String, MP_Fusible As String, mp_90 As String, mp_180 As String, mp_270 As String, mp_450 As String, mp_900 As String, MP_placa As String, ss_720 As String
Dim ljack As Double, l90 As Double, l180 As Double, l270 As Double, l450 As Double, l900 As Double, repite As Double, lgatominmp As Double, lplaca As Double
Dim Distancia As Double, lfija As Double, lfija1 As Double, lfija2 As Double
Dim Punto_inial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double, Punto_aux3(0 To 2) As Double
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String, dato1 As String, extremo1 As String, extremo2 As String, disposicion As String, poste As String, plantalz As String, jack As String
Dim Gcapa As Object
Dim n6000 As Integer, n4500 As Integer, n3000 As Integer, n1500 As Integer, n750 As Integer, nposte As Integer, n450 As Integer, n270 As Integer, n180 As Integer, n90 As Integer, njack As Integer, nfusible As Integer
Dim M20x160A_4 As String, M20x90_4 As String, M20x60_12 As String, M20x50_4 As String, M20x60_4 As String, M20x60A_8 As String, M20x110_4 As String, M20x60_6 As String, M20x60A_6 As String, M16x40_4 As String

Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

Ncapa = "Mega"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Granshor"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150
Ncapa = "Slims"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30

On Error GoTo terminar

rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"
rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
rutampacc = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"
rutass = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlims\"

'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgirogs = 205
lgiro80 = 112
lfusible = 187.5
l90 = 90
l180 = 180
l270 = 270
l450 = 450
l750 = 750
l900 = 900
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
lgatominmp = 435
lgatomin = 1385
ljack = 40
lposte = 375
ltriangulofijo = 1113
ltrasicion = 20

On Error GoTo terminar

kwordList = "Acodalamiento Estabilizador Pata"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo: [Acodalamiento/Estabilizador/Pata]")

kwordList = "Planta Alzado"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
disposicion = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")



poste = ""
jack = "Cero"

If dato1 = "" Or dato1 = "Acodalamiento" Then
    dato1 = "Acodalamiento"
    kwordList = "A B C"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    extremo1 = ThisDrawing.Utility.GetKeyword(vbLf & "Extremo 1(fijo): A (Granshor) B (Doble MP) C (MP + Triángulo fijo): [A/B/C]")
    kwordList = "A B C"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    extremo2 = ThisDrawing.Utility.GetKeyword(vbLf & "Extremo 2(gato): A (Grasnhor) B (Doble MP) C (Triángulo + Gato MP): [A/B/C]")
    kwordList = "No Inicial Final Ambos"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    poste = ThisDrawing.Utility.GetKeyword(vbLf & "¿Poste en algún lado?: [No/Inicial/Final/Ambos]")
    njack = 0
    If extremo2 = "B" Or extremo2 = "C" Then
        kwordList = "Cero Uno Dos"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        jack = ThisDrawing.Utility.GetKeyword(vbLf & "Número de Jack-plates: [Cero/Uno/Dos]")
            If jack = "" Or jack = "Cero" Then
            jack = "Cero"
            njack = 0
            ElseIf jack = "Uno" Then
            njack = 1
            ElseIf jack = "Dos" Then
            njack = 2
            End If
    End If
        
ElseIf dato1 = "Estabilizador" Then
    dato1 = "Estabilizador"
    kwordList = "A B C D"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    extremo1 = ThisDrawing.Utility.GetKeyword(vbLf & "Extremo 1 (Anclaje): A (MP450) B (GS750) C (GS1500) D (Sin arranque específico): [A/B/C/D]")
ElseIf dato1 = "Pata" Then
    kwordList = "A B C"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    extremo2 = ThisDrawing.Utility.GetKeyword(vbLf & "Extremo 2(gato): A (Grasnhor) B (Doble MP) C (Triángulo + Gato MP): [A/B/C]")
    njack = 0
    
    If extremo2 = "B" Or extremo2 = "C" Then
        kwordList = "Cero Uno Dos"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        jack = ThisDrawing.Utility.GetKeyword(vbLf & "Número de Jack-plates: [Cero/Uno/Dos]")
            If jack = "" Or jack = "Cero" Then
            jack = "Cero"
            njack = 0
            ElseIf jack = "Uno" Then
            njack = 1
            ElseIf jack = "Dos" Then
            njack = 2
            End If
    End If
Else
GoTo terminar
End If

If disposicion = "" Or disposicion = "Planta" Then
disposicion = "planta"
plantalz = "PLA"
ElseIf disposicion = "Alzado" Then
disposicion = "alzado"
plantalz = "ALZ"
Else
GoTo terminar
End If

If poste = "" Or poste = "No" Then
nposte = 0
ElseIf poste = "Inicial" Or poste = "Final" Then
nposte = 1
ElseIf poste = "Ambos" Then
nposte = 2
Else
GoTo terminar
End If


M20x90_4 = rutator & "4-M20X90.dwg"
M20x60_12 = rutator & "12-M20X60.dwg"
M20x60_4 = rutator & "4-M20X60.dwg"
M20x50_4 = rutator & "4-M20x50.dwg"
M20x60A_8 = rutator & "8-M20x60A.dwg"
M20x60A_6 = rutator & "6-M20x60A.dwg"
M20x110_4 = rutator & "4-M20x110.dwg"
M20x60_6 = rutator & "6-M20X60.dwg"
M16x40_4 = rutator & "4-M16X60.dwg"
M20x160A_4 = rutator & "4-M20X160A.dwg"

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

    If dato1 = "" Or dato1 = "Acodalamiento" Then

        If extremo1 = "" Or extremo1 = "A" Then
            lfija1 = lgirogs + lfusible + ltriangulofijo
        ElseIf extremo1 = "B" Then
            lfija1 = 2 * l90
        ElseIf extremo1 = "C" Then
            lfija1 = 2 * l90 + ltrasicion + ltriangulofijo
        Else
            GoTo terminar
        End If
    
        If extremo2 = "" Or extremo2 = "A" Then
        lfija2 = lgirogs + lgatomin
        ElseIf extremo2 = "B" Then
        lfija2 = l90 + lgatominmp + ljack * njack
        ElseIf extremo2 = "C" Then
        lfija2 = l90 + lgatominmp + ltrasicion + ltriangulofijo + ljack * njack
        Else
        GoTo terminar
        End If
    
        lfija = lfija1 + lfija2 + nposte * lposte
    
        If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."
        GoTo terminar
        End If

        lpuntal = Distancia - lfija
        n6000 = Fix(lpuntal / l6000)
        lpuntal = lpuntal - n6000 * l6000
        n4500 = Fix(lpuntal / l4500)
        lpuntal = lpuntal - n4500 * l4500
        n3000 = Fix(lpuntal / l3000)
        lpuntal = lpuntal - n3000 * l3000
        n1500 = Fix(lpuntal / l1500)
        lpuntal = lpuntal - n1500 * l1500
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n450 = 0
        n270 = 0
        n180 = 0
        n90 = 0
        
        
            If extremo2 = "" Or extremo2 = "A" Then
                Select Case lpuntal
                Case 0 To 375
                nfusible = 0
                Case 375 To 563
                nfusible = 1
                Case 563 To 750
                nfusible = 2
                Case Else
                MsgBox "Error en regulación del gato, consultar al programador"
                GoTo terminar
                End Select
            ElseIf extremo2 = "B" Then
                n450 = Fix(lpuntal / l450)
                lpuntal = lpuntal - n450 * l450
                n270 = Fix(lpuntal / l270)
                lpuntal = lpuntal - n270 * l270
                n180 = Fix(lpuntal / l180)
                lpuntal = lpuntal - n180 * l180
                n90 = Fix(lpuntal / l90)
                lpuntal = lpuntal - n90 * l90
            ElseIf extremo2 = "C" Then
                n450 = Fix(lpuntal / l450)
                lpuntal = lpuntal - n450 * l450
                n270 = Fix(lpuntal / l270)
                lpuntal = lpuntal - n270 * l270
                n180 = Fix(lpuntal / l180)
                lpuntal = lpuntal - n180 * l180
                n90 = Fix(lpuntal / l90)
                lpuntal = lpuntal - n90 * l90
            Else
                GoTo terminar
            End If
        
            GS_Fusible = rutags & "GS_Fusible_" & disposicion & ".dwg"
        If extremo1 = "" Or extremo1 = "A" Then
        
            GS_Bulon120mm = rutags & "GS_Bulon120mm_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(P1, GS_Bulon120mm, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            GS_Giro = rutags & "GS_Giro_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(P1, GS_Giro, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Punto_inial(0) = P1(0) + lgirogs * Cos(ANG): Punto_inial(1) = P1(1) + lgirogs * Sin(ANG): Punto_inial(2) = P1(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            GS_Triangulo_fijo = rutags & "GS_Triangulo_fijo_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Triangulo_fijo, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Punto_final(0) = Punto_inial(0) + ltriangulofijo * Cos(ANG): Punto_final(1) = Punto_inial(1) + ltriangulofijo * Sin(ANG): Punto_final(2) = Punto_inial(2)
            'Set BlockRef = gcadModel.InsertBlock(Punto_final, M20x60_12, Xs, Ys, Zs, ANG)
            'BlockRef.Layer = "Nonplot"
            'BlockRef.Erase
        ElseIf extremo1 = "B" Then
        
            If disposicion = "alzado" Then
                Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
                Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial(2)
                Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial(2)
                MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Giro, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Giro, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Punto_inial(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
                Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial(2)
                Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial(2)
                MP_Fusible = rutamp & "Mshor90" & plantalz & "fusible.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Fusible, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Fusible, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x50_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x50_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            ElseIf disposicion = "planta" Then
                Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
                MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Giro, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Giro, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Punto_inial(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
                MP_Fusible = rutamp & "Mshor90" & plantalz & "fusible.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Fusible, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Fusible, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x50_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x50_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = P1(2)
            End If
            
        ElseIf extremo1 = "C" Then
            Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
            MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Giro, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"
            Punto_inial(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
            MP_Fusible = rutamp & "Mshor90" & plantalz & "fusible.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Fusible, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Fusible, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x50_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x50_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Punto_inial(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
            GS_transion = rutags & "GS_TransiciónMG_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_transion, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_transion, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60A_8, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60A_8, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_inial(0) = Punto_inial(0) + ltrasicion * Cos(ANG): Punto_inial(1) = Punto_inial(1) + ltrasicion * Sin(ANG): Punto_inial(2) = Punto_inial(2)
            GS_Triangulo_fijo = rutags & "GS_Triangulo_fijo_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Triangulo_fijo, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Punto_final(0) = Punto_inial(0) + ltriangulofijo * Cos(ANG): Punto_final(1) = Punto_inial(1) + ltriangulofijo * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Else
            GoTo terminar
        End If


        If extremo2 = "" Or extremo2 = "A" Then
        
            GS_Bulon120mm = rutags & "GS_Bulon120mm_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(P2, GS_Bulon120mm, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            GS_Giro = rutags & "GS_Giro_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(P2, GS_Giro, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
            Punto_final2(0) = P2(0) + lgirogs * Cos(ANG + PI): Punto_final2(1) = P2(1) + lgirogs * Sin(ANG + PI): Punto_final2(2) = P2(2)
            If nfusible > 0 Then
            i = 0
            Do While i < nfusible
                Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
                Set blockRef = gcadModel.InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Granshor"
                Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final2(0) = Punto_inial2(0) + lfusible * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + lfusible * Sin(ANG + PI): Punto_final2(2) = Punto_inial2(2)
                i = i + 1
            Loop
            End If
            Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            zGS_Husillo = rutags & "zGS_Husillo_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, zGS_Husillo, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"

        ElseIf extremo2 = "B" Then
        
            If disposicion = "alzado" Then
                Punto_inial2(0) = P2(0): Punto_inial2(1) = P2(1): Punto_inial2(2) = P2(2)
                Punto_aux1(0) = Punto_inial2(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial2(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                Punto_aux2(0) = Punto_inial2(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial2(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Giro, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Giro, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Punto_final2(0) = Punto_inial2(0) + l90 * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + l90 * Sin(ANG + PI): Punto_final2(2) = Punto_inial2(2)
                Punto_aux1(0) = Punto_final2(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final2(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final2(2)
                Punto_aux2(0) = Punto_final2(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final2(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final2(2)
                If njack = 2 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final2(0) = Punto_final2(0) - ljack * Cos(ANG): Punto_final2(1) = Punto_final2(1) - ljack * Sin(ANG): Punto_final2(2) = Punto_final2(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                Punto_aux1(0) = Punto_final2(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final2(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final2(2)
                Punto_aux2(0) = Punto_final2(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final2(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final2(2)
                zMP_Base = rutampacc & "zMGBaseGato_naranja.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                
            ElseIf disposicion = "planta" Then
                Punto_inial2(0) = P2(0): Punto_inial2(1) = P2(1): Punto_inial2(2) = P2(2)
                MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial2, MP_Giro, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial2, MP_Giro, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Punto_final2(0) = Punto_inial2(0) + l90 * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + l90 * Sin(ANG + PI): Punto_final2(2) = Punto_inial2(2)
                If njack = 2 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final2(0) = Punto_final2(0) - ljack * Cos(ANG): Punto_final2(1) = Punto_final2(1) - ljack * Sin(ANG): Punto_final2(2) = Punto_final2(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                zMP_Base = rutampacc & "zMGBaseGato_naranja.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_final2, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_final2, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
            End If
            
        ElseIf extremo2 = "C" Then
            Punto_inial2(0) = P2(0): Punto_inial2(1) = P2(1): Punto_inial2(2) = P2(2)
            MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, MP_Giro, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Mega"
            Punto_final2(0) = Punto_inial2(0) + l90 * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + l90 * Sin(ANG + PI): Punto_final2(2) = Punto_inial2(2)
                If njack = 2 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final2(0) = Punto_final2(0) - ljack * Cos(ANG): Punto_final2(1) = Punto_final2(1) - ljack * Sin(ANG): Punto_final2(2) = Punto_final2(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                zMP_Base = rutampacc & "zMGBaseGato_naranja.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_final2, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
        Else
            GoTo terminar
        End If

        If n750 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_750 = rutags & "GS_750_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_750, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If n4500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_4500 = rutags & "GS_4500_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If poste = "Inicial" Or poste = "Ambos" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_Poste = rutags & "GS_Poste_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Poste, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + lposte * Cos(ANG): Punto_final(1) = Punto_inial(1) + lposte * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If n6000 > 0 Then
        i = 0
        GS_6000 = rutags & "GS_6000_" & disposicion & ".dwg"
        Do While i < n6000
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_6000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            i = i + 1
        Loop
        End If

        If poste = "Final" Or poste = "Ambos" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_Poste = rutags & "GS_Poste_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Poste, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + lposte * Cos(ANG): Punto_final(1) = Punto_inial(1) + lposte * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If n3000 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_3000 = rutags & "GS_3000_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If n1500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_1500 = rutags & "GS_1500_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_1500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If extremo2 = "" Or extremo2 = "A" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_Triangulo_gato = rutags & "GS_Triangulo_gato_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Triangulo_gato, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + 1310 * Cos(ANG): Punto_final(1) = Punto_inial(1) + 1310 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    
        ElseIf extremo2 = "B" Then
        
            If disposicion = "alzado" Then
                If n450 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    mp_450 = rutamp & "Mshor450" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If

                If n270 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    mp_270 = rutamp & "Mshor270" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_270, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_270, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If

                If n180 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    mp_180 = rutamp & "Mshor180" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_180, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_180, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
                
                If n90 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    mp_90 = rutamp & "Mshor90" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_90, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_90, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
 
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    
                If njack = 2 Or njack = 1 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Jack, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Jack, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_final(0) + ljack * Cos(ANG): Punto_final(1) = Punto_final(1) + ljack * Sin(ANG): Punto_final(2) = Punto_final(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                Punto_aux1(0) = Punto_final(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final(2)
                Punto_aux2(0) = Punto_final(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final(2)
                zMP_Base = rutampacc & "zMGBaseGato_azul.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, zMP_Base, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, zMP_Base, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Punto_final(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_final(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_final(2) = (Punto_final(2) + Punto_final2(2)) / 2
                Punto_aux1(0) = Punto_final(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final(2)
                Punto_aux2(0) = Punto_final(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final(2)
                MP_Husillo = rutampacc & "MGHusilloGato.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Husillo, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Husillo, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"


            ElseIf disposicion = "planta" Then
                If n450 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    mp_450 = rutamp & "Mshor450" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If

                If n270 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    mp_270 = rutamp & "Mshor270" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If

                If n180 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    mp_180 = rutamp & "Mshor180" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
                
                If n90 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    mp_90 = rutamp & "Mshor90" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
 
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    
                If njack = 2 Or njack = 1 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Jack, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Jack, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_final(0) + ljack * Cos(ANG): Punto_final(1) = Punto_final(1) + ljack * Sin(ANG): Punto_final(2) = Punto_final(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                zMP_Base = rutampacc & "zMGBaseGato_azul.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_final, zMP_Base, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_final, zMP_Base, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Punto_final(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_final(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_final(2) = (Punto_final(2) + Punto_final2(2)) / 2
                MP_Husillo = rutampacc & "MGHusilloGato.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_final, MP_Husillo, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_final, MP_Husillo, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
            End If
    
        ElseIf extremo2 = "C" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + ltriangulofijo * Cos(ANG): Punto_final(1) = Punto_inial(1) + ltriangulofijo * Sin(ANG): Punto_final(2) = Punto_inial(2)
            GS_Triangulo_fijo = rutags & "GS_Triangulo_fijo_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, GS_Triangulo_fijo, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
        
            GS_transion = rutags & "GS_TransiciónMG_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, GS_transion, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x60A_8, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_final(0) + ltrasicion * Cos(ANG): Punto_final(1) = Punto_final(1) + ltrasicion * Sin(ANG): Punto_final(2) = Punto_final(2)
        
        
            If n450 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                mp_450 = rutamp & "Mshor450" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            End If
    
            If n270 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                mp_270 = rutamp & "Mshor270" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            End If

            If n180 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                mp_180 = rutamp & "Mshor180" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            End If
                
            If n90 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                mp_90 = rutamp & "Mshor90" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            End If
 
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    
            If njack = 2 Or njack = 1 Then
                MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Jack, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                If n90 = 0 And n180 = 0 And n270 = 0 And n450 = 0 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x160A_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Else
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                End If
                Punto_final(0) = Punto_final(0) + ljack * Cos(ANG): Punto_final(1) = Punto_final(1) + ljack * Sin(ANG): Punto_final(2) = Punto_final(2)
            ElseIf njack = 0 Or njack = 1 Then
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            End If
            zMP_Base = rutampacc & "zMGBaseGato_azul.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, zMP_Base, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"
         
            Punto_final(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_final(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_final(2) = (Punto_final(2) + Punto_final2(2)) / 2
            MP_Husillo = rutampacc & "MGHusilloGato.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, MP_Husillo, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"
        End If


    ElseIf dato1 = "Estabilizador" Then
        lplaca = 25
        If extremo1 = "" Or extremo1 = "A" Then
            lfija1 = lplaca + l450
        ElseIf extremo1 = "B" Then
            lfija1 = lplaca + l750
        ElseIf extremo1 = "C" Then
            lfija1 = lplaca + l1500
        ElseIf extremo1 = "D" Then
            lfija1 = lplaca
        Else
            GoTo terminar
        End If

        lfija = lfija1
    
        If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."
        GoTo terminar
        End If

        Distancia = Distancia + l750

        lpuntal = Distancia - lfija
        n6000 = Fix(lpuntal / l6000)
        lpuntal = lpuntal - n6000 * l6000
        n4500 = Fix(lpuntal / l4500)
        lpuntal = lpuntal - n4500 * l4500
        n3000 = Fix(lpuntal / l3000)
        lpuntal = lpuntal - n3000 * l3000
        n1500 = Fix(lpuntal / l1500)
        lpuntal = lpuntal - n1500 * l1500
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750


            If disposicion = "alzado" Then
                Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
                Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial(2)
                Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial(2)
                MP_placa = rutags & "GSPlacagato.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_placa, Xs, Ys, Zs, ANG - PI / 2)
                blockRef.Layer = "Granshor"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_placa, Xs, Ys, Zs, ANG - PI / 2)
                blockRef.Layer = "Granshor"
                Punto_inial(0) = Punto_inial(0) + lplaca * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lplaca * Sin(ANG): Punto_inial(2) = Punto_inial(2)
                Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial(2)
                Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial(2)
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60A_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60A_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)

                
                If extremo1 = "" Or extremo1 = "A" Then
                    Punto_aux3(0) = Punto_final(0) + 225 * Cos(ANG): Punto_aux3(1) = Punto_final(1) + 225 * Sin(ANG): Punto_aux3(2) = Punto_final(2)
                    Punto_aux3(0) = Punto_aux3(0) + 360 * Cos(ANG + PI / 2): Punto_aux3(1) = Punto_aux3(1) + 360 * Sin(ANG + PI / 2): Punto_aux3(2) = Punto_aux3(2)
                    ss_720 = rutass & "SS0720.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux3, ss_720, Xs, Ys, Zs, ANG - PI / 2)
                    blockRef.Layer = "Slims"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux3, M16x40_4, Xs, Ys, Zs, ANG - PI / 2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_aux3(0) = Punto_final(0) + 225 * Cos(ANG): Punto_aux3(1) = Punto_final(1) + 225 * Sin(ANG): Punto_aux3(2) = Punto_final(2)
                    Punto_aux3(0) = Punto_aux3(0) - 360 * Cos(ANG + PI / 2): Punto_aux3(1) = Punto_aux3(1) - 360 * Sin(ANG + PI / 2): Punto_aux3(2) = Punto_aux3(2)
                    Set blockRef = gcadModel.InsertBlock(Punto_aux3, M16x40_4, Xs, Ys, Zs, ANG - PI / 2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    mp_450 = rutamp & "Mshor450" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
                                
            ElseIf disposicion = "planta" Then
                Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
                MP_placa = rutags & "GSPlacagato.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_placa, Xs, Ys, Zs, ANG - PI / 2)
                blockRef.Layer = "Granshor"
                Punto_inial(0) = Punto_inial(0) + lplaca * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lplaca * Sin(ANG): Punto_inial(2) = Punto_inial(2)
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60A_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                
                If extremo1 = "" Or extremo1 = "A" Then
                    mp_450 = rutamp & "Mshor450" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
            End If

            If extremo1 = "B" Then
                GS_750 = rutags & "GS_750_" & disposicion & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_750, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Granshor"
                Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            ElseIf extremo1 = "C" Then
                GS_1500 = rutags & "GS_1500_" & disposicion & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_1500, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Granshor"
                Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            ElseIf extremo1 = "D" Then
                Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)
            End If
            
        If n6000 > 0 Then
        i = 0
        GS_6000 = rutags & "GS_6000_" & disposicion & ".dwg"
        Do While i < n6000
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_6000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            i = i + 1
        Loop
        End If

        If n4500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_4500 = rutags & "GS_4500_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If n3000 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_3000 = rutags & "GS_3000_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If n1500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_1500 = rutags & "GS_1500_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_1500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If n750 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_750 = rutags & "GS_750_2_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_750, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

    ElseIf dato1 = "Pata" Then

        If extremo2 = "" Or extremo2 = "A" Then
        lfija2 = lgirogs + lgatomin + lgiro80
        ElseIf extremo2 = "B" Then
        lfija2 = l90 + lgatominmp + ljack * njack + lgiro80
        ElseIf extremo2 = "C" Then
        lfija2 = l90 + lgatominmp + ltrasicion + ltriangulofijo + ljack * njack + lgiro80
        Else
        GoTo terminar
        End If
     
        If Distancia < lfija2 Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija2 & "."
        GoTo terminar
        End If

        lpuntal = Distancia - lfija2
        n6000 = Fix(lpuntal / l6000)
        lpuntal = lpuntal - n6000 * l6000
        n4500 = Fix(lpuntal / l4500)
        lpuntal = lpuntal - n4500 * l4500
        n3000 = Fix(lpuntal / l3000)
        lpuntal = lpuntal - n3000 * l3000
        n1500 = Fix(lpuntal / l1500)
        lpuntal = lpuntal - n1500 * l1500
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n450 = 0
        n270 = 0
        n180 = 0
        n90 = 0
   
        If extremo2 = "" Or extremo2 = "A" Then
            Select Case lpuntal
            Case 0 To 375
            nfusible = 0
            Case 375 To 563
            nfusible = 1
            Case 563 To 750
            nfusible = 2
            Case Else
            MsgBox "Error en regulación del gato, consultar al programador"
            GoTo terminar
            End Select
        ElseIf extremo2 = "B" Then
            n450 = Fix(lpuntal / l450)
            lpuntal = lpuntal - n450 * l450
            n270 = Fix(lpuntal / l270)
            lpuntal = lpuntal - n270 * l270
            n180 = Fix(lpuntal / l180)
            lpuntal = lpuntal - n180 * l180
            n90 = Fix(lpuntal / l90)
            lpuntal = lpuntal - n90 * l90
        ElseIf extremo2 = "C" Then
            n450 = Fix(lpuntal / l450)
            lpuntal = lpuntal - n450 * l450
            n270 = Fix(lpuntal / l270)
            lpuntal = lpuntal - n270 * l270
            n180 = Fix(lpuntal / l180)
            lpuntal = lpuntal - n180 * l180
            n90 = Fix(lpuntal / l90)
            lpuntal = lpuntal - n90 * l90
        Else
            GoTo terminar
        End If

        If disposicion = "alzado" Then
            
            Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
            Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial(2)
            Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial(2)
            GS_Bulon80 = rutags & "GS_Bulon80_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_aux1, GS_Bulon80, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_aux2, GS_Bulon80, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            GS_Giro80 = rutags & "GS_Giro80" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_aux1, GS_Giro80, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_aux2, GS_Giro80, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Punto_inial(0) = Punto_inial(0) + lgiro80 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lgiro80 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
            Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial(2)
            Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)

        ElseIf disposicion = "planta" Then

            Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
            GS_Bulon80 = rutags & "GS_Bulon80_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Bulon80, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Bulon80, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            GS_Giro80 = rutags & "GS_Giro80" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Giro80, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Giro80, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Punto_final(0) = Punto_inial(0) + lgiro80 * Cos(ANG): Punto_final(1) = Punto_inial(1) + lgiro80 * Sin(ANG): Punto_final(2) = Punto_inial(2)

        Else
            GoTo terminar
        End If
        
        If n750 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_750 = rutags & "GS_750_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_750, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If
        
        If n6000 > 0 Then
        i = 0
        GS_6000 = rutags & "GS_6000_" & disposicion & ".dwg"
        Do While i < n6000
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_6000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            i = i + 1
        Loop
        End If

        If n4500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_4500 = rutags & "GS_4500_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If
        
        If n3000 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_3000 = rutags & "GS_3000_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If n1500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_1500 = rutags & "GS_1500_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_1500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If

        If extremo2 = "" Or extremo2 = "A" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            GS_Triangulo_gato = rutags & "GS_Triangulo_gato_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Triangulo_gato, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + 1310 * Cos(ANG): Punto_final(1) = Punto_inial(1) + 1310 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            GS_Bulon120mm = rutags & "GS_Bulon120mm_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(P2, GS_Bulon120mm, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            GS_Giro = rutags & "GS_Giro_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(P2, GS_Giro, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
            Punto_final2(0) = P2(0) + lgirogs * Cos(ANG + PI): Punto_final2(1) = P2(1) + lgirogs * Sin(ANG + PI): Punto_final2(2) = P2(2)
            If nfusible > 0 Then
            i = 0
            Do While i < nfusible
                Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
                GS_Fusible = rutags & "GS_Fusible_" & disposicion & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Granshor"
                Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final2(0) = Punto_inial2(0) + lfusible * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + lfusible * Sin(ANG + PI): Punto_final2(2) = Punto_inial2(2)
                i = i + 1
            Loop
            End If
            Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            zGS_Husillo = rutags & "zGS_Husillo_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, zGS_Husillo, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"

        ElseIf extremo2 = "B" Then
  
            If disposicion = "alzado" Then
                If n450 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    mp_450 = rutamp & "Mshor450" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If

                If n270 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    mp_270 = rutamp & "Mshor270" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_270, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_270, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If

                If n180 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    mp_180 = rutamp & "Mshor180" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_180, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_180, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
                
                If n90 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    mp_90 = rutamp & "Mshor90" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, mp_90, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, mp_90, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
 
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    Punto_aux1(0) = Punto_inial(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                    Punto_aux2(0) = Punto_inial(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                    
                If njack = 2 Or njack = 1 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Jack, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Jack, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_final(0) + ljack * Cos(ANG): Punto_final(1) = Punto_final(1) + ljack * Sin(ANG): Punto_final(2) = Punto_final(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                Punto_aux1(0) = Punto_final(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final(2)
                Punto_aux2(0) = Punto_final(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final(2)
                zMP_Base = rutampacc & "zMGBaseGato_azul.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, zMP_Base, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, zMP_Base, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                
            
            ElseIf disposicion = "planta" Then
                If n450 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    mp_450 = rutamp & "Mshor450" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If

                If n270 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    mp_270 = rutamp & "Mshor270" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If

                If n180 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    mp_180 = rutamp & "Mshor180" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
                
                If n90 > 0 Then
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    mp_90 = rutamp & "Mshor90" & plantalz & ".dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                End If
 
                    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    
                If njack = 2 Or njack = 1 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Jack, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Jack, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final(0) = Punto_final(0) + ljack * Cos(ANG): Punto_final(1) = Punto_final(1) + ljack * Sin(ANG): Punto_final(2) = Punto_final(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                zMP_Base = rutampacc & "zMGBaseGato_azul.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_final, zMP_Base, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_final, zMP_Base, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
            End If
  
            If disposicion = "alzado" Then
                Punto_inial2(0) = P2(0): Punto_inial2(1) = P2(1): Punto_inial2(2) = P2(2)
                Punto_aux1(0) = Punto_inial2(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_inial2(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_inial2(2)
                Punto_aux2(0) = Punto_inial2(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_inial2(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_inial2(2)
                MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Giro, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Giro, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Punto_final2(0) = Punto_inial2(0) + l90 * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + l90 * Sin(ANG + PI): Punto_final2(2) = Punto_inial2(2)
                Punto_aux1(0) = Punto_final2(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final2(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final2(2)
                Punto_aux2(0) = Punto_final2(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final2(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final2(2)
                If njack = 2 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final2(0) = Punto_final2(0) - ljack * Cos(ANG): Punto_final2(1) = Punto_final2(1) - ljack * Sin(ANG): Punto_final2(2) = Punto_final2(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_aux1, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_aux2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                Punto_aux1(0) = Punto_final2(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final2(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final2(2)
                Punto_aux2(0) = Punto_final2(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final2(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final2(2)
                zMP_Base = rutampacc & "zMGBaseGato_naranja.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                
                Punto_final(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_final(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_final(2) = (Punto_final(2) + Punto_final2(2)) / 2
                Punto_aux1(0) = Punto_final(0) + 503 * Cos(ANG + PI / 2): Punto_aux1(1) = Punto_final(1) + 503 * Sin(ANG + PI / 2): Punto_aux1(2) = Punto_final(2)
                Punto_aux2(0) = Punto_final(0) + 503 * Cos(ANG - PI / 2): Punto_aux2(1) = Punto_final(1) + 503 * Sin(ANG - PI / 2): Punto_aux2(2) = Punto_final(2)
                MP_Husillo = rutampacc & "MGHusilloGato.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_aux1, MP_Husillo, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_aux2, MP_Husillo, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                
            ElseIf disposicion = "planta" Then
                Punto_inial2(0) = P2(0): Punto_inial2(1) = P2(1): Punto_inial2(2) = P2(2)
                MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial2, MP_Giro, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial2, MP_Giro, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Punto_final2(0) = Punto_inial2(0) + l90 * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + l90 * Sin(ANG + PI): Punto_final2(2) = Punto_inial2(2)
                If njack = 2 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final2(0) = Punto_final2(0) - ljack * Cos(ANG): Punto_final2(1) = Punto_final2(1) - ljack * Sin(ANG): Punto_final2(2) = Punto_final2(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                zMP_Base = rutampacc & "zMGBaseGato_naranja.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_final2, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_final2, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
                Punto_final(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_final(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_final(2) = (Punto_final(2) + Punto_final2(2)) / 2
                MP_Husillo = rutampacc & "MGHusilloGato.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_final, MP_Husillo, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_final, MP_Husillo, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
            End If
            
        
        ElseIf extremo2 = "C" Then
 
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_12, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_inial(0) + ltriangulofijo * Cos(ANG): Punto_final(1) = Punto_inial(1) + ltriangulofijo * Sin(ANG): Punto_final(2) = Punto_inial(2)
            GS_Triangulo_fijo = rutags & "GS_Triangulo_fijo_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, GS_Triangulo_fijo, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
        
            GS_transion = rutags & "GS_TransiciónMG_" & disposicion & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, GS_transion, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x60A_8, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Punto_final(0) = Punto_final(0) + ltrasicion * Cos(ANG): Punto_final(1) = Punto_final(1) + ltrasicion * Sin(ANG): Punto_final(2) = Punto_final(2)
        
        
            If n450 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                mp_450 = rutamp & "Mshor450" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            End If
    
            If n270 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                mp_270 = rutamp & "Mshor270" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            End If

            If n180 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                mp_180 = rutamp & "Mshor180" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            End If
                
            If n90 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                mp_90 = rutamp & "Mshor90" & plantalz & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_6, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            End If
 
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                    
            If njack = 2 Or njack = 1 Then
                MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Jack, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Mega"
                If n90 = 0 And n180 = 0 And n270 = 0 And n450 = 0 Then
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x160A_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Else
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                Punto_final(0) = Punto_final(0) + ljack * Cos(ANG): Punto_final(1) = Punto_final(1) + ljack * Sin(ANG): Punto_final(2) = Punto_final(2)
            ElseIf njack = 0 Or njack = 1 Then
                Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            End If
            zMP_Base = rutampacc & "zMGBaseGato_azul.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, zMP_Base, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"
         
            Punto_inial2(0) = P2(0): Punto_inial2(1) = P2(1): Punto_inial2(2) = P2(2)
            MP_Giro = rutampacc & "MG_AnguloGiro" & plantalz & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, MP_Giro, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Mega"
            Punto_final2(0) = Punto_inial2(0) + l90 * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + l90 * Sin(ANG + PI): Punto_final2(2) = Punto_inial2(2)
                If njack = 2 Then
                    MP_Jack = rutamp & "MshorJACKPLATE.dwg"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, MP_Jack, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Mega"
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x110_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                    Punto_final2(0) = Punto_final2(0) - ljack * Cos(ANG): Punto_final2(1) = Punto_final2(1) - ljack * Sin(ANG): Punto_final2(2) = Punto_final2(2)
                ElseIf njack = 0 Or njack = 1 Then
                    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x60_4, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                blockRef.Explode
                blockRef.Delete
                End If
                zMP_Base = rutampacc & "zMGBaseGato_naranja.dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_final2, zMP_Base, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Mega"
        
            Punto_final(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_final(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_final(2) = (Punto_final(2) + Punto_final2(2)) / 2
            MP_Husillo = rutampacc & "MGHusilloGato.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, MP_Husillo, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Mega"
            
        Else
            GoTo terminar
        End If
 
    Else
        GoTo terminar
    End If

'If njack > 0 And n90 = 0 And n180 = 0 And n270 = 0 And n450 = 0 And extremo2 = "C" Then
'MsgBox "ATENCIÓN: LA UNICÓN JACK PLATE CON CHAPA DE TRANSICIÓN REQUIERE M20x90 CAB. AVELLANADA, NO TENEMOS ESE TORNILLO"
'GoTo terminar
'End If

Eje1.Layer = "Nonplot"
Loop
terminar:
End Sub











