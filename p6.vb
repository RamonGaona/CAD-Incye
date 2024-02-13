

Sub P6()

Dim rutaps As String, rutap6 As String, rutator As String, rutags As String, rutapl As String
Dim gcadDoc As Object
Dim gcadUtil As Object
Dim gcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim M20x90_16 As String, M20x150_4 As String, M20x160_4 As String, Revisartornilleria As String, M30x100_16 As String, M20x90_4 As String
Dim GS_Bulon120mm As String
Dim GS_Giro As String
Dim GS_Fusible As String
Dim PS_280 As String, PS_560 As String
Dim PS_750 As String, PS_nudo As String
Dim PS_1500 As String
Dim PS_3000 As String, P6_3000 As String
Dim PS_4500 As String, P6_4500 As String
Dim PS_6000 As String
Dim P6_cono As String
Dim PS_Husillo As String
Dim PS_Placa50mm As String, PS_Placa35mm As String
Dim zPS_Gato_Cono As String
Dim zPS_Gato_Tope As String
Dim PS_Gato As String
Dim lgiro As Double
Dim lfusible As Double
Dim l280 As Double, l560 As Double
Dim l750 As Double
Dim l1500 As Double
Dim l3000 As Double
Dim l4500 As Double
Dim l6000 As Double
Dim l50 As Double, l35 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double, lfija2 As Double
Dim lpuntal As Double
Dim lgatomin As Double
Dim lmacho As Double, lcajon As Double
Dim lcono As Double
Dim n6000 As Integer
Dim n4500 As Integer, n4500p6 As Integer
Dim n3000 As Integer, n3000p6 As Integer
Dim n1500 As Integer
Dim n750 As Integer
Dim n280 As Integer, n560 As Integer
Dim nnudo As Integer
Dim nfusible As Integer
Dim n50 As Integer, n35 As Integer, ncono As Integer
Dim nmacho As Integer, ncajon As Integer, ngiro As Integer, n As Integer, lchapones As Integer
Dim blockRef As Object
Dim repite As Double
Dim Punto_inial(0 To 2) As Double
Dim Punto_final(0 To 2) As Double
Dim Punto_inial2(0 To 2) As Double
Dim Punto_final2(0 To 2) As Double
Dim PI As Variant
Dim Eje1 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim ANG As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object

Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

Ncapa = "Pipeshor4S"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor6"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5

rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"
rutap6 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_6\"
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"

'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgiro = 205
lcono = 375
lfusible = 187.5
l280 = 280
l560 = 560
l750 = 750
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
l50 = 50
l35 = 35
l_tope = 325
l_conogato = 170
lgatomin = 620
lmacho = 290
lcajon = 15

On Error GoTo terminar

kwordList = "No Inicial Final Ambos"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Nudo en algún extremo?: [No/Inicial/Final/Ambos]")

kwordList = "Planta Alzado"
dato2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato2 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")

kwordList = "Dos Cero Uno Tres"
dato3 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato3 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce número de chapas de transición de 50mm: [Dos/Cero/Uno/Tres]")

kwordList = "Cero Uno Dos Tres"
dato5 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato5 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce número de chapas de transición de 35mm: [Cero/Uno/Dos/Tres]")

'kwordList = "A B C D"
dato4 = ""
'ThisDrawing.Utility.InitializeUserInput 0, kwordList
'dato4 = ThisDrawing.Utility.GetKeyword(vbLf & "A (giro y gato) B (machon y machón) C (giro y machón) D (giro y cajón + machón): [A/B/C/D]")

If dato1 = "" Or dato1 = "No" Then
nnudo = 0
ElseIf dato1 = "Inicial" Or dato1 = "Final" Then
nnudo = 1
ElseIf dato1 = "Ambos" Then
nnudo = 2
Else
GoTo terminar
End If

If dato2 = "" Or dato2 = "Planta" Then
dato2 = "planta"
ElseIf dato2 = "Alzado" Then
dato2 = "alzado"
Else
GoTo terminar
End If

If dato3 = "" Or dato3 = "Dos" Then
n50 = 2
ElseIf dato3 = "Uno" Then
n50 = 1
ElseIf dato3 = "Tres" Then
n50 = 3
ElseIf dato3 = "Cero" Then
n50 = 0
Else
GoTo terminar
End If

If dato5 = "" Or dato5 = "Cero" Then
n35 = 0
ElseIf dato5 = "Uno" Then
n35 = 1
ElseIf dato5 = "Dos" Then
n35 = 2
ElseIf dato5 = "Tres" Then
n35 = 3
Else
GoTo terminar
End If

lchapones = (n35 * l35) + (n50 * l50)

M20x90_16 = rutator & "16-M20X90.dwg"
M30x100_16 = rutator & "16-M30X100.dwg"
M20x150_4 = rutator & "4-M20X150.dwg"
M20x160_4 = rutator & "4-M20X160.dwg"
M20x90_4 = rutator & "4-M20X90.dwg"
Revisartornilleria = rutator & "4MetricasXrevisar.dwg"

If dato4 = "" Or dato4 = "A" Then
        
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

        lfija2 = InputBox("Introduce longitud en mm de tubo PS6: 3000/4500/6000...", "Longitud PS6", 4500)
        If lfija2 Mod 1500 = 0 Then
        n = (lfija2 / 1500) + 0
            If n = 1 Then n = 2
        lfija2 = n * 1500
        ElseIf lfija2 Mod 1500 <> 0 Then
        n = (lfija2 / 1500) + 1
            If n = 1 Then n = 2
        lfija2 = n * 1500
        End If
        
        If n = 0 Then ncono = 0 Else ncono = 2
        
        lfija = (2 * lgiro) + lfusible + lchapones + lgatomin + (ncono * lcono) + nnudo * l750 + lfija2
        
        If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."
        GoTo terminar
        End If
        
        If n Mod 3 = 0 Then
        n3000p6 = 0
        n4500p6 = n / 3
        ElseIf (n - 2) Mod 3 = 0 Then
        n3000p6 = 1
        n4500p6 = (n - 2) / 3
        ElseIf (n - 4) Mod 3 = 0 Then
        n3000p6 = 2
        n4500p6 = (n - 4) / 3
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

        Select Case lpuntal
        Case 0 To 230
        nfusible = 1
        n280 = 0
        n560 = 0
        Case 230 To 280
        nfusible = 2
        n280 = 0
        n560 = 0
        Case 280 To 510
        nfusible = 1
        n280 = 1
        n560 = 0
        Case 510 To 560
        nfusible = 2
        n280 = 1
        n560 = 0
        Case 560 To 750
        nfusible = 1
        n280 = 0
        n560 = 1
        Case Else
        MsgBox "Longitud no controlada " & lpuntal & "mm, fuera de rango, revisar código"
        GoTo terminar
        End Select

        'Introducir el bulón de 120 mm en los extremos siempre, ángulo de giro, fusible fijo:
        GS_Bulon120mm = rutags & "GS_Bulon120mm_" & dato2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(P1, GS_Bulon120mm, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Granshor"
        Set blockRef = gcadModel.InsertBlock(P2, GS_Bulon120mm, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Granshor"
        GS_Giro = rutags & "GS_Giro_" & dato2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(P1, GS_Giro, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Granshor"
        Set blockRef = gcadModel.InsertBlock(P2, GS_Giro, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Granshor"

        Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
        GS_Fusible = rutags & "GS_Fusible_" & dato2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Granshor"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Punto_final(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_final(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_final(2) = Punto_inial(2)


        If lchapones < 119 And lchapones > 83 Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x160_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Else
        Set blockRef = gcadModel.InsertBlock(Punto_final, Revisartornilleria, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        End If

        If n35 > 0 Then
            i = 0
            Do While i < n35
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_Placa35mm = rutaps & "PS_Placa35mm_" & dato2 & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa35mm, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l35 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l35 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                i = i + 1
            Loop
        End If
        If n50 > 0 Then
            i = 0
            Do While i < n50
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato2 & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                i = i + 1
            Loop
        End If

        If lchapones < 109 And lchapones > 73 Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x150_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Else
        Set blockRef = gcadModel.InsertBlock(Punto_final, Revisartornilleria, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        End If

        If n280 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_280 = rutapl & "PL_280_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_280, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4L"
            Punto_final(0) = Punto_inial(0) + l280 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l280 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        If n560 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_560 = rutaps & "PS_560.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If
 
        If n1500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_1500 = rutaps & "PS_1500_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_1500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        If n4500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_4500 = rutaps & "PS_4500_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If
        
        If dato1 = "Ambos" Or dato1 = "Inicial" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            PS_nudo = rutaps & "PS_nudo_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_final, PS_nudo, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Pipeshor4S"
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If
        
        If n > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            P6_cono = rutap6 & "P6_cono.dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, P6_cono, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor6"
            Punto_final(0) = Punto_inial(0) + lcono * Cos(ANG): Punto_final(1) = Punto_inial(1) + lcono * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M30x100_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        If n3000p6 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            P6_3000 = rutap6 & "P6_3000_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, P6_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor6"
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M30x100_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        If n4500p6 > 0 Then
            i = 0
            Do While i < n4500p6
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            P6_4500 = rutap6 & "P6_4500_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, P6_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor6"
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M30x100_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            i = i + 1
            Loop
        End If

        If n3000p6 = 2 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            P6_3000 = rutap6 & "P6_3000_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, P6_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor6"
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M30x100_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        If n > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Punto_final(0) = Punto_inial(0) + lcono * Cos(ANG): Punto_final(1) = Punto_inial(1) + lcono * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, P6_cono, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Pipeshor6"
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        If dato1 = "Ambos" Or dato1 = "Final" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_nudo = rutaps & "PS_nudo_" & dato2 & ".dwg"
            'Set BlockRef = gcadModel.InsertBlock(Punto_final, PS_nudo, Xs, Ys, Zs, ANG)
            Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_nudo, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        If n6000 > 0 Then
            i = 0
            Do While i < n6000
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_6000 = rutaps & "PS_6000_" & dato2 & ".dwg"
                Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_6000, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                i = i + 1
            Loop
        End If

        If n3000 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_3000 = rutaps & "PS_3000_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        If n750 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_750 = rutaps & "PS_750_" & dato2 & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_750, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If

        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        zPS_Gato_Cono = rutaps & "zPS_Gato_Cono_" & dato2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, zPS_Gato_Cono, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Pipeshor4S"
        Punto_final(0) = Punto_inial(0) + l_conogato * Cos(ANG): Punto_final(1) = Punto_inial(1) + l_conogato * Sin(ANG): Punto_final(2) = Punto_inial(2)

        Punto_inial2(0) = P2(0) - lgiro * Cos(ANG): Punto_inial2(1) = P2(1) - lgiro * Sin(ANG): Punto_inial2(2) = P2(2)
        Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
        If nfusible = 2 Then
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
        End If
        Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
        zPS_Gato_Tope = rutaps & "zPS_Gato_Tope_" & dato2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial2, zPS_Gato_Tope, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Pipeshor4S"
        Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Punto_final2(0) = Punto_inial2(0) - l_tope * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - l_tope * Sin(ANG): Punto_final(2) = Punto_inial2(2)


        Punto_inial(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_inial(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_inial(2) = (Punto_final(2) + Punto_final2(2)) / 2

        PS_Gato = rutaps & "PS_Gato_" & dato2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Gato, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Pipeshor4S"
    
    Eje1.Layer = "Nonplot"
    Loop


ElseIf dato4 = "B" Then






ElseIf dato4 = "C" Then





ElseIf dato4 = "D" Then






Else
GoTo terminar
End If










terminar:
End Sub
