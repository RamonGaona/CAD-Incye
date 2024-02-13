Sub pm()
Dim ruta1 As String
Dim ruta2 As String
Dim ruta3 As String
Dim ruta4 As String
Dim rutapl1 As String
Dim rutapl2 As String
Dim gcadDoc As Object
Dim gcadUtil As Object
Dim gcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim M20x50 As String
Dim M20x90 As String
Dim M20x110 As String
Dim M20x60 As String
Dim M20x90_16 As String
Dim M20x60_4 As String
Dim MP_Giro1 As String
Dim MP_Giro2 As String
Dim MP_Fusible As String
Dim mp_90 As String
Dim mp_180 As String
Dim mp_270 As String
Dim mp_450 As String
Dim mp_900 As String
Dim PS_280 As String
Dim PS_560 As String
Dim PS_750 As String
Dim PS_1500 As String
Dim PS_3000 As String
Dim PS_4500 As String
Dim PS_6000 As String
Dim MP_Husillo As String
Dim zMP_Base As String
Dim PS_Placa50mm As String
Dim MP_Gato As String
Dim MP_Jack As String
Dim lgiro1 As Double
Dim lgiro2 As Double
Dim lfusible As Double
Dim ljack As Double
Dim l280 As Double
Dim l560 As Double
Dim l750 As Double
Dim l1500 As Double
Dim l3000 As Double
Dim l4500 As Double
Dim l6000 As Double
Dim l900 As Double
Dim l450 As Double
Dim l270 As Double
Dim l180 As Double
Dim l90 As Double
Dim l95 As Double
Dim l315 As Double
Dim l50 As Double
Dim l_base As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lgatomin As Double
Dim n6000 As Integer
Dim n4500 As Integer
Dim n3000 As Integer
Dim n1500 As Integer
Dim n750 As Integer
Dim n560 As Integer
Dim n280 As Integer
Dim n900 As Integer
Dim n450 As Integer
Dim n270 As Integer
Dim n180 As Integer
Dim n90 As Integer
Dim njack As Integer
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
Dim dato1 As String
Dim dato2 As String
Dim dato3 As String
Dim dato4 As String
Dim dato5 As String
Dim tipoplaca1 As String
Dim tipoplaca2 As String
Dim plalz As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object

Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

On Error GoTo terminar
repite = 1

Ncapa = "Mega"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Pipeshor4S"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5

'Valores fijos
PI = 4 * Atn(1)
lgiro1 = 90
lgiro2 = 90
lfusible = 90
l280 = 280
l560 = 560
l750 = 750
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
l50 = 50
l900 = 900
l450 = 450
l270 = 270
l180 = 180
l90 = 90
l95 = 95
l315 = 315
lgatomin = 435

kwordList = "Cuña PlacaMP PlacaMPCompacta"
tipoplaca1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
tipoplaca1 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Tipo de placa en extremo1?: [Cuña/PlacaMP/PlacaMPCompacta]")

kwordList = "Cuña PlacaMP PlacaMPCompacta"
tipoplaca2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
tipoplaca2 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Tipo de placa en extremo2?: [Cuña/PlacaMP/PlacaMPCompacta]")

kwordList = "Si No"
dato4 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato4 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Introducir Jack Plate?: [Si/No]")

If tipoplaca1 = "" Or tipoplaca1 = "Cuña" Then
lgiro1 = 90
rutapl1 = "MG_AnguloGiro"
ElseIf tipoplaca1 = "PlacaMP" Then
lgiro1 = 315
rutapl1 = "PL_GCODAL_"
ElseIf tipoplaca1 = "PlacaMPCompacta" Then
lgiro1 = 95
rutapl1 = "PL_GCODAL_C_"
Else
GoTo terminar
End If

If tipoplaca2 = "" Or tipoplaca2 = "Cuña" Then
lgiro2 = 90
rutapl2 = "MG_AnguloGiro"
ElseIf tipoplaca2 = "PlacaMP" Then
lgiro2 = 315
rutapl2 = "PL_GCODAL_"
ElseIf tipoplaca2 = "PlacaMPCompacta" Then
lgiro2 = 95
rutapl2 = "PL_GCODAL_C_"
Else
GoTo terminar
End If

If dato4 = "Si" Or dato4 = "" Then
njack = 1
ElseIf dato4 = "No" Then
njack = 0
Else
GoTo terminar
End If
ljack = njack * 40
lfija = (lgiro1 + lgiro2) + lfusible + (2 * l50) + lgatomin + ljack

kwordList = "S L"
dato2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato2 = "Pshor_4" & ThisDrawing.Utility.GetKeyword(vbLf & "Introduce PS4S o PS4L: [S/L]")

If dato2 = "Pshor_4L" Then
capa = "Pipeshor4L"
dato5 = "PL"
n560 = 0
ElseIf dato2 = "Pshor_4S" Or dato2 = "Pshor_4" Then
dato2 = "Pshor_4S"
capa = "Pipeshor4S"
dato5 = "PS"
Else
GoTo terminar
End If

ruta1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\" & dato2 & "\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
ruta3 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"
ruta4 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"

kwordList = "Planta Alzado"
dato1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")
If dato1 = "" Or dato1 = "Planta" Then
dato1 = "planta"
plalz = "PLA"
ElseIf dato1 = "Alzado" Then
dato1 = "alzado"
plalz = "ALZ"
Else
GoTo terminar
End If

If dato2 = "Pshor_4S" Then
    kwordList = "1500 750 560 280"
    dato3 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    dato3 = ThisDrawing.Utility.GetKeyword(vbLf & "Viga pipeshor de menor longitud en el puntal: [1500/750/560/280]")
ElseIf dato2 = "Pshor_4L" Then
    kwordList = "1500 750 280"
    dato3 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    dato3 = ThisDrawing.Utility.GetKeyword(vbLf & "Viga pipeshor de menor longitud en el puntal: [1500/750/280]")
End If

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

If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
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

If dato2 = "Pshor_4S" Then
    If dato3 = "" Or dato3 = "1500" Then
        n750 = 0
        n560 = 0
        n280 = 0
    ElseIf dato3 = 750 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n280 = 0
    ElseIf dato3 = 560 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n560 = Fix(lpuntal / l560)
        lpuntal = lpuntal - n560 * l560
    ElseIf dato3 = 280 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n560 = Fix(lpuntal / l560)
        lpuntal = lpuntal - n560 * l560
        n280 = Fix(lpuntal / l280)
        lpuntal = lpuntal - n280 * l280
    Else
    GoTo terminar
    End If
ElseIf dato2 = "Pshor_4L" Then
    n560 = 0
    If dato3 = "" Or dato3 = "1500" Then
        n750 = 0
        n280 = 0
    ElseIf dato3 = 750 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n280 = 0
    ElseIf dato3 = 280 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n280 = Fix(lpuntal / l280)
        lpuntal = lpuntal - n280 * l280
    Else
    GoTo terminar
    End If
End If

n900 = Fix(lpuntal / l900)
lpuntal = lpuntal - n900 * l900
n450 = Fix(lpuntal / l450)
lpuntal = lpuntal - n450 * l450
n270 = Fix(lpuntal / l270)
lpuntal = lpuntal - n270 * l270
n180 = Fix(lpuntal / l180)
lpuntal = lpuntal - n180 * l180
n90 = Fix(lpuntal / l90)
lpuntal = lpuntal - n90 * l90

MP_Giro1 = ruta3 & rutapl1 & plalz & ".dwg"
MP_Giro2 = ruta3 & rutapl2 & plalz & ".dwg"

Set blockRef = gcadModel.InsertBlock(P1, MP_Giro1, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_inial(0) = P1(0) + lgiro1 * Cos(ANG): Punto_inial(1) = P1(1) + lgiro1 * Sin(ANG): Punto_inial(2) = P1(2)
MP_Fusible = ruta2 & "Mshor90" & plalz & "fusible.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Fusible, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
M20x50 = ruta4 & "4-M20x50.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x50, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
        blockRef.Explode
        blockRef.Delete
Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)

If n90 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_90 = ruta2 & "Mshor90" & plalz & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x50, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If
M20x60 = ruta4 & "6-M20x60.dwg"
If n180 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_180 = ruta2 & "Mshor180" & plalz & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"

        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n270 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_270 = ruta2 & "Mshor270" & plalz & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Punto_final(0) = Punto_inial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n450 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_450 = ruta2 & "Mshor450" & plalz & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n900 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_900 = ruta2 & "Mshor900" & plalz & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_900, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        Punto_final(0) = Punto_inial(0) + l900 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l900 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
PS_Placa50mm = ruta1 & "PS_Placa50mm_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Pipeshor4S"
M20x90 = ruta4 & "4-M20x90.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
        blockRef.Explode
        blockRef.Delete
Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
M20x110 = ruta4 & "4-M20x110.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, M20x110, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
        blockRef.Explode
        blockRef.Delete
M20x90_16 = ruta4 & "16-M20X90.dwg"

If n280 > 0 Then
    i = 0
    Do While i < n280
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_280 = ruta1 & "PL_280_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_280, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Pipeshor4L"
        If i > 0 Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        End If
        Punto_final(0) = Punto_inial(0) + l280 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l280 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        i = i + 1
    Loop
End If

If n1500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_1500 = ruta1 & dato5 & "_1500_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_1500, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        End If
        Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n3000 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_3000 = ruta1 & dato5 & "_3000_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_3000, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        End If
        Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If


If n4500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_4500 = ruta1 & dato5 & "_4500_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_4500, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Or n3000 > 0 Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        End If
        Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n6000 > 0 Then
    i = 0
    Do While i < n6000
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_6000 = ruta1 & dato5 & "_6000_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_6000, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Or n3000 > 0 Or n4500 > 0 Or i > 0 Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        End If
        Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        i = i + 1
    Loop
End If

If n750 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_750 = ruta1 & dato5 & "_750_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_750, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Or n3000 > 0 Or n4500 > 0 Or n6000 > 0 Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        End If
        Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n560 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_560 = ruta1 & dato5 & "_560_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Or n3000 > 0 Or n4500 > 0 Or n6000 > 0 Or n750 > 0 Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        End If
        Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Pipeshor4S"
Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
        blockRef.Explode
        blockRef.Delete
Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
        blockRef.Explode
        blockRef.Delete
zMP_Base = ruta3 & "zMGBaseGato_azul.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, zMP_Base, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)

Set blockRef = gcadModel.InsertBlock(P2, MP_Giro2, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Mega"
Punto_final2(0) = P2(0) - lgiro2 * Cos(ANG): Punto_final2(1) = P2(1) - lgiro2 * Sin(ANG): Punto_final2(2) = P2(2)

If njack = 1 Then
    MP_Jack = ruta2 & "MshorJACKPLATE.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final2, MP_Jack, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
        blockRef.Explode
        blockRef.Delete
    Punto_final2(0) = Punto_final2(0) - ljack * Cos(ANG): Punto_final2(1) = Punto_final2(1) - ljack * Sin(ANG): Punto_final2(2) = Punto_final2(2)
ElseIf njack = 0 Then
    M20x60_4 = ruta4 & "4-M20x60.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x60_4, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
        blockRef.Explode
        blockRef.Delete
End If
zMP_Base = ruta3 & "zMGBaseGato_naranja.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final2, zMP_Base, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Mega"

Punto_inial(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_inial(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_inial(2) = (Punto_final(2) + Punto_final2(2)) / 2

MP_Husillo = ruta3 & "MGHusilloGato.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_Husillo, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"

Eje1.Layer = "Nonplot"
Loop
terminar:
End Sub


















