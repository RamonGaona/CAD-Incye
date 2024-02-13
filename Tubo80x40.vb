

Sub Tubo80x40()
Dim rutags As String
Dim ruta1 As String
Dim husillo As String
Dim ruta2 As String, rutass As String
Dim gcadDoc As Object
Dim gcadUtil As Object
Dim gcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim Tornillo As String, Tornillo3 As String, ss_husillo As String, ss_gatoizq As String, ss_gatodrc As String
Dim Tornillo1 As String
Dim Tornillo2 As String
Dim Espada1 As String
Dim Angulo As String
Dim Base_drc As String, Base_izq As String
Dim espadags As String
Dim extremo1 As String, extremo2 As String
Dim espadampcorta As String, espadamplarga As String, espadass As String, gato As String, tornillogs1 As String, tornillogs2 As String
Dim langulo As Double, nangulo As Integer, lespadags As Double, lespadampcorta As Double, lespadamplarga As Double, lespadass As Double, lgato As Double
Dim M16x40 As String
Dim M24x110 As String, M20x130 As String
Dim nespadags As Integer, nespadampcorta As Integer, nespadamplarga As Integer, nespadass As Integer, n360os As Integer, ngato As Integer
Dim M20x60 As String
Dim TornilloLibre1 As String
Dim TornilloLibre2 As String
Dim TornilloL1 As String
Dim TornilloL2 As String
Dim M20x90 As String
Dim Tn_Husillo As String
Dim zTn_Base_drc As String
Dim zTn_Base_izq As String
Dim Tn_Husillo2 As String
Dim zTn_Base_drc2 As String
Dim zTn_Base_izq2 As String
Dim Espada As String, gato1 As String, gato2 As String, Gatoref1 As String, Gatoref2 As String
Dim dato1 As String
Dim ss_90 As String
Dim ss_180 As String
Dim ss_360 As String
Dim Tn_400 As String
Dim Tn_800 As String
Dim Tn_2000 As String
Dim Tn_1600 As String, Tn_2800 As String, Tn_2400
Dim Tn_3200 As String
Dim lespada As Double
Dim l540 As Double, l720 As Double, l900 As Double, l2700 As Double, l3600 As Double, l2000 As Double
Dim l400 As Double, l360 As Double, l180 As Double, l90 As Double, l1800 As Double
Dim l800 As Double
Dim l1600 As Double
Dim l3200 As Double, l2800 As Double, l2400 As Double
Dim lpuntal As Double
Dim lgatomin1 As Double, lgatomin2 As Double
Dim laux As Double
Dim n400 As Integer, n360 As Integer, n180 As Integer, n90 As Integer
Dim n800 As Integer
Dim n1600 As Integer
Dim n3200 As Integer, n2800 As Integer, n2400 As Integer, n2000 As Integer
Dim blockRef As Object
Dim repite As Double
Dim plalz As String
Dim gatomin1 As Double, gatomin2 As Double
Dim Punto_inial(0 To 2) As Double
Dim Punto_final(0 To 2) As Double
Dim Punto_inial2(0 To 2) As Double
Dim Punto_final2(0 To 2) As Double
Dim PI As Variant
Dim Eje1 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim lAjustegatos As Double
Dim ANG As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim kwordList As String
Dim i As Integer
Dim lfija As Double
Dim Ncapaslim As String, vigaM As String
Dim capaslim As Object
Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility
On Error GoTo terminar
repite = 1
'Valores constantes
PI = 4 * Atn(1)
lAjustegatos = 96.5
l90 = 90
l180 = 180
l360 = 360
l540 = 540
l400 = 400
l720 = 720
l800 = 800
l900 = 900
l1600 = 1600
l1800 = 1800
l2000 = 2000
l2700 = 2700
l2800 = 2800
l2400 = 2400
l3200 = 3200
l3600 = 3600
lespadags = 150
lespadass = 358
lespadamplarga = 358
lespadampcorta = 158

Ncapaslim = "Slims"
Set capaslim = gcadDoc.Layers.Add(Ncapaslim)
capaslim.color = 30

ruta1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutass = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"

M16x40 = ruta2 & "4-M16X40.dwg"
M24x110 = ruta2 & "1-M24X110.dwg"
M20x130 = ruta2 & "1-M20X130.dwg"

gato1 = ""
Gatoref1 = ""
kwordList = "Ligero Slim No"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
gato1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de gato en el primer extremo: [Ligero/Slim/No]")
    If gato1 = "Ligero" Or gato1 = "" Then
        lgatomin1 = 560
    ElseIf gato1 = "Slim" Then
        lgatomin1 = 420
    ElseIf gato1 = "No" Then
        lgatomin1 = 0
    Else: GoTo terminar
    End If
    
gato2 = ""
Gatoref2 = ""
kwordList = "Ligero Slim No"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
gato2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de gato en el segundo extremo: [Ligero/Slim/No]")
    If gato2 = "Ligero" Or gato2 = "" Then
        lgatomin2 = 560
    ElseIf gato2 = "Slim" Then
        lgatomin2 = 420
    ElseIf gato2 = "No" Then
        lgatomin2 = 0
    Else: GoTo terminar
    End If

If lgatomin1 > 0 Then
    kwordList = "A-Libre B-EspadaSlimAlzado C-EspadaSlimPlanta D-EspadaMpLargaAlzado E-EspadaMpLargaPlanta F-EspadaMpCortaAlzado G-EspadaMpCortaPlanta H-EspadaGranshorAlzado I-EspadaGranshorPlanta"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    extremo1 = ThisDrawing.Utility.GetKeyword(vbLf & "Espada extremo 1: [A-Libre/B-EspadaSlimAlzado/C-EspadaSlimPlanta/D-EspadaMpLargaAlzado/E-EspadaMpLargaPlanta/F-EspadaMpCortaAlzado/G-EspadaMpCortaPlanta/H-EspadaGranshorAlzado/I-EspadaGranshorPlanta]")
        If extremo1 = "" Or extremo1 = "A-Libre" Then
        extremo1 = "A-Libre"
        nespadags = 0
        nespadampcorta = 0
        nespadamplarga = 0
        nespadass = 0
        ElseIf extremo1 = "B-EspadaSlimAlzado" Or extremo1 = "C-EspadaSlimPlanta" Then
        nespadags = 0
        nespadampcorta = 0
        nespadamplarga = 0
        nespadass = 1
        ElseIf extremo1 = "D-EspadaMpLargaAlzado" Or extremo1 = "E-EspadaMpLargaPlanta" Then
        nespadags = 0
        nespadampcorta = 0
        nespadamplarga = 1
        nespadass = 0
        ElseIf extremo1 = "F-EspadaMpCortaAlzado" Or extremo1 = "G-EspadaMpCortaPlanta" Then
        nespadags = 0
        nespadampcorta = 1
        nespadamplarga = 0
        nespadass = 0
        ElseIf extremo1 = "H-EspadaGranshorAlzado" Or extremo1 = "I-EspadaGranshorPlanta" Then
        nespadags = 1
        nespadampcorta = 0
        nespadamplarga = 0
        nespadass = 0
        kwordList = "A-M20x40 B-M20x60"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        tornillogs1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tornillo en la espada Gshor del extremo 1: [A-M20x40/B-M20x60]")
            If tornillogs1 = "A-M20x40" Or tornillogs1 = "" Then
            tornillogs1 = ruta2 & "1-M20X40.dwg"
            ElseIf tornillogs1 = "B-M20x60" Then
            tornillogs1 = ruta2 & "1-M20X60.dwg"
            End If
        End If
Else
    nespadags = 0
    nespadampcorta = 0
    nespadamplarga = 0
    nespadass = 0
End If

If lgatomin1 > 0 Then
    If extremo1 = "" Or extremo1 = "A-Libre" Then
        kwordList = "AM20x90 BM24x110 CM20x110"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        TornilloLibre1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tornillo en el extremo 1: [AM20x90/BM24x110/CM20x110]")
            If TornilloLibre1 = "" Or TornilloLibre1 = "AM20x90" Then
            TornilloLibre1 = "1-M20X90.dwg"
            ElseIf TornilloLibre1 = "BM24x110" Then
            TornilloLibre1 = "1-M24X110.dwg"
            ElseIf TornilloLibre1 = "CM20x110" Then
            TornilloLibre1 = "1-M20X110.dwg"
            Else: GoTo terminar
            End If
    End If
End If

If lgatomin2 > 0 Then
    kwordList = "A-Libre B-EspadaSlimAlzado C-EspadaSlimPlanta D-EspadaMpLargaAlzado E-EspadaMpLargaPlanta F-EspadaMpCortaAlzado G-EspadaMpCortaPlanta H-EspadaGranshorAlzado I-EspadaGranshorPlanta"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    extremo2 = ThisDrawing.Utility.GetKeyword(vbLf & "Espada extremo 2: [A-Libre/B-EspadaSlimAlzado/C-EspadaSlimPlanta/D-EspadaMpLargaAlzado/E-EspadaMpLargaPlanta/F-EspadaMpCortaAlzado/G-EspadaMpCortaPlanta/H-EspadaGranshorAlzado/I-EspadaGranshorPlanta]")
    If extremo2 = "" Or extremo2 = "A-Libre" Then
    extremo2 = "A-Libre"
    nespadags = nespadags + 0
        nespadampcorta = nespadampcorta + 0
        nespadamplarga = nespadamplarga + 0
        nespadass = nespadass + 0
        ElseIf extremo2 = "B-EspadaSlimAlzado" Or extremo2 = "C-EspadaSlimPlanta" Then
        nespadags = nespadags + 0
        nespadampcorta = nespadampcorta + 0
        nespadamplarga = nespadamplarga + 0
        nespadass = nespadass + 1
        ElseIf extremo2 = "D-EspadaMpLargaAlzado" Or extremo2 = "E-EspadaMpLargaPlanta" Then
        nespadags = nespadags + 0
        nespadampcorta = nespadampcorta + 0
        nespadamplarga = nespadamplarga + 1
        nespadass = nespadass + 0
        ElseIf extremo2 = "F-EspadaMpCortaAlzado" Or extremo2 = "G-EspadaMpCortaPlanta" Then
        nespadags = nespadags + 0
        nespadampcorta = nespadampcorta + 1
        nespadamplarga = nespadamplarga + 0
        nespadass = nespadass + 0
        ElseIf extremo2 = "H-EspadaGranshorAlzado" Or extremo2 = "I-EspadaGranshorPlanta" Then
        nespadags = nespadags + 1
        nespadampcorta = nespadampcorta + 0
        nespadamplarga = nespadamplarga + 0
        nespadass = nespadass + 0
        kwordList = "A-M20x40 B-M20x60"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        tornillogs2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tornillo en la espada Gshor del extremo 2: [A-M20x40/B-M20x60]")
        If tornillogs2 = "A-M20x40" Or tornillogs2 = "" Then
            tornillogs2 = ruta2 & "1-M20X40.dwg"
        ElseIf tornillogs2 = "B-M20x60" Then
            tornillogs2 = ruta2 & "1-M20X60.dwg"
        End If
    End If
Else
End If

If lgatomin2 > 0 Then
    If extremo2 = "" Or extremo2 = "A-Libre" Then
        kwordList = "AM20x90 BM24x110 CM20x110"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        TornilloLibre2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tornillo en el extremo 2: [AM20x90/BM24x110/CM20x110]")
        If TornilloLibre2 = "" Or TornilloLibre2 = "AM20x90" Then
            TornilloLibre2 = "1-M20X90.dwg"
        ElseIf TornilloLibre2 = "BM24x110" Then
            TornilloLibre2 = "1-M24X110.dwg"
        ElseIf TornilloLibre2 = "CM20x110" Then
            TornilloLibre2 = "1-M20X110.dwg"
        Else: GoTo terminar
            End If
    End If
End If

lfija = lgatomin1 + lgatomin2 + nespadags * lespadags + nespadampcorta * lespadampcorta + nespadamplarga * lespadamplarga + nespadass * lespadass

kwordList = "3200 2800 2400 2000 1600"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vigaM = ThisDrawing.Utility.GetKeyword(vbLf & "Viga de mayor tamaño: [3200/2800/2400/2000/1600]")

kwordList = "Planta Alzado"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
plalz = ThisDrawing.Utility.GetKeyword(vbLf & "Vista de los tubos en: [Planta/Alzado]")

If plalz = "" Or plalz = "Planta" Then
plalz = "PL"
ElseIf plalz = "Alzado" Then
plalz = ""
Else
GoTo terminar
End If

Do While repite = 1

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
Distancia = Sqr(x ^ 2 + y ^ 2)

If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If

lpuntal = Distancia - lfija
laux = lpuntal
If vigaM = "" Then vigaM = "3200"
If vigaM = 3200 Then
n3200 = Fix(lpuntal / l3200)
lpuntal = lpuntal - n3200 * l3200
n2800 = Fix(lpuntal / l2800)
lpuntal = lpuntal - n2800 * l2800
n2400 = Fix(lpuntal / l2400)
lpuntal = lpuntal - n2400 * l2400
n2000 = Fix(lpuntal / l2000)
lpuntal = lpuntal - n2000 * l2000
n1600 = Fix(lpuntal / l1600)
lpuntal = lpuntal - n1600 * l1600
ElseIf vigaM = 2800 Then
n3200 = 0
n2800 = Fix(lpuntal / l2800)
lpuntal = lpuntal - n2800 * l2800
n2400 = Fix(lpuntal / l2400)
lpuntal = lpuntal - n2400 * l2400
n2000 = Fix(lpuntal / l2000)
lpuntal = lpuntal - n2000 * l2000
n1600 = Fix(lpuntal / l1600)
lpuntal = lpuntal - n1600 * l1600
ElseIf vigaM = 2400 Then
n3200 = 0
n2800 = 0
n2400 = Fix(lpuntal / l2400)
lpuntal = lpuntal - n2400 * l2400
n2000 = Fix(lpuntal / l2000)
lpuntal = lpuntal - n2000 * l2000
n1600 = Fix(lpuntal / l1600)
lpuntal = lpuntal - n1600 * l1600
ElseIf vigaM = 2000 Then
n3200 = 0
n2800 = 0
n2400 = 0
n2000 = Fix(lpuntal / l2000)
lpuntal = lpuntal - n2000 * l2000
n1600 = Fix(lpuntal / l1600)
lpuntal = lpuntal - n1600 * l1600
ElseIf vigaM = 1600 Then
n3200 = 0
n2800 = 0
n2400 = 0
n1600 = Fix(lpuntal / l1600)
lpuntal = lpuntal - n1600 * l1600
Else: GoTo terminar
End If
n800 = Fix(lpuntal / l800)
lpuntal = lpuntal - n800 * l800
n400 = Fix(lpuntal / l400)
lpuntal = lpuntal - n400 * l400
  
If gato1 = "Slim" And gato2 = "Slim" Then
    n360 = 0
    n180 = 0
    n90 = 0
ElseIf gato1 = "Slim" And (gato2 = "Ligero" Or gato2 = "") Then
    n360 = 0
    n180 = 0
    n90 = 0
ElseIf (gato1 = "Ligero" Or gato1 = "") And gato2 = "Slim" Then
    n360 = 0
    n180 = 0
    n90 = 0
ElseIf (gato1 = "Ligero" Or gato1 = "") And (gato2 = "Ligero" Or gato2 = "") Then
    n360 = 0
    n180 = 0
    n90 = 0
ElseIf gato1 = "No" And gato2 = "Slim" Then
    n360 = Fix(lpuntal / l360)
    lpuntal = lpuntal - n360 * l360
    n180 = Fix(lpuntal / l180)
    lpuntal = lpuntal - n180 * l180
    n90 = Fix(lpuntal / l90)
    lpuntal = lpuntal - n90 * l90
ElseIf gato1 = "Slim" And gato2 = "No" Then
    n360 = Fix(lpuntal / l360)
    lpuntal = lpuntal - n360 * l360
    n180 = Fix(lpuntal / l180)
    lpuntal = lpuntal - n180 * l180
    n90 = Fix(lpuntal / l90)
    lpuntal = lpuntal - n90 * l90
ElseIf gato1 = "No" And (gato2 = "Ligero" Or gato2 = "") Then
    n360 = Fix(lpuntal / l360)
    lpuntal = lpuntal - n360 * l360
    n180 = Fix(lpuntal / l180)
    lpuntal = lpuntal - n180 * l180
    n90 = Fix(lpuntal / l90)
    lpuntal = lpuntal - n90 * l90
ElseIf (gato1 = "Ligero" Or gato1 = "") And gato2 = "No" Then
    n360 = Fix(lpuntal / l360)
    lpuntal = lpuntal - n360 * l360
    n180 = Fix(lpuntal / l180)
    lpuntal = lpuntal - n180 * l180
    n90 = Fix(lpuntal / l90)
    lpuntal = lpuntal - n90 * l90
ElseIf gato1 = "No" And gato2 = "No" Then
    n360 = Fix(lpuntal / l360)
    lpuntal = lpuntal - n360 * l360
    n180 = Fix(lpuntal / l180)
    lpuntal = lpuntal - n180 * l180
    n90 = Fix(lpuntal / l90)
    lpuntal = lpuntal - n90 * l90
End If

laux = Distancia - n3200 * l3200 - n2800 * l2800 - n2400 * l2400 - n2000 * l2000 - n1600 * l1600 - n800 * l800 - n400 * l400 - n360 * l360 - n180 * l180 - n90 * l90 - lfija

If (lgatomin1 + lgatomin2) > 561 Then
    laux = laux / 2
End If

If (lgatomin1 + lgatomin2) = 0 Then
    laux = 0
End If

Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)

If extremo1 = "" Or extremo1 = "A-Libre" Then
'No hacer nada
ElseIf extremo1 = "B-EspadaSlimAlzado" Then
    espadass = rutass & "SSEspada.dwg"
    Set blockRef = gcadModel.InsertBlock(P1, espadass, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo1 = "C-EspadaSlimPlanta" Then
    espadass = rutass & "SSPLEspada.dwg"
    Set blockRef = gcadModel.InsertBlock(P1, espadass, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo1 = "D-EspadaMpLargaAlzado" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    espadamplarga = rutass & "SSEspadaLarga.dwg"
    Set blockRef = gcadModel.InsertBlock(P1, espadamplarga, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadamplarga * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadamplarga * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo1 = "E-EspadaMpLargaPlanta" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    espadamplarga = rutass & "SSPLEspadaLarga.dwg"
    Set blockRef = gcadModel.InsertBlock(P1, espadamplarga, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadamplarga * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadamplarga * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo1 = "F-EspadaMpCortaAlzado" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    espadampcorta = rutass & "SSEspadaCorta.dwg"
    Set blockRef = gcadModel.InsertBlock(P1, espadampcorta, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadampcorta * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadampcorta * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo1 = "G-EspadaMpCortaPlanta" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    espadampcorta = rutass & "SSPLEspadaCorta.dwg"
    Set blockRef = gcadModel.InsertBlock(P1, espadampcorta, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadampcorta * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadampcorta * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo1 = "H-EspadaGranshorAlzado" Then
    Set blockRef = gcadModel.InsertBlock(P1, tornillogs1, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    espadags = rutags & "GS_espada.dwg"
    Set blockRef = gcadModel.InsertBlock(P1, espadags, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadags * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadags * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo1 = "I-EspadaGranshorPlanta" Then
    Set blockRef = gcadModel.InsertBlock(P1, tornillogs1, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    espadags = rutags & "GS_PLespada.dwg"
    Set blockRef = gcadModel.InsertBlock(P1, espadags, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadags * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadags * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
Else
GoTo terminar
End If

'####No sé muy bien si hay que poner distancia auxiliar si no hay gato
If lgatomin1 > 0 Then
    If gato1 = "Ligero" Or gato1 = "" Then
        Tn_Husillo = ruta1 & "zTensor80x4_husillo.dwg"
        zTn_Base_drc = ruta1 & "Tensor80x4" & plalz & "_base_gato_drc.dwg"
        zTn_Base_izq = ruta1 & "Tensor80x4" & plalz & "_base_gato_izq.dwg"
        Tornillo = ruta2 & "4-M16X40.dwg"
    ElseIf gato1 = "Slim" Then
        Tn_Husillo = rutass & "zSSHusillo.dwg"
        zTn_Base_drc = rutass & "SS" & plalz & "Gatorefdrc.dwg"
        zTn_Base_izq = rutass & "SS" & plalz & "Gatorefizq.dwg"
        Tornillo = ruta2 & "4-M16X40.dwg"
    ElseIf gato1 = "No" Then
        'No hacer nada
    End If
    If extremo1 = "" Or extremo1 = "A-Libre" Then
        TornilloL1 = ruta2 & TornilloLibre1
        Set blockRef = gcadModel.InsertBlock(Punto_final, TornilloL1, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
    End If
    Set blockRef = gcadModel.InsertBlock(Punto_final, Tn_Husillo, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_inial(0) = Punto_final(0) + laux * Cos(ANG): Punto_inial(1) = Punto_final(1) + laux * Sin(ANG): Punto_final(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)
    Punto_inial(0) = Punto_final(0) + lgatomin1 * Cos(ANG): Punto_inial(1) = Punto_final(1) + lgatomin1 * Sin(ANG): Punto_final(2) = Punto_final(2)
    If gato1 = "Ligero" Or gato1 = "" Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, zTn_Base_drc, Xs, Ys, Zs, ANG)
    ElseIf gato1 = "Slim" Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial, zTn_Base_drc, Xs, Ys, Zs, ANG + PI)
    End If
    blockRef.Layer = "Slims"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, Tornillo, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)
Else
End If


If n3200 > 0 Then
    i = 0
    Do While i < n3200
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Tn_3200 = ruta1 & "Tensor80x4" & plalz & "_3200.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Tn_3200, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l3200 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3200 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n2800 > 0 Then
    i = 0
    Do While i < n2800
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Tn_2800 = ruta1 & "Tensor80x4" & plalz & "_2800.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Tn_2800, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l2800 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l2800 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n2400 > 0 Then
    i = 0
    Do While i < n2400
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Tn_2400 = ruta1 & "Tensor80x4" & plalz & "_2400.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Tn_2400, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l2400 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l2400 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n2000 > 0 Then
    i = 0
    Do While i < n2000
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Tn_2000 = ruta1 & "Tensor80x4" & plalz & "_2000.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Tn_2000, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l2000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l2000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n1600 > 0 Then
    i = 0
    Do While i < n1600
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Tn_1600 = ruta1 & "Tensor80x4" & plalz & "_1600.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Tn_1600, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l1600 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1600 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n800 > 0 Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Tn_800 = ruta1 & "Tensor80x4" & plalz & "_800.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, Tn_800, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + l800 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l800 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n400 > 0 Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Tn_400 = ruta1 & "Tensor80x4" & plalz & "_400.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, Tn_400, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + l400 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l400 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n360 > 0 Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    ss_360 = rutass & "SS" & plalz & "0360.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_360, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + l360 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l360 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n180 > 0 Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    ss_180 = rutass & "SS" & plalz & "0180.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_180, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n90 > 0 Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    ss_90 = rutass & "SS" & plalz & "0090.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If lgatomin2 > 0 Then
    If gato2 = "Ligero" Or gato2 = "" Then
        Tn_Husillo2 = ruta1 & "zTensor80x4_husillo.dwg"
        zTn_Base_drc2 = ruta1 & "Tensor80x4" & plalz & "_base_gato_drc.dwg"
        zTn_Base_izq2 = ruta1 & "Tensor80x4" & plalz & "_base_gato_izq.dwg"
        Tornillo2 = ruta2 & "4-M16X40.dwg"
        ElseIf gato2 = "Slim" Then
        Tn_Husillo2 = rutass & "zSSHusillo.dwg"
        zTn_Base_drc2 = rutass & "SS" & plalz & "Gatorefdrc.dwg"
        zTn_Base_izq2 = rutass & "SS" & plalz & "Gatorefizq.dwg"
        Tornillo2 = ruta2 & "4-M16X40.dwg"
        ElseIf gato2 = "No" Then
        'No hacer nada
    End If
    If gato2 = "Ligero" Or gato2 = "" Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, zTn_Base_izq2, Xs, Ys, Zs, ANG + PI)
    ElseIf gato2 = "Slim" Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, zTn_Base_izq2, Xs, Ys, Zs, ANG + PI)
    End If
    blockRef.Layer = "SLims"
    Punto_inial(0) = Punto_final(0) + laux * Cos(ANG): Punto_inial(1) = Punto_final(1) + laux * Sin(ANG): Punto_final(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)
    Punto_inial(0) = Punto_final(0) + lgatomin2 * Cos(ANG): Punto_inial(1) = Punto_final(1) + lgatomin2 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = gcadModel.InsertBlock(Punto_inial, Tn_Husillo2, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "SLims"
    If extremo2 = "" Or extremo2 = "A-Libre" Then
        TornilloL2 = ruta2 & TornilloLibre2
        Set blockRef = gcadModel.InsertBlock(Punto_inial, TornilloL2, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
    End If
    Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)
Else
End If

If extremo2 = "" Or extremo2 = "A-Libre" Then
'No hacer nada
ElseIf extremo2 = "B-EspadaSlimAlzado" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadass = rutass & "SSEspada.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, espadass, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
ElseIf extremo2 = "C-EspadaSlimPlanta" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadass = rutass & "SSPLEspada.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, espadass, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
ElseIf extremo2 = "D-EspadaMpLargaAlzado" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadamplarga * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadamplarga * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadamplarga = rutass & "SSEspadaLarga.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, espadamplarga, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo2 = "E-EspadaMpLargaPlanta" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadamplarga * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadamplarga * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadamplarga = rutass & "SSPLEspadaLarga.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, espadamplarga, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo2 = "F-EspadaMpCortaAlzado" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadampcorta * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadampcorta * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadampcorta = rutass & "SSEspadaCorta.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, espadampcorta, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo2 = "G-EspadaMpCortaPlanta" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadampcorta * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadampcorta * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadampcorta = rutass & "SSPLEspadaCorta.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, espadampcorta, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo2 = "H-EspadaGranshorAlzado" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadags * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadags * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadags = rutags & "GS_espada.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, espadags, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = gcadModel.InsertBlock(Punto_final, tornillogs2, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf extremo2 = "I-EspadaGranshorPlanta" Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadags * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadags * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadags = rutags & "GS_PLespada.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, espadags, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = gcadModel.InsertBlock(Punto_final, tornillogs2, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
Else
GoTo terminar
End If

Eje1.Layer = "Nonplot"
Loop
terminar:
End Sub


























































