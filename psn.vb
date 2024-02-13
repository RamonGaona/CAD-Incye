''''''' Geometría para el cálculo del ajuste de los puntales pipeshor


Sub mpn(carga As Variant)
Dim ruta As String, rutaps As String, rutapl As String, rutags As String
Dim ruta2 As String
Dim gcadDoc As Object
Dim M20x60_4 As String
Dim M20x50_4 As String
Dim gcadUtil As Object
Dim gcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim line2 As GcadLine
Dim line1 As GcadLine
Dim M20x90 As String
Dim M20x150 As String, M20x110 As String, Var20x250 As String
Dim M20x160 As String
Dim M20x90_16 As String
Dim GS_Bulon120mm As String
Dim GS_Giro As String
Dim GS_Fusible As String
Dim PS_280 As String
Dim PS_750 As String, PS_560 As String
Dim PS_1500 As String
Dim PS_3000 As String
Dim PS_4500 As String
Dim PS_6000 As String
Dim placaanc1 As String
Dim rutaplaca1 As String, rutacajon As String
Dim placaanc2 As String
Dim rutaplaca2 As String
Dim basecajon As String, brazocajon As String
Dim PS_Husillo As String
Dim PS_Placa50mm As String
Dim PS_Placa35mm As String
Dim zPS_Gato_Cono As String
Dim zPS_Gato_Tope As String
Dim PS_Gato As String
Dim lgiro As Double
Dim lfusible As Double
Dim l280 As Double
Dim l750 As Double, l560 As Double
Dim l1500 As Double
Dim l3000 As Double
Dim l4500 As Double
Dim l6000 As Double
Dim l50 As Double, l35 As Double
Dim l145 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lalt1 As Double
Dim lalt2 As Double
Dim lgatomin As Double
Dim lcajonmax As Double
Dim lcajonmin As Double
Dim n6000 As Integer
Dim n4500 As Integer
Dim OffsetDist As Double
Dim n3000 As Integer
Dim n1500 As Integer
Dim n750 As Integer, n560 As Integer
Dim n280 As Integer
Dim nfusible As Integer
Dim blockRef As Object
Dim repite As Double
Dim Punto_inial(0 To 2) As Double
Dim Punto_final(0 To 2) As Double
Dim Punto_inial2(0 To 2) As Double
Dim Punto_final2(0 To 2) As Double
Dim PI As Variant
Dim Eje1 As Object
Dim Eje2 As Object
Dim Eje3 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim ANG As Double
Dim ANG2 As Double
Dim DirBulon1 As Double
Dim DirBulon2 As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim dato1 As String
Dim plcu As String
Dim dato2 As String
Dim dato3 As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object
Dim puntoA As Variant
Dim puntop1 As Variant
Dim puntoB As Variant
Dim puntop2 As Variant
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim x3 As Double
Dim y3 As Double
Dim x4 As Double
Dim y4 As Double
Dim x5 As Double
Dim y5 As Double
Dim PGato1(0 To 2) As Double
Dim PGato2(0 To 2) As Double
Dim PA(0 To 2) As Double
Dim PAP(0 To 2) As Double
Dim PP1(0 To 2) As Double
Dim PP2(0 To 2) As Double
Dim PB(0 To 2) As Double
Dim PBP(0 To 2) As Double
Dim Esq(0 To 2) As Double
Dim userInput As String
Dim muro1menor  As Double
Dim muromoduladome1  As Double
Dim muromoduladoma1 As Double
Dim mod1c As Double
Dim mod1M As Double
Dim muro2menor As Double
Dim muromoduladome2 As Double
Dim muromoduladoma2 As Double
Dim mod2c As Double
Dim mod2M As Double
Dim objgcadDimAligned As GcadDimAligned
Dim TxtPnt(0 To 2) As Double
Dim TxtPnt2(0 To 2) As Double
Dim TxtPnt3(0 To 2) As Double
Dim Pcerca2(0 To 2) As Double
Dim Plejos2(0 To 2) As Double
Dim Pcerca1(0 To 2) As Double
Dim Plejos1(0 To 2) As Double
Dim Slope1 As Double
Dim Slope2 As Double
Dim D_A0 As Double
Dim D_B0 As Double
Dim D_AB As Double
Dim D_ABP As Double
Dim D_Gato As Double
Dim DirMuro1 As Double
Dim DirMuro2 As Double
Dim DirMuro1Inv As Double
Dim DirMuro2Inv As Double
Dim DirPuntal As Double
Dim rutacu As String
Dim DirPuntal2 As Double


Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

Ncapa = "Pipeshor4S"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5
Ncapa = "Granshor"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150

'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgiro = 205
lfusible = 187.5
l145 = 145
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
lcajonmin = 835
lcajonmax = 1022


On Error GoTo terminar
Dim rutamp As String

ruta = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\" & dato2 & "\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
rutacu = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"
rutacajon = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Cajon hidraulico\"

Dim plcue1 As String
Dim plcue2 As String
Dim plcu2 As String



kwordList = "Planta Alzado"
dato1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")
If dato1 = "Planta" Or dato1 = "" Then
dato1 = "Planta"
plcue1 = "Ialzado"
plcue2 = "Dalzado"
plcu2 = "PLA"
ElseIf dato1 = "Alzado" Then
plcue1 = "Ialzado"
plcue2 = "Dalzado"
plcu2 = "ALZ"
Else
GoTo terminar
End If


Do While repite = 1


lalt1 = 0
lalt2 = 0

'''' caso General
'' seleccionar los 4 puntos en sentido horario


puntoA = gcadUtil.GetPoint(, "punto inserción 1ª placa: ")
puntop1 = gcadUtil.GetPoint(puntoA, "punto direccional del muro 1 (convergente): ")

puntoB = gcadUtil.GetPoint(, "punto inserción 2ª placa: ")
puntop2 = gcadUtil.GetPoint(puntoB, "punto direccional del muro 2 (convergente): ")

'PA es el punto de inserción de la primera placa
PA(0) = puntoA(0): PA(1) = puntoA(1): PA(2) = puntoA(2)
PP1(0) = puntop1(0): PP1(1) = puntop1(1): PP1(2) = puntop1(2)

'PB es el punto de inserción de la segunda placa
PB(0) = puntoB(0): PB(1) = puntoB(1): PB(2) = puntoB(2)
PP2(0) = puntop2(0): PP2(1) = puntop2(1): PP2(2) = puntop2(2)

DirMuro1 = gcadUtil.AngleFromXAxis(PA, PP1)
DirMuro2 = gcadUtil.AngleFromXAxis(PB, PP2)
DirPuntal = gcadUtil.AngleFromXAxis(PA, PB)



' conseguir la esquina:
' Calculamos las direcciones de las rectas
Slope1 = Tan(DirMuro1)
Slope2 = Tan(DirMuro2)

If DirMuro1 = DirMuro2 Then
Else
' Calculamos el punto intersección
    Esq(0) = (PB(1) - PA(1) - Slope2 * PB(0) + Slope1 * PA(0)) / (Slope1 - Slope2)
    Esq(1) = PA(1) + Slope1 * (Esq(0) - PA(0))
    Esq(2) = PA(2) ' Assuming the lines are in the same plane
    
End If


If Abs(DirMuro2 - DirMuro1) > PI Then
    If DirMuro2 > DirMuro1 Then
        ' Intercambiar los puntos P1 y P3
        Dim tempP0(0 To 2) As Double
        tempP0(0) = PA(0): tempP0(1) = PA(1): tempP0(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = tempP0(0): PB(1) = tempP0(1): PB(2) = tempP0(2):
        ' Recalcular la dirección del muro 1 y perpendicular al muro
        DirMuro1 = gcadUtil.AngleFromXAxis(PA, Esq)
        ' Recalcular la dirección del muro 2 y perpendicular al muro
        DirMuro2 = gcadUtil.AngleFromXAxis(PB, Esq)
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
    End If
Else
    If DirMuro2 < DirMuro1 Then
        ' Intercambiar los puntos P1 y P3
        Dim tempP(0 To 2) As Double
        tempP(0) = PA(0): tempP(1) = PA(1): tempP(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = tempP(0): PB(1) = tempP(1): PB(2) = tempP(2):
        ' Recalcular la dirección del muro 1 y perpendicular al muro
        DirMuro1 = gcadUtil.AngleFromXAxis(PA, Esq)
        ' Recalcular la dirección del muro 2 y perpendicular al muro
        DirMuro2 = gcadUtil.AngleFromXAxis(PB, Esq)
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
    End If
End If




'If DirMuro2 > DirMuro1 Then
    ' Intercambiar los puntos P1 y P3
    'Dim tempP(0 To 2) As Double
    'tempP(0) = PA(0): tempP(1) = PA(1): tempP(2) = PA(2)
    'PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
    'PB(0) = tempP(0): PB(1) = tempP(1): PB(2) = tempP(2)

    ' Recalcular la dirección del muro 1 y perpendicular al muro
    'DirMuro1 = gcadUtil.AngleFromXAxis(PA, Esq)
    ' Recalcular la dirección del muro 2 y perpendicular al muro
    'DirMuro2 = gcadUtil.AngleFromXAxis(PB, Esq)
    'DirPuntal = gcadUtil.AngleFromXAxis(PA, PB)
'Else
    'DirBulon1 = DirMuro1 - (PI / 2)
    'DirBulon2 = DirMuro2 + (PI / 2)
'End If


DirBulon1 = DirMuro1 - (PI / 2)
DirBulon2 = DirMuro2 + (PI / 2)

'''''Extremo 1 del puntal
Dim AnguloAbsoluto2 As Double
Dim UmbralMin2 As Double
Dim UmbralMax2 As Double

UmbralMin2 = 80 * (PI / 180) ' Convertir 80 grados a radianes
UmbralMax2 = 100 * (PI / 180) ' Convertir 105 grados a radianes

AnguloAbsoluto2 = Abs(DirPuntal - DirMuro1)

If AnguloAbsoluto2 > UmbralMin2 And AnguloAbsoluto2 < UmbralMax2 Then
    ' No pasa nada
    P2(0) = PA(0) + 85 * Cos(DirBulon1): P2(1) = PA(1) + 85 * Sin(DirBulon1): P2(2) = PA(2)
    placaanc2 = "CompactaMP"
Else
    kwordList = "Naranja Azul MP CompactaMP"
    placaanc2 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    placaanc2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de cuña en el primer extremo seleccionado: [Naranja/Azul/MP/CompactaMP]")
    
    ' aquí vienen los condicionales de las diferentes placas de anclaje, según cuál sea la seleccionada, moveremos para encontrar el P1 en este caso
    ' tenemos todas las medidas para P1
    
    If placaanc2 = "" Or placaanc2 = "Naranja" Then
        P2(0) = PA(0) + 239.2 * Cos(DirBulon1): P2(1) = PA(1) + 239.2 * Sin(DirBulon1): P2(2) = PA(2)
        P2(0) = P2(0) + 179.21 * Cos(DirMuro1): P2(1) = P2(1) + 179.21 * Sin(DirMuro1): P2(2) = P2(2)
    ElseIf placaanc2 = "Azul" Then
        P2(0) = PA(0) + 288.7 * Cos(DirBulon1): P2(1) = PA(1) + 288.7 * Sin(DirBulon1): P2(2) = PA(2)
        P2(0) = P2(0) + 213.7 * Cos(DirMuro1): P2(1) = P2(1) + 213.7 * Sin(DirMuro1): P2(2) = P2(2)
    ElseIf placaanc2 = "MP" Then
        P2(0) = PA(0) + 90 * Cos(DirBulon1): P2(1) = PA(1) + 90 * Sin(DirBulon1): P2(2) = PA(2)
    ElseIf placaanc2 = "CompactaMP" Then
        P2(0) = PA(0) + 85 * Cos(DirBulon1): P2(1) = PA(1) + 85 * Sin(DirBulon1): P2(2) = PA(2)
    End If
    
End If


'''''Extremo 2 del puntal
Dim AnguloAbsoluto1 As Double
Dim UmbralMin1 As Double
Dim UmbralMax1 As Double

UmbralMin1 = 80 * (PI / 180) ' Convertir 80 grados a radianes
UmbralMax1 = 100 * (PI / 180) ' Convertir 105 grados a radianes

AnguloAbsoluto1 = Abs((DirPuntal + PI) - DirMuro2)

If AnguloAbsoluto1 > UmbralMin1 And AnguloAbsoluto1 < UmbralMax1 Then
    P1(0) = PB(0) + 85 * Cos(DirBulon2): P1(1) = PB(1) + 85 * Sin(DirBulon2): P1(2) = PB(2)
    placaanc1 = "CompactaMP"
Else
    kwordList = "Naranja Azul MP CompactaMP"
    placaanc1 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    placaanc1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de cuña en el segundo extremo seleccionado: [Naranja/Azul/MP/CompactaMP]")

    ' aquí vienen los condicionales de las diferentes placas de anclaje, según cuál sea la seleccionada, moveremos para encontrar el P1 en este caso
    ' tenemos todas las medidas para P1
    
    If placaanc1 = "" Or placaanc1 = "Naranja" Then
        P1(0) = PB(0) + 239.2 * Cos(DirBulon2): P1(1) = PB(1) + 239.2 * Sin(DirBulon2): P1(2) = PB(2)
        P1(0) = P1(0) + 179.21 * Cos(DirMuro2): P1(1) = P1(1) + 179.21 * Sin(DirMuro2): P1(2) = P1(2)
    ElseIf placaanc1 = "Azul" Then
        P1(0) = PB(0) + 288.7 * Cos(DirBulon2): P1(1) = PB(1) + 288.7 * Sin(DirBulon2): P1(2) = PB(2)
        P1(0) = P1(0) + 213.7 * Cos(DirMuro2): P1(1) = P1(1) + 213.7 * Sin(DirMuro2): P1(2) = P1(2)
    ElseIf placaanc1 = "MP" Then
        P1(0) = PB(0) + 90 * Cos(DirBulon2): P1(1) = PB(1) + 90 * Sin(DirBulon2): P1(2) = PB(2)
    ElseIf placaanc1 = "CompactaMP" Then
        P1(0) = PB(0) + 85 * Cos(DirBulon2): P1(1) = PB(1) + 85 * Sin(DirBulon2): P1(2) = PB(2)
    End If
        
End If




DirPuntal = gcadUtil.AngleFromXAxis(P1, P2)
DirPuntal2 = gcadUtil.AngleFromXAxis(P2, P1)

'''' podemos también añadir las rutas de la placa1 y placa2 como hago aquí abajo pero lo metemos directamente en el condicional que tenemos aquí arriba, para dejar ya cerrada cuál va a ser cada una de las placas
'''' además de añadir las capas. En caso de que sea el ángulo de giro podemos añadirlo también.


If placaanc1 = "Naranja" Or placaanc1 = "" Then
    rutaplaca1 = rutacu & "MG_CunaNar_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "Azul" Then
    rutaplaca1 = rutacu & "MG_CunaAz_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "MP" Then
    rutaplaca1 = rutacu & "PlacaMP_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "CompactaMP" Then
    rutaplaca1 = rutacu & "PlacaMP_C_" & plcue2 & ".dwg"
    capa = "Mega"
End If

If placaanc2 = "Naranja" Or placaanc2 = "" Then
    rutaplaca2 = rutacu & "MG_CunaNar_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "Azul" Then
    rutaplaca2 = rutacu & "MG_CunaAz_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "MP" Then
    rutaplaca2 = rutacu & "PlacaMP_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "CompactaMP" Then
    rutaplaca2 = rutacu & "PlacaMP_C_" & plcue1 & ".dwg"
    capa = "Mega"
End If


Dim agiro As String
agiro = rutacu & "ANGgiro.dwg"

DirMuro1Inv = DirMuro1 - PI
DirMuro2Inv = DirMuro2 - PI

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' Aquí viene la toma de decisiones de si coger un puntal o coger otro::::
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If carga <= 1020 Then
    lfija = 680
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
ElseIf carga > 1020 And carga <= 1200 Then
    lfija = 720
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
ElseIf carga > 1200 Then
    lfija = 760
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
End If

Xs = 1
Ys = 1
Zs = 1


x = P2(0) - P1(0)
y = P2(1) - P1(1)

Distancia = Val(Sqr((x ^ 2 + y ^ 2)))


Dim n5400 As Integer
Dim n2700 As Integer
Dim n1800 As Integer
Dim n900 As Integer
Dim n450 As Integer
Dim n270 As Integer
Dim n180 As Integer
Dim n90 As Integer



lpuntal = Distancia - lfija
n5400 = Fix(lpuntal / 5400)
lpuntal = lpuntal - n5400 * 5400
n2700 = Fix(lpuntal / 2700)
lpuntal = lpuntal - n2700 * 2700
n1800 = Fix(lpuntal / 1800)
lpuntal = lpuntal - n1800 * 1800
n900 = Fix(lpuntal / 900)
lpuntal = lpuntal - n900 * 900
n450 = Fix(lpuntal / 450)
lpuntal = lpuntal - n450 * 450
n270 = Fix(lpuntal / 270)
lpuntal = lpuntal - n270 * 270
n180 = Fix(lpuntal / 180)
lpuntal = lpuntal - n180 * 180
n90 = Fix(lpuntal / 90)
lpuntal = lpuntal - n90 * 90


'''''' aquí OPTIMIZACIÓN

Dim D_P20 As Double
Dim D_P10 As Double
Dim Oprima(0 To 2) As Double
Dim mmm1 As Double, mmm2 As Double, modulacionmenorM1 As Double, modulacionmenorM2 As Double, modulacionmayorM1 As Double, modulacionmayorM2 As Double, mmmayor1 As Double, mmmayor2 As Double


If n90 > 0 Then
    lalt1 = Distancia - 90
    lalt2 = Distancia + 90
    
    If DirMuro1 = DirMuro2 Then
    Else
    ' Calculamos el punto intersección entre P1 y P2 con las direcciones del muro
        Oprima(0) = (P1(1) - P2(1) - Slope2 * P1(0) + Slope1 * P2(0)) / (Slope1 - Slope2)
        Oprima(1) = P2(1) + Slope1 * (Oprima(0) - P2(0))
        Oprima(2) = P2(2) ' Assuming the lines are in the same plane
        
        'Distancia entre P2 y la Oprima
        x1 = Oprima(0) - P2(0)
        y1 = Oprima(1) - P2(1)
        D_P20 = Val(Sqr((x1 ^ 2 + y1 ^ 2)))
        
        'Distancia entre P1 y la Oprima
        x2 = Oprima(0) - P1(0)
        y2 = Oprima(1) - P1(1)
        D_P10 = Val(Sqr((x2 ^ 2 + y2 ^ 2)))
        
        
        ' modulamos hacia la esquina con un puntal más pequeño
        mmm1 = (lalt1 * D_P20) / Distancia
        mmm2 = (lalt1 * D_P10) / Distancia
        
        modulacionmenorM1 = D_P20 - mmm1
        modulacionmenorM2 = D_P10 - mmm2
        
        ' modulamos con un puntal más grande
        mmmayor1 = (lalt2 * D_P20) / Distancia
        mmmayor2 = (lalt2 * D_P10) / Distancia
        
        modulacionmayorM1 = mmmayor1 - D_P20
        modulacionmayorM2 = mmmayor2 - D_P10
        
        If ((modulacionmenorM1 + modulacionmayorM1) < 1400) Or ((modulacionmenorM2 + modulacionmayorM2) < 1400) Then
            
            userInput = InputBox("Elija una de las siguientes opciones:" & vbCrLf & vbCrLf & vbCrLf & "1. Dibujar el puntal seleccionado de longitud " & Distancia & "." & vbCrLf & vbCrLf & vbCrLf & "2. Dibujar un puntal MENOR de " & lalt1 & "mm de longitud más cercano a la esquina" & vbCrLf & vbCrLf & vbCrLf & "3. Dibujar un puntal MAYOR de " & lalt2 & "mm de longitud más alejado de la esquina")

            If userInput = "1" Or userInput = "" Then
            ElseIf userInput = "2" Then
                ' modulamos si elegimos coger el puntal más pequeño
                PA(0) = PA(0) + modulacionmenorM1 * Cos(DirMuro1): PA(1) = PA(1) + modulacionmenorM1 * Sin(DirMuro1): PA(2) = PA(2)
                PB(0) = PB(0) + modulacionmenorM2 * Cos(DirMuro2): PB(1) = PB(1) + modulacionmenorM2 * Sin(DirMuro2): PB(2) = PB(2)
                
                P2(0) = P2(0) + modulacionmenorM1 * Cos(DirMuro1): P2(1) = P2(1) + modulacionmenorM1 * Sin(DirMuro1): P2(2) = P2(2)
                P1(0) = P1(0) + modulacionmenorM2 * Cos(DirMuro2): P1(1) = P1(1) + modulacionmenorM2 * Sin(DirMuro2): P1(2) = P1(2)
                
                n90 = n90 - 1
                
            ElseIf userInput = "3" Then
                ' modulamos si elegimos coger el puntal más grande
                PA(0) = PA(0) + modulacionmayorM1 * Cos(DirMuro1Inv): PA(1) = PA(1) + modulacionmayorM1 * Sin(DirMuro1Inv): PA(2) = PA(2)
                PB(0) = PB(0) + modulacionmayorM2 * Cos(DirMuro2Inv): PB(1) = PB(1) + modulacionmayorM2 * Sin(DirMuro2Inv): PB(2) = PB(2)
                
                P2(0) = P2(0) + modulacionmayorM1 * Cos(DirMuro1Inv): P2(1) = P2(1) + modulacionmayorM1 * Sin(DirMuro1Inv): P2(2) = P2(2)
                P1(0) = P1(0) + modulacionmayorM2 * Cos(DirMuro2Inv): P1(1) = P1(1) + modulacionmayorM2 * Sin(DirMuro2Inv): P1(2) = P1(2)
                
                n90 = n90 - 1
                n180 = n180 + 1
            End If
        Else
        End If

    End If
    

End If


' vamos a colocar la cuña/ angulo de giro en el extremo 1

    Set blockRef = gcadModel.InsertBlock(PA, rutaplaca2, Xs, Ys, Zs, DirMuro1)
    blockRef.Layer = "Mega"
'End If



' vamos a colocar la cuña/angulo de giro en el extremo 2

    Set blockRef = gcadModel.InsertBlock(PB, rutaplaca1, Xs, Ys, Zs, DirMuro2)
    blockRef.Layer = "Mega"



Set Eje1 = gcadModel.AddLine(P1, P2)
ANG = gcadUtil.AngleFromXAxis(P1, P2)
ANG2 = ANG + (PI / 2)



If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If

'Puntos centrales de las placas
PAP(0) = PA(0): PAP(1) = PA(1): PAP(2) = PA(2)
PBP(0) = PB(0): PBP(1) = PB(1): PBP(2) = PB(2)
x4 = PAP(0) - PBP(0)
y4 = PAP(1) - PBP(1)
D_ABP = Val(Sqr((x4 ^ 2 + y4 ^ 2)))

TxtPnt2(0) = PBP(0) + (D_ABP / 2) * Cos(ANG): TxtPnt2(1) = PBP(1) + (D_ABP / 2) * Sin(ANG): TxtPnt2(2) = PBP(2)
TxtPnt2(0) = TxtPnt2(0) + 860 * Cos(ANG2): TxtPnt2(1) = TxtPnt2(1) + 860 * Sin(ANG2): TxtPnt2(2) = TxtPnt2(2)

TxtPnt(0) = P1(0) + (Distancia / 2) * Cos(ANG): TxtPnt(1) = P1(1) + (Distancia / 2) * Sin(ANG): TxtPnt(2) = P1(2)
TxtPnt(0) = TxtPnt(0) + 410 * Cos(ANG2): TxtPnt(1) = TxtPnt(1) + 410 * Sin(ANG2): TxtPnt(2) = TxtPnt(2)

Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(P1, P2, TxtPnt)
objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objgcadDimAligned.StyleName = "MODELO"
objgcadDimAligned.TextStyle = "SIMPLEX"
objgcadDimAligned.VerticalTextPosition = acOutside
objgcadDimAligned.Update

Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(PBP, PAP, TxtPnt2)
objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objgcadDimAligned.StyleName = "MODELO"
objgcadDimAligned.TextStyle = "SIMPLEX"
objgcadDimAligned.VerticalTextPosition = acOutside
objgcadDimAligned.Update


Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
Punto_final(0) = P2(0): Punto_final(1) = P2(1): Punto_final(2) = P2(2)
M20x50_4 = ruta2 & "4-M20X50.dwg"
M20x60_4 = ruta2 & "4-M20X60.dwg"
'
' meter el ángulo de giro y los jackplates + el gato donde haga falta
' gato en el EXTREMO 1

Dim CuMP As String
Dim CuMPc As String


If placaanc1 = "MP" Then
    CuMP = rutacu & "PL_GCODAL_PLA.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, CuMP, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 315 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 315 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
ElseIf placaanc1 = "CompactaMP" Then
    CuMPc = rutacu & "PL_GCODAL_C_PLA.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, CuMPc, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 95 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 95 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Else
    Set blockRef = gcadModel.InsertBlock(Punto_inial, agiro, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
End If
    



''' aquí va el FUSIBLE
Dim mp_fus As String

Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x50_4, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
            blockRef.Explode
            blockRef.Delete
mp_fus = rutamp & "Mshor90" & plcu2 & "fusible.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_fus, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
            blockRef.Explode
            blockRef.Delete




Dim mp_90 As String
Dim mp_270 As String
Dim mp_900 As String
Dim mp_2700 As String
Dim mp_5400 As String
Dim mp_1800 As String
Dim mp_450 As String
Dim mp_180 As String

If n90 > 0 Then
        mp_90 = rutamp & "Mshor90" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n270 > 0 Then
    i = 0
    Do While i < n270
        mp_270 = rutamp & "Mshor270" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 270 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 270 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n900 > 0 Then
        mp_900 = rutamp & "Mshor900" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_900, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 900 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 900 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n2700 > 0 Then
        mp_2700 = rutamp & "Mshor2700" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_2700, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 2700 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 2700 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n5400 > 0 Then
    i = 0
    Do While i < n5400
        mp_5400 = rutamp & "Mshor5400" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_5400, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 5400 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 5400 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n1800 > 0 Then
    i = 0
    Do While i < n1800
        mp_1800 = rutamp & "Mshor1800" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_1800, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 1800 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 1800 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If


If n450 > 0 Then
        mp_450 = rutamp & "Mshor450" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 450 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 450 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If


If n180 > 0 Then
    i = 0
    Do While i < n180
        mp_180 = rutamp & "Mshor180" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 180 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 180 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If carga > 1020 Then
    blockRef.Delete
End If



' primer jack si hace falta

Dim MP_JP As String
Dim M20x110_4 As String
M20x110_4 = ruta2 & "4-M20X110.dwg"

If carga > 1020 Then
        MP_JP = rutamp & "MshorJACKPLATE.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_JP, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 40 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 40 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If
PGato2(0) = Punto_inial(0): PGato2(1) = Punto_inial(1): PGato2(2) = Punto_inial(2)

' base azul
Dim base_azul As String
base_azul = rutacu & "zMGBaseGato_azul.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, base_azul, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_inial(0) = Punto_inial(0) + 150 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 150 * Sin(ANG): Punto_inial(2) = Punto_inial(2)


' ahora vamos al P2 a meter lo que haga falta para respetar la apertura del gato y metemos en orden inverso lo demás
' angulito de giro o la terminación de las cuñas que hagan falta
If placaanc2 = "MP" Then
    CuMP = rutacu & "PL_GCODAL_PLA.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, CuMP, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 315 * Cos(ANG): Punto_final(1) = Punto_final(1) - 315 * Sin(ANG): Punto_final(2) = Punto_final(2)
ElseIf placaanc2 = "CompactaMP" Then
    CuMPc = rutacu & "PL_GCODAL_C_PLA.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, CuMPc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 95 * Cos(ANG): Punto_final(1) = Punto_final(1) - 95 * Sin(ANG): Punto_final(2) = Punto_final(2)
Else
    Set blockRef = gcadModel.InsertBlock(Punto_final, agiro, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 90 * Cos(ANG): Punto_final(1) = Punto_final(1) - 90 * Sin(ANG): Punto_final(2) = Punto_final(2)
End If
    

' segundo jack si hace falta
If carga > 1200 Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x110_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        Punto_final(0) = Punto_final(0) - 40 * Cos(ANG): Punto_final(1) = Punto_final(1) - 40 * Sin(ANG): Punto_final(2) = Punto_final(2)
        MP_JP = rutamp & "MshorJACKPLATE.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, MP_JP, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
End If

PGato1(0) = Punto_final(0): PGato1(1) = Punto_final(1): PGato1(2) = Punto_final(2)

' base naranja del gato
Dim base_naranja As String
base_naranja = rutacu & "zMGBaseGato_naranja.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, base_naranja, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Mega"
If carga < 1200 Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x60_4, Xs, Ys, Zs, ANG)
End If
blockRef.Layer = "Nonplot"
blockRef.Update
            blockRef.Explode
            blockRef.Delete
Punto_final(0) = Punto_final(0) - 150 * Cos(ANG): Punto_final(1) = Punto_final(1) - 150 * Sin(ANG): Punto_final(2) = Punto_final(2)



x1 = Punto_final(0) - Punto_inial(0)
y1 = Punto_final(1) - Punto_inial(1)

Distancia = Val(Sqr((x1 ^ 2 + y1 ^ 2)))
Distancia = Distancia / 2


Punto_inial(0) = Punto_inial(0) + Distancia * Cos(ANG): Punto_inial(1) = Punto_inial(1) + Distancia * Sin(ANG): Punto_inial(2) = Punto_inial(2)


' husillo
Dim husillo As String
husillo = rutacu & "MGHusilloGato.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, husillo, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"

x5 = PGato2(0) - PGato1(0)
y5 = PGato2(1) - PGato1(1)
D_Gato = Val(Sqr((x5 ^ 2 + y5 ^ 2)))

TxtPnt3(0) = PGato1(0) + (D_Gato / 2) * Cos(ANG): TxtPnt3(1) = PGato1(1) + (D_Gato / 2) * Sin(ANG): TxtPnt3(2) = PGato1(2)
TxtPnt3(0) = TxtPnt3(0) - 350 * Cos(ANG2): TxtPnt3(1) = TxtPnt3(1) - 350 * Sin(ANG2): TxtPnt3(2) = TxtPnt3(2)

Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(PGato1, PGato2, TxtPnt3)
objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objgcadDimAligned.StyleName = "MODELO"
objgcadDimAligned.TextStyle = "SIMPLEX"
objgcadDimAligned.VerticalTextPosition = acOutside
objgcadDimAligned.Update


Eje1.Layer = "Nonplot"
Loop

terminar:

End Sub

Sub mpn_homologado(carga As Variant)
Dim ruta As String, rutaps As String, rutapl As String, rutags As String
Dim ruta2 As String
Dim gcadDoc As Object
Dim M20x60_4 As String
Dim M20x50_4 As String
Dim gcadUtil As Object
Dim gcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim line2 As GcadLine
Dim line1 As GcadLine
Dim M20x90 As String
Dim M20x150 As String, M20x110 As String, Var20x250 As String
Dim M20x160 As String
Dim M20x90_16 As String
Dim GS_Bulon120mm As String
Dim GS_Giro As String
Dim GS_Fusible As String
Dim PS_280 As String
Dim PS_750 As String, PS_560 As String
Dim PS_1500 As String
Dim PS_3000 As String
Dim PS_4500 As String
Dim PS_6000 As String
Dim placaanc1 As String
Dim rutaplaca1 As String, rutacajon As String
Dim placaanc2 As String
Dim rutaplaca2 As String
Dim basecajon As String, brazocajon As String
Dim PS_Husillo As String
Dim PS_Placa50mm As String
Dim PS_Placa35mm As String
Dim zPS_Gato_Cono As String
Dim zPS_Gato_Tope As String
Dim PS_Gato As String
Dim lgiro As Double
Dim lfusible As Double
Dim l280 As Double
Dim l750 As Double, l560 As Double
Dim l1500 As Double
Dim l3000 As Double
Dim l4500 As Double
Dim l6000 As Double
Dim l50 As Double, l35 As Double
Dim l145 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lalt1 As Double
Dim lalt2 As Double
Dim lgatomin As Double
Dim lcajonmax As Double
Dim lcajonmin As Double
Dim n6000 As Integer
Dim n4500 As Integer
Dim OffsetDist As Double
Dim n3000 As Integer
Dim n1500 As Integer
Dim n750 As Integer, n560 As Integer
Dim n280 As Integer
Dim nfusible As Integer
Dim blockRef As Object
Dim repite As Double
Dim Punto_inial(0 To 2) As Double
Dim Punto_final(0 To 2) As Double
Dim Punto_inial2(0 To 2) As Double
Dim Punto_final2(0 To 2) As Double
Dim PI As Variant
Dim Eje1 As Object
Dim Eje2 As Object
Dim Eje3 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim ANG As Double
Dim ANG2 As Double
Dim DirBulon1 As Double
Dim DirBulon2 As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim dato1 As String
Dim plcu As String
Dim dato2 As String
Dim dato3 As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object
Dim puntoA As Variant
Dim puntop1 As Variant
Dim puntoB As Variant
Dim puntop2 As Variant
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim x3 As Double
Dim y3 As Double
Dim x4 As Double
Dim y4 As Double
Dim x5 As Double
Dim y5 As Double
Dim PGato1(0 To 2) As Double
Dim PGato2(0 To 2) As Double
Dim PA(0 To 2) As Double
Dim PAP(0 To 2) As Double
Dim PP1(0 To 2) As Double
Dim PP2(0 To 2) As Double
Dim PB(0 To 2) As Double
Dim PBP(0 To 2) As Double
Dim Esq(0 To 2) As Double
Dim userInput As String
Dim muro1menor  As Double
Dim muromoduladome1  As Double
Dim muromoduladoma1 As Double
Dim mod1c As Double
Dim mod1M As Double
Dim muro2menor As Double
Dim muromoduladome2 As Double
Dim muromoduladoma2 As Double
Dim mod2c As Double
Dim mod2M As Double
Dim objgcadDimAligned As GcadDimAligned
Dim TxtPnt(0 To 2) As Double
Dim TxtPnt2(0 To 2) As Double
Dim TxtPnt3(0 To 2) As Double
Dim Pcerca2(0 To 2) As Double
Dim Plejos2(0 To 2) As Double
Dim Pcerca1(0 To 2) As Double
Dim Plejos1(0 To 2) As Double
Dim Slope1 As Double
Dim Slope2 As Double
Dim D_A0 As Double
Dim D_B0 As Double
Dim D_AB As Double
Dim D_ABP As Double
Dim D_Gato As Double
Dim DirMuro1 As Double
Dim DirMuro2 As Double
Dim DirMuro1Inv As Double
Dim DirMuro2Inv As Double
Dim DirPuntal As Double
Dim rutacu As String
Dim DirPuntal2 As Double


Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

Ncapa = "Pipeshor4S"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5
Ncapa = "Granshor"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150

'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgiro = 205
lfusible = 187.5
l145 = 145
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
lcajonmin = 835
lcajonmax = 1022


On Error GoTo terminar
Dim rutamp As String

ruta = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\" & dato2 & "\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
rutacu = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"
rutacajon = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Cajon hidraulico\"

Dim plcue1 As String
Dim plcue2 As String
Dim plcu2 As String



kwordList = "Planta Alzado"
dato1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")
If dato1 = "Planta" Or dato1 = "" Then
dato1 = "Planta"
plcue1 = "Ialzado"
plcue2 = "Dalzado"
plcu2 = "PLA"
ElseIf dato1 = "Alzado" Then
plcue1 = "Ialzado"
plcue2 = "Dalzado"
plcu2 = "ALZ"
Else
GoTo terminar
End If



Do While repite = 1


lalt1 = 0
lalt2 = 0

'''' caso General
'' seleccionar los 4 puntos en sentido horario


puntoA = gcadUtil.GetPoint(, "punto inserción 1ª placa: ")
puntop1 = gcadUtil.GetPoint(puntoA, "punto direccional del muro 1 (convergente): ")

puntoB = gcadUtil.GetPoint(, "punto inserción 2ª placa: ")
puntop2 = gcadUtil.GetPoint(puntoB, "punto direccional del muro 2 (convergente): ")

'PA es el punto de inserción de la primera placa
PA(0) = puntoA(0): PA(1) = puntoA(1): PA(2) = puntoA(2)
PP1(0) = puntop1(0): PP1(1) = puntop1(1): PP1(2) = puntop1(2)

'PB es el punto de inserción de la segunda placa
PB(0) = puntoB(0): PB(1) = puntoB(1): PB(2) = puntoB(2)
PP2(0) = puntop2(0): PP2(1) = puntop2(1): PP2(2) = puntop2(2)

DirMuro1 = gcadUtil.AngleFromXAxis(PA, PP1)
DirMuro2 = gcadUtil.AngleFromXAxis(PB, PP2)
DirPuntal = gcadUtil.AngleFromXAxis(PA, PB)

' conseguir la esquina:
' Calculamos las direcciones de las rectas
Slope1 = Tan(DirMuro1)
Slope2 = Tan(DirMuro2)

If DirMuro1 = DirMuro2 Then
Else
' Calculamos el punto intersección
    Esq(0) = (PB(1) - PA(1) - Slope2 * PB(0) + Slope1 * PA(0)) / (Slope1 - Slope2)
    Esq(1) = PA(1) + Slope1 * (Esq(0) - PA(0))
    Esq(2) = PA(2) ' Assuming the lines are in the same plane
    
End If

DirBulon1 = DirMuro1 - (PI / 2)
DirBulon2 = DirMuro2 + (PI / 2)

'''''Extremo 1 del puntal
Dim AnguloAbsoluto2 As Double
Dim UmbralMin2 As Double
Dim UmbralMax2 As Double

UmbralMin2 = 80 * (PI / 180) ' Convertir 80 grados a radianes
UmbralMax2 = 100 * (PI / 180) ' Convertir 105 grados a radianes

AnguloAbsoluto2 = Abs(DirPuntal - DirMuro1)

If AnguloAbsoluto2 > UmbralMin2 And AnguloAbsoluto2 < UmbralMax2 Then
    ' No pasa nada
    P2(0) = PA(0) + 90 * Cos(DirBulon1): P2(1) = PA(1) + 90 * Sin(DirBulon1): P2(2) = PA(2)
Else
    kwordList = "Naranja Azul MP CompactaMP"
    placaanc2 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    placaanc2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de cuña en el primer extremo seleccionado: [Naranja/Azul/MP/CompactaMP]")
    
    ' aquí vienen los condicionales de las diferentes placas de anclaje, según cuál sea la seleccionada, moveremos para encontrar el P1 en este caso
    ' tenemos todas las medidas para P1
    
    If placaanc2 = "" Or placaanc2 = "Naranja" Then
        P2(0) = PA(0) + 239.2 * Cos(DirBulon1): P2(1) = PA(1) + 239.2 * Sin(DirBulon1): P2(2) = PA(2)
        P2(0) = P2(0) + 179.21 * Cos(DirMuro1): P2(1) = P2(1) + 179.21 * Sin(DirMuro1): P2(2) = P2(2)
    ElseIf placaanc2 = "Azul" Then
        P2(0) = PA(0) + 288.7 * Cos(DirBulon1): P2(1) = PA(1) + 288.7 * Sin(DirBulon1): P2(2) = PA(2)
        P2(0) = P2(0) + 213.7 * Cos(DirMuro1): P2(1) = P2(1) + 213.7 * Sin(DirMuro1): P2(2) = P2(2)
    ElseIf placaanc2 = "MP" Then
        P2(0) = PA(0) + 90 * Cos(DirBulon1): P2(1) = PA(1) + 90 * Sin(DirBulon1): P2(2) = PA(2)
    ElseIf placaanc2 = "CompactaMP" Then
        P2(0) = PA(0) + 85 * Cos(DirBulon1): P2(1) = PA(1) + 85 * Sin(DirBulon1): P2(2) = PA(2)
    End If
    
End If


'''''Extremo 2 del puntal
Dim AnguloAbsoluto1 As Double
Dim UmbralMin1 As Double
Dim UmbralMax1 As Double

UmbralMin1 = 80 * (PI / 180) ' Convertir 80 grados a radianes
UmbralMax1 = 100 * (PI / 180) ' Convertir 105 grados a radianes

AnguloAbsoluto1 = Abs((DirPuntal + PI) - DirMuro2)

If AnguloAbsoluto1 > UmbralMin1 And AnguloAbsoluto1 < UmbralMax1 Then
    P1(0) = PB(0) + 90 * Cos(DirBulon2): P1(1) = PB(1) + 90 * Sin(DirBulon2): P1(2) = PB(2)
Else
    kwordList = "Naranja Azul MP CompactaMP"
    placaanc1 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    placaanc1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de cuña en el segundo extremo seleccionado: [Naranja/Azul/MP/CompactaMP]")

    ' aquí vienen los condicionales de las diferentes placas de anclaje, según cuál sea la seleccionada, moveremos para encontrar el P1 en este caso
    ' tenemos todas las medidas para P1
    
    If placaanc1 = "" Or placaanc1 = "Naranja" Then
        P1(0) = PB(0) + 239.2 * Cos(DirBulon2): P1(1) = PB(1) + 239.2 * Sin(DirBulon2): P1(2) = PB(2)
        P1(0) = P1(0) + 179.21 * Cos(DirMuro2): P1(1) = P1(1) + 179.21 * Sin(DirMuro2): P1(2) = P1(2)
    ElseIf placaanc1 = "Azul" Then
        P1(0) = PB(0) + 288.7 * Cos(DirBulon2): P1(1) = PB(1) + 288.7 * Sin(DirBulon2): P1(2) = PB(2)
        P1(0) = P1(0) + 213.7 * Cos(DirMuro2): P1(1) = P1(1) + 213.7 * Sin(DirMuro2): P1(2) = P1(2)
    ElseIf placaanc1 = "MP" Then
        P1(0) = PB(0) + 90 * Cos(DirBulon2): P1(1) = PB(1) + 90 * Sin(DirBulon2): P1(2) = PB(2)
    ElseIf placaanc1 = "CompactaMP" Then
        P1(0) = PB(0) + 85 * Cos(DirBulon2): P1(1) = PB(1) + 85 * Sin(DirBulon2): P1(2) = PB(2)
    End If
        
End If




DirPuntal = gcadUtil.AngleFromXAxis(P1, P2)
DirPuntal2 = gcadUtil.AngleFromXAxis(P2, P1)

'''' podemos también añadir las rutas de la placa1 y placa2 como hago aquí abajo pero lo metemos directamente en el condicional que tenemos aquí arriba, para dejar ya cerrada cuál va a ser cada una de las placas
'''' además de añadir las capas. En caso de que sea el ángulo de giro podemos añadirlo también.


If placaanc1 = "Naranja" Or placaanc1 = "" Then
    rutaplaca1 = rutacu & "MG_CunaNar_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "Azul" Then
    rutaplaca1 = rutacu & "MG_CunaAz_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "MP" Then
    rutaplaca1 = rutacu & "PlacaMP_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "CompactaMP" Then
    rutaplaca1 = rutacu & "PlacaMP_C_" & plcue2 & ".dwg"
    capa = "Mega"
End If

If placaanc2 = "Naranja" Or placaanc2 = "" Then
    rutaplaca2 = rutacu & "MG_CunaNar_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "Azul" Then
    rutaplaca2 = rutacu & "MG_CunaAz_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "MP" Then
    rutaplaca2 = rutacu & "PlacaMP_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "CompactaMP" Then
    rutaplaca2 = rutacu & "PlacaMP_C_" & plcue1 & ".dwg"
    capa = "Mega"
End If


Dim agiro As String
agiro = rutacu & "ANGgiro.dwg"

DirMuro1Inv = DirMuro1 - PI
DirMuro2Inv = DirMuro2 - PI

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' Aquí viene la toma de decisiones de si coger un puntal o coger otro::::
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If carga <= 1020 Then
    lfija = 680
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
ElseIf carga > 1020 And carga <= 1200 Then
    lfija = 720
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
ElseIf carga > 1200 Then
    lfija = 760
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
End If

Xs = 1
Ys = 1
Zs = 1


x = P2(0) - P1(0)
y = P2(1) - P1(1)

Distancia = Val(Sqr((x ^ 2 + y ^ 2)))


Dim n5400 As Integer
Dim n2700 As Integer
Dim n1800 As Integer
Dim n900 As Integer
Dim n450 As Integer
Dim n270 As Integer
Dim n180 As Integer
Dim n90 As Integer



lpuntal = Distancia - lfija
n5400 = Fix(lpuntal / 5400)
lpuntal = lpuntal - n5400 * 5400
n2700 = Fix(lpuntal / 2700)
lpuntal = lpuntal - n2700 * 2700
n1800 = Fix(lpuntal / 1800)
lpuntal = lpuntal - n1800 * 1800
n900 = Fix(lpuntal / 900)
lpuntal = lpuntal - n900 * 900
n450 = Fix(lpuntal / 450)
lpuntal = lpuntal - n450 * 450
n270 = Fix(lpuntal / 270)
lpuntal = lpuntal - n270 * 270
n180 = Fix(lpuntal / 180)
lpuntal = lpuntal - n180 * 180
n90 = Fix(lpuntal / 90)
lpuntal = lpuntal - n90 * 90


'''''' aquí OPTIMIZACIÓN

Dim D_P20 As Double
Dim D_P10 As Double
Dim Oprima(0 To 2) As Double
Dim mmm1 As Double, mmm2 As Double, modulacionmenorM1 As Double, modulacionmenorM2 As Double, modulacionmayorM1 As Double, modulacionmayorM2 As Double, mmmayor1 As Double, mmmayor2 As Double


If n90 > 0 Then
    lalt1 = Distancia - 90
    lalt2 = Distancia + 90
    
    If DirMuro1 = DirMuro2 Then
    Else
    ' Calculamos el punto intersección entre P1 y P2 con las direcciones del muro
        Oprima(0) = (P1(1) - P2(1) - Slope2 * P1(0) + Slope1 * P2(0)) / (Slope1 - Slope2)
        Oprima(1) = P2(1) + Slope1 * (Oprima(0) - P2(0))
        Oprima(2) = P2(2) ' Assuming the lines are in the same plane
        
        'Distancia entre P2 y la Oprima
        x1 = Oprima(0) - P2(0)
        y1 = Oprima(1) - P2(1)
        D_P20 = Val(Sqr((x1 ^ 2 + y1 ^ 2)))
        
        'Distancia entre P1 y la Oprima
        x2 = Oprima(0) - P1(0)
        y2 = Oprima(1) - P1(1)
        D_P10 = Val(Sqr((x2 ^ 2 + y2 ^ 2)))
        
        
        ' modulamos hacia la esquina con un puntal más pequeño
        mmm1 = (lalt1 * D_P20) / Distancia
        mmm2 = (lalt1 * D_P10) / Distancia
        
        modulacionmenorM1 = D_P20 - mmm1
        modulacionmenorM2 = D_P10 - mmm2
        
        ' modulamos con un puntal más grande
        mmmayor1 = (lalt2 * D_P20) / Distancia
        mmmayor2 = (lalt2 * D_P10) / Distancia
        
        modulacionmayorM1 = mmmayor1 - D_P20
        modulacionmayorM2 = mmmayor2 - D_P10
        
        If ((modulacionmenorM1 + modulacionmayorM1) < 1400) Or ((modulacionmenorM2 + modulacionmayorM2) < 1400) Then
            
            userInput = InputBox("Elija una de las siguientes opciones:" & vbCrLf & vbCrLf & vbCrLf & "1. Dibujar el puntal seleccionado de longitud " & Distancia & "." & vbCrLf & vbCrLf & vbCrLf & "2. Dibujar un puntal MENOR de " & lalt1 & "mm de longitud más cercano a la esquina" & vbCrLf & vbCrLf & vbCrLf & "3. Dibujar un puntal MAYOR de " & lalt2 & "mm de longitud más alejado de la esquina")

            If userInput = "1" Or userInput = "" Then
            ElseIf userInput = "2" Then
                ' modulamos si elegimos coger el puntal más pequeño
                PA(0) = PA(0) + modulacionmenorM1 * Cos(DirMuro1): PA(1) = PA(1) + modulacionmenorM1 * Sin(DirMuro1): PA(2) = PA(2)
                PB(0) = PB(0) + modulacionmenorM2 * Cos(DirMuro2): PB(1) = PB(1) + modulacionmenorM2 * Sin(DirMuro2): PB(2) = PB(2)
                
                P2(0) = P2(0) + modulacionmenorM1 * Cos(DirMuro1): P2(1) = P2(1) + modulacionmenorM1 * Sin(DirMuro1): P2(2) = P2(2)
                P1(0) = P1(0) + modulacionmenorM2 * Cos(DirMuro2): P1(1) = P1(1) + modulacionmenorM2 * Sin(DirMuro2): P1(2) = P1(2)
                
                n90 = n90 - 1
                
            ElseIf userInput = "3" Then
                ' modulamos si elegimos coger el puntal más grande
                PA(0) = PA(0) + modulacionmayorM1 * Cos(DirMuro1Inv): PA(1) = PA(1) + modulacionmayorM1 * Sin(DirMuro1Inv): PA(2) = PA(2)
                PB(0) = PB(0) + modulacionmayorM2 * Cos(DirMuro2Inv): PB(1) = PB(1) + modulacionmayorM2 * Sin(DirMuro2Inv): PB(2) = PB(2)
                
                P2(0) = P2(0) + modulacionmayorM1 * Cos(DirMuro1Inv): P2(1) = P2(1) + modulacionmayorM1 * Sin(DirMuro1Inv): P2(2) = P2(2)
                P1(0) = P1(0) + modulacionmayorM2 * Cos(DirMuro2Inv): P1(1) = P1(1) + modulacionmayorM2 * Sin(DirMuro2Inv): P1(2) = P1(2)
                
                n90 = n90 - 1
                n180 = n180 + 1
            End If
        Else
        End If

    End If
    

End If


' vamos a colocar la cuña/ angulo de giro en el extremo 1


If AnguloAbsoluto2 > UmbralMin2 And AnguloAbsoluto2 < UmbralMax2 Then
     ' insertamos el ángulo de giro
     Set blockRef = gcadModel.InsertBlock(P2, agiro, Xs, Ys, Zs, DirMuro1Inv - PI / 2)
     blockRef.Layer = "Mega"
Else
    ' insertamos la cuña que haga falta
    Set blockRef = gcadModel.InsertBlock(PA, rutaplaca2, Xs, Ys, Zs, DirMuro1)
    blockRef.Layer = "Mega"
End If



' vamos a colocar la cuña/angulo de giro en el extremo 2
'AnguloAbsoluto1 = Abs(DirPuntal - DirMuro2)

If AnguloAbsoluto1 > UmbralMin1 And AnguloAbsoluto1 < UmbralMax1 Then
    ' insertamos el ángulo de giro
    Set blockRef = gcadModel.InsertBlock(P1, agiro, Xs, Ys, Zs, DirMuro2Inv + PI / 2)
    blockRef.Layer = "Mega"
Else
    ' insertamos la cuña que haga falta
    Set blockRef = gcadModel.InsertBlock(PB, rutaplaca1, Xs, Ys, Zs, DirMuro2)
    blockRef.Layer = "Mega"
End If


Set Eje1 = gcadModel.AddLine(P1, P2)
ANG = gcadUtil.AngleFromXAxis(P1, P2)
ANG2 = ANG + (PI / 2)



If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If

'Puntos centrales de las placas
PAP(0) = PA(0): PAP(1) = PA(1): PAP(2) = PA(2)
PBP(0) = PB(0): PBP(1) = PB(1): PBP(2) = PB(2)
x4 = PAP(0) - PBP(0)
y4 = PAP(1) - PBP(1)
D_ABP = Val(Sqr((x4 ^ 2 + y4 ^ 2)))

TxtPnt2(0) = PBP(0) + (D_ABP / 2) * Cos(ANG): TxtPnt2(1) = PBP(1) + (D_ABP / 2) * Sin(ANG): TxtPnt2(2) = PBP(2)
TxtPnt2(0) = TxtPnt2(0) + 860 * Cos(ANG2): TxtPnt2(1) = TxtPnt2(1) + 860 * Sin(ANG2): TxtPnt2(2) = TxtPnt2(2)

TxtPnt(0) = P1(0) + (Distancia / 2) * Cos(ANG): TxtPnt(1) = P1(1) + (Distancia / 2) * Sin(ANG): TxtPnt(2) = P1(2)
TxtPnt(0) = TxtPnt(0) + 410 * Cos(ANG2): TxtPnt(1) = TxtPnt(1) + 410 * Sin(ANG2): TxtPnt(2) = TxtPnt(2)

Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(P1, P2, TxtPnt)
objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objgcadDimAligned.StyleName = "MODELO"
objgcadDimAligned.TextStyle = "SIMPLEX"
objgcadDimAligned.VerticalTextPosition = acOutside
objgcadDimAligned.Update

Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(PBP, PAP, TxtPnt2)
objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objgcadDimAligned.StyleName = "MODELO"
objgcadDimAligned.TextStyle = "SIMPLEX"
objgcadDimAligned.VerticalTextPosition = acOutside
objgcadDimAligned.Update


Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
Punto_final(0) = P2(0): Punto_final(1) = P2(1): Punto_final(2) = P2(2)
M20x50_4 = ruta2 & "4-M20X50.dwg"
M20x60_4 = ruta2 & "4-M20X60.dwg"
'
' meter el ángulo de giro y los jackplates + el gato donde haga falta
' gato en el EXTREMO 1

Dim CuMP As String
Dim CuMPc As String


If placaanc1 = "MP" Then
    CuMP = rutacu & "PL_GCODAL_PLA.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, CuMP, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 315 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 315 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
ElseIf placaanc1 = "CompactaMP" Then
    CuMPc = rutacu & "PL_GCODAL_C_PLA.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, CuMPc, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 95 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 95 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Else
    Set blockRef = gcadModel.InsertBlock(Punto_inial, agiro, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
End If
    




''' aquí va el FUSIBLE
Dim mp_fus As String

Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x50_4, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
            blockRef.Explode
            blockRef.Delete
mp_fus = rutamp & "Mshor90" & plcu2 & "fusible.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_fus, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
blockRef.Update
            blockRef.Explode
            blockRef.Delete



Dim mp_90 As String
Dim mp_270 As String
Dim mp_900 As String
Dim mp_2700 As String
Dim mp_5400 As String
Dim mp_1800 As String
Dim mp_450 As String
Dim mp_180 As String

If n90 > 0 Then
        mp_90 = rutamp & "Mshor90" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n270 > 0 Then
    i = 0
    Do While i < n270
        mp_270 = rutamp & "Mshor270" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 270 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 270 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n900 > 0 Then
        mp_900 = rutamp & "Mshor900" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_900, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 900 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 900 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n2700 > 0 Then
        mp_2700 = rutamp & "Mshor2700" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_2700, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 2700 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 2700 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n5400 > 0 Then
    i = 0
    Do While i < n5400
        mp_5400 = rutamp & "Mshor5400" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_5400, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 5400 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 5400 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n1800 > 0 Then
    i = 0
    Do While i < n1800
        mp_1800 = rutamp & "Mshor1800" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_1800, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 1800 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 1800 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If


If n450 > 0 Then
        mp_450 = rutamp & "Mshor450" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 450 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 450 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If


If n180 > 0 Then
    i = 0
    Do While i < n180
        mp_180 = rutamp & "Mshor180" & plcu2 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 180 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 180 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If carga > 1020 Then
    blockRef.Delete
End If

' primer jack si hace falta

Dim MP_JP As String
Dim M20x110_4 As String
M20x110_4 = ruta2 & "4-M20X110.dwg"

If carga > 1020 Then
        MP_JP = rutamp & "MshorJACKPLATE.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, MP_JP, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 40 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 40 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

' base azul
Dim base_azul As String
base_azul = rutacu & "zMGBaseGato_azul.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, base_azul, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_inial(0) = Punto_inial(0) + 150 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 150 * Sin(ANG): Punto_inial(2) = Punto_inial(2)


' ahora vamos al P2 a meter lo que haga falta para respetar la apertura del gato y metemos en orden inverso lo demás
' angulito de giro o la terminación de las cuñas que hagan falta
If placaanc2 = "MP" Then
    CuMP = rutacu & "PL_GCODAL_PLA.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, CuMP, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 315 * Cos(ANG): Punto_final(1) = Punto_final(1) - 315 * Sin(ANG): Punto_final(2) = Punto_final(2)
ElseIf placaanc2 = "CompactaMP" Then
    CuMPc = rutacu & "PL_GCODAL_C_PLA.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, CuMPc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 95 * Cos(ANG): Punto_final(1) = Punto_final(1) - 95 * Sin(ANG): Punto_final(2) = Punto_final(2)
Else
    Set blockRef = gcadModel.InsertBlock(Punto_final, agiro, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 90 * Cos(ANG): Punto_final(1) = Punto_final(1) - 90 * Sin(ANG): Punto_final(2) = Punto_final(2)
End If
    

' segundo jack si hace falta
If carga > 1200 Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x110_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        Punto_final(0) = Punto_final(0) - 40 * Cos(ANG): Punto_final(1) = Punto_final(1) - 40 * Sin(ANG): Punto_final(2) = Punto_final(2)
        MP_JP = rutamp & "MshorJACKPLATE.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, MP_JP, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
End If

' base naranja del gato
Dim base_naranja As String
base_naranja = rutacu & "zMGBaseGato_naranja.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, base_naranja, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Mega"
If carga < 1200 Then
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x60_4, Xs, Ys, Zs, ANG)
End If
blockRef.Layer = "Nonplot"
blockRef.Update
            blockRef.Explode
            blockRef.Delete
Punto_final(0) = Punto_final(0) - 150 * Cos(ANG): Punto_final(1) = Punto_final(1) - 150 * Sin(ANG): Punto_final(2) = Punto_final(2)



x1 = Punto_final(0) - Punto_inial(0)
y1 = Punto_final(1) - Punto_inial(1)

Distancia = Val(Sqr((x1 ^ 2 + y1 ^ 2)))
Distancia = Distancia / 2

Punto_inial(0) = Punto_inial(0) + Distancia * Cos(ANG): Punto_inial(1) = Punto_inial(1) + Distancia * Sin(ANG): Punto_inial(2) = Punto_inial(2)


' husillo
Dim husillo As String
husillo = rutacu & "MGHusilloGato.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, husillo, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"






Eje1.Layer = "Nonplot"
Loop

terminar:

End Sub



Sub psn()
Dim ruta As String, rutaps As String, rutapl As String, rutags As String
Dim ruta2 As String
Dim gcadDoc As Object
Dim gcadUtil As Object
Dim gcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim line2 As GcadLine
Dim line1 As GcadLine
Dim M20x90 As String
Dim M20x150 As String, M20x110 As String, Var20x250 As String
Dim M20x160 As String
Dim M20x90_16 As String
Dim GS_Bulon120mm As String
Dim GS_Giro As String
Dim GS_Fusible As String
Dim PS_280 As String
Dim PS_750 As String, PS_560 As String
Dim PS_1500 As String
Dim PS_3000 As String
Dim PS_4500 As String
Dim PS_6000 As String
Dim placaanc1 As String
Dim rutaplaca1 As String, rutacajon As String
Dim placaanc2 As String
Dim rutaplaca2 As String
Dim basecajon As String, brazocajon As String
Dim PS_Husillo As String
Dim PS_Placa50mm As String
Dim PS_Placa35mm As String
Dim zPS_Gato_Cono As String
Dim zPS_Gato_Tope As String
Dim PS_Gato As String
Dim lgiro As Double
Dim lfusible As Double
Dim l280 As Double
Dim l750 As Double, l560 As Double
Dim l1500 As Double
Dim l3000 As Double
Dim l4500 As Double
Dim l6000 As Double
Dim l50 As Double, l35 As Double
Dim l145 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lalt1 As Double
Dim lalt2 As Double
Dim lgatomin As Double
Dim lcajonmax As Double
Dim lcajonmin As Double
Dim n6000 As Integer
Dim n4500 As Integer
Dim OffsetDist As Double
Dim n3000 As Integer
Dim n1500 As Integer
Dim n750 As Integer, n560 As Integer
Dim n280 As Integer
Dim nfusible As Integer
Dim blockRef As Object
Dim repite As Double
Dim Punto_inial(0 To 2) As Double
Dim Punto_final(0 To 2) As Double
Dim Punto_inial2(0 To 2) As Double
Dim Punto_final2(0 To 2) As Double
Dim PI As Variant
Dim Eje1 As Object
Dim Eje2 As Object
Dim Eje3 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim ANG As Double
Dim ANG2 As Double
Dim DirBulon1 As Double
Dim DirBulon2 As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim dato1 As String
Dim dato2 As String
Dim dato3 As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object
Dim puntoA As Variant
Dim puntop1 As Variant
Dim puntoB As Variant
Dim puntop2 As Variant
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim x3 As Double
Dim y3 As Double
Dim x4 As Double
Dim y4 As Double
Dim x5 As Double
Dim y5 As Double
Dim PGato1(0 To 2) As Double
Dim PGato2(0 To 2) As Double
Dim PA(0 To 2) As Double
Dim PAP(0 To 2) As Double
Dim PP1(0 To 2) As Double
Dim PP2(0 To 2) As Double
Dim PB(0 To 2) As Double
Dim PBP(0 To 2) As Double
Dim Esq(0 To 2) As Double
Dim userInput As String
Dim muro1menor  As Double
Dim muromoduladome1  As Double
Dim muromoduladoma1 As Double
Dim mod1c As Double
Dim mod1M As Double
Dim muro2menor As Double
Dim muromoduladome2 As Double
Dim muromoduladoma2 As Double
Dim mod2c As Double
Dim mod2M As Double
Dim objgcadDimAligned As GcadDimAligned
Dim TxtPnt(0 To 2) As Double
Dim TxtPnt2(0 To 2) As Double
Dim TxtPnt3(0 To 2) As Double
Dim Pcerca2(0 To 2) As Double
Dim Plejos2(0 To 2) As Double
Dim Pcerca1(0 To 2) As Double
Dim Plejos1(0 To 2) As Double
Dim Slope1 As Double
Dim Slope2 As Double
Dim D_A0 As Double
Dim D_B0 As Double
Dim D_AB As Double
Dim D_ABP As Double
Dim D_Gato As Double
Dim DirMuro1 As Double
Dim DirMuro2 As Double
Dim DirMuro1Inv As Double
Dim DirMuro2Inv As Double
Dim DirPuntal As Double
Dim DirPuntal2 As Double

Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

Ncapa = "Pipeshor4S"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5
Ncapa = "Granshor"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150

'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgiro = 205
lfusible = 187.5
l145 = 145
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
lcajonmin = 835
lcajonmax = 1022

Dim carga As Double

carga = InputBox("Introduce carga soportada por el codal (kN ELU): ", "Carga", 0)

If (carga < 1350) Then
    Call mpn(carga)
    GoTo terminar
ElseIf (carga >= 1350) And (carga < 1500) Then
    lfija = (2 * lgiro) + lfusible + l50 + lgatomin
ElseIf (carga >= 1500) And (carga < 2000) Then
    lfija = (2 * lgiro) + lfusible + l50 + l35 + lgatomin
ElseIf (carga >= 2000) And (carga < 2900) Then
    lfija = (2 * lgiro) + lfusible + (2 * l50) + lgatomin
ElseIf carga >= 2900 Then
    lfija = (2 * lgiro) + lfusible + (5 * l50) + lcajonmin
ElseIf carga = "" Then
    MsgBox "Introduce una carga para continuar"
End If

On Error GoTo terminar

kwordList = "S L"
dato2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato2 = "Pshor_4" & ThisDrawing.Utility.GetKeyword(vbLf & "Introduce PS4S o PS4L: [S/L]")

If dato2 = "Pshor_4L" Then
capa = "Pipeshor4L"
dato3 = "PL"
ElseIf dato2 = "Pshor_4S" Or dato2 = "Pshor_4" Then
dato2 = "Pshor_4S"
capa = "Pipeshor4S"
dato3 = "PS"
Else
GoTo terminar
End If

ruta = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\" & dato2 & "\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"
rutacajon = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Cajon hidraulico\"

kwordList = "Planta Alzado"
dato1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")
If dato1 = "" Or dato1 = "Planta" Then
dato1 = "Planta"
ElseIf dato1 = "Alzado" Then
Else
GoTo terminar
End If


''''''' variables necesarias por meter aquiiiii ?????????????)&/(%/(%/(%/
kwordList = "Granshor Compacta"
placaanc1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
placaanc1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de placa en el primer extremo seleccionado: [Granshor/Compacta]")

kwordList = "Granshor Compacta"
placaanc2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
placaanc2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de placa en el segundo extremo seleccionado: [Granshor/Compacta]")

If placaanc1 = "" Or placaanc1 = "Granshor" Then
    rutaplaca1 = rutags & "GS_PlacaAnclaje_Ialzado.dwg"
    capa = "Granshor"
ElseIf placaanc1 = "Compacta" Then
    rutaplaca1 = rutags & "GS_Placacompacta_Ialzado.dwg"
    capa = "Granshor"
End If

If placaanc2 = "" Or placaanc2 = "Granshor" Then
    rutaplaca2 = rutags & "GS_PlacaAnclaje_Dalzado.dwg"
    capa = "Granshor"
ElseIf placaanc2 = "Compacta" Then
    rutaplaca2 = rutags & "GS_Placacompacta_Dalzado.dwg"
    capa = "Granshor"
End If



Do While repite = 1


lalt1 = 0
lalt2 = 0

'''' caso General
'' seleccionar los 4 puntos en sentido horario


puntoA = gcadUtil.GetPoint(, "punto inserción 1ª placa: ")
puntop1 = gcadUtil.GetPoint(puntoA, "punto direccional del muro 1 (convergente): ")

puntoB = gcadUtil.GetPoint(, "punto inserción 2ª placa: ")
puntop2 = gcadUtil.GetPoint(puntoB, "punto direccional del muro 2 (convergente): ")

'PA es el punto de inserción de la primera placa
PA(0) = puntoA(0): PA(1) = puntoA(1): PA(2) = puntoA(2)
PP1(0) = puntop1(0): PP1(1) = puntop1(1): PP1(2) = puntop1(2)

'PB es el punto de inserción de la segunda placa
PB(0) = puntoB(0): PB(1) = puntoB(1): PB(2) = puntoB(2)
PP2(0) = puntop2(0): PP2(1) = puntop2(1): PP2(2) = puntop2(2)

DirMuro1 = gcadUtil.AngleFromXAxis(PA, PP1)
DirMuro2 = gcadUtil.AngleFromXAxis(PB, PP2)
DirPuntal = gcadUtil.AngleFromXAxis(PA, PB)

' conseguir la esquina:
' Calculamos las direcciones de las rectas
Slope1 = Tan(DirMuro1)
Slope2 = Tan(DirMuro2)

If DirMuro1 = DirMuro2 Then
Else
' Calculamos el punto intersección
    Esq(0) = (PB(1) - PA(1) - Slope2 * PB(0) + Slope1 * PA(0)) / (Slope1 - Slope2)
    Esq(1) = PA(1) + Slope1 * (Esq(0) - PA(0))
    Esq(2) = PA(2) ' Assuming the lines are in the same plane
    
    'Distancia entre A y la esquina
    x1 = Esq(0) - PA(0)
    y1 = Esq(1) - PA(1)
    D_A0 = Val(Sqr((x1 ^ 2 + y1 ^ 2)))
    
    'Distancia entre B y la esquina
    x2 = Esq(0) - PB(0)
    y2 = Esq(1) - PB(1)
    D_B0 = Val(Sqr((x2 ^ 2 + y2 ^ 2)))
    
    'Distacia entre A y B
    x3 = PA(0) - PB(0)
    y3 = PA(1) - PB(1)
    D_AB = Val(Sqr((x3 ^ 2 + y3 ^ 2)))
End If

''''' NUEVO


If Abs(DirMuro2 - DirMuro1) > PI Then
    If DirMuro2 > DirMuro1 Then
        ' Intercambiar los puntos P1 y P3
        Dim tempP0(0 To 2) As Double
        tempP0(0) = PA(0): tempP0(1) = PA(1): tempP0(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = tempP0(0): PB(1) = tempP0(1): PB(2) = tempP0(2):
        ' Recalcular la dirección del muro 1 y perpendicular al muro
        DirMuro1 = gcadUtil.AngleFromXAxis(PA, Esq)
        ' Recalcular la dirección del muro 2 y perpendicular al muro
        DirMuro2 = gcadUtil.AngleFromXAxis(PB, Esq)
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
    End If
Else
    If DirMuro2 < DirMuro1 Then
        ' Intercambiar los puntos P1 y P3
        Dim tempP(0 To 2) As Double
        tempP(0) = PA(0): tempP(1) = PA(1): tempP(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = tempP(0): PB(1) = tempP(1): PB(2) = tempP(2):
        ' Recalcular la dirección del muro 1 y perpendicular al muro
        DirMuro1 = gcadUtil.AngleFromXAxis(PA, Esq)
        ' Recalcular la dirección del muro 2 y perpendicular al muro
        DirMuro2 = gcadUtil.AngleFromXAxis(PB, Esq)
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
    End If
End If

'''' FIN DE NUEVO


DirBulon1 = DirMuro1 - (PI / 2)
DirBulon2 = DirMuro2 + (PI / 2)

P2(0) = PA(0) + l145 * Cos(DirBulon1): P2(1) = PA(1) + l145 * Sin(DirBulon1): P2(2) = PA(2)
P1(0) = PB(0) + l145 * Cos(DirBulon2): P1(1) = PB(1) + l145 * Sin(DirBulon2): P1(2) = PB(2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

DirMuro1Inv = DirMuro1 - PI
DirMuro2Inv = DirMuro2 - PI

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' Aquí viene la toma de decisiones de si coger un puntal o coger otro::::
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If DirMuro1 = DirMuro2 Then
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
            If dato2 = "Pshor_4L" Then
            n280 = 2
            n560 = 0
            ElseIf dato2 = "Pshor_4S" Then
            n280 = 0
            n560 = 1
            End If
        Case Else
        MsgBox "Longitud no controlada " & lpuntal & "mm, fuera de rango, revisar código"
        GoTo terminar
            
    End Select
    DirPuntal = gcadUtil.AngleFromXAxis(P1, P2)
    DirPuntal2 = gcadUtil.AngleFromXAxis(P2, P1)
    If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
        rutaplaca1 = rutaplaca1
        Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
        blockRef.Layer = "Granshor"
    Else
        rutaplaca1 = rutaplaca1
        If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
            rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
        ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
            rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
        End If
        Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
        blockRef.Layer = "Granshor"
    End If
    If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
        rutaplaca2 = rutaplaca2
        Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
        blockRef.Layer = "Granshor"
    Else
        rutaplaca2 = rutaplaca2
        If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
            rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
        ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
            rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
        End If
        Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
        blockRef.Layer = "Granshor"
    End If
Else
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
            If dato2 = "Pshor_4L" Then
            n280 = 2
            n560 = 0
            ElseIf dato2 = "Pshor_4S" Then
            n280 = 0
            n560 = 1
            End If
        Case Else
        MsgBox "Longitud no controlada " & lpuntal & "mm, fuera de rango, revisar código"
        GoTo terminar
            
    End Select



    If dato3 = "PL" Then
        If nfusible = 1 Then
            If n280 = 1 Then
                lalt1 = Distancia - l280
                lalt2 = Distancia + 470
            ElseIf n280 = 2 Then
                lalt1 = Distancia - l280
                lalt2 = Distancia + 190
            End If
        ElseIf nfusible = 2 Then
            If n280 = 1 Then
                lalt1 = Distancia - lfusible - l280 + 150
                lalt2 = Distancia - l280 - lfusible + l750
            ElseIf n280 = 2 Then
                lalt1 = Distancia - lfusible - l280 + 150
                lalt2 = Distancia - 560 - lfusible + l750
            End If
        End If
    ElseIf dato3 = "PS" Then
        If n280 = 1 Then
            If n560 = 1 And n750 = 1 Then
                lalt1 = Distancia - l280 + 190
                lalt2 = Distancia + l280
            ElseIf n560 = 0 And n750 = 1 Then
                lalt1 = Distancia - l280
                lalt2 = Distancia + l280
                
            ElseIf n560 = 1 And n750 = 0 Then
                lalt1 = Distancia - 90
                lalt2 = Distancia + 280
            Else
                lalt1 = Distancia - 280
                lalt2 = Distancia + 280
            End If
        End If
    End If


    '''''''''''''''''''''Cálculo de las posibles modulaciones''''''''''''''''''''''''''''''''''


    '''' MURO 1 '''''''
    muro1menor = (Distancia * D_A0) / D_AB
    muromoduladome1 = (lalt1 * D_A0) / D_AB
    muromoduladoma1 = (lalt2 * D_A0) / D_AB

    mod1c = muro1menor - muromoduladome1
    mod1M = muromoduladoma1 - muro1menor

    Pcerca2(0) = P2(0) + mod1c * Cos(DirMuro1): Pcerca2(1) = P2(1) + mod1c * Sin(DirMuro1): Pcerca2(2) = P2(2)
    Plejos2(0) = P2(0) - mod1M * Cos(DirMuro1): Plejos2(1) = P2(1) - mod1M * Sin(DirMuro1): Plejos2(2) = P2(2)

    '''''' MURO 2 ''''''''''
    muro2menor = (Distancia * D_B0) / D_AB
    muromoduladome2 = (lalt1 * D_B0) / D_AB
    muromoduladoma2 = (lalt2 * D_B0) / D_AB

    mod2c = muro2menor - muromoduladome2
    mod2M = muromoduladoma2 - muro2menor

    Pcerca1(0) = P1(0) + mod2c * Cos(DirMuro2): Pcerca1(1) = P1(1) + mod2c * Sin(DirMuro2): Pcerca1(2) = P1(1)
    Plejos1(0) = P1(0) - mod2M * Cos(DirMuro2): Plejos1(1) = P1(1) - mod2M * Sin(DirMuro2): Plejos1(2) = P1(1)
    

    If ((mod1M + mod1c) < 2000) Or ((mod2c + mod2M) < 2000) Then
        userInput = InputBox("Elija una de las siguientes opciones:" & vbCrLf & vbCrLf & vbCrLf & "1. Dibujar el puntal seleccionado de longitud " & Distancia & "." & vbCrLf & vbCrLf & vbCrLf & "2. Dibujar un puntal MENOR de " & lalt1 & "mm de longitud más cercano a la esquina" & vbCrLf & vbCrLf & vbCrLf & "3. Dibujar un puntal MAYOR de " & lalt2 & "mm de longitud más alejado de la esquina")
        
        If userInput = "1" Or userInput = "" Then
            DirPuntal = gcadUtil.AngleFromXAxis(P1, P2)
            DirPuntal2 = gcadUtil.AngleFromXAxis(P2, P1)
            If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
                rutaplaca1 = rutaplaca1
                Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca1 = rutaplaca1
                If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
                ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
                End If
                Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
                blockRef.Layer = "Granshor"
            End If
            If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
                rutaplaca2 = rutaplaca2
                Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca2 = rutaplaca2
                If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
                ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
                End If
                Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
                blockRef.Layer = "Granshor"
            End If
        ElseIf userInput = "2" Then
            P1(0) = Pcerca1(0): P1(1) = Pcerca1(1): P1(2) = Pcerca1(2)
            P2(0) = Pcerca2(0): P2(1) = Pcerca2(1): P2(2) = Pcerca2(2)
            Distancia = lalt1
            PA(0) = PA(0) + mod1c * Cos(DirMuro1): PA(1) = PA(1) + mod1c * Sin(DirMuro1): PA(2) = PA(2)
            PB(0) = PB(0) + mod2c * Cos(DirMuro2): PB(1) = PB(1) + mod2c * Sin(DirMuro2): PB(2) = PB(0)
            DirPuntal = gcadUtil.AngleFromXAxis(P1, P2)
            DirPuntal2 = gcadUtil.AngleFromXAxis(P2, P1)
            If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
                rutaplaca1 = rutaplaca1
                Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca1 = rutaplaca1
                If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
                ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
                End If
                Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
                blockRef.Layer = "Granshor"
            End If
            If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
                rutaplaca2 = rutaplaca2
                Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca2 = rutaplaca2
                If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
                ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
                End If
                Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
                blockRef.Layer = "Granshor"
            End If

        ElseIf userInput = "3" Then
            P1(0) = Plejos1(0): P1(1) = Plejos1(1): P1(2) = Plejos1(2)
            P2(0) = Plejos2(0): P2(1) = Plejos2(1): P2(2) = Plejos2(2)
            Distancia = lalt2
            PA(0) = PA(0) - mod1M * Cos(DirMuro1): PA(1) = PA(1) - mod1M * Sin(DirMuro1): PA(2) = PA(2)
            PB(0) = PB(0) - mod2M * Cos(DirMuro2): PB(1) = PB(1) - mod2M * Sin(DirMuro2): PB(2) = PB(0)
            DirPuntal = gcadUtil.AngleFromXAxis(P1, P2)
            DirPuntal2 = gcadUtil.AngleFromXAxis(P2, P1)
            If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
                rutaplaca1 = rutaplaca1
                Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca1 = rutaplaca1
                If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
                ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
                End If
                Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
                blockRef.Layer = "Granshor"
            End If
            If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
                rutaplaca2 = rutaplaca2
                Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca2 = rutaplaca2
                If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
                ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
                End If
                Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
                blockRef.Layer = "Granshor"
            End If

        End If
    Else
        DirPuntal = gcadUtil.AngleFromXAxis(P1, P2)
        DirPuntal2 = gcadUtil.AngleFromXAxis(P2, P1)
        If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
            rutaplaca1 = rutaplaca1
            Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
            blockRef.Layer = "Granshor"
        Else
            rutaplaca1 = rutaplaca1
            If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
                rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
            ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
                rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
            End If
            Set blockRef = gcadModel.InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
            blockRef.Layer = "Granshor"
        End If
        If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
            rutaplaca2 = rutaplaca2
            Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Granshor"
        Else
            rutaplaca2 = rutaplaca2
            If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
                rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
            ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
                rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
            End If
            Set blockRef = gcadModel.InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
            blockRef.Layer = "Granshor"
        End If

    End If
End If


''''''''''' Dibujamos el puntal con los antiguos PS
''''''' antiguo PS

Set Eje1 = gcadModel.AddLine(P1, P2)
ANG = gcadUtil.AngleFromXAxis(P1, P2)
ANG2 = ANG + (PI / 2)

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

'Puntos centrales de las placas
PAP(0) = PA(0) - 50 * Cos(DirMuro1): PAP(1) = PA(1) - 50 * Sin(DirMuro1): PAP(2) = PA(2)
PBP(0) = PB(0) - 50 * Cos(DirMuro2): PBP(1) = PB(1) - 50 * Sin(DirMuro2): PBP(2) = PB(2)
x4 = PAP(0) - PBP(0)
y4 = PAP(1) - PBP(1)
D_ABP = Val(Sqr((x4 ^ 2 + y4 ^ 2)))

TxtPnt2(0) = PBP(0) + (D_ABP / 2) * Cos(ANG): TxtPnt2(1) = PBP(1) + (D_ABP / 2) * Sin(ANG): TxtPnt2(2) = PBP(2)
TxtPnt2(0) = TxtPnt2(0) + 860 * Cos(ANG2): TxtPnt2(1) = TxtPnt2(1) + 860 * Sin(ANG2): TxtPnt2(2) = TxtPnt2(2)

TxtPnt(0) = P1(0) + (Distancia / 2) * Cos(ANG): TxtPnt(1) = P1(1) + (Distancia / 2) * Sin(ANG): TxtPnt(2) = P1(2)
TxtPnt(0) = TxtPnt(0) + 410 * Cos(ANG2): TxtPnt(1) = TxtPnt(1) + 410 * Sin(ANG2): TxtPnt(2) = TxtPnt(2)

Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(P1, P2, TxtPnt)
objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objgcadDimAligned.StyleName = "MODELO"
objgcadDimAligned.TextStyle = "SIMPLEX"
objgcadDimAligned.VerticalTextPosition = acOutside
objgcadDimAligned.Update

Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(PBP, PAP, TxtPnt2)
objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objgcadDimAligned.StyleName = "MODELO"
objgcadDimAligned.TextStyle = "SIMPLEX"
objgcadDimAligned.VerticalTextPosition = acOutside
objgcadDimAligned.Update






'Introducir el bulón de 120 mm en los extremos siempre, ángulo de giro, fusible fijo y chapas de 50mm:
GS_Bulon120mm = rutags & "GS_Bulon120mm_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(P1, GS_Bulon120mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
Set blockRef = gcadModel.InsertBlock(P2, GS_Bulon120mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
GS_Giro = rutags & "GS_Giro_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(P1, GS_Giro, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
Set blockRef = gcadModel.InsertBlock(P2, GS_Giro, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Granshor"


If (carga < 1350) Then
    MsgBox "Los puntales mixtos actualmente han de lanzarse con el comando PM"
ElseIf (carga >= 1350) And (carga < 1500) Then
    Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
    GS_Fusible = rutags & "GS_Fusible_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"
    M20x90 = ruta2 & "4-M20X90.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
    M20x110 = ruta2 & "4-M20X110.dwg"
    M20x160 = ruta2 & "4-M20X160.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x160, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf (carga >= 1500) And (carga < 2000) Then
    Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
    GS_Fusible = rutags & "GS_Fusible_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"
    M20x90 = ruta2 & "4-M20X90.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
    PS_Placa35mm = rutaps & "PS_Placa35mm_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa35mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    M20x150 = ruta2 & "4-M20X150.dwg"
    M20x160 = ruta2 & "4-M20X160.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x160, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_inial(0) + 35 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 35 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x150, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf (carga >= 2000) And (carga < 2900) Then
    Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
    GS_Fusible = rutags & "GS_Fusible_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"
    M20x90 = ruta2 & "4-M20X90.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    M20x150 = ruta2 & "4-M20X150.dwg"
    M20x160 = ruta2 & "4-M20X160.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x160, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, M20x150, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
ElseIf carga >= 2900 Then
    Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
    GS_Fusible = rutags & "GS_Fusible_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"
    M20x90 = ruta2 & "4-M20X90.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Var20x250 = ruta2 & "1VarM20X250.dwg"
    M20x160 = ruta2 & "4-M20X160.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x160, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    ' aquí van las varillas
    Set blockRef = gcadModel.InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = gcadModel.InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = gcadModel.InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = gcadModel.InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
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
        If dato2 = "Pshor_4L" Then
        n280 = 2
        n560 = 0
        ElseIf dato2 = "Pshor_4S" Then
        n280 = 0
        n560 = 1
        End If
    Case Else
    MsgBox "Longitud no controlada " & lpuntal & "mm, fuera de rango, revisar código"
    GoTo terminar
        
End Select



M20x90_16 = ruta2 & "16-M20X90.dwg"

If n280 > 0 Then
    i = 0
    Do While i < n280
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_280 = rutapl & "PL_280_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_280, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Pipeshor4L"
        Punto_final(0) = Punto_inial(0) + l280 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l280 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If n560 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_560 = ruta & "PS_560.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n1500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_1500 = ruta & dato3 & "_1500_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_1500, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n3000 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_3000 = ruta & dato3 & "_3000_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_3000, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n4500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_4500 = ruta & dato3 & "_4500_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_4500, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n6000 > 0 Then
    i = 0
    Do While i < n6000
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_6000 = ruta & dato3 & "_6000_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_6000, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If


If n750 > 0 Then
    i = 0
    Do While i < n750
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_750 = ruta & dato3 & "_750_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_750, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

If carga >= 2900 Then
    blockRef.Delete
End If
    

If carga < 2900 Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    PGato1(0) = Punto_inial(0): PGato1(1) = Punto_inial(1): PGato1(2) = Punto_inial(2)
    zPS_Gato_Cono = rutaps & "zPS_Gato_Cono_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, zPS_Gato_Cono, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Pipeshor4S"
    'M20x90_16 = ruta & "16M20x90.dwg"
    'Set BlockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
    'BlockRef.Layer = "TORNILLERIA"
    Punto_final(0) = Punto_inial(0) + l_conogato * Cos(ANG): Punto_final(1) = Punto_inial(1) + l_conogato * Sin(ANG): Punto_final(2) = Punto_inial(2)



    Punto_inial2(0) = P2(0) - lgiro * Cos(ANG): Punto_inial2(1) = P2(1) - lgiro * Sin(ANG): Punto_inial2(2) = P2(2)
    Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
        If nfusible = 2 Then
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
        End If
    Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
    zPS_Gato_Tope = rutaps & "zPS_Gato_Tope_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial2, zPS_Gato_Tope, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Pipeshor4S"
    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x90, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Punto_final2(0) = Punto_inial2(0) - l_tope * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - l_tope * Sin(ANG): Punto_final(2) = Punto_inial2(2)
    
    PGato2(0) = Punto_inial2(0): PGato2(1) = Punto_inial2(1): PGato2(2) = Punto_inial2(2)

    x5 = PGato2(0) - PGato1(0)
    y5 = PGato2(1) - PGato1(1)
    D_Gato = Val(Sqr((x5 ^ 2 + y5 ^ 2)))

    TxtPnt3(0) = PGato1(0) + (D_Gato / 2) * Cos(ANG): TxtPnt3(1) = PGato1(1) + (D_Gato / 2) * Sin(ANG): TxtPnt3(2) = PGato1(2)
    TxtPnt3(0) = TxtPnt3(0) - 350 * Cos(ANG2): TxtPnt3(1) = TxtPnt3(1) - 350 * Sin(ANG2): TxtPnt3(2) = TxtPnt3(2)

    Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(PGato1, PGato2, TxtPnt3)
    objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
    objgcadDimAligned.StyleName = "MODELO"
    objgcadDimAligned.TextStyle = "SIMPLEX"
    objgcadDimAligned.VerticalTextPosition = acOutside
    objgcadDimAligned.Update


    Punto_inial(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_inial(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_inial(2) = (Punto_final(2) + Punto_final2(2)) / 2


    PS_Gato = rutaps & "PS_Gato_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Gato, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
Else
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    PGato1(0) = Punto_inial(0) + 100 * Cos(ANG): PGato1(1) = Punto_inial(1) + 100 * Sin(ANG): PGato1(2) = Punto_inial(2)
    'dos chapones de 50 y 4 varillas 20x250
    Set blockRef = gcadModel.InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = gcadModel.InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = gcadModel.InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = gcadModel.InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    ' base del cajón
    basecajon = rutacajon & "cajonh_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial, basecajon, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + 810 * Cos(ANG): Punto_final(1) = Punto_inial(1) + 810 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    
    'nos vamos al final del codal a meter el giro y acercarnos al cajón, haya uno o dos fusibles
    Punto_inial2(0) = P2(0) - lgiro * Cos(ANG): Punto_inial2(1) = P2(1) - lgiro * Sin(ANG): Punto_inial2(2) = P2(2)
    Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
        If nfusible = 2 Then
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
            Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
        End If
    Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
    'metemos aquí el cajón hidráulico
    brazocajon = rutacajon & "modcajon_" & dato1 & ".dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inial2, brazocajon, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Pipeshor4S"
    Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x90, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Nonplot"
    blockRef.Update
            blockRef.Explode
            blockRef.Delete
    
    
    
    PGato2(0) = Punto_inial2(0): PGato2(1) = Punto_inial2(1): PGato2(2) = Punto_inial2(2)

    x5 = PGato2(0) - PGato1(0)
    y5 = PGato2(1) - PGato1(1)
    D_Gato = Val(Sqr((x5 ^ 2 + y5 ^ 2)))

    TxtPnt3(0) = PGato1(0) + (D_Gato / 2) * Cos(ANG): TxtPnt3(1) = PGato1(1) + (D_Gato / 2) * Sin(ANG): TxtPnt3(2) = PGato1(2)
    TxtPnt3(0) = TxtPnt3(0) - 350 * Cos(ANG2): TxtPnt3(1) = TxtPnt3(1) - 350 * Sin(ANG2): TxtPnt3(2) = TxtPnt3(2)

    Set objgcadDimAligned = ThisDrawing.ModelSpace.AddDimAligned(PGato1, PGato2, TxtPnt3)
    objgcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
    objgcadDimAligned.StyleName = "MODELO"
    objgcadDimAligned.TextStyle = "SIMPLEX"
    objgcadDimAligned.VerticalTextPosition = acOutside
    objgcadDimAligned.Update
End If


  

 
Eje1.Layer = "Nonplot"
Loop

terminar:

End Sub

























