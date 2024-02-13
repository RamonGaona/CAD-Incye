
Sub ps()
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
Dim M20x90 As String
Dim M20x150 As String
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
Dim PS_Husillo As String
Dim PS_Placa50mm As String
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
Dim l50 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lalt1 As Double
Dim lalt2 As Double
Dim lgatomin As Double
Dim n6000 As Integer
Dim n4500 As Integer
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
l280 = 280
l560 = 560
l750 = 750
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
l50 = 50
l_tope = 325
l_conogato = 170
lgatomin = 620
lfija = (2 * lgiro) + lfusible + (2 * l50) + lgatomin

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

Dim blockCount As Object
Set blockCount = CreateObject("Scripting.Dictionary")
Dim blockNames As New Collection
Dim blk As Variant
Dim blockLayers As Object
Set blockLayers = CreateObject("Scripting.Dictionary")

'Introducir el bulón de 120 mm en los extremos siempre, ángulo de giro, fusible fijo y chapas de 50mm:
GS_Bulon120mm = rutags & "GS_Bulon120mm_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(P1, GS_Bulon120mm, Xs, Ys, Zs, ANG)
Dim bulon As String
bulon = "GS_Bulon120mm_" & dato1 & ""
If Not blockCount.Exists(bulon) Then
    blockCount.Add bulon, 1
Else
    blockCount(bulon) = blockCount(bulon) + 1
End If
blockNames.Add bulon
blockRef.Layer = "Granshor"
blockLayers.Add bulon, blockRef.Layer
Set blockRef = gcadModel.InsertBlock(P2, GS_Bulon120mm, Xs, Ys, Zs, ANG)
If Not blockCount.Exists(bulon) Then
    blockCount.Add bulon, 1
Else
    blockCount(bulon) = blockCount(bulon) + 1
End If
blockNames.Add bulon
blockRef.Layer = "Granshor"
If Not blockLayers.Exists(bulon) Then
    blockLayers.Add bulon, blockRef.Layer
End If
GS_Giro = rutags & "GS_Giro_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(P1, GS_Giro, Xs, Ys, Zs, ANG)
    Dim Giro As String
    Giro = "GS_Giro_" & dato1 & ""
    If Not blockCount.Exists(Giro) Then
        blockCount.Add Giro, 1
    Else
      blockCount(Giro) = blockCount(Giro) + 1
    End If
    blockNames.Add Giro
blockRef.Layer = "Granshor"
If Not blockLayers.Exists(Giro) Then
    blockLayers.Add Giro, blockRef.Layer
End If
Set blockRef = gcadModel.InsertBlock(P2, GS_Giro, Xs, Ys, Zs, ANG + PI)
    Giro = "GS_Giro_" & dato1 & ""
    If Not blockCount.Exists(Giro) Then
        blockCount.Add Giro, 1
    Else
      blockCount(Giro) = blockCount(Giro) + 1
    End If
    blockNames.Add Giro
blockRef.Layer = "Granshor"
If Not blockLayers.Exists(Giro) Then
    blockLayers.Add Giro, blockRef.Layer
End If


Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
GS_Fusible = rutags & "GS_Fusible_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    Dim fusible As String
    fusible = "GS_Fusible_" & dato1 & ""
    If Not blockCount.Exists(fusible) Then
        blockCount.Add fusible, 1
    Else
      blockCount(fusible) = blockCount(fusible) + 1
    End If
    blockNames.Add fusible
blockRef.Layer = "Granshor"
If Not blockLayers.Exists(fusible) Then
    blockLayers.Add fusible, blockRef.Layer
End If

M20x90 = ruta2 & "4-M20X90.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)

    Dim m2090 As String
    m2090 = "4M20X90"
    If Not blockCount.Exists(m2090) Then
        blockCount.Add m2090, 1
    Else
      blockCount(m2090) = blockCount(m2090) + 1
    End If
    blockNames.Add m2090
blockRef.Layer = "Nonplot"

If Not blockLayers.Exists(m2090) Then
    blockLayers.Add m2090, blockRef.Layer
End If
blockRef.Update
            blockRef.Explode
            blockRef.Delete
Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    Dim placa50 As String
    placa50 = "PS_Placa50mm_" & dato1 & ""
    If Not blockCount.Exists(placa50) Then
        blockCount.Add placa50, 1
    Else
      blockCount(placa50) = blockCount(placa50) + 1
    End If
    blockNames.Add placa50
blockRef.Layer = "Pipeshor4S"
If Not blockLayers.Exists(placa50) Then
    blockLayers.Add placa50, blockRef.Layer
End If

M20x150 = ruta2 & "4-M20X150.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x150, Xs, Ys, Zs, ANG)

    Dim m20150 As String
    m20150 = "4M20X150"
    If Not blockCount.Exists(m20150) Then
        blockCount.Add m20150, 1
    Else
      blockCount(m20150) = blockCount(m20150) + 1
    End If
    blockNames.Add m20150
blockRef.Layer = "Nonplot"

If Not blockLayers.Exists(m20150) Then
    blockLayers.Add m20150, blockRef.Layer
End If
blockRef.Update
            blockRef.Explode
            blockRef.Delete

Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    If Not blockCount.Exists(placa50) Then
        blockCount.Add placa50, 1
    Else
      blockCount(placa50) = blockCount(placa50) + 1
    End If
    blockNames.Add placa50
blockRef.Layer = "Pipeshor4S"
If Not blockLayers.Exists(placa50) Then
    blockLayers.Add placa50, blockRef.Layer
End If

M20x160 = ruta2 & "4-M20X160.dwg"
Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
Set blockRef = gcadModel.InsertBlock(Punto_final, M20x160, Xs, Ys, Zs, ANG)

    Dim m20160 As String
    m20160 = "4M20X160"
    If Not blockCount.Exists(m20160) Then
        blockCount.Add m20160, 1
    Else
      blockCount(m20160) = blockCount(m20160) + 1
    End If
    blockNames.Add m20160
blockRef.Layer = "Nonplot"

If Not blockLayers.Exists(m20160) Then
    blockLayers.Add m20160, blockRef.Layer
End If
blockRef.Update
            blockRef.Explode
            blockRef.Delete


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
Dim m2090_16 As String
m2090_16 = "16M20X90"

If n280 > 0 Then
    i = 0
    Do While i < n280
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_280 = rutapl & "PL_280_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_280, Xs, Ys, Zs, ANG)
            Dim ps280 As String
            ps280 = "PL_280_" & dato1 & ""
            If Not blockCount.Exists(ps280) Then
                blockCount.Add ps280, 1
            Else
              blockCount(ps280) = blockCount(ps280) + 1
            End If
            blockNames.Add ps280
        blockRef.Layer = "Pipeshor4L"
        If Not blockLayers.Exists(ps280) Then
            blockLayers.Add ps280, blockRef.Layer
        End If
        
        Punto_final(0) = Punto_inial(0) + l280 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l280 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)

            If Not blockCount.Exists(m2090_16) Then
                blockCount.Add m2090_16, 1
            Else
              blockCount(m2090_16) = blockCount(m2090_16) + 1
            End If
            blockNames.Add m2090_16
        blockRef.Layer = "Nonplot"
        
        If Not blockLayers.Exists(m2090_16) Then
            blockLayers.Add m2090_16, blockRef.Layer
        End If
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
            Dim ps560 As String
            ps560 = "PL_560_" & dato1 & ""
            If Not blockCount.Exists(ps560) Then
                blockCount.Add ps560, 1
            Else
              blockCount(ps560) = blockCount(ps560) + 1
            End If
            blockNames.Add ps560
        blockRef.Layer = capa
        If Not blockLayers.Exists(ps560) Then
            blockLayers.Add ps560, blockRef.Layer
        End If
        Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)

            If Not blockCount.Exists(m2090_16) Then
                blockCount.Add m2090_16, 1
            Else
              blockCount(m2090_16) = blockCount(m2090_16) + 1
            End If
            blockNames.Add m2090_16
        blockRef.Layer = "Nonplot"
        
        If Not blockLayers.Exists(m2090_16) Then
            blockLayers.Add m2090_16, blockRef.Layer
        End If
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n1500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_1500 = ruta & dato3 & "_1500_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_1500, Xs, Ys, Zs, ANG)
            Dim ps1500 As String
            ps1500 = "PL_1500_" & dato1 & ""
            If Not blockCount.Exists(ps1500) Then
                blockCount.Add ps1500, 1
            Else
              blockCount(ps1500) = blockCount(ps1500) + 1
            End If
            blockNames.Add ps1500
        blockRef.Layer = capa
        If Not blockLayers.Exists(ps1500) Then
            blockLayers.Add ps1500, blockRef.Layer
        End If
        Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)

            If Not blockCount.Exists(m2090_16) Then
                blockCount.Add m2090_16, 1
            Else
              blockCount(m2090_16) = blockCount(m2090_16) + 1
            End If
            blockNames.Add m2090_16
        blockRef.Layer = "Nonplot"
        
        If Not blockLayers.Exists(m2090_16) Then
            blockLayers.Add m2090_16, blockRef.Layer
        End If
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n3000 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_3000 = ruta & dato3 & "_3000_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_3000, Xs, Ys, Zs, ANG)
            Dim ps3000 As String
            ps3000 = "PL_3000_" & dato1 & ""
            If Not blockCount.Exists(ps3000) Then
                blockCount.Add ps3000, 1
            Else
              blockCount(ps3000) = blockCount(ps3000) + 1
            End If
            blockNames.Add ps3000
        blockRef.Layer = capa
        If Not blockLayers.Exists(ps3000) Then
            blockLayers.Add ps3000, blockRef.Layer
        End If
        Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)

            If Not blockCount.Exists(m2090_16) Then
                blockCount.Add m2090_16, 1
            Else
              blockCount(m2090_16) = blockCount(m2090_16) + 1
            End If
            blockNames.Add m2090_16
        blockRef.Layer = "Nonplot"
        
        If Not blockLayers.Exists(m2090_16) Then
            blockLayers.Add m2090_16, blockRef.Layer
        End If
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
End If

If n4500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_4500 = ruta & dato3 & "_4500_" & dato1 & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_4500, Xs, Ys, Zs, ANG)
            Dim ps4500 As String
            ps4500 = "PL_4500_" & dato1 & ""
            If Not blockCount.Exists(ps4500) Then
                blockCount.Add ps4500, 1
            Else
              blockCount(ps4500) = blockCount(ps4500) + 1
            End If
            blockNames.Add ps4500
        blockRef.Layer = capa
        If Not blockLayers.Exists(ps4500) Then
            blockLayers.Add ps4500, blockRef.Layer
        End If
        Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)

            If Not blockCount.Exists(m2090_16) Then
                blockCount.Add m2090_16, 1
            Else
              blockCount(m2090_16) = blockCount(m2090_16) + 1
            End If
            blockNames.Add m2090_16
        blockRef.Layer = "Nonplot"
        
        If Not blockLayers.Exists(m2090_16) Then
            blockLayers.Add m2090_16, blockRef.Layer
        End If
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
            Dim ps6000 As String
            ps6000 = "PL_6000_" & dato1 & ""
            If Not blockCount.Exists(ps6000) Then
                blockCount.Add ps6000, 1
            Else
              blockCount(ps6000) = blockCount(ps6000) + 1
            End If
            blockNames.Add ps6000
        blockRef.Layer = capa
        If Not blockLayers.Exists(ps6000) Then
            blockLayers.Add ps6000, blockRef.Layer
        End If
        Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)

            If Not blockCount.Exists(m2090_16) Then
                blockCount.Add m2090_16, 1
            Else
              blockCount(m2090_16) = blockCount(m2090_16) + 1
            End If
            blockNames.Add m2090_16
        blockRef.Layer = "Nonplot"
    
        If Not blockLayers.Exists(m2090_16) Then
            blockLayers.Add m2090_16, blockRef.Layer
        End If
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
            Dim ps750 As String
            ps750 = "PL_750_" & dato1 & ""
            If Not blockCount.Exists(ps750) Then
                blockCount.Add ps750, 1
            Else
              blockCount(ps750) = blockCount(ps750) + 1
            End If
            blockNames.Add ps750
        blockRef.Layer = capa
        If Not blockLayers.Exists(ps750) Then
            blockLayers.Add ps750, blockRef.Layer
        End If
        Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)

            If Not blockCount.Exists(m2090_16) Then
                blockCount.Add m2090_16, 1
            Else
              blockCount(m2090_16) = blockCount(m2090_16) + 1
            End If
            blockNames.Add m2090_16
        blockRef.Layer = "Nonplot"
        
        If Not blockLayers.Exists(m2090_16) Then
            blockLayers.Add m2090_16, blockRef.Layer
        End If
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        i = i + 1
    Loop
End If

Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
zPS_Gato_Cono = rutaps & "zPS_Gato_Cono_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, zPS_Gato_Cono, Xs, Ys, Zs, ANG + PI)
    Dim zpsgatocono As String
    zpsgatocono = "zPS_Gato_Cono_" & dato1 & ""
        If Not blockCount.Exists(zpsgatocono) Then
            blockCount.Add zpsgatocono, 1
        Else
          blockCount(zpsgatocono) = blockCount(zpsgatocono) + 1
        End If
        blockNames.Add zpsgatocono
blockRef.Layer = "Pipeshor4S"
If Not blockLayers.Exists(zpsgatocono) Then
    blockLayers.Add zpsgatocono, blockRef.Layer
End If
'M20x90_16 = ruta & "16M20x90.dwg"
'Set BlockRef = gcadModel.InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
'BlockRef.Layer = "TORNILLERIA"
Punto_final(0) = Punto_inial(0) + l_conogato * Cos(ANG): Punto_final(1) = Punto_inial(1) + l_conogato * Sin(ANG): Punto_final(2) = Punto_inial(2)

Punto_inial2(0) = P2(0) - lgiro * Cos(ANG): Punto_inial2(1) = P2(1) - lgiro * Sin(ANG): Punto_inial2(2) = P2(2)
Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
    If nfusible = 2 Then
        Set blockRef = gcadModel.InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
            If Not blockCount.Exists(fusible) Then
                blockCount.Add fusible, 1
            Else
              blockCount(fusible) = blockCount(fusible) + 1
            End If
            blockNames.Add fusible
        blockRef.Layer = "Granshor"
        If Not blockLayers.Exists(fusible) Then
            blockLayers.Add fusible, blockRef.Layer
        End If
        Set blockRef = gcadModel.InsertBlock(Punto_inial2, M20x90, Xs, Ys, Zs, ANG + PI)

            If Not blockCount.Exists(m2090) Then
                blockCount.Add m2090, 1
            Else
              blockCount(m2090) = blockCount(m2090) + 1
            End If
            blockNames.Add m2090
        blockRef.Layer = "Nonplot"
        
        If Not blockLayers.Exists(m2090) Then
            blockLayers.Add m2090, blockRef.Layer
        End If
        blockRef.Update
            blockRef.Explode
            blockRef.Delete
        Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
    End If
Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
zPS_Gato_Tope = rutaps & "zPS_Gato_Tope_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial2, zPS_Gato_Tope, Xs, Ys, Zs, ANG + PI)
    Dim zpsgatotope As String
    zpsgatotope = "zPS_Gato_Tope_" & dato1 & ""
    If Not blockCount.Exists(zpsgatotope) Then
        blockCount.Add zpsgatotope, 1
    Else
      blockCount(zpsgatotope) = blockCount(zpsgatotope) + 1
    End If
    blockNames.Add zpsgatotope
blockRef.Layer = "Pipeshor4S"
If Not blockLayers.Exists(zpsgatotope) Then
    blockLayers.Add zpsgatotope, blockRef.Layer
End If
Set blockRef = gcadModel.InsertBlock(Punto_final2, M20x90, Xs, Ys, Zs, ANG + PI)

    If Not blockCount.Exists(m2090) Then
        blockCount.Add m2090, 1
    Else
      blockCount(m2090) = blockCount(m2090) + 1
    End If
    blockNames.Add m2090
blockRef.Layer = "Nonplot"

If Not blockLayers.Exists(m2090) Then
    blockLayers.Add m2090, blockRef.Layer
End If
blockRef.Update
            blockRef.Explode
            blockRef.Delete
Punto_final2(0) = Punto_inial2(0) - l_tope * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - l_tope * Sin(ANG): Punto_final(2) = Punto_inial2(2)


Punto_inial(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_inial(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_inial(2) = (Punto_final(2) + Punto_final2(2)) / 2

PS_Gato = rutaps & "PS_Gato_" & dato1 & ".dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inial, PS_Gato, Xs, Ys, Zs, ANG)
    Dim psgato As String
    psgato = "PS_Gato_" & dato1 & ""
    If Not blockCount.Exists(psgato) Then
        blockCount.Add psgato, 1
    Else
      blockCount(psgato) = blockCount(psgato) + 1
    End If
    blockNames.Add psgato
blockRef.Layer = "Pipeshor4S"
If Not blockLayers.Exists(psgato) Then
    blockLayers.Add psgato, blockRef.Layer
End If
        
If dato3 = "PL" Then
    If nfusible = 1 Then
        If n280 = 1 Then
            lalt1 = Distancia - l280
            lalt2 = Distancia + 470
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        ElseIf n280 = 2 Then
            lalt1 = Distancia - l280
            lalt2 = Distancia + 190
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt2 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt1 & "."
        End If
    ElseIf nfusible = 2 Then
        If n280 = 1 Then
            lalt1 = Distancia - lfusible - l280 + 150
            lalt2 = Distancia - l280 - lfusible + l750
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt2 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt1 & "."
        ElseIf n280 = 2 Then
            lalt1 = Distancia - lfusible - l280 + 150
            lalt2 = Distancia - 560 - lfusible + l750
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt2 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt1 & "."
        End If
    End If
ElseIf dato3 = "PS" Then
    If n280 = 1 Then
        If n560 = 1 And n750 = 1 Then
            lalt1 = Distancia - l280 + 190
            lalt2 = Distancia + l280
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        ElseIf n560 = 0 And n750 = 1 Then
            lalt1 = Distancia - l280
            lalt2 = Distancia + l280
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        ElseIf n560 = 1 And n750 = 0 Then
            lalt1 = Distancia - 90
            lalt2 = Distancia + 280
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        Else
            lalt1 = Distancia - 280
            lalt2 = Distancia + 280
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        End If
    End If
End If
        
Eje1.Layer = "Nonplot"

Dim uniqueBlocks As Object
Set uniqueBlocks = CreateObject("Scripting.Dictionary")

' Llenar el diccionario de bloques únicos y sus recuentos
For Each blk In blockNames
    If Not uniqueBlocks.Exists(blk) Then
        uniqueBlocks.Add blk, blockCount(blk)
    End If
Next blk

' Crear una nueva instancia de Excel
Dim excelApp As Object
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = True

' Crear un nuevo libro de Excel
Dim wb As Object
Set wb = excelApp.Workbooks.Add

' Seleccionar la primera hoja de Excel
Dim ws As Object
Set ws = wb.Sheets(1)

' Establecer encabezados de columna
ws.Cells(1, 1).Value = "Nombre de bloque"
ws.Cells(1, 2).Value = "Cantidad"
ws.Cells(1, 3).Value = "CAPA"

' Llenar los datos en la hoja de Excel
Dim row As Integer
row = 2

For Each blk In uniqueBlocks
    ws.Cells(row, 1).Value = blk
    ws.Cells(row, 2).Value = blockLayers(blk)
    ws.Cells(row, 3).Value = blockCount(blk)
    row = row + 1
Next blk

' Ajustar el ancho de las columnas
ws.Columns("A:C").AutoFit

' Guardar el libro de Excel
wb.SaveAs "Ruta\de\tu\archivo.xlsx"

' Cerrar la instancia de Excel
wb.Close
excelApp.Quit


Loop


terminar:

End Sub











