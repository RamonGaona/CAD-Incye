''''''' Geometría para el cálculo del ajuste de los puntales pipeshor
Option Explicit
' MTN : Multitensor

Sub mtn()
' luego cada subfunción es llamada con la primera selección de puntos sobre el plano de gcad
Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim rutall As String, rutamp As String, rutator As String, rutampacc As String, rutass As String
Dim PI As Variant, Distancia As Double
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim repite As Double
Dim punto1 As Variant, punto2 As Variant
Dim kwordList As String

Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

On Error GoTo terminar
repite = 1

Do While repite = 1

punto1 = gcadUtil.GetPoint(, "1º Punto: ")
punto2 = gcadUtil.GetPoint(punto1, "2º Punto: ")

x = punto2(0) - punto1(0)
y = punto2(1) - punto1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

Debug.Print punto1(0); punto1(1); punto1(2)
Debug.Print punto2(0); punto2(1); punto2(2)

If (Distancia >= 500) And (Distancia <= 778) Then
    Call tensor_c(punto1, punto2)
ElseIf (Distancia >= 915) And (Distancia <= 1160) Then
    Call tensor_l(punto1, punto2)
ElseIf (Distancia > 1160) And (Distancia <= 1500) Then
    Call tensor_xl(punto1, punto2)
ElseIf (Distancia > 1500) And (Distancia <= 1900) Then
    Call tensor_xl2(punto1, punto2)
ElseIf (Distancia > 1900) And (Distancia < 3200) Then
' zona gris donde según la carga puede entrar un Telescópico o un Tubo80 pequeño
    If Distancia > 2970 Then
        kwordList = "Telescopico Tubo80"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        
        Dim teleocuadrado As String
        teleocuadrado = ThisDrawing.Utility.GetKeyword(vbLf & "Según la carga, ¿qué tensor deseas introducir?: [Telescopico/Tubo80]")
        
        If teleocuadrado = "" Or teleocuadrado = "Telescopico" Then
            Call tensor_telesc(punto1, punto2)
        ElseIf teleocuadrado = "Tubo80" Then
            Call Tubo80x40_tensor(punto1, punto2)
        End If
    Else
        Call tensor_telesc(punto1, punto2)
    End If
ElseIf (Distancia >= 3200) And (Distancia <= 6320) Then
    Call Tubo80x40_tensor(punto1, punto2)
ElseIf (Distancia > 6320) And (Distancia < 9000) Then
    Call ssc_tensor(punto1, punto2)
ElseIf (Distancia >= 9000) And (Distancia < 12000) Then
' zona gris donde según la carga puede entrar un SS o un Lolashor
    If Distancia > 10000 Then
        kwordList = "SS Lola"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        
        Dim ssolola As String
        ssolola = ThisDrawing.Utility.GetKeyword(vbLf & "Según la carga, ¿qué deseas introducir?: [SS/Lola]")
        
        If ssolola = "" Or ssolola = "SS" Then
            Call ssclargo_tensor(punto1, punto2)
        ElseIf ssolola = "Lola" Then
            Call ll_tensor(punto1, punto2)
        End If
    Else
        Call ssclargo_tensor(punto1, punto2)
    End If

    
ElseIf Distancia >= 12000 Then
    Call ll_tensor(punto1, punto2)
Else
    MsgBox "No existen tensores de esta longitud"
End If

Loop
terminar:
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub tensor_c(punto1 As Variant, punto2 As Variant)

Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim rutatensor As String
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double
Dim Gcapa As Object
Dim Ncapa As String
Dim PI As Variant, Distancia As Double
Dim PuntoM(0 To 2) As Double, husillo_c As String, cuerpo_c As String


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

rutatensor = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensores\"

Dim rutator As String
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
Dim Bulond19 As String
Bulond19 = rutator & "1M19_BULOND19.dwg"

On Error GoTo terminar

PI = 4 * Atn(1)

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
Punto_final(0) = Punto_inicial(0) + Distancia * Cos(ANG): Punto_final(1) = Punto_inicial(1) + Distancia * Sin(ANG): Punto_final(2) = Punto_inicial(2)

PuntoM(0) = Punto_inicial(0) + (Distancia / 2) * Cos(ANG): PuntoM(1) = Punto_inicial(1) + (Distancia / 2) * Sin(ANG): PuntoM(2) = Punto_inicial(2)

husillo_c = rutatensor & "SSHUSILLOTENSORCORTO.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inicial, husillo_c, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"

husillo_c = rutatensor & "SSHUSILLOTENSORCORTO.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, husillo_c, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Slims"

cuerpo_c = rutatensor & "SSCUERPOTENSORCORTO.dwg"
Set blockRef = gcadModel.InsertBlock(PuntoM, cuerpo_c, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"


Set blockRef = gcadModel.InsertBlock(P1, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Set blockRef = gcadModel.InsertBlock(P2, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"

terminar:
End Sub




Sub tensor_l(punto1 As Variant, punto2 As Variant)

Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim rutatensor As String
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double
Dim Gcapa As Object
Dim Ncapa As String
Dim PI As Variant, Distancia As Double
Dim PuntoM(0 To 2) As Double, husillo_l As String, cuerpo_l As String


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

rutatensor = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensores\"

Dim rutator As String
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
Dim Bulond19 As String
Bulond19 = rutator & "1M19_BULOND19.dwg"

On Error GoTo terminar

PI = 4 * Atn(1)

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
Punto_final(0) = Punto_inicial(0) + Distancia * Cos(ANG): Punto_final(1) = Punto_inicial(1) + Distancia * Sin(ANG): Punto_final(2) = Punto_inicial(2)

PuntoM(0) = Punto_inicial(0) + (Distancia / 2) * Cos(ANG): PuntoM(1) = Punto_inicial(1) + (Distancia / 2) * Sin(ANG): PuntoM(2) = Punto_inicial(2)

husillo_l = rutatensor & "SSHUSILLOTENSORLARGO.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inicial, husillo_l, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"

husillo_l = rutatensor & "SSHUSILLOTENSORLARGO.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, husillo_l, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Slims"

cuerpo_l = rutatensor & "SSCUERPOTENSORLARGO.dwg"
Set blockRef = gcadModel.InsertBlock(PuntoM, cuerpo_l, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"


Set blockRef = gcadModel.InsertBlock(P1, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Set blockRef = gcadModel.InsertBlock(P2, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"

terminar:
End Sub



Sub tensor_xl(punto1 As Variant, punto2 As Variant)

Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim rutatensor As String
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double
Dim Gcapa As Object
Dim Ncapa As String
Dim PI As Variant, Distancia As Double
Dim PuntoM(0 To 2) As Double, husillo_xl As String, cuerpo_xl As String


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

rutatensor = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensores\"

Dim rutator As String
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
Dim Bulond19 As String
Bulond19 = rutator & "1M19_BULOND19.dwg"

On Error GoTo terminar

PI = 4 * Atn(1)

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
Punto_final(0) = Punto_inicial(0) + Distancia * Cos(ANG): Punto_final(1) = Punto_inicial(1) + Distancia * Sin(ANG): Punto_final(2) = Punto_inicial(2)

PuntoM(0) = Punto_inicial(0) + (Distancia / 2) * Cos(ANG): PuntoM(1) = Punto_inicial(1) + (Distancia / 2) * Sin(ANG): PuntoM(2) = Punto_inicial(2)

husillo_xl = rutatensor & "SSHUSILLOTENSORXL.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inicial, husillo_xl, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"

husillo_xl = rutatensor & "SSHUSILLOTENSORXL.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, husillo_xl, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Slims"

cuerpo_xl = rutatensor & "SSCUERPOTENSORXL.dwg"
Set blockRef = gcadModel.InsertBlock(PuntoM, cuerpo_xl, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"

Set blockRef = gcadModel.InsertBlock(P1, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Set blockRef = gcadModel.InsertBlock(P2, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"

terminar:
End Sub

Sub tensor_xl2(punto1 As Variant, punto2 As Variant)

Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim rutatensor As String
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double
Dim Gcapa As Object
Dim Ncapa As String
Dim PI As Variant, Distancia As Double
Dim PuntoM(0 To 2) As Double, husillo_xl2 As String, cuerpo_xl2 As String


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

Dim rutator As String
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
Dim Bulond19 As String
Bulond19 = rutator & "1M19_BULOND19.dwg"

rutatensor = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensores\"

On Error GoTo terminar

PI = 4 * Atn(1)

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
Punto_final(0) = Punto_inicial(0) + Distancia * Cos(ANG): Punto_final(1) = Punto_inicial(1) + Distancia * Sin(ANG): Punto_final(2) = Punto_inicial(2)

PuntoM(0) = Punto_inicial(0) + (Distancia / 2) * Cos(ANG): PuntoM(1) = Punto_inicial(1) + (Distancia / 2) * Sin(ANG): PuntoM(2) = Punto_inicial(2)

husillo_xl2 = rutatensor & "SSHUSILLOTENSORXL.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_inicial, husillo_xl2, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"

husillo_xl2 = rutatensor & "SSHUSILLOTENSORXL.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, husillo_xl2, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Slims"

cuerpo_xl2 = rutatensor & "SSCUERPOTENSOR1500a1900.dwg"
Set blockRef = gcadModel.InsertBlock(PuntoM, cuerpo_xl2, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"

Set blockRef = gcadModel.InsertBlock(P1, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Set blockRef = gcadModel.InsertBlock(P2, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"

terminar:
End Sub

Sub tensor_telesc(punto1 As Variant, punto2 As Variant)

Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim rutatensor As String
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double
Dim Gcapa As Object
Dim Ncapa As String
Dim PI As Variant, Distancia As Double
Dim PuntoM(0 To 2) As Double, husillo_telesc As String, cuerpo_telesc As String

Dim rutator As String
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
Dim Bulond19 As String
Bulond19 = rutator & "1M19_BULOND19.dwg"

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

rutatensor = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensores\"

On Error GoTo terminar

PI = 4 * Atn(1)

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
Punto_final(0) = Punto_inicial(0) + Distancia * Cos(ANG): Punto_final(1) = Punto_inicial(1) + Distancia * Sin(ANG): Punto_final(2) = Punto_inicial(2)

PuntoM(0) = Punto_inicial(0) + (Distancia / 2) * Cos(ANG): PuntoM(1) = Punto_inicial(1) + (Distancia / 2) * Sin(ANG): PuntoM(2) = Punto_inicial(2)


If Distancia <= 2250 Then
    husillo_telesc = rutatensor & "SSHusilloTelesc.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inicial, husillo_telesc, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    
    husillo_telesc = rutatensor & "SSHusilloTelesc.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, husillo_telesc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    
    cuerpo_telesc = rutatensor & "SSCUERPOTELESCOPICO1900a3200P1.dwg"
    Set blockRef = gcadModel.InsertBlock(PuntoM, cuerpo_telesc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
ElseIf (Distancia > 2250) And (Distancia < 2550) Then
    husillo_telesc = rutatensor & "SSHusilloTelesc.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inicial, husillo_telesc, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    
    husillo_telesc = rutatensor & "SSHusilloTelesc.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, husillo_telesc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    
    cuerpo_telesc = rutatensor & "SSCUERPOTELESCOPICO1900a3200P2.dwg"
    Set blockRef = gcadModel.InsertBlock(PuntoM, cuerpo_telesc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
ElseIf (Distancia >= 2250) And (Distancia < 2850) Then
    husillo_telesc = rutatensor & "SSHusilloTelesc.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inicial, husillo_telesc, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    
    husillo_telesc = rutatensor & "SSHusilloTelesc.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, husillo_telesc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    
    cuerpo_telesc = rutatensor & "SSCUERPOTELESCOPICO1900a3200P3.dwg"
    Set blockRef = gcadModel.InsertBlock(PuntoM, cuerpo_telesc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
ElseIf (Distancia >= 2850) And (Distancia < 3200) Then
    husillo_telesc = rutatensor & "SSHusilloTelesc.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_inicial, husillo_telesc, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    
    husillo_telesc = rutatensor & "SSHusilloTelesc.dwg"
    Set blockRef = gcadModel.InsertBlock(Punto_final, husillo_telesc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    
    cuerpo_telesc = rutatensor & "SSCUERPOTELESCOPICO1900a3200P4.dwg"
    Set blockRef = gcadModel.InsertBlock(PuntoM, cuerpo_telesc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
End If

Set blockRef = gcadModel.InsertBlock(P1, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Set blockRef = gcadModel.InsertBlock(P2, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"

terminar:
End Sub



Sub ll_tensor(punto1 As Variant, punto2 As Variant)

Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim rutall As String, rutamp As String, rutator As String, rutampacc As String, rutass As String
Dim PI As Variant
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

Dim rutatensor As String

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

kwordList = "A B"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
disposicion = ThisDrawing.Utility.GetKeyword(vbLf & "Disposición: A (Planta) B (Alzado): [A/B]")

If disposicion = "" Or disposicion = "A" Then
vista = "PL"
vistamp = "PLA"
ElseIf disposicion = "B" Then
vista = ""
vistamp = "ALZ"
Else
GoTo terminar
End If



lfija = 2 * lcono + lgatominss + ladaptador + 415


M24x110 = rutator & "1-M24X110.dwg"
M16x40_4 = rutator & "4-M16X40.dwg"
M20x60_4 = rutator & "4-M20X60.dwg"
M16x40_8 = rutator & "8-M16X40.dwg"
M16x40_16 = rutator & "16-M16X40.dwg"
Dim Bulond19 As String
Bulond19 = rutator & "1M19_BULOND19.dwg"
Dim Bulond23 As String
Bulond23 = rutator & "1M23_BULOND23.dwg"


    'Geometría:
    'punto1 = gcadUtil.GetPoint(, "1º Punto: ")
    'punto2 = gcadUtil.GetPoint(punto1, "2º Punto: ")
    
    
For i = 0 To 2
  P1(i) = 0
  P2(i) = 0
Next i

Dim ANG2 As Double
Dim ANG3 As Double
Dim PuntoMedio As Double
Dim P2prov(0 To 2) As Double
Dim PStrut(0 To 2) As Double
Dim PStrut1 As Variant
Dim PStrut2 As Variant

P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
P2prov(0) = punto2(0): P2prov(1) = punto2(1): P2prov(2) = punto2(2)

PStrut1 = gcadUtil.GetPoint(, "1 anclaje del Strut Adaptor: ")
PStrut2 = gcadUtil.GetPoint(PStrut1, "2 anclaje del Strut Adaptor: ")

Dim Pbulon1(0 To 2) As Double
Dim Pbulon2(0 To 2) As Double

Pbulon1(0) = PStrut1(0): Pbulon1(1) = PStrut1(1): Pbulon1(2) = PStrut1(2)
Pbulon2(0) = PStrut2(0): Pbulon2(1) = PStrut2(1): Pbulon2(2) = PStrut2(2)

ANG2 = gcadUtil.AngleFromXAxis(PStrut2, PStrut1)

PStrut(0) = PStrut2(0) + 90 * Cos(ANG2): PStrut(1) = PStrut2(1) + 90 * Sin(ANG2): PStrut(2) = PStrut2(2)
ANG3 = gcadUtil.AngleFromXAxis(P2prov, PStrut)
P2(0) = PStrut(0) + 104 * Cos(ANG3): P2(1) = PStrut(1) + 104 * Sin(ANG3): P2(2) = PStrut(2)

       Set Eje1 = gcadModel.AddLine(P1, P2)
       ANG = gcadUtil.AngleFromXAxis(P1, P2)
    
       x = P2(0) - P1(0)
       y = P2(1) - P1(1)
       Xs = 1
       Ys = 1
       Zs = 1
       Distancia = Val(Sqr((x ^ 2 + y ^ 2)))
       
       Set blockRef = gcadModel.InsertBlock(Pbulon1, Bulond19, Xs, Ys, Zs, ANG)
       blockRef.Layer = "Nonplot"
       Set blockRef = gcadModel.InsertBlock(Pbulon2, Bulond19, Xs, Ys, Zs, ANG)
       blockRef.Layer = "Nonplot"
       Set blockRef = gcadModel.InsertBlock(P2, Bulond23, Xs, Ys, Zs, ANG)
       blockRef.Layer = "Nonplot"
        
       Punto_inicial(0) = P1(0): Punto_inicial(1) = P1(1): Punto_inicial(2) = P1(2)
       Punto_final(0) = Punto_inicial(0): Punto_final(1) = Punto_inicial(1): Punto_final(2) = Punto_inicial(2)
    
    
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
    
               n540 = Fix(lpuntal / l540)
               lpuntal = lpuntal - n540 * l540
               n360 = Fix(lpuntal / l360)
               lpuntal = lpuntal - n360 * l360
               n180 = Fix(lpuntal / l180)
               lpuntal = lpuntal - n180 * l180
               lregulacion = (Distancia - n2200 * l2200 - n1100 * l1100 - n550 * l550 - n540 * l540 - n360 * l360 - n180 * l180 + lgatominss - lfija) / 2
           
            'If extremo1 = "B" Then
            'ss_espada = rutass & "SS" & vista & "Espada.dwg"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_espada, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Slims"
            'Punto_inicial(0) = Punto_inicial(0) + lespada * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + lespada * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            'End If
    
            ss_husillo = rutass & "zSSHusillo.dwg"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_husillo, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Slims"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, M24x110, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Nonplot"

            'Punto_inicial(0) = Punto_inicial(0) + lregulacion * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + lregulacion * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            'Punto_final(0) = Punto_inicial(0): Punto_final(1) = Punto_inicial(1): Punto_final(2) = Punto_inicial(2)
        
            ss_gatoizq = rutass & "SS" & vista & "Gatorefizq.dwg"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_gatoizq, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Slims"
            'Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_4, Xs, Ys, Zs, ANG)
            'blockRef.Layer = "Nonplot"
            Dim ss_tubop As String
            ss_tubop = rutator & "TuboPivote" & vista & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_tubop, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Slims"
            
            Punto_inicial(0) = Punto_inicial(0) + 415 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + 415 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
            Punto_final(0) = Punto_inicial(0): Punto_final(1) = Punto_inicial(1): Punto_final(2) = Punto_inicial(2)
            
            Dim ss_antigiro As String
            ss_antigiro = rutatensor & "ESPantigiro" & vista & ".dwg"
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, ss_antigiro, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Slims"
            
            Set blockRef = gcadModel.InsertBlock(Punto_inicial, M16x40_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete

        ' aquí va la espada antigiro en todos los casos
        
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
          
            ss_adaptador = rutass & "SSAdaptador" & vista & ".dwg"
            Set blockRef = gcadModel.InsertBlock(P2, ss_adaptador, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Slims"
            
            P2(0) = P2(0) - ladaptador * Cos(ANG): P2(1) = P2(1) - ladaptador * Sin(ANG): P2(2) = P2(2)
            
            Set blockRef = gcadModel.InsertBlock(P2, ss_husillo, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Slims"
            
            ss_llave = rutass & "SSLlave.dwg"
            Set blockRef = gcadModel.InsertBlock(P2, ss_llave, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Slims"
            
            Set blockRef = gcadModel.InsertBlock(P2, M24x110, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            
        
Dim strut As String
strut = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensores\Strut.dwg"
Set blockRef = gcadModel.InsertBlock(PStrut, strut, Xs, Ys, Zs, ANG3 - (PI) / 2)
blockRef.Layer = "Slims"
    
    
Eje1.Layer = "Nonplot"

terminar:
End Sub

Sub ssc_tensor(punto1 As Variant, punto2 As Variant)
Dim rutass As String
Dim rutat As String, rutags As String
Dim gcadDoc As Object
Dim gcadUtil As Object
Dim gcadModel As Object
Dim x As Double
Dim y As Double
Dim z As Double
Dim M16x40 As String, M24x110 As String, M20x130 As String
Dim VarM16x166 As String
Dim Angulo As String, espadags As String, espadampcorta As String, espadamplarga As String, espadass As String, gato As String, tornillogs1 As String, tornillogs2 As String, husillo As String, Base_drc As String, Base_izq As String
Dim ss_90 As String
Dim ss_180 As String
Dim ss_360 As String
Dim ss_360os As String
Dim ss_540 As String
Dim ss_720 As String
Dim ss_900 As String
Dim ss_1800 As String
Dim ss_2700 As String
Dim ss_3600 As String
Dim langulo As Double
Dim nangulo As Integer
Dim l90 As Double, lespadags As Double, lespadampcorta As Double, lespadamplarga As Double, lespadass As Double, lgato As Double
Dim l180 As Double
Dim l360 As Double
Dim l540 As Double
Dim l720 As Double
Dim l900 As Double
Dim l1800 As Double
Dim l2700 As Double
Dim l3600 As Double
Dim n90 As Integer, nespadags As Integer, nespadampcorta As Integer, nespadamplarga As Integer, nespadass As Integer, n360os As Integer, ngato As Integer
Dim n180 As Integer
Dim n360 As Integer
Dim n540 As Integer
Dim n720 As Integer
Dim n900 As Integer
Dim n1800 As Integer
Dim n2700 As Integer
Dim n3600 As Integer
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
Dim extremo1 As String, extremo2 As String
Dim vigaM As String
Dim vigamenor As String, naranja As String
Dim offsecc As String
Dim lpuntal As Double, laux As Double
Dim plalz As String
Dim kwordList As String
Dim i As Integer
Dim lfija As Double
Dim Ncapaslim As String
Dim capaslim As Object

Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

On Error GoTo terminar
repite = 1

Ncapaslim = "Slims"
Set capaslim = gcadDoc.Layers.Add(Ncapaslim)
capaslim.color = 30

rutass = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\"
rutat = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"

VarM16x166 = rutat & "2VarM16X166.dwg"
M16x40 = rutat & "4-M16X40.dwg"
M24x110 = rutat & "1-M24X110.dwg"
M20x130 = rutat & "1-M20X130.dwg"
Dim Bulond23 As String
Bulond23 = rutat & "1M23_BULOND23.dwg"

'Valores fijos
PI = 4 * Atn(1)
langulo = 119.5
l90 = 90
l180 = 180
l360 = 360
l540 = 540
l720 = 720
l900 = 900
l1800 = 1800
l2700 = 2700
l3600 = 3600
lespadags = 150
lespadass = 358
lespadamplarga = 358
lespadampcorta = 158
lgato = 420

nespadags = 0
nespadampcorta = 0
nespadamplarga = 0
nangulo = 0
n360os = 0
nespadass = 2
ngato = 2


kwordList = "Planta Alzado"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
plalz = ThisDrawing.Utility.GetKeyword(vbLf & "Vista de las Slim en: [Planta/Alzado]")

If plalz = "" Or plalz = "Planta" Then
plalz = "PL"
espadass = rutass & "SSPLEspada.dwg"
ElseIf plalz = "Alzado" Then
plalz = ""
espadass = rutass & "SSEspada.dwg"
Else
GoTo terminar
End If

kwordList = "3600 2700 1800"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vigaM = ThisDrawing.Utility.GetKeyword(vbLf & "Viga de mayor tamaño: [3600/2700/1800]")

kwordList = "90 180 360"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vigamenor = ThisDrawing.Utility.GetKeyword(vbLf & "Viga de menor tamaño: [90/180/360]")

kwordList = "Galvanizada Pintada"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
naranja = ThisDrawing.Utility.GetKeyword(vbLf & "Vigas Slim: [Galvanizada/Pintada]")

If naranja = "" Or naranja = "Galvanizada" Then
naranja = ""
ElseIf naranja = "Pintada" Then
naranja = "N"
Else
GoTo terminar
End If

lfija = nangulo * langulo + n360os * l360 + nespadags * lespadags + nespadampcorta * lespadampcorta + nespadamplarga * lespadamplarga + nespadass * lespadass + ngato * lgato
 




'Geometría:
'punto1 = gcadUtil.GetPoint(, "1º Punto: ")
'punto2 = gcadUtil.GetPoint(punto1, "2º Punto: ")
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
        MsgBox "Medida " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If

lpuntal = Distancia - lfija
If vigaM = "" Then vigaM = "3600"
If vigaM = 3600 Or vigaM = "" Then
n3600 = Fix(lpuntal / l3600)
lpuntal = lpuntal - n3600 * l3600
n2700 = Fix(lpuntal / l2700)
lpuntal = lpuntal - n2700 * l2700
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
ElseIf vigaM = 2700 Then
n3600 = 0
n2700 = Fix(lpuntal / l2700)
lpuntal = lpuntal - n2700 * l2700
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
ElseIf vigaM = 1800 Then
n3600 = 0
n2700 = 0
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
Else: GoTo terminar
End If

n900 = Fix(lpuntal / l900)
lpuntal = lpuntal - n900 * l900
n720 = Fix(lpuntal / l720)
lpuntal = lpuntal - n720 * l720
n540 = Fix(lpuntal / l540)
lpuntal = lpuntal - n540 * l540

If vigamenor = "" Then vigamenor = "90"

If vigamenor = 360 Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = 0
n90 = 0
ElseIf vigamenor = 180 Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = Fix(lpuntal / l180)
lpuntal = lpuntal - n180 * l180
n90 = 0
ElseIf vigamenor = 90 Or vigamenor = "" Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = Fix(lpuntal / l180)
lpuntal = lpuntal - n180 * l180
n90 = Fix(lpuntal / l90)
lpuntal = lpuntal - n90 * l90
Else: GoTo terminar
End If

If ngato = 0 Then
    laux = 0
ElseIf ngato = 1 Then
    If lpuntal > 200 Then
        MsgBox "La abertura del gato " & laux & "mm, es mayor que la máxima admisible de 620mm, el puntal dibujado no llega al segundo extremo solicitado (se dibuja abertura de 620mm)"
        laux = 620
    Else
        laux = lpuntal + lgato
    End If
ElseIf ngato = 2 Then
    laux = lpuntal / 2 + lgato
End If

Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)


    
Set blockRef = gcadModel.InsertBlock(P1, espadass, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"
Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
'Set BlockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
'BlockRef.Layer = "Nonplot"

If plalz = "" Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, Bulond23, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_final(0) + laux * Cos(ANG): Punto_final(1) = Punto_final(1) + laux * Sin(ANG): Punto_final(2) = Punto_final(2)
        Base_izq = rutass & "SSGatorefizq.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, Base_izq, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If plalz = "PL" Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, Bulond23, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_final(0) + laux * Cos(ANG): Punto_final(1) = Punto_final(1) + laux * Sin(ANG): Punto_final(2) = Punto_final(2)
        Base_izq = rutass & "SSPLGatorefizq.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, Base_izq, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n3600 > 0 Then
    i = 0
    Do While i < n3600
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_3600 = rutass & "SS" & plalz & "3600" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_3600, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l3600 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3600 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n2700 > 0 Then
    i = 0
    Do While i < n2700
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_2700 = rutass & "SS" & plalz & "2700" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_2700, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l2700 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l2700 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n1800 > 0 Then
    i = 0
    Do While i < n1800
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_1800 = rutass & "SS" & plalz & "1800" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_1800, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l1800 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1800 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n900 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_900 = rutass & "SS" & plalz & "0900" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_900, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l900 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l900 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n720 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_720 = rutass & "SS" & plalz & "0720" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_720, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l720 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l720 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n540 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_540 = rutass & "SS" & plalz & "0540" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_540, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l540 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l540 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n360 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_360 = rutass & "SS" & plalz & "0360" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_360, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l360 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l360 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n180 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_180 = rutass & "SS" & plalz & "0180" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_180, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n90 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_90 = rutass & "SS" & plalz & "0090" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_90, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If plalz = "" Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Base_drc = rutass & "SSGatorefdrc.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Base_drc, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Slims"
End If

If plalz = "PL" Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Base_drc = rutass & "SSPLGatorefdrc.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Base_drc, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Slims"
End If


Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
Set blockRef = gcadModel.InsertBlock(Punto_final, espadass, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Slims"

Set blockRef = gcadModel.InsertBlock(Punto_inial, Bulond23, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"


Eje1.Layer = "Nonplot"

terminar:
End Sub


Sub ssclargo_tensor(punto1 As Variant, punto2 As Variant)
Dim rutass As String
Dim rutat As String, rutags As String
Dim gcadDoc As Object
Dim gcadUtil As Object
Dim gcadModel As Object
Dim x As Double
Dim y As Double
Dim z As Double
Dim x1 As Double
Dim y1 As Double
Dim z1 As Double
Dim M16x40 As String, M24x110 As String, M20x130 As String
Dim VarM16x166 As String
Dim Angulo As String, espadags As String, espadampcorta As String, espadamplarga As String, espadass As String, gato As String, tornillogs1 As String, tornillogs2 As String, husillo As String, Base_drc As String, Base_izq As String
Dim ss_90 As String
Dim ss_180 As String
Dim ss_360 As String
Dim ss_360os As String
Dim ss_540 As String
Dim ss_720 As String
Dim ss_900 As String
Dim ss_1800 As String
Dim ss_2700 As String
Dim ss_3600 As String
Dim langulo As Double
Dim ladaptadorg As Double
Dim nangulo As Integer
Dim l90 As Double, lespadags As Double, lespadampcorta As Double, lespadamplarga As Double, lespadass As Double, lgato As Double
Dim l180 As Double
Dim l360 As Double
Dim l540 As Double
Dim l720 As Double
Dim l900 As Double
Dim l1800 As Double
Dim l2700 As Double
Dim l3600 As Double
Dim n90 As Integer, nespadags As Integer, nespadampcorta As Integer, nespadamplarga As Integer, nespadass As Integer, n360os As Integer, ngato As Integer
Dim n180 As Integer
Dim n360 As Integer
Dim n540 As Integer
Dim n720 As Integer
Dim n900 As Integer
Dim n1800 As Integer
Dim n2700 As Integer
Dim n3600 As Integer
Dim blockRef As Object
Dim repite As Double
Dim Punto_inial(0 To 2) As Double
Dim Punto_final(0 To 2) As Double
Dim Punto_inial2(0 To 2) As Double
Dim Punto_final2(0 To 2) As Double
Dim PStrut(0 To 2) As Double
Dim PStrut1 As Variant
Dim PStrut2 As Variant
Dim PI As Variant
Dim Eje1 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim ANG As Double
Dim ANG2 As Double
Dim ANG3 As Double
Dim adaptadorg As String
Dim Distancia As Double
Dim PuntoMedio As Double
Dim P2prov(0 To 2) As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim extremo1 As String, extremo2 As String
Dim vigaM As String
Dim vigamenor As String, naranja As String
Dim offsecc As String
Dim lpuntal As Double, laux As Double
Dim plalz As String
Dim kwordList As String
Dim i As Integer
Dim lfija As Double
Dim Ncapaslim As String
Dim capaslim As Object

Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

On Error GoTo terminar
repite = 1

Ncapaslim = "Slims"
Set capaslim = gcadDoc.Layers.Add(Ncapaslim)
capaslim.color = 30

rutass = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\"
rutat = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"

VarM16x166 = rutat & "2VarM16X166.dwg"
M16x40 = rutat & "4-M16X40.dwg"
M24x110 = rutat & "1-M24X110.dwg"
M20x130 = rutat & "1-M20X130.dwg"

Dim Bulond19 As String
Bulond19 = rutat & "1M19_BULOND19.dwg"
Dim Bulond23 As String
Bulond23 = rutat & "1M23_BULOND23.dwg"

'Valores fijos
PI = 4 * Atn(1)
langulo = 119.5
l90 = 90
l180 = 180
l360 = 360
l540 = 540
l720 = 720
l900 = 900
l1800 = 1800
l2700 = 2700
l3600 = 3600
lespadags = 150
lespadass = 358
lespadamplarga = 358
lespadampcorta = 158
lgato = 420
ladaptadorg = 117.25

nespadags = 0
nespadampcorta = 0
nespadamplarga = 0
nangulo = 0
n360os = 0
nespadass = 1
ngato = 2


kwordList = "Planta Alzado"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
plalz = ThisDrawing.Utility.GetKeyword(vbLf & "Vista de las Slim en: [Planta/Alzado]")

If plalz = "" Or plalz = "Planta" Then
plalz = "PL"
espadass = rutass & "SSPLEspada.dwg"
ElseIf plalz = "Alzado" Then
plalz = ""
espadass = rutass & "SSEspada.dwg"
Else
GoTo terminar
End If

kwordList = "3600 2700 1800"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vigaM = ThisDrawing.Utility.GetKeyword(vbLf & "Viga de mayor tamaño: [3600/2700/1800]")

kwordList = "90 180 360"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vigamenor = ThisDrawing.Utility.GetKeyword(vbLf & "Viga de menor tamaño: [90/180/360]")

kwordList = "Galvanizada Pintada"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
naranja = ThisDrawing.Utility.GetKeyword(vbLf & "Vigas Slim: [Galvanizada/Pintada]")

If naranja = "" Or naranja = "Galvanizada" Then
naranja = ""
ElseIf naranja = "Pintada" Then
naranja = "N"
Else
GoTo terminar
End If

lfija = nangulo * langulo + n360os * l360 + nespadags * lespadags + nespadampcorta * lespadampcorta + nespadamplarga * lespadamplarga + nespadass * lespadass + ngato * lgato + ladaptadorg
 


'Geometría:
'punto1 = gcadUtil.GetPoint(, "1º Punto: ")
'punto2 = gcadUtil.GetPoint(punto1, "2º Punto: ")
P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
P2prov(0) = punto2(0): P2prov(1) = punto2(1): P2prov(2) = punto2(2)

PStrut1 = gcadUtil.GetPoint(, "1 anclaje del Strut Adaptor: ")
PStrut2 = gcadUtil.GetPoint(PStrut1, "2 anclaje del Strut Adaptor: ")

Dim Pbulon1(0 To 2) As Double
Dim Pbulon2(0 To 2) As Double

Pbulon1(0) = PStrut1(0): Pbulon1(1) = PStrut1(1): Pbulon1(2) = PStrut1(2)
Pbulon2(0) = PStrut2(0): Pbulon2(1) = PStrut2(1): Pbulon2(2) = PStrut2(2)



ANG2 = gcadUtil.AngleFromXAxis(PStrut2, PStrut1)

PStrut(0) = PStrut2(0) + 90 * Cos(ANG2): PStrut(1) = PStrut2(1) + 90 * Sin(ANG2): PStrut(2) = PStrut2(2)
ANG3 = gcadUtil.AngleFromXAxis(P2prov, PStrut)
P2(0) = PStrut(0) + 104 * Cos(ANG3): P2(1) = PStrut(1) + 104 * Sin(ANG3): P2(2) = PStrut(2)



Set Eje1 = gcadModel.AddLine(P1, P2)
ANG = gcadUtil.AngleFromXAxis(P1, P2)


x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

Set blockRef = gcadModel.InsertBlock(Pbulon1, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Set blockRef = gcadModel.InsertBlock(Pbulon2, Bulond19, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Set blockRef = gcadModel.InsertBlock(P2, Bulond23, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"

If Distancia < lfija Then
        MsgBox "Medida " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If

lpuntal = Distancia - lfija
If vigaM = "" Then vigaM = "3600"
If vigaM = 3600 Or vigaM = "" Then
n3600 = Fix(lpuntal / l3600)
lpuntal = lpuntal - n3600 * l3600
n2700 = Fix(lpuntal / l2700)
lpuntal = lpuntal - n2700 * l2700
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
ElseIf vigaM = 2700 Then
n3600 = 0
n2700 = Fix(lpuntal / l2700)
lpuntal = lpuntal - n2700 * l2700
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
ElseIf vigaM = 1800 Then
n3600 = 0
n2700 = 0
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
Else: GoTo terminar
End If

n900 = Fix(lpuntal / l900)
lpuntal = lpuntal - n900 * l900
n720 = Fix(lpuntal / l720)
lpuntal = lpuntal - n720 * l720
n540 = Fix(lpuntal / l540)
lpuntal = lpuntal - n540 * l540

If vigamenor = "" Then vigamenor = "90"

If vigamenor = 360 Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = 0
n90 = 0
ElseIf vigamenor = 180 Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = Fix(lpuntal / l180)
lpuntal = lpuntal - n180 * l180
n90 = 0
ElseIf vigamenor = 90 Or vigamenor = "" Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = Fix(lpuntal / l180)
lpuntal = lpuntal - n180 * l180
n90 = Fix(lpuntal / l90)
lpuntal = lpuntal - n90 * l90
Else: GoTo terminar
End If

If ngato = 0 Then
    laux = 0
ElseIf ngato = 1 Then
    If lpuntal > 200 Then
        MsgBox "La abertura del gato " & laux & "mm, es mayor que la máxima admisible de 620mm, el puntal dibujado no llega al segundo extremo solicitado (se dibuja abertura de 620mm)"
        laux = 620
    Else
        laux = lpuntal + lgato
    End If
ElseIf ngato = 2 Then
    laux = lpuntal / 2 + lgato
End If

Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)


    
Set blockRef = gcadModel.InsertBlock(P1, espadass, Xs, Ys, Zs, ANG)
blockRef.Layer = "Slims"
Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
'Set BlockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
'BlockRef.Layer = "Nonplot"

If plalz = "" Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, Bulond23, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_final(0) + laux * Cos(ANG): Punto_final(1) = Punto_final(1) + laux * Sin(ANG): Punto_final(2) = Punto_final(2)
        Base_izq = rutass & "SSGatorefizq.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, Base_izq, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
End If

If plalz = "PL" Then
        Set blockRef = gcadModel.InsertBlock(Punto_final, Bulond23, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_final(0) + laux * Cos(ANG): Punto_final(1) = Punto_final(1) + laux * Sin(ANG): Punto_final(2) = Punto_final(2)
        Base_izq = rutass & "SSPLGatorefizq.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, Base_izq, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
End If

If n3600 > 0 Then
    i = 0
    Do While i < n3600
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_3600 = rutass & "SS" & plalz & "3600" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_3600, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l3600 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3600 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
        i = i + 1
    Loop
End If

If n2700 > 0 Then
    i = 0
    Do While i < n2700
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_2700 = rutass & "SS" & plalz & "2700" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_2700, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l2700 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l2700 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
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
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_1800 = rutass & "SS" & plalz & "1800" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_1800, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l1800 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1800 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
        i = i + 1
    Loop
End If

If n900 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_900 = rutass & "SS" & plalz & "0900" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_900, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l900 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l900 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
End If

If n720 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_720 = rutass & "SS" & plalz & "0720" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_720, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l720 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l720 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
End If

If n540 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_540 = rutass & "SS" & plalz & "0540" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_540, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l540 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l540 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
End If

If n360 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_360 = rutass & "SS" & plalz & "0360" & naranja & ".dwg"
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
        ss_180 = rutass & "SS" & plalz & "0180" & naranja & ".dwg"
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
        ss_90 = rutass & "SS" & plalz & "0090" & naranja & ".dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, ss_90, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = gcadModel.InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
End If

If plalz = "" Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Base_drc = rutass & "SSGatorefdrc.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Base_drc, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Slims"
        Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
End If

If plalz = "PL" Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Base_drc = rutass & "SSPLGatorefdrc.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_inial, Base_drc, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = gcadModel.InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Slims"
        Set blockRef = gcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
End If


Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)

Dim ss_llave As String
ss_llave = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSLlave.dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, ss_llave, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Slims"


Punto_final(0) = Punto_inial(0) + 117.25 * Cos(ANG): Punto_final(1) = Punto_inial(1) + 117.25 * Sin(ANG): Punto_final(2) = Punto_inial(2)
adaptadorg = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSAdaptador" & plalz & ".dwg"
Set blockRef = gcadModel.InsertBlock(Punto_final, adaptadorg, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Slims"


Dim strut As String
strut = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensores\Strut.dwg"
Set blockRef = gcadModel.InsertBlock(PStrut, strut, Xs, Ys, Zs, ANG3 - (PI) / 2)
blockRef.Layer = "Slims"


Eje1.Layer = "Nonplot"
MsgBox "Posible necesidad de arriostrar lateralmente"
terminar:
End Sub


Sub Tubo80x40_tensor(punto1 As Variant, punto2 As Variant)
Dim rutags As String
Dim ruta1 As String
Dim husillo As String
Dim ruta2 As String, rutass As String
Dim gcadDoc As Object
Dim gcadUtil As Object
Dim gcadModel As Object
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
Dim Bulond19 As String
Bulond19 = ruta2 & "1M19_BULOND19.dwg"

gato1 = "Ligero"
lgatomin1 = 560

    
gato2 = "Ligero"
lgatomin2 = 560

nespadags = 0
nespadampcorta = 0
nespadamplarga = 0
nespadass = 0



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


'punto1 = gcadUtil.GetPoint(, "1º Punto: ")
'punto2 = gcadUtil.GetPoint(punto1, "2º Punto: ")
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
    Set blockRef = gcadModel.InsertBlock(Punto_final, Tn_Husillo, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Set blockRef = gcadModel.InsertBlock(Punto_final, Bulond19, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
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
    Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)
    Set blockRef = gcadModel.InsertBlock(Punto_final, Bulond19, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
Else
End If


Eje1.Layer = "Nonplot"

terminar:
End Sub












