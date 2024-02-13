
Sub cu()
Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, Eje1 As Object, blockRef As Object
Dim rutags As String, rutamp As String, rutator As String, rutampacc As String, rutacuña As String, capa As String
Dim Gcapa As Object
Dim Ncapa As String, cuña As String, lado As String, tipo1 As String, tipo2 As String, disposicion As String, kwordList As String, M20x90_4 As String
Dim repite As Double, ANG As Double, x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, P1(0 To 2) As Double, P2(0 To 2) As Double, Punto_inial(0 To 2) As Double
Dim punto1 As Variant, punto2 As Variant, PI As Variant



Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
Set gcadModel = gcadDoc.ModelSpace
Set gcadUtil = gcadDoc.Utility

Ncapa = "Mega"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Granshor"
Set Gcapa = gcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150

On Error GoTo terminar

rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"
rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
rutampacc = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"

PI = 4 * Atn(1)

    kwordList = "Naranja Azul Verde Granshor Compacta MuroMP"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    cuña = ThisDrawing.Utility.GetKeyword(vbLf & "Cuña a introducir: [Naranja/Azul/Verde/Granshor/Compacta/MuroMP]")

    kwordList = "Alzado Planta Sección"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    disposicion = ThisDrawing.Utility.GetKeyword(vbLf & "Lado del puntal: [Alzado/Planta/Sección]")

    If cuña = "" Or cuña = "Naranja" Then
        cuña = rutampacc & "MG_CunaNar_"
        capa = "Mega"
    ElseIf cuña = "Azul" Then
        cuña = rutampacc & "MG_CunaAz_"
        capa = "Mega"
    ElseIf cuña = "Verde" Then
        cuña = rutampacc & "MG_CunaVe_"
        capa = "Mega"
    ElseIf cuña = "Granshor" Then
        cuña = rutags & "GS_PlacaAnclaje_"
        capa = "Granshor"
    ElseIf cuña = "Compacta" Then
        cuña = rutags & "GS_Placacompacta_"
        capa = "Granshor"
    ElseIf cuña = "MuroMP" Then
        cuña = rutampacc & "PlacaMP_"
        capa = "Mega"
        kwordList = "Compacta Normal"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        tipo1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de cuña: [Compacta/Normal]")
        If tipo1 = "Compacta" Then
            cuña = cuña & "C_"
        ElseIf tipo1 = "" Or tipo1 = "Normal" Then
            cuña = cuña
        End If
    Else
        GoTo terminar
    End If

    If disposicion = "" Or disposicion = "Alzado" Then
        disposicion = "alzado"

        kwordList = "Derecha Izquierda"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        lado = ThisDrawing.Utility.GetKeyword(vbLf & "Lado del puntal: [Derecha/Izquierda]")

        If lado = "" Or lado = "Derecha" Then
            cuña = cuña & "Dalzado.dwg"
        ElseIf lado = "Izquierda" Then
            cuña = cuña & "Ialzado.dwg"
        Else
            GoTo terminar
        End If

    ElseIf disposicion = "Planta" Then
        cuña = cuña & "planta.dwg"
    ElseIf disposicion = "Sección" Then
        cuña = cuña & "seccion.dwg"
    Else
        GoTo terminar
    End If


repite = 1
Do While repite = 1

        punto1 = gcadUtil.GetPoint(, "1º Punto: ")
        punto2 = gcadUtil.GetPoint(punto1, "2º Punto: ")
        P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
        P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)

        'Set Eje1 = gcadModel.AddLine(P1, P2)
        ANG = gcadUtil.AngleFromXAxis(P1, P2)

        x = P2(0) - P1(0)
        y = P2(1) - P1(1)
        Xs = 1
        Ys = 1
        Zs = 1

    M20x90_4 = rutator & "4-M20X90.dwg"
    
        Set blockRef = gcadModel.InsertBlock(P1, cuña, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa


    If cuña = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\MG_CunaNar_Dalzado.dwg" Then
        Punto_inial(0) = P1(0) + 115.56 * Cos(ANG) + 175.56 * Cos(ANG + PI / 2): Punto_inial(1) = P1(1) + 115.56 * Sin(ANG) + 175.56 * Sin(ANG + PI / 2): Punto_inial(2) = P1(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    ElseIf cuña = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\MG_CunaNar_Ialzado.dwg" Then
        Punto_inial(0) = P1(0) + 115.56 * Cos(ANG) + 175.56 * Cos(ANG - PI / 2): Punto_inial(1) = P1(1) + 115.56 * Sin(ANG) + 175.56 * Sin(ANG - PI / 2): Punto_inial(2) = P1(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    ElseIf cuña = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\MG_CunaAz_Dalzado.dwg" Then
        Punto_inial(0) = P1(0) + 150.06 * Cos(ANG) + 225.06 * Cos(ANG + PI / 2): Punto_inial(1) = P1(1) + 150.06 * Sin(ANG) + 225.06 * Sin(ANG + PI / 2): Punto_inial(2) = P1(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    ElseIf cuña = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\MG_CunaAz_Ialzado.dwg" Then
        Punto_inial(0) = P1(0) + 150.06 * Cos(ANG) + 225.06 * Cos(ANG - PI / 2): Punto_inial(1) = P1(1) + 150.06 * Sin(ANG) + 225.06 * Sin(ANG - PI / 2): Punto_inial(2) = P1(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    ElseIf cuña = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\MG_CunaVe_Dalzado.dwg" Then
        Punto_inial(0) = P1(0) + 150.06 * Cos(ANG) + 225.06 * Cos(ANG + PI / 2): Punto_inial(1) = P1(1) + 150.06 * Sin(ANG) + 225.06 * Sin(ANG + PI / 2): Punto_inial(2) = P1(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    ElseIf cuña = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\MG_CunaVe_Ialzado.dwg" Then
        Punto_inial(0) = P1(0) + 150.06 * Cos(ANG) + 225.06 * Cos(ANG - PI / 2): Punto_inial(1) = P1(1) + 150.06 * Sin(ANG) + 225.06 * Sin(ANG - PI / 2): Punto_inial(2) = P1(2)
        Set blockRef = gcadModel.InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    
    End If
Loop
terminar:
End Sub


