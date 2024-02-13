Public Sub GetUserSelection()
    
    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object
    Dim ss As GcadSelectionSet
    Dim block As GcadBlockReference
            
    Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility

    On Error Resume Next
    Set ss = gcadDoc.SelectionSets("SS1")
    On Error GoTo 0

    If ss Is Nothing Then
        Set ss = gcadDoc.SelectionSets.Add("SS1")
    Else
        ss.Clear
    End If

    ss.SelectOnScreen
    
    If ss.Count = 0 Then
        MsgBox "No block selected. Operation cancelled."
        ss.Delete
        Exit Sub
    End If

    'Return first selected block reference
    Set block = ss(0)
    
    Set GetUserSelection = block
    
End Sub


Public Sub sust()
    On Error GoTo terminar

    Dim tu As String
    kwordList = "PS SS Tubo80 MP GS"
    tu = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    tu = ThisDrawing.Utility.GetKeyword(vbLf & "Familia de materiales a sustituir: [PS/SS/Tubo80/MP/GS]")
    
    If tu = "PS" Then
        Call DivPS
    ElseIf tu = "SS" Then
        Call DivSS
    ElseIf tu = "Tubo80" Then
        Call DivTN
    ElseIf tu = "GS" Then
        Call DivGS
    'ElseIf tu = "Lola" Then
        'Call DivLS
    ElseIf tu = "MP" Then
        Call DivMP
    End If
terminar:
    
End Sub

Sub DivPS()

    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, blockRef As Object
    Dim M20x90_16 As String
    Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    ' Paso 1: Crear un nuevo conjunto de selección
    Dim ss As GcadSelectionSet
    On Error Resume Next
    Set ss = ThisDrawing.SelectionSets("SS1")
    On Error GoTo 0

    If ss Is Nothing Then
        Set ss = ThisDrawing.SelectionSets.Add("SS1")
    Else
        ss.Clear
    End If

    ' Solicitar al usuario que seleccione los bloques en la pantalla
    ss.SelectOnScreen
    

    If ss.Count = 0 Then
        MsgBox "No se seleccionó ningún bloque. La operación ha sido cancelada."
        ss.Delete
        Exit Sub
    End If
    
    ' Paso 3: Obtener el bloque seleccionado
    Dim block As GcadBlockReference
    Dim block_inicial As GcadBlockReference
    Set block_inicial = ss.Item(0)
    
    Dim nombre_inicial As String
    nombre_inicial = block_inicial.effectiveName
    
        rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
        
        M20x90_16 = rutator & "16-M20X90.dwg"

    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '                       DECISIÓN PIPESHOR 4L
    '
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim ui_pl6000 As String, ui_pl4500 As String, ui_pl3000 As String, ui_pl1500 As String
    
    If nombre_inicial = "PL_6000_alzado" Or nombre_inicial = "PL_6000_planta" Then
        ui_pl6000 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  4500 + 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 3000 + 3000" & vbCrLf & vbCrLf & vbCrLf & "3). 4 x 1500")
    ElseIf nombre_inicial = "PL_4500_planta" Or nombre_inicial = "PL_4500_alzado" Then
        ui_pl4500 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  3000 + 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 3 x 1500" & vbCrLf & vbCrLf & vbCrLf & "3). 3000 + 750 + 750" & vbCrLf & vbCrLf & vbCrLf & "4. 6 x 750")
    ElseIf nombre_inicial = "PL_3000_planta" Or nombre_inicial = "PL_3000_alzado" Then
        ui_pl3000 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  2 x 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 4 x 750")
    ElseIf nombre_inicial = "PL_1500_Planta" Or nombre_inicial = "PL_1500_alzado" Then
        ui_pl1500 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  2 x 750")
    End If

    ' BLOQUES -------------
    Dim b_pl4500_al As String, b_pl4500_pl As String, b_pl3000_al As String, b_pl3000_pl As String, b_pl1500_al As String, b_pl1500_pl As String, b_pl750_pl As String, b_pl750_al As String
    
    b_pl4500_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_4500_alzado.dwg"
    b_pl4500_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_4500_planta.dwg"
    b_pl3000_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_3000_alzado.dwg"
    b_pl3000_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_3000_planta.dwg"
    b_pl1500_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_1500_alzado.dwg"
    b_pl1500_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_1500_planta.dwg"
    b_pl750_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_750_alzado.dwg"
    b_pl750_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_750_planta.dwg"
    
         
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '                       DECISIÓN PIPESHOR 4S
    '
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim ui_ps6000 As String, ui_ps4500 As String, ui_ps3000 As String, ui_ps1500 As String
    
    If nombre_inicial = "PS_6000_alzado" Or nombre_inicial = "PS_6000_planta" Then
        ui_ps6000 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  4500 + 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 3000 + 3000" & vbCrLf & vbCrLf & vbCrLf & "3). 4 x 1500")
    ElseIf nombre_inicial = "PS_4500_planta" Or nombre_inicial = "PS_4500_alzado" Then
        ui_ps4500 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  3000 + 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 3 x 1500" & vbCrLf & vbCrLf & vbCrLf & "3). 3000 + 750 + 750" & vbCrLf & vbCrLf & vbCrLf & "4. 6 x 750")
    ElseIf nombre_inicial = "PS_3000_planta" Or nombre_inicial = "PS_3000_alzado" Then
        ui_ps3000 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  2 x 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 4 x 750")
    ElseIf nombre_inicial = "PS_1500_planta" Or nombre_inicial = "PS_1500_alzado" Then
        ui_ps1500 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  2 x 750")
    End If
    
    ' BLOQUES -------------
    Dim b_ps4500_al As String, b_ps4500_pl As String, b_ps3000_al As String, b_ps3000_pl As String, b_ps1500_al As String, b_ps1500_pl As String, b_ps750_pl As String, b_ps750_al As String, b_ps560_pl As String, b_ps560_al As String, b_ps280_pl As String, b_ps280_al As String
    
    b_ps4500_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_4500_alzado.dwg"
    b_ps4500_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_4500_planta.dwg"
    b_ps3000_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_3000_alzado.dwg"
    b_ps3000_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_3000_planta.dwg"
    b_ps1500_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_1500_alzado.dwg"
    b_ps1500_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_1500_planta.dwg"
    b_ps750_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_750_alzado.dwg"
    b_ps750_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_750_planta.dwg"
    b_ps560_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_560.dwg"
    b_ps560_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\PS_560.dwg"
    b_ps280_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_280_alzado.dwg"
    b_ps280_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\PL_280_planta.dwg"
    
    
    'Dim i As Long
    'i = 0
    'For i = 0 To ss.Count
       ' Set block = ss.Item(i)
    Dim obj As GcadEntity
    For Each obj In ss
        If TypeOf obj Is GcadBlockReference Then
            Set block = obj
        ElseIf TypeOf obj Is GcadLine Then
            GoTo terminar
        End If
        Dim effectiveName As String
        effectiveName = block.effectiveName
        
        Dim insertionPoint As Variant
        insertionPoint = block.insertionPoint
        
        Dim orientation As Double
        orientation = block.Rotation
        
        Dim Xs As Integer, Ys As Integer, Zs As Integer
    
        Xs = 1
        Ys = 1
        Zs = 1
        
        '------------------------------------- PIP 4L ----------------------------------------------------------------------------------------------------------------------
        ' Pipeshor 4L 6000
        If effectiveName = "PL_6000_alzado" Then
            block.Delete
            If ui_pl6000 = "1" Or ui_pl6000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl4500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 4500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 4500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl6000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl6000 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "PL_6000_planta" Then
            block.Delete
            If ui_pl6000 = "1" Or ui_pl6000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl4500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 4500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 4500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl6000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl6000 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' Pipeshor 4L 4500
        ElseIf effectiveName = "PL_4500_planta" Then
            block.Delete
            If ui_pl4500 = "1" Or ui_pl4500 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl4500 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl4500 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl4500 = "4" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "PL_4500_alzado" Then
            block.Delete
            If ui_pl4500 = "1" Or ui_pl4500 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl4500 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl4500 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl4500 = "4" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' Pipeshor 4L 3000
        ElseIf effectiveName = "PL_3000_planta" Then
            block.Delete
            If ui_pl3000 = "1" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl3000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "PL_3000_alzado" Then
            block.Delete
            If ui_pl3000 = "1" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_pl3000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' Pipeshor 4L 1500
        ElseIf effectiveName = "PL_1500_alzado" Then
            block.Delete
            If ui_pl1500 = "1" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "PL_1500_Planta" Then
            block.Delete
            If ui_pl1500 = "1" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_pl750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4L"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        '----------------------------------------------------- PIP 4S ----------------------------------------------------------------------------------------------------
        ' Pipeshor 4S 6000
        ElseIf effectiveName = "PS_6000_alzado" Then
            block.Delete
            If ui_ps6000 = "1" Or ui_ps6000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps4500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 4500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 4500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps6000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps6000 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "PS_6000_planta" Then
            block.Delete
            If ui_ps6000 = "1" Or ui_ps6000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps4500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 4500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 4500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps6000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps6000 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If

        ' Pipeshor 4S 4500
        ElseIf effectiveName = "PS_4500_planta" Then
            block.Delete
            If ui_ps4500 = "1" Or ui_ps4500 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps4500 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps4500 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps4500 = "4" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "PS_4500_alzado" Then
            block.Delete
            If ui_ps4500 = "1" Or ui_ps4500 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps4500 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps4500 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps4500 = "4" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' Pipeshor 4S 3000
        ElseIf effectiveName = "PS_3000_planta" Then
            block.Delete
            If ui_ps3000 = "1" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps3000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "PS_3000_alzado" Then
            block.Delete
            If ui_ps3000 = "1" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ps3000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' Pipeshor 4S 1500
        ElseIf effectiveName = "PS_1500_alzado" Then
            block.Delete
            If ui_ps1500 = "1" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "PS_1500_planta" Then
            block.Delete
            If ui_ps1500 = "1" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Pipeshor4S"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' Pipeshor 4S 560
        ElseIf effectiveName = "PS_560" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps280_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Pipeshor4S"
            insertionPoint(0) = insertionPoint(0) + 280 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 280 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps280_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Pipeshor4S"
            insertionPoint(0) = insertionPoint(0) + 280 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 280 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        ElseIf effectiveName = "PS_560_Planta" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps280_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Pipeshor4S"
            insertionPoint(0) = insertionPoint(0) + 280 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ps280_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Pipeshor4S"
            insertionPoint(0) = insertionPoint(0) + 280 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        End If
terminar:
    Next obj
    
    

    
    ' Paso 7: Limpiar el conjunto de selección al finalizar
    ss.Delete
End Sub

Sub DivSS()

    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, blockRef As Object
    Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    ' Paso 1: Crear un nuevo conjunto de selección
    Dim ss As GcadSelectionSet
    On Error Resume Next
    Set ss = ThisDrawing.SelectionSets("SS1")
    On Error GoTo 0

    If ss Is Nothing Then
        Set ss = ThisDrawing.SelectionSets.Add("SS1")
    Else
        ss.Clear
    End If

    ' Solicitar al usuario que seleccione los bloques en la pantalla
    ss.SelectOnScreen
    

    If ss.Count = 0 Then
        MsgBox "No se seleccionó ningún bloque. La operación ha sido cancelada."
        ss.Delete
        Exit Sub
    End If
    
    ' Paso 3: Obtener el bloque seleccionado
    Dim block As GcadBlockReference
    Dim block_inicial As GcadBlockReference
    Set block_inicial = ss.Item(0)
    
    Dim nombre_inicial As String
    nombre_inicial = block_inicial.effectiveName

    Dim b_ss3600 As String, b_ss3600n As String, b_sspl3600 As String, b_sspl3600n As String
    Dim b_ss2700 As String, b_ss2700n As String, b_sspl2700 As String, b_sspl2700n As String
    Dim b_ss1800 As String, b_ss1800n As String, b_sspl1800 As String, b_sspl1800n As String
    Dim b_ss0900 As String, b_ss0900n As String, b_sspl0900 As String, b_sspl0900n As String
    Dim b_ss0720 As String, b_ss0720n As String, b_sspl0720 As String, b_sspl0720n As String
    Dim b_ss0540 As String, b_ss0540n As String, b_sspl0540 As String, b_sspl0540n As String
    Dim b_ss0360 As String, b_ss0360n As String, b_sspl0360 As String, b_sspl0360n As String
    Dim b_ss0180 As String, b_ss0180n As String, b_sspl0180 As String, b_sspl0180n As String
    Dim b_ss0090 As String, b_ss0090n As String, b_sspl0090 As String, b_sspl0090n As String
    Dim M16x40 As String

    ' BLOQUES -------------
    b_ss3600 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS3600.dwg"
    b_ss3600n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS3600N.dwg"
    b_sspl3600 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL3600.dwg"
    b_sspl3600n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL3600N.dwg"
    b_ss2700 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS2700.dwg"
    b_ss2700n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS2700N.dwg"
    b_sspl2700 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL2700.dwg"
    b_sspl2700n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL2700N.dwg"
    b_ss1800 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS1800.dwg"
    b_ss1800n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS1800N.dwg"
    b_sspl1800 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL1800.dwg"
    b_sspl1800n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL1800N.dwg"
    b_ss0900 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0900.dwg"
    b_ss0900n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0900N.dwg"
    b_sspl0900 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0900.dwg"
    b_sspl0900n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0900N.dwg"
    b_ss0720 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0720.dwg"
    b_ss0720n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0720N.dwg"
    b_sspl0720 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0720.dwg"
    b_sspl0720n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0720N.dwg"
    b_ss0540 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0540.dwg"
    b_ss0540n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0540N.dwg"
    b_sspl0540 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0540.dwg"
    b_sspl0540n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0540N.dwg"
    b_ss0360 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0360.dwg"
    b_ss0360n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0360N.dwg"
    b_sspl0360 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0360.dwg"
    b_sspl0360n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0360N.dwg"
    b_ss0180 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0180.dwg"
    b_ss0180n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0180N.dwg"
    b_sspl0180 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0180.dwg"
    b_sspl0180n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0180N.dwg"
    b_ss0090 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0090.dwg"
    b_ss0090n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SS0090N.dwg"
    b_sspl0090 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0090.dwg"
    b_sspl0090n = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\SSPL0090N.dwg"
    rutat = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    
    
    M16x40 = rutat & "4-M16X40.dwg"



    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '                       DECISIÓN SUPERSLIM
    '
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim ui_ss3600 As String, ui_ss2700 As String, ui_ss1800 As String, ui_ss900 As String, ui_ss720 As String, ui_ss540 As String, ui_ss360 As String, ui_ss180 As String
    
    
    If nombre_inicial = "SS3600" Or nombre_inicial = "SSPL3600" Or nombre_inicial = "SS3600N" Or nombre_inicial = "SSPL3600N" Then
        ui_ss3600 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  2700 + 900" & vbCrLf & vbCrLf & vbCrLf & "2). 2 x 1800" & vbCrLf & vbCrLf & vbCrLf & "3). 4 x 900")
    ElseIf nombre_inicial = "SS2700" Or nombre_inicial = "SSPL2700" Or nombre_inicial = "SS2700N" Or nombre_inicial = "SSPL2700N" Then
        ui_ss2700 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  1800 + 900" & vbCrLf & vbCrLf & vbCrLf & "2). 3 x 900" & vbCrLf & vbCrLf & vbCrLf & "3). 3 x 720 + 540")
    ElseIf nombre_inicial = "SS1800" Or nombre_inicial = "SSPL1800" Or nombre_inicial = "SS1800N" Or nombre_inicial = "SSPL1800N" Then
        ui_ss1800 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  2 x 900" & vbCrLf & vbCrLf & vbCrLf & "2). 720 + 2 x 540")
    ElseIf nombre_inicial = "SS0900" Or nombre_inicial = "SSPL0900" Or nombre_inicial = "SS0900N" Or nombre_inicial = "SSPL0900N" Then
        ui_ss900 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  720 + 180" & vbCrLf & vbCrLf & vbCrLf & "2). 2 x 360 + 180" & vbCrLf & vbCrLf & vbCrLf & "3). 720 + 2 x 90")
    ElseIf nombre_inicial = "SS0720" Or nombre_inicial = "SS0720N" Or nombre_inicial = "SSPL0720" Or nombre_inicial = "SSPL0720N" Then
        ui_ss720 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  2 x 360" & vbCrLf & vbCrLf & vbCrLf & "2). 540 + 180")
    ElseIf nombre_inicial = "SS0540" Or nombre_inicial = "SSPL0540" Or nombre_inicial = "SS0540N" Or nombre_inicial = "SSPL0540N" Then
        ui_ss540 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  360 + 180" & vbCrLf & vbCrLf & vbCrLf & "2). 360 + 2 x 90" & vbCrLf & vbCrLf & vbCrLf & "3). 3 x 180")
    ElseIf nombre_inicial = "SS0360" Or nombre_inicial = "SSPL0360" Or nombre_inicial = "SS0360N" Or nombre_inicial = "SSPL0360N" Then
        ui_ss360 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1).  2 x 180" & vbCrLf & vbCrLf & vbCrLf & "2). 4 x 90")
    End If
    
    



    Dim obj As GcadEntity
    For Each obj In ss
        If TypeOf obj Is GcadBlockReference Then
            Set block = obj
        ElseIf TypeOf obj Is GcadLine Then
            GoTo terminar
        End If
        Dim effectiveName As String
        effectiveName = block.effectiveName
        
        Dim insertionPoint As Variant
        insertionPoint = block.insertionPoint
        
        Dim orientation As Double
        orientation = block.Rotation
        
        Dim Xs As Integer, Ys As Integer, Zs As Integer
    
        Xs = 1
        Ys = 1
        Zs = 1




        '--------------- SUPERSLIM -----------------------------------------------------------------------------------------------------
        ' SuperSlim 3600 GALVANIZADA -----------
        If effectiveName = "SS3600" Then
            block.Delete
            If ui_ss3600 = "1" Or ui_ss3600 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss2700, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss3600 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss1800, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss1800, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss3600 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SS3600N" Then ' NARANJA 3600
            block.Delete
            If ui_ss3600 = "1" Or ui_ss3600 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss2700n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss3600 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss1800n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss1800n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss3600 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL3600" Then ' GALVANIZADA 3600 PLANTA
            block.Delete
            If ui_ss3600 = "1" Or ui_ss3600 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl2700, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss3600 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl1800, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl1800, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss3600 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL3600N" Then ' NARANJA 3600 PLANTA
            block.Delete
            If ui_ss3600 = "1" Or ui_ss3600 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl2700n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss3600 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl1800n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl1800n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss3600 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
            
        ' SuperSlim 2700 GALVANIZADA -----------
        ElseIf effectiveName = "SS2700" Then
            block.Delete
            If ui_ss2700 = "1" Or ui_ss2700 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss1800, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss2700 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss2700 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0540, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SS2700N" Then ' NARANJA 2700
            block.Delete
            If ui_ss2700 = "1" Or ui_ss2700 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss1800n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss2700 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss2700 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0540n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL2700" Then ' GALVANIZADA 2700 PLANTA
            block.Delete
            If ui_ss2700 = "1" Or ui_ss2700 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl1800, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss2700 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss2700 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0540, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL2700N" Then ' NARANJA 2700 PLANTA
            block.Delete
            If ui_ss2700 = "1" Or ui_ss2700 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl1800n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss2700 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss2700 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0540n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' SuperSlim 1800 GALVANIZADA -----------
        ElseIf effectiveName = "SS1800" Then
            block.Delete
            If ui_ss1800 = "1" Or ui_ss1800 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss1800 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0540, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0540, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SS1800N" Then ' NARANJA 1800
            block.Delete
            If ui_ss1800 = "1" Or ui_ss1800 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss1800 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0540n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0540n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL1800" Then ' GALVANIZADA 1800 PLANTA
            block.Delete
            If ui_ss1800 = "1" Or ui_ss1800 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss1800 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0540, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0540, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL1800N" Then ' NARANJA 1800 PLANTA
            block.Delete
            If ui_ss1800 = "1" Or ui_ss1800 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0900n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss1800 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0540n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0540n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' SuperSlim 900 GALVANIZADA ----------------------
        ElseIf effectiveName = "SS0900" Then
            block.Delete
            If ui_ss900 = "1" Or ui_ss900 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss900 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss900 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SS0900N" Then ' NARANJA 900
            block.Delete
            If ui_ss900 = "1" Or ui_ss900 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss900 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss900 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL0900" Then ' GALVANIZADA 900 PLANTA
            block.Delete
            If ui_ss900 = "1" Or ui_ss900 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss900 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss900 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL0900N" Then ' NARANJA 900 PLANTA
            block.Delete
            If ui_ss900 = "1" Or ui_ss900 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss900 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss900 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0720n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 720 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 720 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' SuperSlim 720 GALVANIZADA -----------
        ElseIf effectiveName = "SS0720" Then
            block.Delete
            If ui_ss720 = "1" Or ui_ss720 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss720 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0540, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SS0720N" Then ' NARANJA 720
            block.Delete
            If ui_ss720 = "1" Or ui_ss720 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss720 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0540n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL0720" Then ' GALVANIZADA 720 PLANTA
            block.Delete
            If ui_ss720 = "1" Or ui_ss720 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss720 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0540, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL0720N" Then ' NARANJA 720 PLANTA
            block.Delete
            If ui_ss720 = "1" Or ui_ss720 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss720 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0540n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 540 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 540 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' SuperSlim 540 GALVANIZADA -----------
        ElseIf effectiveName = "SS0540" Then
            block.Delete
            If ui_ss540 = "1" Or ui_ss540 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss540 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SS0540N" Then ' NARANJA 540
            block.Delete
            If ui_ss540 = "1" Or ui_ss540 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss540 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL0540" Then ' GALVANIZADA 540 PLANTA
            block.Delete
            If ui_ss540 = "1" Or ui_ss540 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss540 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                çblockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL0540N" Then ' NARANJA 540 PLANTA
            block.Delete
            If ui_ss540 = "1" Or ui_ss540 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss540 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0360n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 360 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 360 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' SuperSlim 360 GALVANIZADA -----------
        ElseIf effectiveName = "SS0360" Then
            block.Delete
            If ui_ss360 = "1" Or ui_ss360 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss360 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SS0360N" Then ' NARANJA 360
            block.Delete
            If ui_ss360 = "1" Or ui_ss360 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss360 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL0360" Then ' GALVANIZADA 360 PLANTA
            block.Delete
            If ui_ss360 = "1" Or ui_ss360 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss360 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "SSPL0360N" Then ' NARANJA 360 PLANTA
            block.Delete
            If ui_ss360 = "1" Or ui_ss360 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0180n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ss360 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' SuperSlim 180 GALVANIZADA -----------
        ElseIf effectiveName = "SS0180" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        ElseIf effectiveName = "SS0180N" Then ' NARANJA 180
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ss0090n, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        ElseIf effectiveName = "SSPL0180" Then ' GALVANIZADA 180 PLANTA
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        ElseIf effectiveName = "SSPL0180N" Then ' NARANJA 180 PLANTA
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_sspl0090n, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)

        End If
terminar:
    Next obj
    
    ss.Delete
End Sub

Sub DivTN()

    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, blockRef As Object
    Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    ' Paso 1: Crear un nuevo conjunto de selección
    Dim ss As GcadSelectionSet
    
    Dim M16x40 As String
    
    On Error Resume Next
    Set ss = ThisDrawing.SelectionSets("SS1")
    On Error GoTo 0

    If ss Is Nothing Then
        Set ss = ThisDrawing.SelectionSets.Add("SS1")
    Else
        ss.Clear
    End If

    ' Solicitar al usuario que seleccione los bloques en la pantalla
    ss.SelectOnScreen
    

    If ss.Count = 0 Then
        MsgBox "No se seleccionó ningún bloque. La operación ha sido cancelada."
        ss.Delete
        Exit Sub
    End If
    
    ' Paso 3: Obtener el bloque seleccionado
    Dim block As GcadBlockReference
    Dim block_inicial As GcadBlockReference
    Set block_inicial = ss.Item(0)
    
    Dim nombre_inicial As String
    nombre_inicial = block_inicial.effectiveName
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '                       DECISIÓN TUBO 80
    '
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim ui_tn3200 As String, ui_tn2800 As String, ui_tn2400 As String, ui_tn2000 As String, ui_tn1600 As String, ui_tn800 As String
    
    If nombre_inicial = "Tensor80x4_3200" Or nombre_inicial = "Tensor80x4PL_3200" Then
        ui_tn3200 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2 x 1600" & vbCrLf & vbCrLf & vbCrLf & "2). 2800 + 400" & vbCrLf & vbCrLf & vbCrLf & "3). 2400 + 800")
    ElseIf nombre_inicial = "Tensor80x4_2800" Or nombre_inicial = "Tensor80x4PL_2800" Then
        ui_tn2800 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2400 + 400" & vbCrLf & vbCrLf & vbCrLf & "2). 2000 + 800" & vbCrLf & vbCrLf & vbCrLf & "3). 1600 + 800 + 400")
    ElseIf nombre_inicial = "Tensor80x4_2400" Or nombre_inicial = "Tensor80x4PL_2400" Then
        ui_tn2400 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2000 + 400" & vbCrLf & vbCrLf & vbCrLf & "2). 1600 + 800" & vbCrLf & vbCrLf & vbCrLf & "3). 3 x 800")
    ElseIf nombre_inicial = "Tensor80x4_2000" Or nombre_inicial = "Tensor80x4PL_2000" Then
        ui_tn2000 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 1600 + 400" & vbCrLf & vbCrLf & vbCrLf & "2). 2 x 800 + 400")
    ElseIf nombre_inicial = "Tensor80x4_1600" Or nombre_inicial = "Tensor80x4PL_1600" Then
        ui_tn1600 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2 x 800" & vbCrLf & vbCrLf & vbCrLf & "2). 4 x 400")
    End If

    ' BLOQUES -------------
    Dim b_tn2800_al As String, b_tn2800_pl As String, b_tn2400_al As String, b_tn2400_pl As String, b_tn2000_al As String, b_tn2000_pl As String, b_tn1600_pl As String, b_tn1600_al As String, b_tn800_pl As String, b_tn800_al As String, b_tn400_pl As String, b_tn400_al As String
    
    b_tn2800_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4_2800.dwg"
    b_tn2800_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4PL_2800.dwg"
    b_tn2400_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4_2400.dwg"
    b_tn2400_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4PL_2400.dwg"
    b_tn2000_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4_2000.dwg"
    b_tn2000_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4PL_2000.dwg"
    b_tn1600_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4_1600.dwg"
    b_tn1600_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4PL_1600.dwg"
    b_tn800_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4_800.dwg"
    b_tn800_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4PL_800.dwg"
    b_tn400_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4_400.dwg"
    b_tn400_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Tensor80x4\Tensor80x4PL_400.dwg"
    ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    
    M16x40 = ruta2 & "4-M16X40.dwg"

    'Dim i As Long
    'i = 0
    'For i = 0 To ss.Count
       ' Set block = ss.Item(i)
    Dim obj As GcadEntity
    For Each obj In ss
        If TypeOf obj Is GcadBlockReference Then
            Set block = obj
        ElseIf TypeOf obj Is GcadLine Then
            GoTo terminar
        End If
        Dim effectiveName As String
        effectiveName = block.effectiveName
        
        Dim insertionPoint As Variant
        insertionPoint = block.insertionPoint
        
        Dim orientation As Double
        orientation = block.Rotation
        
        Dim Xs As Integer, Ys As Integer, Zs As Integer
    
        Xs = 1
        Ys = 1
        Zs = 1
        
        '------------------------------------- TUBO 80 ----------------------------------------------------------------------------------------------------------------------
        ' Tubo 3200
        If effectiveName = "Tensor80x4_3200" Then
            block.Delete
            If ui_tn3200 = "1" Or ui_tn3200 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn3200 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn3200 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Tensor80x4PL_3200" Then
            block.Delete
            If ui_tn3200 = "1" Or ui_tn3200 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn3200 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn3200 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' Tubo 2800
        ElseIf effectiveName = "Tensor80x4_2800" Then
            block.Delete
            If ui_tn2800 = "1" Or ui_tn2800 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2800 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2800 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Tensor80x4PL_2800" Then
            block.Delete
            If ui_tn2800 = "1" Or ui_tn2800 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2800 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2800 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' Tubo 2400
        ElseIf effectiveName = "Tensor80x4_2400" Then
            block.Delete
            If ui_tn2400 = "1" Or ui_tn2400 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2400 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2400 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Tensor80x4PL_2400" Then
            block.Delete
            If ui_tn2400 = "1" Or ui_tn2400 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn2000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 2000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2400 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2400 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' Tubo 2000
        ElseIf effectiveName = "Tensor80x4_2000" Then
            block.Delete
            If ui_tn2000 = "1" Or ui_tn2000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Tensor80x4PL_2000" Then
            block.Delete
            If ui_tn2000 = "1" Or ui_tn2000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn1600_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 1600 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1600 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn2000 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            

        ' Tubo 1600
        ElseIf effectiveName = "Tensor80x4_1600" Then
            block.Delete
            If ui_tn1600 = "1" Or ui_tn1600 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                blockRef.Layer = "Slims"
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_al, Xs, Ys, Zs, orientation)
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn1600 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Tensor80x4PL_1600" Then
            block.Delete
            If ui_tn1600 = "1" Or ui_tn1600 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_tn1600 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Slims"
                insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' Tubo 800
        ElseIf effectiveName = "Tensor80x4_800" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        ElseIf effectiveName = "Tensor80x4PL_800" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M16x40, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_tn400_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Slims"
            insertionPoint(0) = insertionPoint(0) + 400 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 400 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        End If
terminar:
    Next obj
    
    

    
    ' Paso 7: Limpiar el conjunto de selección al finalizar
    ss.Delete
End Sub

Sub DivGS()

    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, blockRef As Object
    Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    
    Dim M20x60 As String
    
    
    ' Paso 1: Crear un nuevo conjunto de selección
    Dim ss As GcadSelectionSet
    On Error Resume Next
    Set ss = ThisDrawing.SelectionSets("SS1")
    On Error GoTo 0

    If ss Is Nothing Then
        Set ss = ThisDrawing.SelectionSets.Add("SS1")
    Else
        ss.Clear
    End If

    ' Solicitar al usuario que seleccione los bloques en la pantalla
    ss.SelectOnScreen
    

    If ss.Count = 0 Then
        MsgBox "No se seleccionó ningún bloque. La operación ha sido cancelada."
        ss.Delete
        Exit Sub
    End If
    
    ' Paso 3: Obtener el bloque seleccionado
    Dim block As GcadBlockReference
    Dim block_inicial As GcadBlockReference
    Set block_inicial = ss.Item(0)
    
    Dim nombre_inicial As String
    nombre_inicial = block_inicial.effectiveName
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '                       DECISIÓN GS
    '
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim ui_gs6000 As String, ui_gs4500 As String, ui_gs3000 As String
    
    If nombre_inicial = "GS_6000_planta" Or nombre_inicial = "GS_6000_alzado" Then
        ui_gs6000 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 4500 + 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 2 x 3000" & vbCrLf & vbCrLf & vbCrLf & "3). 4 x 1500")
    ElseIf nombre_inicial = "GS_4500_planta" Or nombre_inicial = "GS_4500_alzado" Then
        ui_gs4500 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 3000 + 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 3 x 1500")
    ElseIf nombre_inicial = "GS_3000_planta" Or nombre_inicial = "GS_3000_alzado" Then
        ui_gs3000 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2 x 1500" & vbCrLf & vbCrLf & vbCrLf & "2). 4 x 750")
    End If

    ' BLOQUES -------------
    Dim b_gs4500_al As String, b_gs4500_pl As String, b_gs3000_al As String, b_gs3000_pl As String, b_gs1500_al As String, b_gs1500_pl As String, b_gs750_pl As String, b_gs750_al As String
    
    b_gs4500_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_4500_alzado.dwg"
    b_gs4500_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_4500_planta.dwg"
    b_gs3000_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_3000_alzado.dwg"
    b_gs3000_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_3000_planta.dwg"
    b_gs1500_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_1500_alzado.dwg"
    b_gs1500_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_1500_planta.dwg"
    b_gs750_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_750_alzado.dwg"
    b_gs750_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_750_planta.dwg"
    
    ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    
    M20x60 = ruta2 & "12-M20X60.dwg"
    

    'Dim i As Long
    'i = 0
    'For i = 0 To ss.Count
       ' Set block = ss.Item(i)
    Dim obj As GcadEntity
    For Each obj In ss
        If TypeOf obj Is GcadBlockReference Then
            Set block = obj
        ElseIf TypeOf obj Is GcadLine Then
            GoTo terminar
        End If
        Dim effectiveName As String
        effectiveName = block.effectiveName
        
        Dim insertionPoint As Variant
        insertionPoint = block.insertionPoint
        
        Dim orientation As Double
        orientation = block.Rotation
        
        Dim Xs As Integer, Ys As Integer, Zs As Integer
    
        Xs = 1
        Ys = 1
        Zs = 1
        
        '------------------------------------- GS ----------------------------------------------------------------------------------------------------------------------
        ' GS 6000
        If effectiveName = "GS_6000_alzado" Then
            block.Delete
            If ui_gs6000 = "1" Or ui_gs6000 = "a" Or ui_gs6000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs4500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 4500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 4500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_gs6000 = "2" Or ui_gs6000 = "b" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_gs6000 = "3" Or ui_gs6000 = "c" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "GS_6000_planta" Then
            block.Delete
            If ui_gs6000 = "1" Or ui_gs6000 = "a" Or ui_gs6000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs4500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 4500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 4500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_gs6000 = "2" Or ui_gs6000 = "b" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_gs6000 = "3" Or ui_gs6000 = "c" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' GS 4500
        ElseIf effectiveName = "GS_4500_alzado" Then
            block.Delete
            If ui_gs4500 = "1" Or ui_gs4500 = "A" Or ui_gs4500 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs3000_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_gs4500 = "2" Or ui_gs4500 = "B" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "GS_4500_planta" Then
            block.Delete
            If ui_gs4500 = "1" Or ui_gs4500 = "A" Or ui_gs4500 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs3000_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 3000 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 3000 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_gs4500 = "2" Or ui_gs4500 = "B" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' GS 3000
        ElseIf effectiveName = "GS_3000_alzado" Then
            block.Delete
            If ui_gs3000 = "1" Or ui_gs3000 = "" Or ui_gs3000 = "A" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_gs3000 = "2" Or ui_gs3000 = "B" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "GS_3000_planta" Then
            block.Delete
            If ui_gs3000 = "1" Or ui_gs3000 = "A" Or ui_gs3000 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs1500_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 1500 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1500 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_gs3000 = "2" Or ui_gs3000 = "B" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Granshor"
                insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' GS 1500
        ElseIf effectiveName = "GS_1500_alzado" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Granshor"
            insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Granshor"
            insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        ElseIf effectiveName = "GS_1500_planta" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Granshor"
            insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_gs750_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Granshor"
            insertionPoint(0) = insertionPoint(0) + 750 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 750 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        End If
terminar:
    Next obj
    
    

    
    ' Paso 7: Limpiar el conjunto de selección al finalizar
    ss.Delete
End Sub

Sub DivLS()

    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, blockRef As Object
    Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    ' Paso 1: Crear un nuevo conjunto de selección
    Dim ss As GcadSelectionSet
    On Error Resume Next
    Set ss = ThisDrawing.SelectionSets("SS1")
    On Error GoTo 0

    If ss Is Nothing Then
        Set ss = ThisDrawing.SelectionSets.Add("SS1")
    Else
        ss.Clear
    End If

    ' Solicitar al usuario que seleccione los bloques en la pantalla
    ss.SelectOnScreen
    

    If ss.Count = 0 Then
        MsgBox "No se seleccionó ningún bloque. La operación ha sido cancelada."
        ss.Delete
        Exit Sub
    End If
    
    ' Paso 3: Obtener el bloque seleccionado
    Dim block As GcadBlockReference
    Dim block_inicial As GcadBlockReference
    Set block_inicial = ss.Item(0)
    
    Dim nombre_inicial As String
    nombre_inicial = block_inicial.effectiveName
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '                       DECISIÓN LOLASHOR
    '
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim ui_ls2200 As String
    
    If nombre_inicial = "Lola_2200" Or nombre_inicial = "Lola_2200PL" Then
        ui_ls2200 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2 x 1100" & vbCrLf & vbCrLf & vbCrLf & "2). 4 x 550")
    End If

    ' BLOQUES -------------
    Dim b_ls1100_al As String, b_ls1100_pl As String, b_ls550_al As String, b_ls550_pl As String
    
    b_ls1100_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Lola\Lola_1100.dwg"
    b_ls1100_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Lola\Lola_1100PL.dwg"
    b_ls550_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Lola\Lola_550.dwg"
    b_ls550_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Lola\Lola_550PL.dwg"
    

    'Dim i As Long
    'i = 0
    'For i = 0 To ss.Count
       ' Set block = ss.Item(i)
    Dim obj As GcadEntity
    For Each obj In ss
        If TypeOf obj Is GcadBlockReference Then
            Set block = obj
        ElseIf TypeOf obj Is GcadLine Then
            GoTo terminar
        End If
        Dim effectiveName As String
        effectiveName = block.effectiveName
        
        Dim insertionPoint As Variant
        insertionPoint = block.insertionPoint
        
        Dim orientation As Double
        orientation = block.Rotation
        
        Dim Xs As Integer, Ys As Integer, Zs As Integer
    
        Xs = 1
        Ys = 1
        Zs = 1
        
        '------------------------------------- LS ----------------------------------------------------------------------------------------------------------------------
        ' LOLASHOR 2200
        If effectiveName = "Lola_2200" Then
            block.Delete
            If ui_ls2200 = "1" Or ui_ls2200 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls1100_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 1100 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1100 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls1100_al, Xs, Ys, Zs, orientation)
                'Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls1100_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 1100 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1100 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ls2200 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
                'Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
                'Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
                'Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
                'Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Lola_2200PL" Then
            block.Delete
            If ui_ls2200 = "1" Or ui_ls2200 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls1100_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 1100 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1100 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls1100_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 1100 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1100 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_ls2200 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Lolashor"
                insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' LOLASHOR 1100
        ElseIf effectiveName = "Lola_1100" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Lolashor"
            insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Lolashor"
            insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        ElseIf effectiveName = "Lola_1100PL" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Lolashor"
            insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_ls550_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Lolashor"
            insertionPoint(0) = insertionPoint(0) + 550 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 550 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        End If

terminar:
    Next obj
    
    

    
    ' Paso 7: Limpiar el conjunto de selección al finalizar
    ss.Delete
End Sub


Sub DivMP()

    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, blockRef As Object
    Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    
    Dim M20x60 As String
    
    ' Paso 1: Crear un nuevo conjunto de selección
    Dim ss As GcadSelectionSet
    On Error Resume Next
    Set ss = ThisDrawing.SelectionSets("SS1")
    On Error GoTo 0

    If ss Is Nothing Then
        Set ss = ThisDrawing.SelectionSets.Add("SS1")
    Else
        ss.Clear
    End If

    ' Solicitar al usuario que seleccione los bloques en la pantalla
    ss.SelectOnScreen
    

    If ss.Count = 0 Then
        MsgBox "No se seleccionó ningún bloque. La operación ha sido cancelada."
        ss.Delete
        Exit Sub
    End If
    
    ' Paso 3: Obtener el bloque seleccionado
    Dim block As GcadBlockReference
    Dim block_inicial As GcadBlockReference
    Set block_inicial = ss.Item(0)
    
    Dim nombre_inicial As String
    nombre_inicial = block_inicial.effectiveName
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '                       DECISIÓN MEGAPROP
    '
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim ui_mp5400 As String, ui_mp2700 As String, ui_mp1800 As String, ui_mp900 As String, ui_mp450 As String, ui_mp270 As String
    
    If nombre_inicial = "Mshor5400PLA" Or nombre_inicial = "Mshor5400ALZ" Then
        ui_mp5400 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2 x 2700" & vbCrLf & vbCrLf & vbCrLf & "2). 3 x 1800" & vbCrLf & vbCrLf & vbCrLf & "3). 2700 + 1800 + 900")
    ElseIf nombre_inicial = "Mshor2700PLA" Or nombre_inicial = "Mshor2700ALZ" Then
        ui_mp2700 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 1800 + 900" & vbCrLf & vbCrLf & vbCrLf & "2). 3 x 900" & vbCrLf & vbCrLf & vbCrLf & "3). 1800 + 2 x 450")
    ElseIf nombre_inicial = "Mshor1800PLA" Or nombre_inicial = "Mshor1800ALZ" Then
        ui_mp1800 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2 x 900" & vbCrLf & vbCrLf & vbCrLf & "2). 4 x 450")
    ElseIf nombre_inicial = "Mshor900PLA" Or nombre_inicial = "Mshor900ALZ" Then
        ui_mp900 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 2 x 450" & vbCrLf & vbCrLf & vbCrLf & "2). 2 x 270 + 2 x 180" & vbCrLf & vbCrLf & vbCrLf & "3). 3 x 270 + 90")
    ElseIf nombre_inicial = "Mshor450PLA" Or nombre_inicial = "Mshor450ALZ" Then
        ui_mp450 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 270 + 180" & vbCrLf & vbCrLf & vbCrLf & "2). 2 x 180 + 90" & vbCrLf & vbCrLf & vbCrLf & "3). 5 x 90")
    ElseIf nombre_inicial = "Mshor270PLA" Or nombre_inicial = "Mshor270ALZ" Then
        ui_mp270 = InputBox("Elige una de las alternativas: " & vbCrLf & vbCrLf & vbCrLf & "1). 180 + 90" & vbCrLf & vbCrLf & vbCrLf & "2). 3 x 90")
    End If

    ' BLOQUES -------------
    Dim b_mp90_al As String, b_mp90_pl As String, b_mp180_pl As String, b_mp180_al As String, b_mp270_al As String, b_mp270_pl As String, b_mp450_pl As String, b_mp450_al As String, b_mp900_al As String, b_mp900_pl As String, b_mp1800_al As String, b_mp1800_pl As String, b_mp2700_al As String, b_mp2700_pl As String, b_mp5400_al As String, b_mp5400_pl As String

    b_mp5400_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor5400ALZ.dwg"
    b_mp5400_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor5400PLA.dwg"
    b_mp2700_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor2700ALZ.dwg"
    b_mp2700_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor2700PLA.dwg"
    b_mp1800_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor1800ALZ.dwg"
    b_mp1800_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor1800PLA.dwg"
    b_mp900_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor900ALZ.dwg"
    b_mp900_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor900PLA.dwg"
    b_mp450_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor450ALZ.dwg"
    b_mp450_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor450PLA.dwg"
    b_mp270_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor270ALZ.dwg"
    b_mp270_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor270PLA.dwg"
    b_mp180_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor180ALZ.dwg"
    b_mp180_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor180PLA.dwg"
    b_mp90_al = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor90ALZ.dwg"
    b_mp90_pl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\Mshor90PLA.dwg"

    ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    
    M20x60 = ruta2 & "4-M20X60.dwg"

    'Dim i As Long
    'i = 0
    'For i = 0 To ss.Count
       ' Set block = ss.Item(i)
    Dim obj As GcadEntity
    For Each obj In ss
        If TypeOf obj Is GcadBlockReference Then
            Set block = obj
        ElseIf TypeOf obj Is GcadLine Then
            GoTo terminar
        End If
        Dim effectiveName As String
        effectiveName = block.effectiveName
        
        Dim insertionPoint As Variant
        insertionPoint = block.insertionPoint
        
        Dim orientation As Double
        orientation = block.Rotation
        
        Dim Xs As Integer, Ys As Integer, Zs As Integer
    
        Xs = 1
        Ys = 1
        Zs = 1
        
        '------------------------------------- MEGAPROP ----------------------------------------------------------------------------------------------------------------------
        ' MEGAPROP 5400
        If effectiveName = "Mshor5400ALZ" Then
            block.Delete
            If ui_mp5400 = "1" Or ui_mp5400 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp2700_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp2700_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp5400 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp5400 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp2700_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Mshor5400PLA" Then
            block.Delete
            If ui_mp5400 = "1" Or ui_mp5400 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp2700_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp2700_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp5400 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp5400 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp2700_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 2700 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 2700 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        
        ' MEGAPROP 2700
        ElseIf effectiveName = "Mshor2700ALZ" Then
            block.Delete
            If ui_mp2700 = "1" Or ui_mp2700 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp2700 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp2700 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Mshor2700PLA" Then
            block.Delete
            If ui_mp2700 = "1" Or ui_mp2700 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp2700 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp2700 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' MEGAPROP 2700
        ElseIf effectiveName = "Mshor2700ALZ" Then
            block.Delete
            If ui_mp2700 = "1" Or ui_mp2700 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp2700 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp2700 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Mshor2700PLA" Then
            block.Delete
            If ui_mp2700 = "1" Or ui_mp2700 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp2700 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp2700 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp1800_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 1800 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 1800 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' MEGAPROP 1800
        ElseIf effectiveName = "Mshor1800ALZ" Then
            block.Delete
            If ui_mp1800 = "1" Or ui_mp1800 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp1800 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Mshor1800PLA" Then
            block.Delete
            If ui_mp1800 = "1" Or ui_mp1800 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp900_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 900 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 900 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp1800 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' MEGAPROP 900
        ElseIf effectiveName = "Mshor900ALZ" Then
            block.Delete
            If ui_mp900 = "1" Or ui_mp900 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp900 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp900 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Mshor900PLA" Then
            block.Delete
            If ui_mp900 = "1" Or ui_mp900 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp450_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 450 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 450 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp900 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp900 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_pl, Xs, Ys, Zs, orientation)
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_pl, Xs, Ys, Zs, orientation)
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_pl, Xs, Ys, Zs, orientation)
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' MEGAPROP 450
        ElseIf effectiveName = "Mshor450ALZ" Then
            block.Delete
            If ui_mp450 = "1" Or ui_mp450 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp450 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp450 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Mshor450PLA" Then
            block.Delete
            If ui_mp450 = "1" Or ui_mp450 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp270_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 270 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 270 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp450 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp450 = "3" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' MEGAPROP 270
        ElseIf effectiveName = "Mshor270ALZ" Then
            block.Delete
            If ui_mp270 = "1" Or ui_mp270 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp270 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
        ElseIf effectiveName = "Mshor270PLA" Then
            block.Delete
            If ui_mp270 = "1" Or ui_mp270 = "" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp180_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 180 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 180 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            ElseIf ui_mp270 = "2" Then
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
                Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
                blockRef.Layer = "Mega"
                insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            End If
            
        ' MEGAPROP 180
        ElseIf effectiveName = "Mshor180ALZ" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Mega"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_al, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Mega"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        ElseIf effectiveName = "Mshor180PLA" Then
            block.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Mega"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
            Set blockRef = gcadModel.InsertBlock(insertionPoint, M20x60, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Set blockRef = gcadModel.InsertBlock(insertionPoint, b_mp90_pl, Xs, Ys, Zs, orientation)
            blockRef.Layer = "Mega"
            insertionPoint(0) = insertionPoint(0) + 90 * Cos(orientation): insertionPoint(1) = insertionPoint(1) + 90 * Sin(orientation): insertionPoint(2) = insertionPoint(2)
        End If
            
terminar:
    Next obj
    
    
    
    ss.Delete
End Sub















