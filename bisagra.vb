                           
Sub bisagra()
    Dim doc As Object
    Set doc = ThisDrawing
    
    Dim blockRef As Object
    
    Set gcadDoc = GetObject(, "Gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    
    Ncapa = "Mega"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 30
    Ncapa = "Granshor"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 150
    Ncapa = "Pipeshor4S"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 7
    Ncapa = "Pipeshor6"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 7
    Ncapa = "Pipeshor4L"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 5
    Ncapa = "Slims"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 30
    
    Dim intPoint As Variant
    Dim PA(0 To 2) As Double
    Dim PP1(0 To 2) As Double
    Dim PB(0 To 2) As Double
    Dim PA2(0 To 2) As Double
    Dim PB2(0 To 2) As Double
    Dim Esq(0 To 2) As Double
    Dim Esqa(0 To 2) As Double
    Dim Esqb(0 To 2) As Double
    Dim Esqi(0 To 2) As Double
    Dim Esqt(0 To 2) As Double
    Dim DirMuro1 As Double
    Dim DirMuro2 As Double
    Dim DirMuro1a As Double
    Dim DirMuro2a As Double
    Dim Slope1 As Double
    Dim Slope2 As Double
    Dim pl As String
    Dim pr As String
    Dim va As String
    Dim Xs As Double
    Dim Ys As Double
    Dim Zs As Double
    Dim distanciaA As Double
    Dim distanciaB As Double
    Dim xa As Double
    Dim ya As Double
    Dim xb As Double
    Dim yb As Double
    Dim dato1 As String
    Dim dato2 As String
    Dim perf As String, jr As String, jr3 As String
    Dim lon As String
    Dim alma450 As String, alma As String, junta As String
    Dim lperfil As Double
    Dim n4570 As Integer
    Dim n4500 As Integer
    Dim n4100 As Integer
    Dim n4030 As Integer
    Dim n3070 As Integer
    Dim n3000 As Integer
    Dim n1500 As Integer
    Dim nil15000 As Integer
    Dim n900 As Integer, n6070 As Integer, n6000 As Integer, ni15000 As Integer, ni10500 As Integer
    Dim Mp90 As Integer, Mp180 As Integer, Mp450 As Integer, Mp270 As Integer
    Dim lP4570JR As Double
    Dim lP4500 As Double
    Dim lP3070JR As Double
    Dim lP3000 As Double
    Dim lMp90 As Double, lMp270 As Double, lMp450 As Double, lMp180 As Double
    Dim ruta2 As String
    Dim lfija As Double, lbisagra As Double, lP900 As Double
    Dim rutaperf As String, rutamp As String
    Dim repite As Integer
    Dim Esqaa As Variant, Esqtt As Variant
    Dim GetAngleBetweenLines As Double
    Dim Ptemp(0 To 2) As Double
    Dim lfijaH300 As Double, lfijaH300JR As Double, lfijaH450 As Double, lfijaH600 As Double, lfijaH600JR As Double
    
    On Error GoTo terminar
    
    rutaperf = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Perfiles\"
    rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
    ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    repite = 1
    
    'Valores fijos
    PI = 4 * Atn(1)
    hbisa = 630
    vbisa = 160
    lbisagra = 630
    lP900 = 900
    lP450 = 4500
    lP600N = 4030
    lP600JR = 4100
    lP300N = 4500
    lP4570JR = 4570
    lP4500 = 4500
    lP3070JR = 3070
    lP3000 = 3000
    lP6070JR = 6070
    lP6000 = 6000
    lP1500 = 1500
    li15000 = 15000
    li10500 = 10500
    lP4030 = 4030
    lP4100 = 4100
    lMp90 = 90
    lMp180 = 180
    lMp270 = 270
    lMp450 = 450
    
    lfijaH300 = lbisagra + lP900 + 80
    lfijaH300JR = lbisagra + lP3070JR + 80
    lfijaH450 = lbisagra + lP3000 + 80
    lfijaH600JR = lbisagra + lP4100 + 80
    lfijaH600 = lbisagra + lP4030 + 80
    
    On Error GoTo terminar
    
    Do While repite = 1
    
        ' Dibujar la primera línea
        Dim puntoInicioLinea1 As Variant
        Dim puntoFinLinea1 As Variant
        
        puntoInicioLinea1 = doc.Utility.GetPoint(, "Selecciona el punto de inicio del primer muro: ")
        puntoFinLinea1 = doc.Utility.GetPoint(puntoInicioLinea1, "Selecciona el punto de intersección de los dos muros: ")
            
        ' Dibujar la segunda línea
        Dim puntoInicioLinea2 As Variant
        Dim puntoFinLinea2 As Variant
        
        puntoInicioLinea2 = doc.Utility.GetPoint(puntoFinLinea1, "Selecciona el punto de inicio del segundo muro: ")
          
        'PA es el punto de inserción de la primera linea
        PA(0) = puntoInicioLinea1(0): PA(1) = puntoInicioLinea1(1): PA(2) = puntoInicioLinea1(2)
        PP1(0) = puntoFinLinea1(0): PP1(1) = puntoFinLinea1(1): PP1(2) = puntoFinLinea1(2)

        'PB es el punto de inserción de la segunda linea
        PB(0) = puntoInicioLinea2(0): PB(1) = puntoInicioLinea2(1): PB(2) = puntoInicioLinea2(2)

        'Obtener el angulo de las lineas
        DirMuro1 = gcadUtil.AngleFromXAxis(PA, PP1)
        DirMuro2 = gcadUtil.AngleFromXAxis(PB, PP1)

        If Abs(DirMuro2 - DirMuro1) > PI Then
        
            If DirMuro2 > DirMuro1 Then
                        
            Ptemp(0) = PA(0): Ptemp(1) = PA(1): Ptemp(2) = PA(2)
            PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
            PB(0) = Ptemp(0): PB(1) = Ptemp(1): PB(2) = Ptemp(2)
        
            DirMuro1 = gcadUtil.AngleFromXAxis(PA, PP1)
            DirMuro2 = gcadUtil.AngleFromXAxis(PB, PP1)
            
            End If
        
        ElseIf DirMuro2 < DirMuro1 Then
                
        Ptemp(0) = PA(0): Ptemp(1) = PA(1): Ptemp(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = Ptemp(0): PB(1) = Ptemp(1): PB(2) = Ptemp(2)
        
        DirMuro1 = gcadUtil.AngleFromXAxis(PA, PP1)
        DirMuro2 = gcadUtil.AngleFromXAxis(PB, PP1)
        
                
        End If
               
        Angulo = Abs(DirMuro1 - DirMuro2)
        
        GetAngleBetweenLines = Angulo * (180 / PI)
                
       If GetAngleBetweenLines < 81 Then
            MsgBox "El ángulo entre las líneas es: " & GetAngleBetweenLines & " grados, menor que 81 grados", vbInformation
            Exit Sub
        
        End If
        
        ' conseguir la esquina:
        ' Calculamos las direcciones de las rectas
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
                
        'Giro de direccion 90 grados para hallar perpendicular
        DirMuro1a = DirMuro1 - ((PI) / 2)
        DirMuro2a = DirMuro2 - ((PI) / 2)
        
        'Puntos paralelos
        PA2(0) = PA(0) + 90 * Cos(DirMuro1a): PA2(1) = PA(1) + 90 * Sin(DirMuro1a): PA2(2) = PA(2)
        PB2(0) = PB(0) - 90 * Cos(DirMuro2a): PB2(1) = PB(1) - 90 * Sin(DirMuro2a): PB2(2) = PB(2)
           
        If DirMuro1 = DirMuro2 Then
            MsgBox "Las líneas son paralelas, gestionar según sea necesario"
            Exit Sub
                
        ElseIf DirMuro1 = (PI / 2) Or DirMuro1 = ((3 * PI) / 2) Then
        
           ' Calculamos el punto intersección paralelo sin Tangente
            Esqi(0) = (PA2(1) - PB2(1) - Slope1 * PA2(0) + Slope2 * PB2(0)) / (Slope2 - Slope1)
            Esqi(1) = PB2(1) + Slope2 * (Esqi(0) - PB2(0))
            Esqi(2) = PB2(2) ' Assuming the lines are in the same plane
            
            ' Obtener el punto de intersección
            On Error Resume Next
            Set intPoint = doc.ModelSpace.AddPoint(PA2)
            Set intPoint = doc.ModelSpace.AddPoint(PB2)
            On Error GoTo 0
            
        Else
            
            ' Calculamos el punto intersección paralelo
            Esqi(0) = (PB2(1) - PA2(1) - Slope2 * PB2(0) + Slope1 * PA2(0)) / (Slope1 - Slope2)
            Esqi(1) = PA2(1) + Slope1 * (Esqi(0) - PA2(0))
            Esqi(2) = PA2(2) ' Assuming the lines are in the same plane
            
            ' Obtener el punto de intersección
            On Error Resume Next
            Set intPoint = doc.ModelSpace.AddPoint(PA2)
            Set intPoint = doc.ModelSpace.AddPoint(PB2)
            On Error GoTo 0
            
        End If
        
                
        ' Manejo de errores
        If Err.Number <> 0 Then
                MsgBox "Error al intentar dibujar el punto de intersección: " & Err.Description, vbExclamation
                Exit Sub
        End If
            
            'calculo lado A
            xa = PA2(0) - Esqi(0)
            ya = PA2(1) - Esqi(1)
            
            'calculo lado B
            xb = PB2(0) - Esqi(0)
            yb = PB2(1) - Esqi(1)
            
            Xs = 1
            Ys = 1
            Zs = 1
            distanciaA = Val(Sqr((xa ^ 2 + ya ^ 2)))
            distanciaB = Val(Sqr((xb ^ 2 + yb ^ 2)))
                  
                                         
            pr = rutaperf & "BIS_FIN_AL.dwg"
            pl = rutaperf & "BIS_PRI_AL.dwg"
            va = rutaperf & "Incye_600JR_4000_AL.dwg"
            v300jr4570 = rutaperf & "Incye_300JR_4570_AL.dwg"
            v300jr3070 = rutaperf & "Incye_300JR_3070_AL.dwg"
            v300jr6070 = rutaperf & "Incye_300JR_6070_AL.dwg"
            v300n6000 = rutaperf & "Incye_300_6000_AL.dwg"
            v300n4500 = rutaperf & "Incye_300_4500_AL.dwg"
            v300jr4500 = rutaperf & "Incye_300_4500_AL.dwg"
            v300n3000 = rutaperf & "Incye_300_3000_AL.dwg"
            v300n1500 = rutaperf & "Incye_300_1500_AL.dwg"
            v300n900 = rutaperf & "Incye_300_900_AL.dwg"
            v450SAjr4500 = rutaperf & "Incye_450SAJR_4500_AL.dwg"
            v450SAjr6000 = rutaperf & "Incye_450SAJR_6000_AL.dwg"
            v450SAjr3000 = rutaperf & "Incye_450SAJR_3000_AL.dwg"
            v450jr4500 = rutaperf & "Incye_450JR_4500_AL.dwg"
            v450jr6000 = rutaperf & "Incye_450JR_6000_AL.dwg"
            v450jr3000 = rutaperf & "Incye_450JR_3000_AL.dwg"
            v600jr4000 = rutaperf & "Incye_600JR_4000_AL.dwg"
            v600n4000 = rutaperf & "Incye_600_4000_AL.dwg"
            Mpshor450 = rutamp & "Mshor450ALZ.dwg"
            Mpshor270 = rutamp & "Mshor270ALZ.dwg"
            Mpshor180 = rutamp & "Mshor180ALZ.dwg"
            Mpshor90 = rutamp & "Mshor90ALZ.dwg"
            
            On Error GoTo terminar
            
            'punto inserccion muro 1
            Esqa(0) = Esqi(0) + vbisa * Cos(DirMuro1a): Esqa(1) = Esqi(1) + vbisa * Sin(DirMuro1a): Esqa(2) = Esqi(2)
            
            Esqa(0) = Esqa(0) - hbisa * Cos(DirMuro1): Esqa(1) = Esqa(1) - hbisa * Sin(DirMuro1): Esqa(2) = Esqa(2)
                
            Set blockRef = doc.ModelSpace.InsertBlock(Esqa, pr, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Perfiles INCYE"
            
            M30x100_2 = ruta2 & "2-M30x100{10.9}.dwg"
            M30x100_3 = ruta2 & "3-M30x100{10.9}.dwg"
            M20x130_6 = ruta2 & "6-M20x130.dwg"
            M20x90_10 = ruta2 & "10-M20x90.dwg"
            M20x130_12 = ruta2 & "12-M20x130.dwg"
            M20x90_6 = ruta2 & "6-M20x90.dwg"
            
            'Ubicacion tornillo en junta reforzada
            Esqt(0) = Esqa(0) + 290 * Cos(DirMuro1a): Esqt(1) = Esqa(1) + 290 * Sin(DirMuro1a): Esqt(2) = Esqa(2)
                        
            kwordList = "600 450 300"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            perf = ThisDrawing.Utility.GetKeyword(vbLf & "Viga HEB lado A: [300/450/600]")
            
        If perf = "300" Or perf = "" Then
        
            Call HEB300MuroA(Esqa, DirMuro1a, Esqt, DirMuro1, distanciaA)
               
        ElseIf perf = "450" Then
        
        If distanciaA < lfijaH450 Then
            MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH450 & "mm."""
                         
            GoTo terminar
        End If
        
        kwordList = "Triple Simple"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        alma450 = ThisDrawing.Utility.GetKeyword(vbLf & "Alma?: [Triple/Simple]")
        Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_2, Xs, Ys, Zs, DirMuro1 + PI)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        If alma450 = "Triple" Or alma450 = "" Then
        
            'Inserccion primer perfil obligatorio
            Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v450jr4500, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Perfiles INCYE"
            Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                        
            lperfil = distanciaA - lbisagra - lP4500
            ni15000 = Fix(lperfil / li15000)
            lperfil = lperfil - ni15000 * li15000
            ni10500 = Fix(lperfil / li10500)
            lperfil = lperfil - ni10500 * li10500
            n6000 = Fix(lperfil / lP6000)
            lperfil = lperfil - n6000 * lP6000
            n4500 = Fix(lperfil / lP4500)
            lperfil = lperfil - n4500 * lP4500
            n3000 = Fix(lperfil / lP3000)
            lperfil = lperfil - n3000 * lP3000
            Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)
            
            If ni15000 > 0 Then
                    i = 0
                    Do While i < ni15000
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450jr6000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP6000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v450jr4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v450jr4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    i = i + 1
                    Loop
                    
            End If
            
            If ni10500 > 0 Then
                    i = 0
                    Do While i < ni10500
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450jr6000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP6000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v450jr4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    i = i + 1
                    Loop
                    
            End If
            
            If n6000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450jr6000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP6000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    
                End If
        
                If n4500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450jr4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    
                End If
                
            If n3000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450jr3000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP3000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    
                End If
                
                'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 110 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 110 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
        
        ElseIf alma450 = "Simple" Then
            
            'Inserccion primer perfil obligatorio
            Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v450SAjr4500, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Perfiles INCYE"
            Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                        
            lperfil = distanciaA - lbisagra - lP4500
            ni15000 = Fix(lperfil / li15000)
            lperfil = lperfil - ni15000 * li15000
            ni10500 = Fix(lperfil / li10500)
            lperfil = lperfil - ni10500 * li10500
            n6000 = Fix(lperfil / lP6000)
            lperfil = lperfil - n6000 * lP6000
            n4500 = Fix(lperfil / lP4500)
            lperfil = lperfil - n4500 * lP4500
            n3000 = Fix(lperfil / lP3000)
            lperfil = lperfil - n3000 * lP3000
            Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)
            
            If ni15000 > 0 Then
                    i = 0
                    Do While i < ni15000
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450SAjr6000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP6000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v450SAjr4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v450SAjr4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    i = i + 1
                    Loop
                    
            End If
            
            If ni10500 > 0 Then
                    i = 0
                    Do While i < ni10500
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450SAjr6000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP6000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v450SAjr4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    i = i + 1
                    Loop
                    
            End If
            
            If n6000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450SAjr6000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP6000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    
                End If
        
            If n4500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450SAjr4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    
                End If
                
            If n3000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v450SAjr3000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_10, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP3000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    
                End If
                
                'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 110 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 110 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
        End If
        
    ElseIf perf = "600" Then
        kwordList = "Sí No"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        jr = ThisDrawing.Utility.GetKeyword(vbLf & "Juntas reforzadas?: [Sí/No]")
        If jr = "Sí" Or jr = "" Then
        
        If distanciaA < lfijaH600JR Then
            MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH600JR & "mm."""
            GoTo terminar
        End If
            
            'Inserccion primer perfil obligatorio
            Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v600jr4000, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Perfiles INCYE"
            Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_12, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            Esqt(0) = Esqt(0) + 150 * Cos(DirMuro1a): Esqt(1) = Esqt(1) + 150 * Sin(DirMuro1a): Esqt(2) = Esqt(2)
            Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_2, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            Esqa(0) = Esqa(0) - lP4100 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4100 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            Esqt(0) = Esqt(0) - lP4100 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4100 * Sin(DirMuro1): Esqt(2) = Esqt(2)
            
            
            lperfil = distanciaA - lbisagra - lP4100
            n4100 = Fix(lperfil / lP4100)
            lperfil = lperfil - n4100 * lP4100
            Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)
            
            If n4100 > 0 Then
                    i = 0
                    Do While i < n4100
                    Set blockRef = gcadModel.InsertBlock(Esqa, v600jr4000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_12, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4100 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4100 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4100 * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4100 * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    i = i + 1
                    Loop
                    
            End If
            
            'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 110 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 110 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
        ElseIf jr = "No" Then
        
        If distanciaA < lfijaH600 Then
            MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH600 & "mm."""
            GoTo terminar
        End If
        
            'Nivelar HEB600 normal
            Esqa(0) = Esqa(0) + 50 * Cos(DirMuro1a): Esqa(1) = Esqa(1) + 50 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
            
            'Inserccion primer perfil obligatorio
            Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v600n4000, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Perfiles INCYE"
            Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_12, Xs, Ys, Zs, DirMuro1 + PI)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            Esqa(0) = Esqa(0) - lP4030 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4100 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            
            lperfil = distanciaA - lbisagra - lP4030
            n4030 = Fix(lperfil / lP4030)
            lperfil = lperfil - n4030 * lP4030
            Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)
            
            If n4030 > 0 Then
                    i = 0
                    Do While i < n4030
                    Set blockRef = gcadModel.InsertBlock(Esqa, v600n4000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_12, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4030 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4030 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    i = i + 1
                    Loop
                    
            End If
            
            'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 160 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 160 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
        End If
    End If
            
            On Error GoTo terminar
            
            'punto de inserccion muro 2
            Esqb(0) = Esqi(0) - hbisa * Cos(DirMuro2): Esqb(1) = Esqi(1) - hbisa * Sin(DirMuro2): Esqb(2) = Esqi(2)
            
            Esqb(0) = Esqb(0) - vbisa * Cos(DirMuro2a): Esqb(1) = Esqb(1) - vbisa * Sin(DirMuro2a): Esqb(2) = Esqb(2)
            
            'Ubicacion tornillo en junta reforzada
            Esqt(0) = Esqb(0) - 290 * Cos(DirMuro2a): Esqt(1) = Esqb(1) - 290 * Sin(DirMuro2a): Esqt(2) = Esqb(2)
            
            Set blockRef = doc.ModelSpace.InsertBlock(Esqb, pl, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Perfiles INCYE"
                                   
            kwordList = "600 450 300"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            perf = ThisDrawing.Utility.GetKeyword(vbLf & "Viga HEB lado B: [300/450/600]")
            If perf = "300" Or perf = "" Then
                    
            'Nivelar HEB300
            Esqb(0) = Esqb(0) + 100 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 100 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
            kwordList = "Sí No"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            jr3 = ThisDrawing.Utility.GetKeyword(vbLf & "Juntas reforzadas?: [Sí/No]")
            If jr3 = "Sí" Or jr3 = "" Then
            If distanciaB < lfijaH300JR Then
            MsgBox "Medida Muro B  " & distanciaB & "mm, menor que el mínimo necesario de " & lfijaH300JR & "mm."""
            GoTo terminar
            End If
            kwordList = "3070 4570 6070"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [3070/4570/6070]")
            Esqt(0) = Esqt(0) + 150 * Cos(DirMuro2a): Esqt(1) = Esqt(1) + 150 * Sin(DirMuro2a): Esqt(2) = Esqt(2)
            Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_2, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            
            If lon = "3070" Then
            'Inserccion primer perfil obligatorio
                Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqb(0) = Esqb(0) - lP3070JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3070JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v300jr3070, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Perfiles INCYE"
                Esqt(0) = Esqt(0) - lP3070JR * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP3070JR * Sin(DirMuro2): Esqt(2) = Esqt(2)
                                
                lperfil = distanciaB - lbisagra - lP3070JR
                n3070 = Fix(lperfil / lP3070JR)
                lperfil = lperfil - n3070 * lP3070JR
                Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)

                If n3070 > 0 Then
                    i = 0
                    Do While i < n3070
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP3070JR * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP3070JR * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP3070JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3070JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300jr3070, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 10 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 10 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
            
            
            ElseIf lon = "4570" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqb(0) = Esqb(0) - lP4570JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4570JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v300jr4570, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Perfiles INCYE"
                Esqt(0) = Esqt(0) - lP4570JR * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4570JR * Sin(DirMuro2): Esqt(2) = Esqt(2)
                                            
                lperfil = distanciaB - lbisagra - lP4570JR
                n4570 = Fix(lperfil / lP4570JR)
                lperfil = lperfil - n4570 * lP4570JR
                n3070 = Fix(lperfil / lP3070JR)
                lperfil = lperfil - n3070 * lP3070JR
                Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)
                
                If n4570 > 0 Then
                    i = 0
                    Do While i < n4570
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4570JR * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4570JR * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4570JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4570JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300jr4570, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                End If
                
                If n3070 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP3070JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3070JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300jr3070, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                                        
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 10 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 10 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
            
            ElseIf lon = "6070" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqb(0) = Esqb(0) - lP6070JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6070JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v300jr6070, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Perfiles INCYE"
                
                Esqt(0) = Esqt(0) - lP6070JR * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP6070JR * Sin(DirMuro2): Esqt(2) = Esqt(2)
                                
                lperfil = distanciaB - lbisagra - lP6070JR
                n6070 = Fix(lperfil / lP6070JR)
                lperfil = lperfil - n6070 * lP6070JR
                n4570 = Fix(lperfil / lP4570JR)
                lperfil = lperfil - n4570 * lP4570JR
                n3070 = Fix(lperfil / lP3070JR)
                lperfil = lperfil - n3070 * lP3070JR
                Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)

                If n6070 > 0 Then
                    i = 0
                    Do While i < n6070
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6070JR * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP6070JR * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP6070JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6070JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300jr6070, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
                End If
                
                If n4570 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4570JR * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4570JR * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4570JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4570JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300jr4570, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    
                End If
                
                If n3070 > 0 Then
                Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP3070JR * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3070JR * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300jr3070, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                                        
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 10 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 10 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            End If
            
        ElseIf jr3 = "No" Then
        
        If distanciaB < lfijaH300 Then
            MsgBox "Medida Muro B  " & distanciaB & "mm, menor que el mínimo necesario de " & lfijaH300 & "mm."""
            GoTo terminar
        End If
        
            kwordList = "900 1500 3000 4500 6000"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [900/1500/3000/4500/6000]")
            
            If lon = "6000" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqb(0) = Esqb(0) - lP6000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v300n6000, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Perfiles INCYE"
            
                lperfil = distanciaB - lbisagra - lP6000
                n6000 = Fix(lperfil / lP6000)
                lperfil = lperfil - n6000 * lP6000
                n4500 = Fix(lperfil / lP4500)
                lperfil = lperfil - n4500 * lP4500
                n3000 = Fix(lperfil / lP3000)
                lperfil = lperfil - n3000 * lP3000
                n1500 = Fix(lperfil / lP1500)
                lperfil = lperfil - n1500 * lP1500
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)

                If n6000 > 0 Then
                    i = 0
                    Do While i < n6000
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP6000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n6000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
                End If
                
                If n4500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    
                End If
                
                If n3000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP3000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n3000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    
                End If
                    
                If n1500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP1500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP1500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n1500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                
                End If
                
                If n900 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP900 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP900 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n900, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 10 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 10 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
            ElseIf lon = "4500" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v300n4500, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Perfiles INCYE"
                            
                lperfil = distanciaB - lbisagra - lP4500
                n4500 = Fix(lperfil / lP4500)
                lperfil = lperfil - n4500 * lP4500
                n3000 = Fix(lperfil / lP3000)
                lperfil = lperfil - n3000 * lP3000
                n1500 = Fix(lperfil / lP1500)
                lperfil = lperfil - n1500 * lP1500
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)

                If n4500 > 0 Then
                    i = 0
                    Do While i < n4500
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
                End If
                
                If n3000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP3000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n3000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                                        
                End If
                    
                If n1500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP1500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP1500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n1500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                                    
                End If
                
                If n900 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP900 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP900 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n900, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                                    
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 10 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 10 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
            ElseIf lon = "3000" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqb(0) = Esqb(0) - lP3000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v300n3000, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Perfiles INCYE"
                            
                lperfil = distanciaB - lbisagra - lP3000
                n3000 = Fix(lperfil / lP3000)
                lperfil = lperfil - n3000 * lP3000
                n1500 = Fix(lperfil / lP1500)
                lperfil = lperfil - n1500 * lP1500
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)

                If n3000 > 0 Then
                    i = 0
                    Do While i < n3000
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP3000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n3000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
                End If
                
                If n1500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP1500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP1500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n1500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                                    
                End If
                
                If n900 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP900 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP900 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n900, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                                    
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 10 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 10 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
                        
            ElseIf lon = "1500" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqb(0) = Esqb(0) - lP1500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP1500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v300n1500, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Perfiles INCYE"
                            
                lperfil = distanciaB - lbisagra - lP1500
                n1500 = Fix(lperfil / lP1500)
                lperfil = lperfil - n1500 * lP1500
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)

                If n1500 > 0 Then
                    i = 0
                    Do While i < n1500
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP1500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP1500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n1500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
                End If
                            
                If n900 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP900 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP900 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n900, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 10 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 10 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
            
            
            ElseIf lon = "900" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqb(0) = Esqb(0) - lP900 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP900 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v300n900, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Perfiles INCYE"
                            
                lperfil = distanciaB - lbisagra - lP900
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)

                If n900 > 0 Then
                    i = 0
                    Do While i < n900
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_6, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP900 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP900 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v300n900, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 10 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 10 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
                            
            End If
            
        End If
        
    ElseIf perf = "450" Then
    If distanciaB < lfijaH450 Then
            MsgBox "Medida Muro B  " & distanciaB & "mm, menor que el mínimo necesario de " & lfijaH450 & "mm."""
            GoTo terminar
        End If
        kwordList = "Triple Simple"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        alma450 = ThisDrawing.Utility.GetKeyword(vbLf & "Alma?: [Triple/Simple]")
        Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_2, Xs, Ys, Zs, DirMuro2)
        blockRef.Layer = "Nonplot"
        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
        If alma450 = "Triple" Or alma450 = "" Then
        
            'Inserccion primer perfil obligatorio
            Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
            Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v450jr4500, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Perfiles INCYE"
            Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                        
            lperfil = distanciaB - lbisagra - lP4500
            ni15000 = Fix(lperfil / li15000)
            lperfil = lperfil - ni15000 * li15000
            ni10500 = Fix(lperfil / li10500)
            lperfil = lperfil - ni10500 * li10500
            n6000 = Fix(lperfil / lP6000)
            lperfil = lperfil - n6000 * lP6000
            n4500 = Fix(lperfil / lP4500)
            lperfil = lperfil - n4500 * lP4500
            n3000 = Fix(lperfil / lP3000)
            lperfil = lperfil - n3000 * lP3000
            Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)
            
            If ni15000 > 0 Then
                    i = 0
                    Do While i < ni15000
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP6000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450jr6000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v450jr4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v450jr4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
            End If
            
            If ni10500 > 0 Then
                    i = 0
                    Do While i < ni10500
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP6000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450jr6000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v450jr4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
            End If
            
            If n6000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP6000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450jr6000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    
                End If
        
            If n4500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450jr4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    
                End If
                
            If n3000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP3000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450jr3000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP3000 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP3000 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 110 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 110 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
        ElseIf alma450 = "Simple" Then
            
            'Inserccion primer perfil obligatorio
            Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
            Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v450SAjr4500, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Perfiles INCYE"
            Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                        
            lperfil = distanciaB - lbisagra - lP4500
            ni15000 = Fix(lperfil / li15000)
            lperfil = lperfil - ni15000 * li15000
            ni10500 = Fix(lperfil / li10500)
            lperfil = lperfil - ni10500 * li10500
            n6000 = Fix(lperfil / lP6000)
            lperfil = lperfil - n6000 * lP6000
            n4500 = Fix(lperfil / lP4500)
            lperfil = lperfil - n4500 * lP4500
            n3000 = Fix(lperfil / lP3000)
            lperfil = lperfil - n3000 * lP3000
            Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)
            
            If ni15000 > 0 Then
                    i = 0
                    Do While i < ni15000
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP6000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450SAjr6000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v450SAjr4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v450SAjr4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
            End If
            
            If ni10500 > 0 Then
                    i = 0
                    Do While i < ni10500
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP6000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450SAjr6000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v450SAjr4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
            End If
            
            If n6000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP6000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP6000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450SAjr6000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6000 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP6000 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    
                End If
        
            If n4500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP4500 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4500 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450SAjr4500, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4500 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4500 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    
                End If
                
            If n3000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_10, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP3000 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP3000 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v450SAjr3000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP3000 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP3000 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    
                End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 110 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 110 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
        End If
    
        
    ElseIf perf = "600" Then
        kwordList = "Sí No"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        jr = ThisDrawing.Utility.GetKeyword(vbLf & "Juntas reforzadas?: [Sí/No]")
        If jr = "Sí" Or jr = "" Then
        
        If distanciaB < lfijaH600JR Then
            MsgBox "Medida Muro B  " & distanciaB & "mm, menor que el mínimo necesario de " & lfijaH600JR & "mm."""
            GoTo terminar
        End If
            
            'Inserccion primer perfil obligatorio
            Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_12, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            Esqb(0) = Esqb(0) - lP4100 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4100 * Sin(DirMuro2): Esqb(2) = Esqb(2)
            Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v600jr4000, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Perfiles INCYE"
            Esqt(0) = Esqt(0) - 150 * Cos(DirMuro2a): Esqt(1) = Esqt(1) - 150 * Sin(DirMuro2a): Esqt(2) = Esqt(2)
            Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_2, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            Esqt(0) = Esqt(0) - lP4100 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4100 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                        
            lperfil = distanciaB - lbisagra - lP4100
            n4100 = Fix(lperfil / lP4100)
            lperfil = lperfil - n4100 * lP4100
            Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)
            
            If n4100 > 0 Then
                    i = 0
                    Do While i < n4100
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_12, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4100 * Cos(DirMuro2): Esqt(1) = Esqt(1) - lP4100 * Sin(DirMuro2): Esqt(2) = Esqt(2)
                    Esqb(0) = Esqb(0) - lP4100 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4100 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v600jr4000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
            End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 110 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 110 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
        ElseIf jr = "No" Then
        
        If distanciaB < lfijaH600 Then
            MsgBox "Medida Muro B  " & distanciaB & "mm, menor que el mínimo necesario de " & lfijaH600 & "mm."""
            GoTo terminar
        End If
        
            'Nivelar HEB600 normal
            Esqb(0) = Esqb(0) - 50 * Cos(DirMuro2a): Esqb(1) = Esqb(1) - 50 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
            
            'Inserccion primer perfil obligatorio
            Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_12, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Nonplot"
            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            Esqb(0) = Esqb(0) - lP4030 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4030 * Sin(DirMuro2): Esqb(2) = Esqb(2)
            Set blockRef = doc.ModelSpace.InsertBlock(Esqb, v600n4000, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Perfiles INCYE"
            
            lperfil = distanciaB - lbisagra - lP4030
            n4030 = Fix(lperfil / lP4030)
            lperfil = lperfil - n4030 * lP4030
            Mp450 = Fix(lperfil / lMp450)
                Mp270 = Fix(lperfil / lMp270)
                Mp180 = Fix(lperfil / lMp180)
                Mp90 = Fix(lperfil / lMp90)
            
            If n4030 > 0 Then
                    i = 0
                    Do While i < n4030
                    Set blockRef = gcadModel.InsertBlock(Esqb, M20x130_12, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqb(0) = Esqb(0) - lP4030 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lP4030 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                    Set blockRef = gcadModel.InsertBlock(Esqb, v600n4000, Xs, Ys, Zs, DirMuro2)
                    blockRef.Layer = "Perfiles INCYE"
                    i = i + 1
                    Loop
                    
            End If
                
                'Nivelar Megapro
                            Esqb(0) = Esqb(0) + 160 * Cos(DirMuro2a): Esqb(1) = Esqb(1) + 160 * Sin(DirMuro2a): Esqb(2) = Esqb(2)
                        
                            Call MegaproLadoB(Esqb, DirMuro2, Mp90, Mp180, Mp270, Mp450, lperfil)
            
        End If
    End If
            
  
    Loop
terminar:
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------




Sub HEB300MuroA(Esqa() As Double, DirMuro1a As Double, Esqt() As Double, DirMuro1 As Double, distanciaA As Double)
    
    Dim doc As Object
    Set doc = ThisDrawing
    
    Dim blockRef As Object
    
    Set gcadDoc = GetObject(, "Gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    
    Ncapa = "Mega"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 30
    Ncapa = "Granshor"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 150
    Ncapa = "Pipeshor4S"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 7
    Ncapa = "Pipeshor6"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 7
    Ncapa = "Pipeshor4L"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 5
    Ncapa = "Slims"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 30

    Dim intPoint As Variant
    
    Dim PA(0 To 2) As Double
    Dim PP1(0 To 2) As Double
    Dim PB(0 To 2) As Double
    Dim PA2(0 To 2) As Double
    Dim PB2(0 To 2) As Double
    Dim Esq(0 To 2) As Double
    Dim Esqb(0 To 2) As Double
    Dim Esqi(0 To 2) As Double
    Dim Slope1 As Double
    Dim Slope2 As Double
    Dim pl As String
    Dim pr As String
    Dim va As String
    Dim v300jr4570 As String
    Dim v300n4500 As String
    Dim v300jr6070 As String
    Dim v300n6000 As String
    Dim v300n3000 As String
    Dim v300jr3070 As String
    Dim v300n1500 As String
    Dim v300n900 As String
    Dim Xs As Double
    Dim Ys As Double
    Dim Zs As Double
    Dim distanciaB As Double
    Dim xa As Double
    Dim ya As Double
    Dim xb As Double
    Dim yb As Double
    Dim perf As String, jr As String, jr3 As String
    Dim lon As String
    Dim alma450 As String, alma As String, junta As String
    Dim lperfil As Double
    Dim n4570 As Integer
    Dim n4500 As Integer
    Dim n4100 As Integer
    Dim n4030 As Integer
    Dim n3070 As Integer
    Dim n3000 As Integer
    Dim n1500 As Integer
    Dim nil15000 As Integer
    Dim n900 As Integer, n6070 As Integer, n6000 As Integer, ni15000 As Integer, ni10500 As Integer
    Dim Mp90 As Integer, Mp180 As Integer, Mp450 As Integer, Mp270 As Integer
    Dim lP4570JR As Double
    Dim lP4500 As Double
    Dim lP3070JR As Double
    Dim lP3000 As Double
    Dim lMp90 As Double, lMp270 As Double, lMp450 As Double, lMp180 As Double
    Dim ruta2 As String
    Dim lfija As Double, lbisagra As Double, lP900 As Double
    Dim rutaperf As String, rutamp As String
    Dim repite As Integer
    Dim lfijaH300 As Double, lfijaH300JR As Double, lfijaH450 As Double, lfijaH600 As Double, lfijaH600JR As Double
    
    rutaperf = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Perfiles\"
    rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
    ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    repite = 1
    
    'Valores fijos
    PI = 4 * Atn(1)
    hbisa = 630
    vbisa = 160
    lbisagra = 630
    lP900 = 900
    lP450 = 4500
    lP600N = 4030
    lP600JR = 4100
    lP300N = 4500
    lP4570JR = 4570
    lP4500 = 4500
    lP3070JR = 3070
    lP3000 = 3000
    lP6070JR = 6070
    lP6000 = 6000
    lP1500 = 1500
    li15000 = 15000
    li10500 = 10500
    lP4030 = 4030
    lP4100 = 4100
    lMp90 = 90
    lMp180 = 180
    lMp270 = 270
    lMp450 = 450
    lfijaH300 = lbisagra + lP900 + 80
    lfijaH300JR = lbisagra + lP3070JR + 80
    
            Xs = 1
            Ys = 1
            Zs = 1
    
            pr = rutaperf & "BIS_FIN_AL.dwg"
            pl = rutaperf & "BIS_PRI_AL.dwg"
            va = rutaperf & "Incye_600JR_4000_AL.dwg"
            v300jr4570 = rutaperf & "Incye_300JR_4570_AL.dwg"
            v300jr3070 = rutaperf & "Incye_300JR_3070_AL.dwg"
            v300jr6070 = rutaperf & "Incye_300JR_6070_AL.dwg"
            v300n6000 = rutaperf & "Incye_300_6000_AL.dwg"
            v300n4500 = rutaperf & "Incye_300_4500_AL.dwg"
            v300jr4500 = rutaperf & "Incye_300_4500_AL.dwg"
            v300n3000 = rutaperf & "Incye_300_3000_AL.dwg"
            v300n1500 = rutaperf & "Incye_300_1500_AL.dwg"
            v300n900 = rutaperf & "Incye_300_900_AL.dwg"
            v450SAjr4500 = rutaperf & "Incye_450SAJR_4500_AL.dwg"
            v450SAjr6000 = rutaperf & "Incye_450SAJR_6000_AL.dwg"
            v450SAjr3000 = rutaperf & "Incye_450SAJR_3000_AL.dwg"
            v450jr4500 = rutaperf & "Incye_450JR_4500_AL.dwg"
            v450jr6000 = rutaperf & "Incye_450JR_6000_AL.dwg"
            v450jr3000 = rutaperf & "Incye_450JR_3000_AL.dwg"
            v600jr4000 = rutaperf & "Incye_600JR_4000_AL.dwg"
            v600n4000 = rutaperf & "Incye_600_4000_AL.dwg"
            Mpshor450 = rutamp & "Mshor450ALZ.dwg"
            Mpshor270 = rutamp & "Mshor270ALZ.dwg"
            Mpshor180 = rutamp & "Mshor180ALZ.dwg"
            Mpshor90 = rutamp & "Mshor90ALZ.dwg"
            
            M30x100_2 = ruta2 & "2-M30x100{10.9}.dwg"
            M30x100_3 = ruta2 & "3-M30x100{10.9}.dwg"
            M20x130_6 = ruta2 & "6-M20x130.dwg"
            M20x90_10 = ruta2 & "10-M20x90.dwg"
            M20x130_12 = ruta2 & "12-M20x130.dwg"
            M20x90_6 = ruta2 & "6-M20x90.dwg"
    
            lfija = lbisagra + lP900 + 80
            
            'Nivelar HEB300
            Esqa(0) = Esqa(0) - 100 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 100 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
            kwordList = "Sí No"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            jr3 = ThisDrawing.Utility.GetKeyword(vbLf & "Juntas reforzadas?: [Sí/No]")
                If jr3 = "Sí" Or jr3 = "" Then
                
                    If distanciaA < lfijaH300JR Then
                    MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300JR & "mm."""
                    GoTo terminar
                    End If
                
                kwordList = "3070 4570 6070"
                ThisDrawing.Utility.InitializeUserInput 0, kwordList
                lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [3070/4570/6070]")
                Esqt(0) = Esqt(0) - 150 * Cos(DirMuro1a): Esqt(1) = Esqt(1) - 150 * Sin(DirMuro1a): Esqt(2) = Esqt(2)
                Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_2, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
            
                    If lon = "3070" Then
                    'Inserccion primer perfil obligatorio
                    Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v300jr3070, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Esqa(0) = Esqa(0) - lP3070JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3070JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Esqt(0) = Esqt(0) - lP3070JR * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP3070JR * Sin(DirMuro1): Esqt(2) = Esqt(2)
            
                    lperfil = distanciaA - lbisagra - lP3070JR
                    n3070 = Fix(lperfil / lP3070JR)
                    lperfil = lperfil - n3070 * lP3070JR
                    Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)
                                                        
                        If n3070 > 0 Then
                        i = 0
                        Do While i < n3070
                        Set blockRef = gcadModel.InsertBlock(Esqa, v300jr3070, Xs, Ys, Zs, DirMuro1 + PI)
                        blockRef.Layer = "Perfiles INCYE"
                        Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                        blockRef.Layer = "Nonplot"
                        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                        Esqa(0) = Esqa(0) - lP3070JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3070JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                        Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                        blockRef.Layer = "Nonplot"
                        blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                        Esqt(0) = Esqt(0) - lP3070JR * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP3070JR * Sin(DirMuro1): Esqt(2) = Esqt(2)
                        i = i + 1
                        Loop
                    
                        End If
                        
                            'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 10 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 10 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                                            
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
                        ElseIf lon = "4570" Then
            
                'Inserccion primer perfil obligatorio
                Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v300jr4570, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Perfiles INCYE"
                Esqa(0) = Esqa(0) - lP4570JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4570JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                Esqt(0) = Esqt(0) - lP4570JR * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4570JR * Sin(DirMuro1): Esqt(2) = Esqt(2)
                                
                lperfil = distanciaA - lbisagra - lP4570JR
                n4570 = Fix(lperfil / lP4570JR)
                lperfil = lperfil - n4570 * lP4570JR
                n3070 = Fix(lperfil / lP3070JR)
                lperfil = lperfil - n3070 * lP3070JR
                Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)
                
                    If n4570 > 0 Then
                    i = 0
                    Do While i < n4570
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300jr4570, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4570JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4570JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4570JR * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4570JR * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    i = i + 1
                    Loop
                    End If
                
                    If n3070 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300jr3070, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP3070JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3070JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    
                    End If
                    
                    'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 10 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 10 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
                ElseIf lon = "6070" Then
            
                'Inserccion primer perfil obligatorio
                Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v300jr6070, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Perfiles INCYE"
                Esqa(0) = Esqa(0) - lP6070JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6070JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                Esqt(0) = Esqt(0) - lP6070JR * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP6070JR * Sin(DirMuro1): Esqt(2) = Esqt(2)
                            
                lperfil = distanciaA - lbisagra - lP6070JR
                n6070 = Fix(lperfil / lP6070JR)
                lperfil = lperfil - n6070 * lP6070JR
                n4570 = Fix(lperfil / lP4570JR)
                lperfil = lperfil - n4570 * lP4570JR
                n3070 = Fix(lperfil / lP3070JR)
                lperfil = lperfil - n3070 * lP3070JR
                Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)

                    If n6070 > 0 Then
                    i = 0
                    Do While i < n6070
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300jr6070, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP6070JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6070JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP6070JR * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP6070JR * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    i = i + 1
                    Loop
                        
                    End If
                
                    If n4570 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300jr4570, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4570JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4570JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqt(0) = Esqt(0) - lP4570JR * Cos(DirMuro1): Esqt(1) = Esqt(1) - lP4570JR * Sin(DirMuro1): Esqt(2) = Esqt(2)
                    
                    
                    End If
                
                    If n3070 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300jr3070, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP3070JR * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3070JR * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                                        
                    End If
                    
                    'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 10 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 10 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
                End If
                
                ElseIf jr3 = "No" Then
                
                If distanciaA < lfijaH300 Then
                MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300 & "mm."""
                GoTo terminar
                End If
                
            kwordList = "900 1500 3000 4500 6000"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [900/1500/3000/4500/6000]")
            
            If lon = "6000" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v300n6000, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Perfiles INCYE"
                Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqa(0) = Esqa(0) - lP6000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            
                lperfil = distanciaA - lbisagra - lP6000
                n6000 = Fix(lperfil / lP6000)
                lperfil = lperfil - n6000 * lP6000
                n4500 = Fix(lperfil / lP4500)
                lperfil = lperfil - n4500 * lP4500
                n3000 = Fix(lperfil / lP3000)
                lperfil = lperfil - n3000 * lP3000
                n1500 = Fix(lperfil / lP1500)
                lperfil = lperfil - n1500 * lP1500
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)

                If n6000 > 0 Then
                    i = 0
                    Do While i < n6000
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n6000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP6000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP6000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    i = i + 1
                    Loop
                    
                End If
                
                If n4500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    
                End If
                
                If n3000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n3000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP3000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    
                End If
                    
                If n1500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n1500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP1500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP1500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                
                End If
                
                If n900 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n900, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP900 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP900 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                
                End If
                
                'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 10 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 10 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
            ElseIf lon = "4500" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v300n4500, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Perfiles INCYE"
                Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            
                lperfil = distanciaA - lbisagra - lP4500
                n4500 = Fix(lperfil / lP4500)
                lperfil = lperfil - n4500 * lP4500
                n3000 = Fix(lperfil / lP3000)
                lperfil = lperfil - n3000 * lP3000
                n1500 = Fix(lperfil / lP1500)
                lperfil = lperfil - n1500 * lP1500
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)

                If n4500 > 0 Then
                    i = 0
                    Do While i < n4500
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n4500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP4500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP4500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    i = i + 1
                    Loop
                    
                End If
                
                If n3000 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n3000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP3000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    
                End If
                    
                If n1500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n1500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP1500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP1500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                
                End If
                
                If n900 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n900, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP900 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP900 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                
                End If
                
                'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 10 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 10 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
                        
            ElseIf lon = "3000" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v300n3000, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Perfiles INCYE"
                Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqa(0) = Esqa(0) - lP3000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            
                lperfil = distanciaA - lbisagra - lP3000
                n3000 = Fix(lperfil / lP3000)
                lperfil = lperfil - n3000 * lP3000
                n1500 = Fix(lperfil / lP1500)
                lperfil = lperfil - n1500 * lP1500
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)

                If n3000 > 0 Then
                    i = 0
                    Do While i < n3000
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n3000, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP3000 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP3000 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    i = i + 1
                    Loop
                    
                End If
                
                If n1500 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n1500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP1500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP1500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                
                End If
                
                If n900 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n900, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP900 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP900 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                
                End If
                
                'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 10 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 10 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
            ElseIf lon = "1500" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v300n1500, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Perfiles INCYE"
                Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqa(0) = Esqa(0) - lP1500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP1500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            
                lperfil = distanciaA - lbisagra - lP1500
                n1500 = Fix(lperfil / lP1500)
                lperfil = lperfil - n1500 * lP1500
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)

                If n1500 > 0 Then
                    i = 0
                    Do While i < n1500
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n1500, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP1500 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP1500 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    i = i + 1
                    Loop
                    
                End If
                            
                If n900 > 0 Then
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n900, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP900 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP900 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                
                End If
                
                'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 10 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 10 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
                        
            ElseIf lon = "900" Then
            
            'Inserccion primer perfil obligatorio
                Set blockRef = doc.ModelSpace.InsertBlock(Esqa, v300n900, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Perfiles INCYE"
                Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                Esqa(0) = Esqa(0) - lP900 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP900 * Sin(DirMuro1): Esqa(2) = Esqa(2)
            
                lperfil = distanciaA - lbisagra - lP900
                n900 = Fix(lperfil / lP900)
                lperfil = lperfil - n900 * lP900
                Mp450 = Fix(lperfil / lMp450)
                    Mp270 = Fix(lperfil / lMp270)
                    Mp180 = Fix(lperfil / lMp180)
                    Mp90 = Fix(lperfil / lMp90)

                If n900 > 0 Then
                    i = 0
                    Do While i < n900
                    Set blockRef = gcadModel.InsertBlock(Esqa, v300n900, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Perfiles INCYE"
                    Set blockRef = gcadModel.InsertBlock(Esqa, M20x130_6, Xs, Ys, Zs, DirMuro1 + PI)
                    blockRef.Layer = "Nonplot"
                    blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                    Esqa(0) = Esqa(0) - lP900 * Cos(DirMuro1): Esqa(1) = Esqa(1) - lP900 * Sin(DirMuro1): Esqa(2) = Esqa(2)
                    i = i + 1
                    Loop
                    
                End If
                
                'Nivelar Megapro
                            Esqa(0) = Esqa(0) - 10 * Cos(DirMuro1a): Esqa(1) = Esqa(1) - 10 * Sin(DirMuro1a): Esqa(2) = Esqa(2)
                        
                            Call MegaproLadoA(Esqa, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
            
            End If
            
        End If
 
terminar:
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub MegaproLadoA(Esqa() As Double, DirMuro1 As Double, Mp90 As Integer, Mp180 As Integer, Mp450 As Integer, Mp270 As Integer, lperfil As Double)
    
    Dim doc As Object
    Set doc = ThisDrawing
    
    Dim blockRef As Object
    
    Set gcadDoc = GetObject(, "Gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
       
    Dim Xs As Double
    Dim Ys As Double
    Dim Zs As Double
    Dim lMp90 As Double, lMp270 As Double, lMp450 As Double, lMp180 As Double
    Dim ruta2 As String
    Dim rutamp As String
    Dim repite As Integer
    
    rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
    ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    repite = 1
    
    'Valores fijos
    PI = 4 * Atn(1)
    lMp90 = 90
    lMp180 = 180
    lMp270 = 270
    lMp450 = 450
    
            Xs = 1
            Ys = 1
            Zs = 1
            
            Mpshor450 = rutamp & "Mshor450ALZ.dwg"
            Mpshor270 = rutamp & "Mshor270ALZ.dwg"
            Mpshor180 = rutamp & "Mshor180ALZ.dwg"
            Mpshor90 = rutamp & "Mshor90ALZ.dwg"
            
            M20x90_6 = ruta2 & "6-M20x90.dwg"
                    
            Mp450 = Fix(lperfil / lMp450)
            Mp270 = Fix(lperfil / lMp270)
            Mp180 = Fix(lperfil / lMp180)
            Mp90 = Fix(lperfil / lMp90)

                            If Mp450 > 0 Then
                        
                            Set blockRef = gcadModel.InsertBlock(Esqa, Mpshor450, Xs, Ys, Zs, DirMuro1 + PI)
                            blockRef.Layer = "Mega"
                            Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_6, Xs, Ys, Zs, DirMuro1 + PI)
                            blockRef.Layer = "Nonplot"
                            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                            
                            ElseIf Mp270 > 0 Then
                        
                            Set blockRef = gcadModel.InsertBlock(Esqa, Mpshor270, Xs, Ys, Zs, DirMuro1 + PI)
                            blockRef.Layer = "Mega"
                            Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_6, Xs, Ys, Zs, DirMuro1 + PI)
                            blockRef.Layer = "Nonplot"
                            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                            
                            ElseIf Mp180 > 0 Then
                        
                            Set blockRef = gcadModel.InsertBlock(Esqa, Mpshor180, Xs, Ys, Zs, DirMuro1 + PI)
                            blockRef.Layer = "Mega"
                            Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_6, Xs, Ys, Zs, DirMuro1 + PI)
                            blockRef.Layer = "Nonplot"
                            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                            
                            ElseIf Mp90 > 0 Then
                        
                            Set blockRef = gcadModel.InsertBlock(Esqa, Mpshor90, Xs, Ys, Zs, DirMuro1 + PI)
                            blockRef.Layer = "Mega"
                            Set blockRef = gcadModel.InsertBlock(Esqa, M20x90_6, Xs, Ys, Zs, DirMuro1 + PI)
                            blockRef.Layer = "Nonplot"
                            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                            
                            End If

terminar:
End Sub

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub MegaproLadoB(Esqb() As Double, DirMuro2 As Double, Mp90 As Integer, Mp180 As Integer, Mp450 As Integer, Mp270 As Integer, lperfil As Double)
    
    Dim doc As Object
    Set doc = ThisDrawing
    
    Dim blockRef As Object
    
    Set gcadDoc = GetObject(, "Gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
       
    Dim Xs As Double
    Dim Ys As Double
    Dim Zs As Double
    Dim lMp90 As Double, lMp270 As Double, lMp450 As Double, lMp180 As Double
    Dim ruta2 As String
    Dim rutamp As String
    Dim repite As Integer
    
    rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
    ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    repite = 1
    
    'Valores fijos
    PI = 4 * Atn(1)
    lMp90 = 90
    lMp180 = 180
    lMp270 = 270
    lMp450 = 450
    
            Xs = 1
            Ys = 1
            Zs = 1
            
            Mpshor450 = rutamp & "Mshor450ALZ.dwg"
            Mpshor270 = rutamp & "Mshor270ALZ.dwg"
            Mpshor180 = rutamp & "Mshor180ALZ.dwg"
            Mpshor90 = rutamp & "Mshor90ALZ.dwg"
            
            M20x90_6 = ruta2 & "6-M20x90.dwg"
                    
            Mp450 = Fix(lperfil / lMp450)
            Mp270 = Fix(lperfil / lMp270)
            Mp180 = Fix(lperfil / lMp180)
            Mp90 = Fix(lperfil / lMp90)
            
                            If Mp450 > 0 Then
                            Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_6, Xs, Ys, Zs, DirMuro2)
                            blockRef.Layer = "Nonplot"
                            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                            Esqb(0) = Esqb(0) - lMp450 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lMp450 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                            Set blockRef = gcadModel.InsertBlock(Esqb, Mpshor450, Xs, Ys, Zs, DirMuro2)
                            blockRef.Layer = "Mega"
                        
                            ElseIf Mp270 > 0 Then
                            Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_6, Xs, Ys, Zs, DirMuro2)
                            blockRef.Layer = "Nonplot"
                            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                            Esqb(0) = Esqb(0) - lMp270 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lMp270 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                            Set blockRef = gcadModel.InsertBlock(Esqb, Mpshor270, Xs, Ys, Zs, DirMuro2)
                            blockRef.Layer = "Mega"
                        
                            ElseIf Mp180 > 0 Then
                            Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_6, Xs, Ys, Zs, DirMuro2)
                            blockRef.Layer = "Nonplot"
                            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                            Esqb(0) = Esqb(0) - lMp180 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lMp180 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                            Set blockRef = gcadModel.InsertBlock(Esqb, Mpshor180, Xs, Ys, Zs, DirMuro2)
                            blockRef.Layer = "Mega"
                        
                            ElseIf Mp90 > 0 Then
                            Set blockRef = gcadModel.InsertBlock(Esqb, M20x90_6, Xs, Ys, Zs, DirMuro2)
                            blockRef.Layer = "Nonplot"
                            blockRef.Update
                    blockRef.Explode
                    blockRef.Delete
                            Esqb(0) = Esqb(0) - lMp90 * Cos(DirMuro2): Esqb(1) = Esqb(1) - lMp90 * Sin(DirMuro2): Esqb(2) = Esqb(2)
                            Set blockRef = gcadModel.InsertBlock(Esqb, Mpshor90, Xs, Ys, Zs, DirMuro2)
                            blockRef.Layer = "Mega"
                        
                            End If
                            
terminar:
End Sub
                            


