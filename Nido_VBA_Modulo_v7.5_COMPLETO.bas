'=======================================================================
' NIDO MANAGER PROYECTOS - Modulo VBA Principal
' Nido Arquitectura v7.4
'=======================================================================
Option Explicit

Const PREFIJO_CLI   As String = "CLI-"
Const PREFIJO_OBR   As String = "OBR-"
Const RUTA_ONEDRIVE As String = "C:\Users\juanc\OneDrive - CAMPANA\MANGER NIDO\"

'=======================================================================
' COTIZACION USD OFICIAL (promedio compra/venta, fuente vinculada a BCRA)
' Fuente: dolarapi.com/v1/dolares/oficial (bancos y casas de cambio
' autorizadas por el BCRA). Cachea el ultimo valor obtenido en
' OBRAS_ACTIVAS!O1 por si no hay internet en el momento de recalcular.
'=======================================================================
Function ObtenerCotizacionUSD() As Double
    Dim http      As Object
    Dim resp      As String
    Dim compra    As Double
    Dim venta     As Double
    Dim p1        As Long
    Dim p2        As Long
    Dim wsCache   As Worksheet

    On Error GoTo ErrorFetch
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.SetTimeouts 5000, 5000, 5000, 5000
    http.Open "GET", "https://dolarapi.com/v1/dolares/oficial", False
    http.setRequestHeader "User-Agent", "NidoManager-VBA"
    http.send
    If http.Status <> 200 Then GoTo ErrorFetch
    resp = http.responseText

    p1 = InStr(resp, Chr(34) & "compra" & Chr(34))
    If p1 = 0 Then GoTo ErrorFetch
    p1 = InStr(p1, resp, ":") + 1
    p2 = InStr(p1, resp, ",")
    compra = Val(Mid(resp, p1, p2 - p1))

    p1 = InStr(resp, Chr(34) & "venta" & Chr(34))
    If p1 = 0 Then GoTo ErrorFetch
    p1 = InStr(p1, resp, ":") + 1
    p2 = InStr(p1, resp, ",")
    venta = Val(Mid(resp, p1, p2 - p1))

    If compra <= 0 Or venta <= 0 Then GoTo ErrorFetch

    ObtenerCotizacionUSD = (compra + venta) / 2

    ' Guardar en cache por si la proxima vez falla la conexion
    On Error Resume Next
    Set wsCache = ThisWorkbook.Sheets("OBRAS_ACTIVAS")
    If Not wsCache Is Nothing Then
        wsCache.Range("N1").Value = "Cotizacion USD oficial (prom. C/V):"
        wsCache.Range("O1").Value = ObtenerCotizacionUSD
        wsCache.Range("N2").Value = "Actualizada:"
        wsCache.Range("O2").Value = Now
        wsCache.Range("O2").NumberFormat = "DD/MM/YYYY HH:MM"
    End If
    On Error GoTo 0
    Exit Function

ErrorFetch:
    ' Sin conexion o la API no respondio: usar el ultimo valor cacheado
    On Error Resume Next
    Set wsCache = ThisWorkbook.Sheets("OBRAS_ACTIVAS")
    If Not wsCache Is Nothing Then
        If IsNumeric(wsCache.Range("O1").Value) Then
            ObtenerCotizacionUSD = CDbl(wsCache.Range("O1").Value)
        Else
            ObtenerCotizacionUSD = 0
        End If
    Else
        ObtenerCotizacionUSD = 0
    End If
    On Error GoTo 0
End Function

'=======================================================================
' MACRO 1: REGISTRAR CLIENTE
'=======================================================================
Sub MacroRegistrarCliente()

    Dim wsForm        As Worksheet
    Dim wsReg         As Worksheet
    Dim codCli        As String
    Dim nroCli        As Long
    Dim maxNum        As Long
    Dim i             As Long
    Dim cod           As String
    Dim apellido      As String
    Dim dni           As String
    Dim carpeta       As String
    Dim nombreCarpeta As String
    Dim msgCarpeta    As String
    Dim carpetaAnterior As String
    Dim dniFinal      As String
    Dim nuevaFila     As Long
    Dim resp          As Integer
    Dim bgColor       As Long

    Set wsForm = ThisWorkbook.Sheets("ALT_CLIENTE")
    Set wsReg = ThisWorkbook.Sheets("CLIENTES")

    apellido = Trim(CStr(wsForm.Cells(5, 3).Value))
    If Len(apellido) = 0 Then
        MsgBox "El campo APELLIDO / RAZON SOCIAL es obligatorio.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If

    dni = Trim(CStr(wsForm.Cells(7, 3).Value))

    If dni <> "" Then
        For i = 5 To 204
            If Trim(CStr(wsReg.Cells(i, 4).Value)) = dni Then
                resp = MsgBox("Ya existe un cliente con ese DNI/CUIT:" & vbCrLf & _
                    wsReg.Cells(i, 2).Value & vbCrLf & "Registrarlo igual?", _
                    vbQuestion + vbYesNo, "Nido Manager")
                If resp = vbNo Then GoTo Salir
                Exit For
            End If
        Next i
    End If

    nuevaFila = 0
    For i = 5 To 204
        If Trim(CStr(wsReg.Cells(i, 1).Value)) = "" Then
            nuevaFila = i
            Exit For
        End If
    Next i
    If nuevaFila = 0 Then
        MsgBox "La tabla esta llena (200 registros maximo).", vbExclamation, "Nido Manager"
        GoTo Salir
    End If

    maxNum = 0
    For i = 5 To 204
        cod = CStr(wsReg.Cells(i, 1).Value)
        If Left(cod, Len(PREFIJO_CLI)) = PREFIJO_CLI Then
            If IsNumeric(Mid(cod, Len(PREFIJO_CLI) + 1)) Then
                If CLng(Mid(cod, Len(PREFIJO_CLI) + 1)) > maxNum Then
                    maxNum = CLng(Mid(cod, Len(PREFIJO_CLI) + 1))
                End If
            End If
        End If
    Next i
    nroCli = maxNum + 1
    codCli = PREFIJO_CLI & Format(nroCli, "0000")

    wsReg.Cells(nuevaFila, 1).Value = codCli
    wsReg.Cells(nuevaFila, 2).Value = apellido
    wsReg.Cells(nuevaFila, 3).Value = wsForm.Cells(6, 3).Value
    wsReg.Cells(nuevaFila, 4).Value = dni
    wsReg.Cells(nuevaFila, 5).Value = wsForm.Cells(8, 3).Value
    wsReg.Cells(nuevaFila, 6).Value = wsForm.Cells(9, 3).Value
    wsReg.Cells(nuevaFila, 7).Value = wsForm.Cells(10, 3).Value
    wsReg.Cells(nuevaFila, 8).Value = wsForm.Cells(11, 3).Value
    wsReg.Cells(nuevaFila, 9).Value = wsForm.Cells(12, 3).Value
    wsReg.Cells(nuevaFila, 10).Value = wsForm.Cells(13, 3).Value
    wsReg.Cells(nuevaFila, 11).Value = "Activo"
    wsReg.Cells(nuevaFila, 13).Value = codCli & " | " & apellido

    If nuevaFila Mod 2 = 0 Then bgColor = RGB(254, 243, 220) Else bgColor = RGB(255, 255, 255)
    For i = 1 To 13
        With wsReg.Cells(nuevaFila, i)
            .Font.Name = "Abel"
            .Font.Size = 9
            .Font.Color = RGB(90, 98, 110)
            .Interior.Color = bgColor
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(221, 224, 227)
        End With
    Next i
    With wsReg.Cells(nuevaFila, 13)
        .Font.Italic = True
        .Interior.Color = RGB(254, 243, 220)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = RGB(90, 98, 110)
    End With

    carpetaAnterior = Trim(CStr(wsForm.Cells(16, 3).Value))
    If carpetaAnterior <> "" Then
        carpeta = carpetaAnterior
        If Dir(carpeta, vbDirectory) <> "" Then
            msgCarpeta = "Carpeta existente vinculada: " & carpetaAnterior
        Else
            On Error Resume Next: MkDir carpeta: On Error GoTo 0
            If Dir(carpeta, vbDirectory) <> "" Then
                msgCarpeta = "Carpeta creada en ruta indicada: " & carpeta
            Else
                msgCarpeta = "ATENCION: No se pudo acceder a la carpeta indicada."
            End If
        End If
    Else
        dniFinal = IIf(dni <> "", dni, "SIN-DNI")
        nombreCarpeta = LimpiarNombre(apellido) & "_" & LimpiarNombre(dniFinal)
        carpeta = RUTA_ONEDRIVE & "Clientes\" & nombreCarpeta
        If Dir(RUTA_ONEDRIVE & "Clientes\", vbDirectory) = "" Then
            On Error Resume Next: MkDir RUTA_ONEDRIVE & "Clientes\": On Error GoTo 0
        End If
        If Dir(carpeta, vbDirectory) = "" Then
            On Error Resume Next
            MkDir carpeta
            If Dir(carpeta, vbDirectory) <> "" Then
                MkDir carpeta & "\Contratos"
                MkDir carpeta & "\Planos"
                MkDir carpeta & "\Fotos"
                MkDir carpeta & "\Presupuestos"
                MkDir carpeta & "\Pagos"
                msgCarpeta = "Carpeta creada: " & nombreCarpeta
            Else
                msgCarpeta = "No se pudo crear la carpeta (verifica la ruta OneDrive)"
            End If
            On Error GoTo 0
        Else
            msgCarpeta = "La carpeta ya existia: " & nombreCarpeta
        End If
    End If
    wsReg.Cells(nuevaFila, 12).Value = carpeta

    Call LimpiarFormCliente(wsForm)
    MsgBox "Cliente registrado correctamente." & vbCrLf & _
           "Codigo: " & codCli & vbCrLf & _
           "Fila: " & nuevaFila & vbCrLf & vbCrLf & msgCarpeta, _
           vbInformation, "Nido Manager"

Salir:
    Set wsForm = Nothing: Set wsReg = Nothing
End Sub

'=======================================================================
' MACRO 2: REGISTRAR OBRA
'=======================================================================
Sub MacroRegistrarObra()

    Dim wsForm          As Worksheet
    Dim wsReg           As Worksheet
    Dim wsCli           As Worksheet
    Dim codObr          As String
    Dim nroObr          As Long
    Dim maxO            As Long
    Dim i               As Long
    Dim j               As Integer
    Dim codO            As String
    Dim fechaInicio     As Date
    Dim fechaFin        As Date
    Dim plazo           As Long
    Dim codCliIngresado As String
    Dim nombreCliente   As String
    Dim valorC6         As String
    Dim hayEtapas       As Boolean
    Dim totalHon        As Double
    Dim respHon         As Integer
    Dim presupuesto     As Double
    Dim nroEtapa        As Integer
    Dim etapaNombre     As String
    Dim etapaPct        As Double
    Dim rowEtH          As Long
    Dim ultFilaH        As Long
    Dim ultFilaMO       As Long
    Dim wsEt            As Worksheet
    Dim existeHoja      As Boolean
    Dim sh              As Worksheet
    Dim hdrs            As Variant
    Dim h               As Integer
    Dim cH              As Object
    Dim moData(1, 1)    As Variant
    Dim nuevaFila       As Long
    Dim bgColor         As Long
    Dim nombreObra      As String
    Dim rawPctV         As Double
    Dim dblPct          As Double
    Dim moneda          As String
    Dim intentosOD2     As Integer
    Dim maxIntentosOD2  As Integer

    Set wsForm = ThisWorkbook.Sheets("ALT_OBRA")
    Set wsReg = ThisWorkbook.Sheets("OBRAS")
    Set wsCli = ThisWorkbook.Sheets("CLIENTES")

    nombreObra = Trim(CStr(wsForm.Cells(5, 3).Value))
    valorC6 = Trim(CStr(wsForm.Cells(6, 3).Value))

    If Len(nombreObra) = 0 Then
        MsgBox "El NOMBRE DE LA OBRA es obligatorio.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If
    If Len(valorC6) = 0 Then
        MsgBox "El CLIENTE es obligatorio. Selecciona de la lista desplegable.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If
    If Len(Trim(CStr(wsForm.Cells(9, 3).Value))) = 0 Then
        MsgBox "La FECHA DE INICIO es obligatoria.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If
    If Not IsNumeric(wsForm.Cells(11, 3).Value) Or CDbl(wsForm.Cells(11, 3).Value) = 0 Then
        MsgBox "El PRESUPUESTO es obligatorio y debe ser mayor a cero.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If

    ' Fila 12 = MONEDA (nueva), filas 13+ se corrieron
    moneda = Trim(CStr(wsForm.Cells(12, 3).Value))
    If moneda = "" Then moneda = "ARS"

    If InStr(valorC6, " | ") > 0 Then
        codCliIngresado = Trim(Left(valorC6, InStr(valorC6, " | ") - 1))
    Else
        codCliIngresado = valorC6
    End If

    nombreCliente = ""
    For i = 5 To 204
        If Trim(CStr(wsCli.Cells(i, 1).Value)) = codCliIngresado Then
            nombreCliente = wsCli.Cells(i, 2).Value
            Exit For
        End If
    Next i
    If nombreCliente = "" Then
        MsgBox "El codigo '" & codCliIngresado & "' no existe en CLIENTES." & vbCrLf & _
               "Registra el cliente primero.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If

    hayEtapas = False
    totalHon = 0
    For i = 23 To 25
        rawPctV = 0
        On Error Resume Next
        rawPctV = CDbl(wsForm.Cells(i, 4).Value)
        On Error GoTo 0
        If rawPctV > 0 Then
            If rawPctV <= 1 Then rawPctV = rawPctV * 100
            hayEtapas = True
            totalHon = totalHon + rawPctV
        End If
    Next i
    If hayEtapas And Abs(totalHon - 100) > 0.01 Then
        respHon = MsgBox("Honorarios: los % suman " & totalHon & "% (deben ser 100%)." & vbCrLf & _
                         "Registrar igual?", vbQuestion + vbYesNo, "Nido Manager")
        If respHon = vbNo Then GoTo Salir
    End If

    nuevaFila = 0
    For i = 5 To 204
        If Trim(CStr(wsReg.Cells(i, 1).Value)) = "" Then
            nuevaFila = i
            Exit For
        End If
    Next i
    If nuevaFila = 0 Then
        MsgBox "La tabla de obras esta llena.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If

    maxO = 0
    For i = 5 To 204
        codO = CStr(wsReg.Cells(i, 1).Value)
        If Left(codO, Len(PREFIJO_OBR)) = PREFIJO_OBR Then
            If IsNumeric(Mid(codO, Len(PREFIJO_OBR) + 1)) Then
                If CLng(Mid(codO, Len(PREFIJO_OBR) + 1)) > maxO Then
                    maxO = CLng(Mid(codO, Len(PREFIJO_OBR) + 1))
                End If
            End If
        End If
    Next i
    nroObr = maxO + 1
    codObr = PREFIJO_OBR & Format(nroObr, "0000")

    fechaInicio = CDate(wsForm.Cells(9, 3).Value)
    plazo = 0
    If IsNumeric(wsForm.Cells(10, 3).Value) Then plazo = CLng(wsForm.Cells(10, 3).Value)

    ' Escribir en hoja OBRAS
    ' Cols: 1=COD | 2=NOMBRE | 3=COD_CLI | 4=CLIENTE | 5=TIPO | 6=DIRECCION |
    '       7=FECHA INICIO | 8=PLAZO | 9=FECHA FIN | 10=PRESUPUESTO |
    '       11=MONEDA | 12=ESTRUCTURA COBRO | 13=ESTADO
    wsReg.Cells(nuevaFila, 1).Value = codObr
    wsReg.Cells(nuevaFila, 2).Value = nombreObra
    wsReg.Cells(nuevaFila, 3).Value = codCliIngresado
    wsReg.Cells(nuevaFila, 4).Value = nombreCliente
    wsReg.Cells(nuevaFila, 5).Value = wsForm.Cells(7, 3).Value
    wsReg.Cells(nuevaFila, 6).Value = wsForm.Cells(8, 3).Value
    wsReg.Cells(nuevaFila, 7).Value = fechaInicio
    wsReg.Cells(nuevaFila, 7).NumberFormat = "DD/MM/YYYY"
    wsReg.Cells(nuevaFila, 8).Value = plazo
    If plazo > 0 Then
        fechaFin = fechaInicio + plazo
        wsReg.Cells(nuevaFila, 9).Value = fechaFin
        wsReg.Cells(nuevaFila, 9).NumberFormat = "DD/MM/YYYY"
    End If
    wsReg.Cells(nuevaFila, 10).Value = CDbl(wsForm.Cells(11, 3).Value)
    wsReg.Cells(nuevaFila, 10).NumberFormat = "$#,##0;($#,##0);-"
    wsReg.Cells(nuevaFila, 11).Value = moneda
    wsReg.Cells(nuevaFila, 12).Value = wsForm.Cells(13, 3).Value
    wsReg.Cells(nuevaFila, 13).Value = "Activa"

    If nuevaFila Mod 2 = 0 Then bgColor = RGB(254, 243, 220) Else bgColor = RGB(255, 255, 255)
    For i = 1 To 13
        With wsReg.Cells(nuevaFila, i)
            .Font.Name = "Abel": .Font.Size = 9: .Font.Color = RGB(90, 98, 110)
            .Interior.Color = bgColor
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(221, 224, 227)
        End With
    Next i
    wsReg.Cells(nuevaFila, 1).HorizontalAlignment = xlCenter
    wsReg.Cells(nuevaFila, 13).HorizontalAlignment = xlCenter

    existeHoja = False
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "ETAPAS_COBRO" Then existeHoja = True
    Next sh
    If Not existeHoja Then
        Set wsEt = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("OBRAS"))
        wsEt.Name = "ETAPAS_COBRO"
        hdrs = Array("COD. OBRA", "OBRA", "COD. CLIENTE", "CLIENTE", "N ETAPA", "NOMBRE ETAPA", "TIPO", "% COBRO", "MONTO ($)", "ESTADO")
        For h = 0 To 9
            Set cH = wsEt.Cells(1, h + 1)
            cH.Value = hdrs(h): cH.Font.Bold = True: cH.Font.Name = "Abel": cH.Font.Size = 9
            cH.Interior.Color = RGB(251, 180, 43): cH.HorizontalAlignment = xlCenter
        Next h
        wsEt.Columns("A").ColumnWidth = 12: wsEt.Columns("B").ColumnWidth = 28
        wsEt.Columns("C").ColumnWidth = 13: wsEt.Columns("D").ColumnWidth = 24
        wsEt.Columns("E").ColumnWidth = 10: wsEt.Columns("F").ColumnWidth = 26
        wsEt.Columns("G").ColumnWidth = 14: wsEt.Columns("H").ColumnWidth = 10
        wsEt.Columns("I").ColumnWidth = 16: wsEt.Columns("J").ColumnWidth = 14
    Else
        Set wsEt = ThisWorkbook.Sheets("ETAPAS_COBRO")
    End If

    presupuesto = CDbl(wsForm.Cells(11, 3).Value)
    nroEtapa = 0

    ' Etapas corridas a filas 23-25 por insercion de MONEDA en fila 12
    For rowEtH = 23 To 25
        etapaNombre = Trim(CStr(wsForm.Cells(rowEtH, 3).Value))
        dblPct = 0
        On Error Resume Next
        dblPct = CDbl(wsForm.Cells(rowEtH, 4).Value)
        On Error GoTo 0
        If dblPct > 0 And dblPct <= 1 Then dblPct = dblPct * 100
        If etapaNombre <> "" And dblPct > 0 Then
            nroEtapa = nroEtapa + 1
            etapaPct = dblPct
            ultFilaH = wsEt.Cells(wsEt.Rows.Count, 1).End(xlUp).Row + 1
            If ultFilaH < 2 Then ultFilaH = 2
            wsEt.Cells(ultFilaH, 1) = codObr: wsEt.Cells(ultFilaH, 2) = nombreObra
            wsEt.Cells(ultFilaH, 3) = codCliIngresado: wsEt.Cells(ultFilaH, 4) = nombreCliente
            wsEt.Cells(ultFilaH, 5) = nroEtapa: wsEt.Cells(ultFilaH, 6) = etapaNombre
            wsEt.Cells(ultFilaH, 7) = "Honorario"
            wsEt.Cells(ultFilaH, 8) = etapaPct / 100: wsEt.Cells(ultFilaH, 8).NumberFormat = "0.0%"
            wsEt.Cells(ultFilaH, 9) = presupuesto * (etapaPct / 100): wsEt.Cells(ultFilaH, 9).NumberFormat = "$#,##0;($#,##0);-"
            wsEt.Cells(ultFilaH, 10) = "Pendiente"
            Call FormatoFilaEt(wsEt, ultFilaH)
        End If
    Next rowEtH

    moData(0, 0) = "Subcontrato Mano de Obra": moData(0, 1) = 97
    moData(1, 0) = "Subcontrato":              moData(1, 1) = 3
    For j = 0 To 1
        nroEtapa = nroEtapa + 1: etapaPct = moData(j, 1)
        ultFilaMO = wsEt.Cells(wsEt.Rows.Count, 1).End(xlUp).Row + 1
        If ultFilaMO < 2 Then ultFilaMO = 2
        wsEt.Cells(ultFilaMO, 1) = codObr: wsEt.Cells(ultFilaMO, 2) = nombreObra
        wsEt.Cells(ultFilaMO, 3) = codCliIngresado: wsEt.Cells(ultFilaMO, 4) = nombreCliente
        wsEt.Cells(ultFilaMO, 5) = nroEtapa: wsEt.Cells(ultFilaMO, 6) = moData(j, 0)
        wsEt.Cells(ultFilaMO, 7) = "Mano de Obra"
        wsEt.Cells(ultFilaMO, 8) = etapaPct / 100: wsEt.Cells(ultFilaMO, 8).NumberFormat = "0.0%"
        wsEt.Cells(ultFilaMO, 9) = presupuesto * (etapaPct / 100): wsEt.Cells(ultFilaMO, 9).NumberFormat = "$#,##0;($#,##0);-"
        wsEt.Cells(ultFilaMO, 10) = "Pendiente"
        Call FormatoFilaEt(wsEt, ultFilaMO)
    Next j

    Application.CalculateFullRebuild

    ' -- Actualizar OBRAS_ACTIVAS local ---------------------------
    ' Cols: 1=COD | 2=OBRA | 3=COD_CLI | 4=NOMBRE_CLI | 5=PRESUPUESTO |
    '       6=COBRADO | 7=EGRESOS | 8=RESULTADO | 9=DEBE | 10=%COBRADO |
    '       11=DIAS RESTANTES | 12=ESTADO
    Dim wsOA   As Worksheet
    Dim iiOA   As Long
    Dim ultOA  As Long
    Dim presOA As Double
    Set wsOA = ThisWorkbook.Sheets("OBRAS_ACTIVAS")
    ultOA = 0
    For iiOA = 5 To 200
        If Trim(CStr(wsOA.Cells(iiOA, 1).Value)) = "" Then
            ultOA = iiOA
            Exit For
        End If
    Next iiOA
    If ultOA = 0 Then ultOA = 5
    presOA = CDbl(wsForm.Cells(11, 3).Value)
    wsOA.Cells(ultOA, 1).Value = codObr
    wsOA.Cells(ultOA, 2).Value = nombreObra
    wsOA.Cells(ultOA, 3).Value = codCliIngresado
    wsOA.Cells(ultOA, 4).Value = nombreCliente
    wsOA.Cells(ultOA, 5).Value = presOA
    wsOA.Cells(ultOA, 5).NumberFormat = "$#,##0;($#,##0);-"
    wsOA.Cells(ultOA, 6).Value = 0
    wsOA.Cells(ultOA, 6).NumberFormat = "$#,##0;($#,##0);-"
    wsOA.Cells(ultOA, 7).Value = 0
    wsOA.Cells(ultOA, 7).NumberFormat = "$#,##0;($#,##0);-"
    wsOA.Cells(ultOA, 8).Value = presOA
    wsOA.Cells(ultOA, 8).NumberFormat = "$#,##0;($#,##0);-"
    wsOA.Cells(ultOA, 9).Value = presOA
    wsOA.Cells(ultOA, 9).NumberFormat = "$#,##0;($#,##0);-"
    wsOA.Cells(ultOA, 10).Value = 0
    wsOA.Cells(ultOA, 10).NumberFormat = "0.0%"
    wsOA.Cells(ultOA, 11).Value = ""
    wsOA.Cells(ultOA, 12).Value = "Activa"
    Set wsOA = Nothing

    ' -- Sincronizar OBRAS_ACTIVAS con OneDrive -------------------
    Dim wbOD2   As Workbook
    Dim wsOD2   As Worksheet
    Dim rutaOD2 As String
    Dim ultOD2  As Long
    Dim iiOD2   As Long
    Dim presOD2 As Double
    Dim msgSync As String

    rutaOD2 = RUTA_ONEDRIVE & "Nido_Manager_Sistema_v4.xlsx"

    ' Verificar que el archivo existe y NO esta abierto en browser
    If Dir(rutaOD2) = "" Then
        msgSync = "ADVERTENCIA: No se encontro el archivo en OneDrive." & vbCrLf & _
                  "Ruta: " & rutaOD2 & vbCrLf & _
                  "La obra se guardo localmente. Sincroniza manualmente."
        GoTo MostrarMensaje
    End If

    ' Verificar si ya esta abierto (conflicto con browser)
    Dim wbYaAbierto As Workbook
    Dim estaAbierto As Boolean
    estaAbierto = False
    For Each wbYaAbierto In Workbooks
        If InStr(LCase(wbYaAbierto.FullName), "nido_manager_sistema_v4") > 0 And _
           wbYaAbierto.Name <> ThisWorkbook.Name Then
            estaAbierto = True
            Exit For
        End If
    Next wbYaAbierto

    If estaAbierto Then
        ' Usar el libro ya abierto
        Set wbOD2 = wbYaAbierto
        Set wsOD2 = wbOD2.Sheets("OBRAS_ACTIVAS")
    Else
        ' Abrir el archivo, con reintento automatico ante hipo de red/token
        Application.ScreenUpdating = False
        maxIntentosOD2 = 3
        For intentosOD2 = 1 To maxIntentosOD2
            On Error Resume Next
            Set wbOD2 = Nothing
            Set wbOD2 = Workbooks.Open(rutaOD2)
            On Error GoTo 0
            If Not wbOD2 Is Nothing Then Exit For
            If intentosOD2 < maxIntentosOD2 Then
                Application.StatusBar = "Reintentando abrir OneDrive (" & intentosOD2 & "/" & maxIntentosOD2 & ")..."
                Application.Wait Now + TimeValue("0:00:03")
            End If
        Next intentosOD2
        Application.StatusBar = False
        If wbOD2 Is Nothing Then
            msgSync = "ADVERTENCIA: No se pudo abrir el archivo OneDrive despues de " & maxIntentosOD2 & " intentos." & vbCrLf & _
                      "Cerralo del browser si esta abierto, verifica tu conexion, y usa MacroImportarPagosOneDrive luego."
            Application.ScreenUpdating = True
            GoTo MostrarMensaje
        End If
        Set wsOD2 = wbOD2.Sheets("OBRAS_ACTIVAS")
    End If

    ultOD2 = 0
    For iiOD2 = 5 To 200
        If Trim(CStr(wsOD2.Cells(iiOD2, 1).Value)) = "" Then
            ultOD2 = iiOD2
            Exit For
        End If
    Next iiOD2
    If ultOD2 = 0 Then ultOD2 = 5

    presOD2 = CDbl(wsForm.Cells(11, 3).Value)
    wsOD2.Cells(ultOD2, 1).Value = codObr
    wsOD2.Cells(ultOD2, 2).Value = nombreObra
    wsOD2.Cells(ultOD2, 3).Value = codCliIngresado
    wsOD2.Cells(ultOD2, 4).Value = nombreCliente
    wsOD2.Cells(ultOD2, 5).Value = presOD2
    wsOD2.Cells(ultOD2, 5).NumberFormat = "$#,##0;($#,##0);-"
    wsOD2.Cells(ultOD2, 6).Value = 0
    wsOD2.Cells(ultOD2, 6).NumberFormat = "$#,##0;($#,##0);-"
    wsOD2.Cells(ultOD2, 7).Value = 0
    wsOD2.Cells(ultOD2, 7).NumberFormat = "$#,##0;($#,##0);-"
    wsOD2.Cells(ultOD2, 8).Value = presOD2
    wsOD2.Cells(ultOD2, 8).NumberFormat = "$#,##0;($#,##0);-"
    wsOD2.Cells(ultOD2, 9).Value = presOD2
    wsOD2.Cells(ultOD2, 9).NumberFormat = "$#,##0;($#,##0);-"
    wsOD2.Cells(ultOD2, 10).Value = 0
    wsOD2.Cells(ultOD2, 10).NumberFormat = "0.0%"
    wsOD2.Cells(ultOD2, 11).Value = ""
    wsOD2.Cells(ultOD2, 12).Value = "Activa"

    If Not estaAbierto Then
        wbOD2.Save
        wbOD2.Close
    End If
    Application.ScreenUpdating = True
    Set wsOD2 = Nothing: Set wbOD2 = Nothing
    msgSync = "OneDrive sincronizado correctamente."

MostrarMensaje:
    wsForm.Cells(37, 2).Value = "  OBRA REGISTRADA - Codigo: " & codObr & "  |  " & nombreCliente
    Call LimpiarFormObra(wsForm)

    MsgBox "Obra registrada correctamente." & vbCrLf & _
           "Codigo: " & codObr & vbCrLf & _
           "Cliente: " & nombreCliente & vbCrLf & _
           "Moneda: " & moneda & vbCrLf & _
           nroEtapa & " etapas registradas." & vbCrLf & vbCrLf & _
           msgSync, vbInformation, "Nido Manager"

Salir:
    Set wsForm = Nothing: Set wsReg = Nothing: Set wsCli = Nothing: Set wsEt = Nothing
End Sub

'=======================================================================
' MACRO 3: MARCAR ETAPA COBRADA
'=======================================================================
Sub MacroMarcarEtapaCobrada()
    Dim wsEt    As Worksheet
    Dim codObr  As String
    Dim nroEt   As String
    Dim i       As Long
    Dim ultFila As Long

    On Error Resume Next
    Set wsEt = ThisWorkbook.Sheets("ETAPAS_COBRO")
    On Error GoTo 0
    If wsEt Is Nothing Then
        MsgBox "La hoja ETAPAS_COBRO no existe. Registra una obra primero.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If

    codObr = InputBox("Codigo de obra (Ej: OBR-0001):", "Marcar Etapa Cobrada")
    If Trim(codObr) = "" Then GoTo Salir
    nroEt = InputBox("Numero de etapa a marcar como cobrada:", "Marcar Etapa Cobrada")
    If Not IsNumeric(nroEt) Then GoTo Salir

    ultFila = wsEt.Cells(wsEt.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultFila
        If CStr(wsEt.Cells(i, 1).Value) = Trim(codObr) And CStr(wsEt.Cells(i, 5).Value) = Trim(nroEt) Then
            wsEt.Cells(i, 10).Value = "Cobrada"
            wsEt.Cells(i, 10).Interior.Color = RGB(234, 244, 238)
            wsEt.Cells(i, 10).Font.Color = RGB(46, 125, 79)
            wsEt.Cells(i, 10).Font.Bold = True
            MsgBox "Etapa " & nroEt & " de " & codObr & " marcada como Cobrada.", vbInformation, "Nido Manager"
            GoTo Salir
        End If
    Next i
    MsgBox "No se encontro esa etapa.", vbExclamation, "Nido Manager"
Salir:
    Set wsEt = Nothing
End Sub

'=======================================================================
' MACRO 4: CAMBIAR ESTADO DE OBRA
'=======================================================================
Sub MacroCambiarEstadoObra()
    Dim wsOb   As Worksheet
    Dim codObr As String
    Dim opcion As String
    Dim i      As Long

    Set wsOb = ThisWorkbook.Sheets("OBRAS")
    codObr = InputBox("Codigo de la obra (Ej: OBR-0001):", "Cambiar Estado")
    If Trim(codObr) = "" Then GoTo Salir

    For i = 5 To 204
        If Trim(CStr(wsOb.Cells(i, 1).Value)) = Trim(codObr) Then
            opcion = InputBox("Estado actual: " & wsOb.Cells(i, 12).Value & vbCrLf & vbCrLf & _
                "1 - Activa" & vbCrLf & "2 - En pausa" & vbCrLf & _
                "3 - Finalizada" & vbCrLf & "4 - Cancelada", "Nuevo estado")
            Select Case Trim(opcion)
                Case "1": wsOb.Cells(i, 12).Value = "Activa"
                Case "2": wsOb.Cells(i, 12).Value = "En pausa"
                Case "3": wsOb.Cells(i, 12).Value = "Finalizada"
                Case "4": wsOb.Cells(i, 12).Value = "Cancelada"
                Case Else
                    If opcion <> "" Then MsgBox "Opcion invalida.", vbExclamation, "Nido Manager"
                    GoTo Salir
            End Select
            Application.CalculateFullRebuild
            MsgBox "Estado actualizado: " & wsOb.Cells(i, 12).Value, vbInformation, "Nido Manager"
            GoTo Salir
        End If
    Next i
    MsgBox "No se encontro la obra: " & codObr, vbExclamation, "Nido Manager"
Salir:
    Set wsOb = Nothing
End Sub

'=======================================================================
' MACRO 5: ACTUALIZAR OBRAS ACTIVAS
'=======================================================================
Sub MacroActualizarObrasActivas()
    Call RecalcularBalanceObras
    Application.CalculateFullRebuild
    ThisWorkbook.Sheets("OBRAS_ACTIVAS").Activate
    MsgBox "Balance actualizado.", vbInformation, "Nido Manager"
End Sub

'=======================================================================
' RECALCULAR BALANCE: suma pagos por obra y actualiza COBRADO/RESULTADO/DEBE/%COBRADO
' (07/07/2026 - agregado para reflejar cobros automaticamente en OBRAS_ACTIVAS)
'=======================================================================
Private Sub RecalcularBalanceObras()
    Dim wsOA         As Worksheet
    Dim wsPag        As Worksheet
    Dim wsObr        As Worksheet
    Dim ultOA        As Long
    Dim ultPag       As Long
    Dim ultObr       As Long
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim codObra      As String
    Dim totalCobrado As Double
    Dim presupuesto  As Double
    Dim egresos      As Double
    Dim monedaObra   As String
    Dim monedaPago   As String
    Dim montoPago    As Double
    Dim montoSumar   As Double
    Dim cotizacion   As Double
    Dim sinConvertir As Long

    On Error Resume Next
    Set wsOA = ThisWorkbook.Sheets("OBRAS_ACTIVAS")
    Set wsPag = ThisWorkbook.Sheets("PAGOS")
    Set wsObr = ThisWorkbook.Sheets("OBRAS")
    On Error GoTo 0
    If wsOA Is Nothing Or wsPag Is Nothing Then GoTo Salir

    ultOA = wsOA.Cells(wsOA.Rows.Count, 1).End(xlUp).Row
    ultPag = wsPag.Cells(wsPag.Rows.Count, 1).End(xlUp).Row
    If ultOA < 5 Then GoTo Salir

    cotizacion = ObtenerCotizacionUSD()
    sinConvertir = 0

    ultObr = 0
    If Not wsObr Is Nothing Then ultObr = wsObr.Cells(wsObr.Rows.Count, 1).End(xlUp).Row

    For i = 5 To ultOA
        codObra = Trim(CStr(wsOA.Cells(i, 1).Value))
        If codObra <> "" Then

            ' --- Buscar la moneda oficial de esta obra en la hoja OBRAS ---
            monedaObra = "ARS"
            If ultObr >= 2 Then
                For k = 2 To ultObr
                    If Trim(CStr(wsObr.Cells(k, 1).Value)) = codObra Then
                        monedaObra = Trim(CStr(wsObr.Cells(k, 11).Value))
                        If monedaObra = "" Then monedaObra = "ARS"
                        Exit For
                    End If
                Next k
            End If

            totalCobrado = 0
            If ultPag >= 5 Then
                For j = 5 To ultPag
                    If Trim(CStr(wsPag.Cells(j, 3).Value)) = codObra Then
                        montoPago = Val(wsPag.Cells(j, 7).Value)
                        monedaPago = Trim(CStr(wsPag.Cells(j, 8).Value))
                        If monedaPago = "" Then monedaPago = "ARS"

                        If monedaPago = monedaObra Then
                            montoSumar = montoPago
                        ElseIf cotizacion > 0 Then
                            If monedaPago = "USD" And monedaObra = "ARS" Then
                                montoSumar = montoPago * cotizacion
                            ElseIf monedaPago = "ARS" And monedaObra = "USD" Then
                                montoSumar = montoPago / cotizacion
                            Else
                                montoSumar = montoPago  ' combinacion no ARS/USD, no se convierte
                            End If
                        Else
                            ' No hay cotizacion disponible (sin internet y sin cache) - no se puede convertir
                            montoSumar = 0
                            sinConvertir = sinConvertir + 1
                        End If

                        totalCobrado = totalCobrado + montoSumar
                    End If
                Next j
            End If

            presupuesto = 0: On Error Resume Next: presupuesto = CDbl(wsOA.Cells(i, 5).Value): On Error GoTo 0
            egresos = 0: On Error Resume Next: egresos = CDbl(wsOA.Cells(i, 7).Value): On Error GoTo 0

            wsOA.Cells(i, 6).Value = totalCobrado
            wsOA.Cells(i, 6).NumberFormat = "$#,##0;($#,##0);-"
            wsOA.Cells(i, 8).Value = presupuesto - egresos
            wsOA.Cells(i, 8).NumberFormat = "$#,##0;($#,##0);-"
            wsOA.Cells(i, 9).Value = presupuesto - totalCobrado
            wsOA.Cells(i, 9).NumberFormat = "$#,##0;($#,##0);-"
            If presupuesto > 0 Then
                wsOA.Cells(i, 10).Value = totalCobrado / presupuesto
            Else
                wsOA.Cells(i, 10).Value = 0
            End If
            wsOA.Cells(i, 10).NumberFormat = "0.0%"
        End If
    Next i

    If sinConvertir > 0 Then
        MsgBox "Atencion: " & sinConvertir & " pago(s) no se pudieron convertir de moneda " & _
               "porque no hay conexion a internet ni cotizacion en cache." & vbCrLf & _
               "Esos pagos NO se sumaron al balance esta vez. Volve a correr la macro " & _
               "cuando tengas conexion.", vbExclamation, "Nido Manager"
    End If

Salir:
    Set wsOA = Nothing: Set wsPag = Nothing: Set wsObr = Nothing
End Sub

'=======================================================================
' MACRO 6: IMPORTAR PAGOS DESDE ONEDRIVE  (CORREGIDA 07/07/2026)
'=======================================================================
Sub MacroImportarPagosOneDrive()

    Dim rutaOneDrive  As String
    Dim wbOD          As Workbook
    Dim wsOD          As Worksheet
    Dim wsPagos       As Worksheet
    Dim ultFilaOD     As Long
    Dim ultFilaLocal  As Long
    Dim i             As Long
    Dim j             As Long
    Dim copiados      As Long
    Dim fechaOD       As String
    Dim idPagoOD      As String
    Dim yaExiste      As Boolean
    Dim wbYaAbierto   As Workbook
    Dim estaAbierto   As Boolean
    Dim cerrarAlFinal As Boolean
    Dim intentos      As Integer
    Dim maxIntentos   As Integer

    rutaOneDrive = RUTA_ONEDRIVE & "Nido_Manager_Sistema_v4.xlsx"

    If Dir(rutaOneDrive) = "" Then
        MsgBox "No se encontro el archivo en OneDrive." & vbCrLf & _
               "Verifica que OneDrive este sincronizado.", vbExclamation, "Nido Manager"
        GoTo Salir
    End If

    Application.ScreenUpdating = False

    ' --- Chequear si ya esta abierto y reusarlo (evita conflicto de apertura) ---
    estaAbierto = False
    For Each wbYaAbierto In Workbooks
        If InStr(LCase(wbYaAbierto.FullName), "nido_manager_sistema_v4") > 0 And _
           wbYaAbierto.Name <> ThisWorkbook.Name Then
            estaAbierto = True
            Exit For
        End If
    Next wbYaAbierto

    If estaAbierto Then
        Set wbOD = wbYaAbierto
        cerrarAlFinal = False
    Else
        maxIntentos = 3
        For intentos = 1 To maxIntentos
            On Error Resume Next
            Set wbOD = Nothing
            Set wbOD = Workbooks.Open(rutaOneDrive, ReadOnly:=True)
            On Error GoTo 0
            If Not wbOD Is Nothing Then Exit For
            If intentos < maxIntentos Then
                Application.StatusBar = "Reintentando abrir OneDrive (" & intentos & "/" & maxIntentos & ")..."
                Application.Wait Now + TimeValue("0:00:03")
            End If
        Next intentos
        Application.StatusBar = False
        If wbOD Is Nothing Then
            MsgBox "No se pudo abrir el archivo de OneDrive despues de " & maxIntentos & " intentos." & vbCrLf & _
                   "Cerralo del navegador/otra ventana si esta abierto, verifica tu conexion a internet, " & _
                   "y volve a intentar en un minuto.", _
                   vbExclamation, "Nido Manager"
            Application.ScreenUpdating = True
            GoTo Salir
        End If
        cerrarAlFinal = True
    End If

    Set wsOD = wbOD.Sheets("PAGOS")
    Set wsPagos = ThisWorkbook.Sheets("PAGOS")

    ultFilaOD = wsOD.Cells(wsOD.Rows.Count, 1).End(xlUp).Row

    ' --- CORREGIDO: usar End(xlUp) en vez de buscar celda vacia ---
    ultFilaLocal = wsPagos.Cells(wsPagos.Rows.Count, 1).End(xlUp).Row
    If ultFilaLocal < 4 Then ultFilaLocal = 4

    copiados = 0

    For i = 5 To ultFilaOD
        If Trim(CStr(wsOD.Cells(i, 1).Value)) = "" Then GoTo SiguienteFila
        If InStr(LCase(CStr(wsOD.Cells(i, 4).Value)), "total") > 0 Then GoTo SiguienteFila

        fechaOD = CStr(wsOD.Cells(i, 1).Value)
        idPagoOD = Trim(CStr(wsOD.Cells(i, 2).Value))
        yaExiste = False

        If idPagoOD <> "" Then
            ' --- Deduplicacion robusta por ID unico (columna B) ---
            For j = 5 To ultFilaLocal
                If Trim(CStr(wsPagos.Cells(j, 2).Value)) = idPagoOD Then
                    yaExiste = True
                    Exit For
                End If
            Next j
        Else
            ' --- Respaldo para filas viejas sin ID (datos previos a este fix) ---
            For j = 5 To ultFilaLocal
                If CStr(wsPagos.Cells(j, 1).Value) = fechaOD And _
                   CStr(wsPagos.Cells(j, 4).Value) = CStr(wsOD.Cells(i, 4).Value) And _
                   CStr(wsPagos.Cells(j, 7).Value) = CStr(wsOD.Cells(i, 7).Value) And _
                   CStr(wsPagos.Cells(j, 9).Value) = CStr(wsOD.Cells(i, 9).Value) And _
                   CStr(wsPagos.Cells(j, 10).Value) = CStr(wsOD.Cells(i, 10).Value) And _
                   CStr(wsPagos.Cells(j, 11).Value) = CStr(wsOD.Cells(i, 11).Value) Then
                    yaExiste = True
                    Exit For
                End If
            Next j
        End If

        If Not yaExiste Then
            ultFilaLocal = ultFilaLocal + 1
            For j = 1 To 12
                wsPagos.Cells(ultFilaLocal, j).Value = wsOD.Cells(i, j).Value
            Next j
            copiados = copiados + 1
        End If

SiguienteFila:
    Next i

    If cerrarAlFinal Then wbOD.Close SaveChanges:=False
    Application.ScreenUpdating = True

    If copiados > 0 Then
        Call RecalcularBalanceObras
        Application.CalculateFullRebuild
        MsgBox copiados & " cobro(s) nuevo(s) importado(s) desde OneDrive." & vbCrLf & _
               "Balance de OBRAS_ACTIVAS actualizado.", vbInformation, "Nido Manager"
    Else
        MsgBox "No hay cobros nuevos para importar.", vbInformation, "Nido Manager"
    End If

Salir:
    Set wbOD = Nothing: Set wsOD = Nothing: Set wsPagos = Nothing
End Sub

'=======================================================================
' MACRO 7: LIMPIAR FORMATO DE PAGOS (fechas inconsistentes + moneda vacia)
' (10/07/2026 - Punto 2 pendiente: normaliza fechas texto y completa ARS por defecto)
'=======================================================================
Sub MacroLimpiarFormatoPagos()
    Dim wsPag    As Worksheet
    Dim ultPag   As Long
    Dim i        As Long
    Dim valFecha As Variant
    Dim txtFecha As String
    Dim fechaOK  As Date
    Dim corregidasFecha As Long
    Dim corregidasMoneda As Long
    Dim resp     As Integer

    Set wsPag = ThisWorkbook.Sheets("PAGOS")
    ultPag = wsPag.Cells(wsPag.Rows.Count, 1).End(xlUp).Row
    If ultPag < 5 Then
        MsgBox "No hay datos en PAGOS para limpiar.", vbInformation, "Nido Manager"
        GoTo Salir
    End If

    resp = MsgBox("Se van a revisar " & (ultPag - 4) & " fila(s) en PAGOS:" & vbCrLf & _
                  "- Fechas con guiones (03-07-26) se convertiran a DD/MM/AAAA" & vbCrLf & _
                  "- Moneda vacia se completara con ARS" & vbCrLf & vbCrLf & _
                  "Continuar?", vbQuestion + vbYesNo, "Nido Manager")
    If resp = vbNo Then GoTo Salir

    corregidasFecha = 0
    corregidasMoneda = 0

    For i = 5 To ultPag
        ' --- Normalizar fecha ---
        valFecha = wsPag.Cells(i, 1).Value
        If Not IsDate(valFecha) And Not IsEmpty(valFecha) Then
            txtFecha = Trim(CStr(valFecha))
            txtFecha = Replace(txtFecha, "-", "/")
            On Error Resume Next
            fechaOK = CDate(txtFecha)
            On Error GoTo 0
            If fechaOK <> 0 Then
                wsPag.Cells(i, 1).Value = fechaOK
                wsPag.Cells(i, 1).NumberFormat = "DD/MM/YYYY"
                corregidasFecha = corregidasFecha + 1
            End If
        ElseIf IsDate(valFecha) Then
            wsPag.Cells(i, 1).NumberFormat = "DD/MM/YYYY"
        End If

        ' --- Completar moneda vacia con ARS ---
        If Trim(CStr(wsPag.Cells(i, 8).Value)) = "" And Trim(CStr(wsPag.Cells(i, 7).Value)) <> "" Then
            wsPag.Cells(i, 8).Value = "ARS"
            corregidasMoneda = corregidasMoneda + 1
        End If
    Next i

    MsgBox "Limpieza completada:" & vbCrLf & _
           corregidasFecha & " fecha(s) normalizada(s)" & vbCrLf & _
           corregidasMoneda & " moneda(s) completada(s) con ARS", vbInformation, "Nido Manager"

Salir:
    Set wsPag = Nothing
End Sub

'=======================================================================
' HELPERS
'=======================================================================
Private Function LimpiarNombre(txt As String) As String
    Dim res As String, c As String, ii As Long
    res = ""
    For ii = 1 To Len(txt)
        c = Mid(txt, ii, 1)
        If InStr("/\:*?""<>|. ", c) > 0 Then
            res = res & "-"
        Else
            res = res & c
        End If
    Next ii
    Do While InStr(res, "--") > 0
        res = Replace(res, "--", "-")
    Loop
    If Len(res) > 0 Then
        If Right(res, 1) = "-" Then res = Left(res, Len(res) - 1)
    End If
    LimpiarNombre = res
End Function

Private Sub AplicarFormatoFila(rng As Range, nCols As Long)
    Dim col As Long
    For col = 1 To nCols
        With rng.Cells(1, col)
            .Font.Name = "Abel": .Font.Size = 9: .Font.Color = RGB(90, 98, 110)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(221, 224, 227)
            If rng.Row Mod 2 = 0 Then .Interior.Color = RGB(254, 243, 220) Else .Interior.Color = RGB(255, 255, 255)
        End With
    Next col
End Sub

Private Sub FormatoFilaEt(ws As Worksheet, fila As Long)
    Dim col As Integer
    For col = 1 To 10
        With ws.Cells(fila, col)
            .Font.Name = "Abel": .Font.Size = 9: .Font.Color = RGB(90, 98, 110)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(221, 224, 227)
            If fila Mod 2 = 0 Then .Interior.Color = RGB(254, 243, 220) Else .Interior.Color = RGB(255, 255, 255)
        End With
    Next col
End Sub

Private Sub LimpiarFormCliente(ws As Worksheet)
    Dim rowL As Long
    For rowL = 5 To 16
        On Error Resume Next
        ws.Cells(rowL, 3).Value = ""
        On Error GoTo 0
    Next rowL
    On Error Resume Next
    ws.Cells(20, 2).Value = ""
    On Error GoTo 0
End Sub

Private Sub LimpiarFormObra(ws As Worksheet)
    Dim rowL As Long
    For rowL = 5 To 18
        On Error Resume Next
        ws.Cells(rowL, 3).Value = ""
        On Error GoTo 0
    Next rowL
    For rowL = 23 To 25
        On Error Resume Next
        ws.Cells(rowL, 4).Value = ""
        On Error GoTo 0
    Next rowL
    On Error Resume Next
    ws.Cells(37, 2).Value = ""
    On Error GoTo 0
End Sub

Sub Auto_Open()
    ThisWorkbook.Sheets("INICIO").Activate
End Sub
