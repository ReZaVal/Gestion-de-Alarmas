    Option Explicit

    ' Constantes para colores
    Const COLOR_BLANCO As Long = 16777215  ' RGB(255, 255, 255)
    Const COLOR_OLIVA As Long = 32896      ' RGB(128, 128, 0)
    Const COLOR_CIAN As Long = 16776960    ' RGB(0, 255, 255)
    Const COLOR_VERDE As Long = 65280      ' RGB(0, 255, 0)
    Const COLOR_MAGENTA As Long = 16711935 ' RGB(255, 0, 255)
    Const COLOR_AMARILLO As Long = 65535   ' RGB(255, 255, 0)
    Const COLOR_ROJO As Long = 255         ' RGB(255, 0, 0)

    Sub FilterAndDeleteRows()
        ' Declarar variables
        Dim ws As Worksheet, wsSV As Worksheet, ws2 As Worksheet
        Dim ws3 As Worksheet, wsReporte As Worksheet, wsData As Worksheet
        Dim lastRow As Long, lastColumn As Long
        Dim lastRowSV As Long, lastColumnSV As Long
        Dim lastRow2 As Long, lastRow3 As Long, lastRowTodo As Long
        Dim rng As Range
        Dim r As Long, c As Long, i As Integer
        Dim colorFilters As Variant, criteriaFilters As Variant
        Dim cnt As Long, cnt3 As Long
        Dim dictCounts As Object

        On Error GoTo ErrHandler

        ' Paso 1: Preparar hojas
        Call PrepararHojas(ws, wsSV, ws2, ws3, wsReporte, wsData)

        ' Paso 2: Verificar datos en "Todo"
        If Not VerificarDatosTodo(ws) Then
            MsgBox "La hoja 'Todo' debe tener al menos 3 filas de datos.", vbExclamation
            Exit Sub
        End If

        ' Paso 3: Limpiar datos iniciales
        Call LimpiarDatosIniciales(ws, wsReporte, wsData)

        ' Paso 4: Filtrar y eliminar en "Todo"
        Call FiltrarEliminarTodo(ws, lastRow, lastColumn, rng)

        ' Paso 5: Procesar "Todo_SV"
        Call ProcesarTodoSV(ws, wsSV, lastRowSV, lastColumnSV)

        ' Paso 6: Procesar hoja "2"
        Call ProcesarHoja2(ws, wsSV, ws2, lastRow2, lastRowTodo, dictCounts)

        ' Paso 7: Procesar hoja "3"
        Call ProcesarHoja3(ws, ws2, ws3, lastRow3, lastRowTodo, dictCounts)

        ' Paso 8: Generar reporte
        Call GenerarReporte(ws2, ws3, wsReporte, lastRow2, lastRow3)

        ' Paso 9: Llenar datos en hoja "Data" por horas y colores
        Call LlenarDataPorHorasYColores(ws, wsData)

        ' Paso 10: Limpiar celdas innecesarias en "Reporte Diario"
        wsReporte.Range("I36:I45").ClearContents
        wsReporte.Range("I49:I52").ClearContents

        ' Paso 11: Restaurar configuración y confirmar
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "Proceso completado correctamente", vbInformation

        Exit Sub

    ErrHandler:
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "Error en el proceso: " & Err.Description, vbCritical
    End Sub

    Sub PrepararHojas(ByRef ws As Worksheet, ByRef wsSV As Worksheet, ByRef ws2 As Worksheet, ByRef ws3 As Worksheet, ByRef wsReporte As Worksheet, ByRef wsData As Worksheet)
        ' Desactivar actualizaciones
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        ' Asignar o crear hojas
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("Todo")
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = "Todo"
        End If
        Set ws2 = ThisWorkbook.Worksheets("2")
        If ws2 Is Nothing Then
            Set ws2 = ThisWorkbook.Worksheets.Add
            ws2.Name = "2"
        End If
        Set ws3 = ThisWorkbook.Worksheets("3")
        If ws3 Is Nothing Then
            Set ws3 = ThisWorkbook.Worksheets.Add
            ws3.Name = "3"
        End If
        Set wsReporte = ThisWorkbook.Worksheets("Reporte Diario")
        If wsReporte Is Nothing Then
            Set wsReporte = ThisWorkbook.Worksheets.Add
            wsReporte.Name = "Reporte Diario"
        End If
        Set wsData = ThisWorkbook.Worksheets("DataMensual")
        If wsData Is Nothing Then
            Set wsData = ThisWorkbook.Worksheets.Add
            wsData.Name = "DataMensual"
        End If
        Set wsSV = ThisWorkbook.Worksheets("Todo_SV")
        If wsSV Is Nothing Then
            Set wsSV = ThisWorkbook.Worksheets.Add(After:=ws)
            wsSV.Name = "Todo_SV"
            wsSV.Visible = xlSheetHidden
        End If
        On Error GoTo 0
    End Sub

    Function VerificarDatosTodo(ws As Worksheet) As Boolean
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        VerificarDatosTodo = (lastRow >= 3)
    End Function

    Sub LimpiarDatosIniciales(ws As Worksheet, wsReporte As Worksheet, wsData As Worksheet)
        ' Limpiar rango en "Reporte Diario"
        wsReporte.Range("D49:G53").ClearContents

        ' Borrar información en "DataMensual"
        Dim lastRow As Long
        lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
        If lastRow >= 2 Then wsData.Rows("2:" & lastRow).ClearContents

        ' Eliminar columnas E:H en "Todo"
        ws.Columns("E:H").Delete Shift:=xlToLeft
    End Sub

    Sub FiltrarEliminarTodo(ws As Worksheet, ByRef lastRow As Long, ByRef lastColumn As Long, ByRef rng As Range)
        ' Declarar variables locales
        Dim r As Long
        Dim colorPermitido As Boolean
        Dim deleteRange As Range
        
        ' Definir rango
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))

        ' Eliminar filas que NO tengan los colores permitidos (amarillo, verde, rojo, magenta)
        ' Iterar de atrás hacia adelante para evitar cambios de índice
        For r = lastRow To 2 Step -1
            colorPermitido = False
            
            ' Verificar si la celda en columna A tiene uno de los colores permitidos
            Select Case ws.Cells(r, 1).Font.Color
                Case COLOR_AMARILLO, COLOR_VERDE, COLOR_ROJO, COLOR_MAGENTA
                    colorPermitido = True
            End Select
            
            ' Si no tiene color permitido, marcar la fila para eliminarla
            If Not colorPermitido Then
                If deleteRange Is Nothing Then
                    Set deleteRange = ws.Rows(r)
                Else
                    Set deleteRange = Union(deleteRange, ws.Rows(r))
                End If
            End If
        Next r

        ' Eliminar filas en una sola operación para mejorar rendimiento
        If Not deleteRange Is Nothing Then
            deleteRange.Delete
        End If

        ' Actualizar rango
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
    End Sub

    Sub ProcesarTodoSV(ws As Worksheet, wsSV As Worksheet, ByRef lastRowSV As Long, ByRef lastColumnSV As Long)
        ' Copiar datos de "Todo" a "Todo_SV"
        wsSV.Cells.Clear
        Dim lastRow As Long, lastColumn As Long
        Dim r As Long
        Dim deleteRange As Range
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Copy wsSV.Cells(1, 1)

        ' Eliminar filas con COLOR_VERDE en la columna D (Description) de "Todo_SV"
        lastRowSV = wsSV.Cells(wsSV.Rows.Count, "A").End(xlUp).Row
        lastColumnSV = wsSV.Cells(1, wsSV.Columns.Count).End(xlToLeft).Column
        ' r ya está declarada como variable global en el módulo o subrutina principal
        For r = lastRowSV To 2 Step -1
            ' Eliminar filas con color verde en columna D (Description)
            ' o cuya columna C (Tag) empiece con "SLDS-LD" (alarmas LDS)
            If wsSV.Cells(r, 4).Font.Color = COLOR_VERDE Or Left(wsSV.Cells(r, 3).Value, 7) = "SLDS-LD" Then
                If deleteRange Is Nothing Then
                    Set deleteRange = wsSV.Rows(r)
                Else
                    Set deleteRange = Union(deleteRange, wsSV.Rows(r))
                End If
            End If
        Next r

        If Not deleteRange Is Nothing Then
            deleteRange.Delete
        End If
    End Sub

    Sub ProcesarHoja2(ws As Worksheet, wsSV As Worksheet, ws2 As Worksheet, ByRef lastRow2 As Long, lastRowTodo As Long, dictCounts As Object)
        ' Copiar columnas B:D de "Todo_SV" a "2" (en A:C)
        ws2.Cells.Clear
        
        ' Identificar la última fila con datos (letras) en "Todo_SV"
        ' Buscar en todas las columnas para asegurar que se captura la última fila con contenido
        Dim lastRowSV As Long
        Dim lastColSV As Long
        Dim r As Long
        lastRowSV = wsSV.Cells(wsSV.Rows.Count, "A").End(xlUp).Row  ' Última fila con datos en col A
        lastColSV = wsSV.Cells(1, wsSV.Columns.Count).End(xlToLeft).Column  ' Última columna con datos
        
        ' Asegurar que se copia desde la fila de encabezado hasta la última fila de datos
        If lastRowSV >= 1 Then
            wsSV.Range("B1:D" & lastRowSV).Copy ws2.Range("A1")
        End If

        ' Eliminar duplicados en "2" ANTES de procesar (considerando el color)
        lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        If lastRow2 > 1 Then
            ' Crear columna auxiliar E con el color de la fuente para diferenciar alarmas
            For r = 2 To lastRow2
                ws2.Cells(r, 5).Value = ws2.Cells(r, 1).Font.Color  ' Capturar color en columna E
            Next r
            
            ' Eliminar duplicados incluyendo la columna de color (A:E)
            ws2.Range("A1:E" & lastRow2).RemoveDuplicates Columns:=Array(1, 2, 3, 5), Header:=xlYes
            
            ' Recalcular lastRow2 después de deduplicación
            lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
            
            ' Eliminar la columna auxiliar E
            ws2.Columns(5).Delete
        End If

        ' Contar ocurrencias usando Dictionary para rendimiento
        Set dictCounts = CreateObject("Scripting.Dictionary")
        lastRowTodo = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        Dim cell As Range
        For Each cell In ws.Range("D2:D" & lastRowTodo)
            If cell.Font.Color <> COLOR_VERDE Then
                If Not dictCounts.Exists(cell.Value) Then
                    dictCounts.Add cell.Value, 1
                Else
                    dictCounts(cell.Value) = dictCounts(cell.Value) + 1
                End If
            End If
        Next cell

        ' Asignar conteos a "2"
        lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        For r = 2 To lastRow2
            If dictCounts.Exists(ws2.Cells(r, "C").Value) Then
                ws2.Cells(r, "D").Value = dictCounts(ws2.Cells(r, "C").Value)
            Else
                ws2.Cells(r, "D").Value = 0
            End If
        Next r

        ' Ordenar "2" por col D descendente
        If lastRow2 > 1 Then
            With ws2.Sort
                .SortFields.Clear
                .SortFields.Add Key:=ws2.Range("D2:D" & lastRow2), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange ws2.Range("A1:D" & lastRow2)
                .Header = xlYes
                .Apply
            End With
        End If
    End Sub

    Sub ProcesarHoja3(ws As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ByRef lastRow3 As Long, lastRowTodo As Long, dictCounts As Object)
        ' PASO 1: Copiar TODA la hoja "2" a hoja "3" sin filtrar aún
        Dim lastRow2 As Long
        Dim r As Long
        Dim deleteRange As Range
        lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        ws3.Cells.Clear
        ws2.Range("A1:D" & lastRow2).Copy ws3.Range("A1")
        lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

        ' PASO 2: Filtrar para mantener SOLO las alarmas magentas en columna C (Description) de "3"
        ' Esto asegura que se copien TODAS las alarmas magentas que tienen color magenta en la columna Description
        ' r ya está declarada como variable global en el módulo o subrutina principal
        For r = lastRow3 To 2 Step -1
            If ws3.Cells(r, 3).Font.Color <> COLOR_MAGENTA Then
                If deleteRange Is Nothing Then
                    Set deleteRange = ws3.Rows(r)
                Else
                    Set deleteRange = Union(deleteRange, ws3.Rows(r))
                End If
            End If
        Next r

        If Not deleteRange Is Nothing Then
            deleteRange.Delete
        End If
        lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

        ' PASO 2B: Eliminar duplicados en "3" ANTES de procesar (considerando el color)
        If lastRow3 > 1 Then
            ' Crear columna auxiliar E con el color de la fuente en columna C
            For r = 2 To lastRow3
                ws3.Cells(r, 5).Value = ws3.Cells(r, 3).Font.Color
            Next r
            ' Eliminar duplicados incluyendo el color (A:C + E)
            ws3.Range("A1:E" & lastRow3).RemoveDuplicates Columns:=Array(1, 2, 3, 5), Header:=xlYes
            ' Eliminar la columna auxiliar E
            ws3.Columns(5).Delete
        End If
        lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

        ' PASO 3: Recalcular conteos de magentas desde ws (Todo) para asegurar precisión
        Set dictCounts = CreateObject("Scripting.Dictionary")
        lastRowTodo = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        Dim cell As Range
        For Each cell In ws.Range("D2:D" & lastRowTodo)
            If cell.Font.Color = COLOR_MAGENTA Then
                If Not dictCounts.Exists(cell.Value) Then
                    dictCounts.Add cell.Value, 1
                Else
                    dictCounts(cell.Value) = dictCounts(cell.Value) + 1
                End If
            End If
        Next cell

        ' PASO 4: Asignar conteos precisos a "3" basados en el contenido de columna C
        For r = 2 To lastRow3
            If dictCounts.Exists(ws3.Cells(r, "C").Value) Then
                ws3.Cells(r, "D").Value = dictCounts(ws3.Cells(r, "C").Value)
            Else
                ws3.Cells(r, "D").Value = 0
            End If
        Next r

        ' PASO 5: Ordenar "3" por col D descendente
        If lastRow3 > 1 Then
            With ws3.Sort
                .SortFields.Clear
                .SortFields.Add Key:=ws3.Range("D2:D" & lastRow3), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange ws3.Range("A1:D" & lastRow3)
                .Header = xlYes
                .Apply
            End With
        End If
    End Sub

    Sub GenerarReporte(ws2 As Worksheet, ws3 As Worksheet, wsReporte As Worksheet, lastRow2 As Long, lastRow3 As Long)
        ' Copiar rangos a "Reporte Diario" manteniendo bordes, alineación y formatos numéricos
        Dim srcTop As Range
        Dim dstTop As Range
        Dim nfTop As Variant, haTop As Variant, vaTop As Variant
        Dim r As Long, c As Long
        Set srcTop = ws2.Range("A2:D11")
        Set dstTop = wsReporte.Range("D36:G45")
        nfTop = dstTop.NumberFormat
        haTop = dstTop.HorizontalAlignment
        vaTop = dstTop.VerticalAlignment
        dstTop.Value = srcTop.Value
        dstTop.NumberFormat = nfTop
        dstTop.HorizontalAlignment = haTop
        dstTop.VerticalAlignment = vaTop
        For r = 1 To srcTop.Rows.Count
            For c = 1 To srcTop.Columns.Count
                dstTop.Cells(r, c).Font.Color = srcTop.Cells(r, c).Font.Color
            Next c
        Next r

        ' Copiar color de texto de D36:D45 a C36:C45 en "Reporte Diario"
        For r = 0 To 9
            wsReporte.Range("C36").Offset(r, 0).Font.Color = wsReporte.Range("D36").Offset(r, 0).Font.Color
        Next r

        ' Copiar alarmas magentas (A2:D6 de hoja 3) a D49:G53 de Reporte Diario
        Dim srcMagenta As Range
        Dim dstMagenta As Range
        Set srcMagenta = ws3.Range("A2:D6")
        Set dstMagenta = wsReporte.Range("D49:G53")
        Dim nfMagenta As Variant, haMagenta As Variant, vaMagenta As Variant
        nfMagenta = dstMagenta.NumberFormat
        haMagenta = dstMagenta.HorizontalAlignment
        vaMagenta = dstMagenta.VerticalAlignment
        dstMagenta.ClearContents
        dstMagenta.Value = srcMagenta.Value
        dstMagenta.NumberFormat = nfMagenta
        dstMagenta.HorizontalAlignment = haMagenta
        dstMagenta.VerticalAlignment = vaMagenta
        For r = 1 To srcMagenta.Rows.Count
            For c = 1 To srcMagenta.Columns.Count
                dstMagenta.Cells(r, c).Font.Color = srcMagenta.Cells(r, c).Font.Color
            Next c
        Next r
        
        ' Mantener el formato existente en "Reporte Diario" (no modificar formatos aquí)
    End Sub

    Sub LlenarDataPorHorasYColores(wsTodo As Worksheet, wsData As Worksheet)
        ' Subrutina para llenar la hoja "Data" con conteo de alarmas por hora y color
        ' Estructura: Horas (0-23), Alarmas Verde (B), Amarillas (C), Rojas (D), Magentas (E)
        ' OPTIMIZADO: Utiliza un único bucle sobre los datos (O(n) en lugar de O(n*24))
        
        Dim contadores(0 To 23, 1 To 4) As Long  ' [hora, colorIndex] colorIndex: 1=verde, 2=amarillo, 3=rojo, 4=magenta
        Dim r As Long, ultimaFila As Long, horaActual As Integer
        Dim colorCelda As Long
        Dim fila As Long
        Dim cellValue As Variant
        
        On Error GoTo ErrHandler
        
        ' Limpiar datos anteriores en "Data" (mantener solo encabezados)
        ultimaFila = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
        If ultimaFila > 1 Then
            wsData.Range("A2:E" & ultimaFila).ClearContents
        End If
        
        ' Agregar encabezados si no existen
        If wsData.Cells(1, 1).Value = "" Then
            wsData.Cells(1, 1).Value = "Horas"
            wsData.Cells(1, 2).Value = "Alarmas Verde"
            wsData.Cells(1, 3).Value = "Baja"
            wsData.Cells(1, 4).Value = "Alta"
            wsData.Cells(1, 5).Value = "Crítica"
            
            ' Formato de encabezados
            With wsData.Range("A1:E1")
                .Font.Bold = True
                .Interior.Color = RGB(0, 0, 0)
                .Font.Color = RGB(255, 255, 255)
                .HorizontalAlignment = xlCenter
            End With
        End If
        
        ' Obtener la última fila de datos en "Todo"
        ultimaFila = wsTodo.Cells(wsTodo.Rows.Count, "A").End(xlUp).Row
        
        ' OPTIMIZACIÓN: Un único bucle sobre todos los datos
        ' Extraer hora y color, acumular en array
        For r = 2 To ultimaFila
            cellValue = wsTodo.Cells(r, 1).Value
            
            ' Validar que el valor sea una fecha válida antes de extraer la hora
            If IsDate(cellValue) Or IsNumeric(cellValue) Then
                On Error Resume Next
                horaActual = Hour(cellValue)
                On Error GoTo ErrHandler
                
                ' Si la hora está en rango válido (0-23)
                If horaActual >= 0 And horaActual <= 23 Then
                    colorCelda = wsTodo.Cells(r, 1).Font.Color
                    
                    ' Incrementar contador según color
                    Select Case colorCelda
                        Case COLOR_VERDE
                            contadores(horaActual, 1) = contadores(horaActual, 1) + 1
                        Case COLOR_AMARILLO
                            contadores(horaActual, 2) = contadores(horaActual, 2) + 1
                        Case COLOR_ROJO
                            contadores(horaActual, 3) = contadores(horaActual, 3) + 1
                        Case COLOR_MAGENTA
                            contadores(horaActual, 4) = contadores(horaActual, 4) + 1
                    End Select
                End If
            End If
        Next r
        
        ' Llenar datos en "Data" desde el array (sin iteraciones adicionales)
        For horaActual = 0 To 23
            fila = horaActual + 2
            wsData.Cells(fila, 1).Value = horaActual
            wsData.Cells(fila, 2).Value = contadores(horaActual, 1)
            wsData.Cells(fila, 3).Value = contadores(horaActual, 2)
            wsData.Cells(fila, 4).Value = contadores(horaActual, 3)
            wsData.Cells(fila, 5).Value = contadores(horaActual, 4)
        Next horaActual
        
        ' Aplicar colores a las columnas según corresponda
        With wsData.Range("B2:B25")
            .Font.Color = COLOR_VERDE
            .Font.Bold = True
        End With
        
        With wsData.Range("C2:C25")
            .Font.Color = COLOR_AMARILLO
            .Font.Bold = True
        End With
        
        With wsData.Range("D2:D25")
            .Font.Color = COLOR_ROJO
            .Font.Bold = True
        End With
        
        With wsData.Range("E2:E25")
            .Font.Color = COLOR_MAGENTA
            .Font.Bold = True
        End With
        
        Exit Sub
        
    ErrHandler:
        MsgBox "Error al llenar datos en Data: " & Err.Description, vbExclamation
    End Sub

