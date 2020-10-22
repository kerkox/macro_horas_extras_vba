Sub procesar_horas_extras()

    Dim range_cells As String
    Dim texto_horas_extras As String
    Dim texto_cedula  As String
    Dim hojaActual As Worksheet
    range_cells = "A1:T30"
    Dim range_destino As Range
    Dim row_index_destino As Integer
    Dim col_index_destino_cedula As String
    Dim col_index_destino_horas As String
    Dim nombre_hoja_destino As String
    Dim valido As Boolean
    nombre_hoja_destino = "DATOS_HORAS_EXTRAS"
    
    'Application.ScreenUpdating = False
    
    texto_horas_extras = "TOTAL POR TIPO DE HORA EXTRA"
    texto_cedula = "CEDULA"
    row_index_destino = 2
    col_index_destino_cedula = "A"
    col_index_destino_horas = "B"
    
    For Each hojaActual In Worksheets
        If Not hojaActual Is Nothing And hojaActual.Name <> "DATOS_HORAS_EXTRAS" Then
            valido = encontrar_horas_extras(hojaActual, nombre_hoja_destino, range_cells, texto_horas_extras, row_index_destino, col_index_destino_horas)
            If valido Then
                Call encontrar_y_copiar_cedula(hojaActual, nombre_hoja_destino, range_cells, texto_cedula, row_index_destino, col_index_destino_cedula)
                Call generar_formula_texto_para_planos(nombre_hoja_destino, row_index_destino)
                row_index_destino = row_index_destino + 1
            End If
        End If
    Next
    Worksheets(nombre_hoja_destino).Activate
    MsgBox ("Todo ha sido procesado")
    Exit Sub
End Sub

Function encontrar_horas_extras(hoja As Worksheet, _
                                nombre_hoja_destino As String, _
                                rango_busqueda As String, _
                                texto_horas_extras As String, _
                                row_index_destino As Integer, _
                                col_index_destino As String) As Boolean
    Dim range_horas_extras As Range
    Dim range_horas_extras_valor As Range
    Dim col_range As Integer
    Dim row_range As Integer
    
    hoja.Activate
    hoja.Select
    
    'Busqueda
    Set range_horas_extras = Range(rango_busqueda).Find(texto_horas_extras)
    'Se valida si se encontro el valor de cedula
    If range_horas_extras Is Nothing Then
        'MsgBox ("No se encontro la cedula")
        encontrar_horas_extras = False
    Else
        col_range = range_horas_extras.Column
        row_range = range_horas_extras.Row
        Set range_horas_extras_valor = Range(Cells(row_range, col_range + 3), Cells(row_range, col_range + 6))
        range_horas_extras_valor.Copy
        Worksheets(nombre_hoja_destino).Activate
        'Worksheets("DATOS_HORAS_EXTRAS").Select
        Range(col_index_destino & row_index_destino).PasteSpecial Paste:=xlPasteValues
        encontrar_horas_extras = True
        'Worksheets("DATOS_HORAS_EXTRAS").Select
        'Range(
        'MsgBox ("El valor encontrado de la celda es:")
    End If
End Function

Function encontrar_y_copiar_cedula(hoja As Worksheet, _
                                    nombre_hoja_destino As String, _
                                    rango_busqueda As String, _
                                    texto_cedula As String, _
                                    row_index_destino As Integer, _
                                    col_index_destino As String) As Boolean
    Dim range_cedula As Range
    Dim range_cedula_valor As Range
    Dim col_range As Integer
    Dim row_range As Integer
    
    hoja.Activate
    hoja.Select
    
    'Busqueda de cedula
    Set range_cedula = Range(rango_busqueda).Find(texto_cedula)
    'Se valida si se encontro el valor de cedula
    If range_cedula Is Nothing Then
        'MsgBox ("No se encontro la cedula")
        encontrar_y_copiar_cedula = False
    Else
        col_range = range_cedula.Column
        row_range = range_cedula.Row
        Dim temp As String
        temp = Cells(row_range + 2, col_range).Text
        temp = Replace(temp, ",", "")
        temp = Replace(temp, ".", "")
        'Set range_cedula_valor = Range(Cells(row_range + 1, col_range), Cells(row_range + 1, col_range))
        Worksheets(nombre_hoja_destino).Activate
        Worksheets(nombre_hoja_destino).Select
        Range(col_index_destino & row_index_destino).Value = temp
        Range(col_index_destino & row_index_destino).NumberFormat = "General"
        'MsgBox ("El valor encontrado de la celda es:")
        encontrar_y_copiar_cedula = True
    End If
End Function

Sub generar_formula_texto_para_planos(nombre_hoja_destino As String, row_index As Integer)
    Dim columnas_formulas As Variant
    Dim index_col As Integer
    Dim columna As String
    
    Worksheets(nombre_hoja_destino).Activate
    Worksheets(nombre_hoja_destino).Select
    
    columnas_formulas = Array("G", "H", "I", "J")
    formulas_lista = Array("=+CONCATENATE(RC[-6],"";"",RC[-5])", _
                            "=+CONCATENATE(RC[-7],"";"",RC[-5])", _
                            "=+CONCATENATE(RC[-8],"";"",RC[-5])", _
                            "=+CONCATENATE(RC[-9],"";"",RC[-5])")
    
    index_col = -5
    For col = 0 To 3
        columna = columnas_formulas(col) & row_index
        Range(columna).Delete
        Range(columna).Select
        ActiveCell.FormulaR1C1 = formulas_lista(col)
        Selection.Columns.AutoFit
    Next
End Sub


Sub Generar_archivos_planos()
    
    Dim path As String
    path = Range("M1").Value
    Call generar_archivo_columna("G1", path)
    Call generar_archivo_columna("H1", path)
    Call generar_archivo_columna("I1", path)
    Call generar_archivo_columna("J1", path)
    
  
    
End Sub

Sub generar_archivo_columna(columna As String, path As String)
    nombre_archivo = Range(columna).Value
    ruta = path & "\" & nombre_archivo & ".txt"
    
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set a = fs.CreateTextFile(ruta, True)
    
    
    ultima_fila = Cells(Rows.Count, 1).End(xlUp).Row
    Dim index_columna As Integer
    index_columna = Range(columna).Column
    Open ruta For Output As #1
    
    For i = 2 To ultima_fila
        If validar_horas_mayor_cero(Cells(i, index_columna).Value) Then
            Print #1, Cells(i, index_columna).Value
        End If
    Next
    
    Close #1
End Sub

Function validar_horas_mayor_cero(valor As String) As Boolean
    Dim arrayValores As Variant
    
    arrayValores = Split(valor, ";")
    
    validar_horas_mayor_cero = arrayValores(1) <> "0"
    
End Function


