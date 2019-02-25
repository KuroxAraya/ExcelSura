Attribute VB_Name = "moduloAmp"
Option Explicit
Sub duplicadorVac_Click()
    Application.ScreenUpdating = False
    ' Application.Visible = False
    
    ' LECTURA ARCHIVO FUENTE
    Dim libroDir As Office.FileDialog
    Dim libroFuente As Workbook
    Set libroDir = Application.FileDialog(msoFileDialogFilePicker)
        With libroDir
            .AllowMultiSelect = False
            .Title = "SELECCIONE BASE DE ARCHIVO FUENTE VACACIONES:"
            .Filters.Clear
            If .Show = True Then
                Set libroFuente = Application.Workbooks.Open(libroDir.SelectedItems(1))
            Else
                MsgBox "Proceso cancelado", vbOKOnly
                
                Exit Sub
            End If
        End With

    ' CREACIÓN ARCHIVO DESTINO
    Dim libroDestino As Workbook
    Set libroDestino = Workbooks.Add
        With libroDestino
        .Title = "Datos procesados de la Base de Vacaciones"
        .SaveAs Filename:="libroDestino.xlsx" ' , FileFormat:="xlOpenXMLWorkbookMacroEnabled"
        End With
    
    ' DECLARACIÓN PLANILLA ARCHIVO FUENTE n CORRESPONDE A LA POSICIÓN DE LA PLANILLA
    Dim planillaFuente As Worksheet
    Set planillaFuente = libroFuente.Worksheets(1)
    planillaFuente.Name = "hojaFuente"
    
    ' DECLARACIÓN Y CREACIÓN PLANILLA DESTINO
    Dim planillaDestino As Worksheet
    Set planillaDestino = libroDestino.Worksheets(1)
    planillaDestino.Name = "hojaDest"
    
    ' BORRA PLANILLAS EXTRAS EN ARCHIVO NUEVO
    Dim dSheet As Worksheet
    For Each dSheet In Worksheets
        Select Case dSheet.Name
        Case "Hoja1", "Sheet1"
            Application.DisplayAlerts = False
            dSheet.Delete
        End Select
    Next dSheet
    
    'COPIADOR ENCABEZADO
    'libroFuente.Sheets(5).Cells(7, "A").EntireRow.Copy Destination:=libroDestino.Sheets(1).Range("A" & Rows.Count).End(xlUp).Offset(0)
    libroFuente.Sheets(1).Range("A1:R1").Copy Destination:=libroDestino.Sheets(1).Range("A1")
    libroDestino.Sheets(1).Range("A1:R1").ClearFormats
    
    Dim filaFuenteUltima As Long
    filaFuenteUltima = planillaFuente.Cells(planillaFuente.Rows.Count, "B").End(xlUp).Row
    
    Dim filaIndiceFuente As Long
    
    Dim filaIndiceDestino As Long
    filaIndiceDestino = 1 ' SALTO ENCABEZADO
    
    Dim fechaInicio As Variant
    Dim fechaFin As Variant
    Dim fechaIndice As Date
    
        ' FOR LOOP DESDE FILA INICIAL HASTA ÚLTIMA
        For filaIndiceFuente = 8 To filaFuenteUltima ' EMPIEZA EN SEGUNDA LÍNEA DE ARCHIVO FUENTE
        fechaInicio = planillaFuente.Cells(filaIndiceFuente, "I").Value
        fechaFin = planillaFuente.Cells(filaIndiceFuente, "J").Value
        
        ' IF VALIDADOR, SI LAS FECHAS ESTÁN MAL INGRESADAS MUESTRA LA CELDA CON LOS DATOS ERRÓNEOS, POSTERIORMENTE VERIFICA QUE LA FECHA INICIAL SEA MENOR A LA FINAL
            If Not IsDate(fechaInicio) Or Not IsDate(fechaFin) Then
                MsgBox ("Fecha invalida en la fila " & filaIndiceFuente & " columna " & planillaFuente.Name & ".")
                Application.Goto planillaFuente.Cells(filaIndiceFuente, "I").Value
                Exit Sub
            ElseIf fechaInicio > fechaFin Then
                MsgBox ("fecha inicio sobrepasa la fecha final")
                Application.Goto planillaFuente.Cells(filaIndiceFuente, "J").ClearFormats.Value
                Exit Sub
            End If
        ' COPIA DE DATOS, EL LOOP RECORRE CADA LINEA Y LA DUPLICA SEGÚN LA DEFINICIÓN CORRESPONDIENTE A LA COPIA DEL DATO EN LA hojaFuente Y LUEGO EN LA hojaDestino
            For fechaIndice = fechaInicio To fechaFin
                'LA FECHA ÍNDICE SE DEFINE CON EL NÚMERO 1 PARA SALTARSE EL ENCABEZADO, LUEGO SUMA 1 CADA REPASO PARA LEER LA FILA SIGUIENTE
                filaIndiceDestino = filaIndiceDestino + 1
                'SI EXISTE UNA FÓRMULA O PROCESO ESPECIAL EN COLUMNAS, SE DEBEN COMENTAR PARA EVITAR EL SOLAPAMIENTO DE DATOS
                planillaDestino.Cells(filaIndiceDestino, "A").Value = planillaFuente.Cells(filaIndiceFuente, "A").Value
                planillaDestino.Cells(filaIndiceDestino, "B").Value = planillaFuente.Cells(filaIndiceFuente, "B").Value
                planillaDestino.Cells(filaIndiceDestino, "C").Value = planillaFuente.Cells(filaIndiceFuente, "C").Value
                planillaDestino.Cells(filaIndiceDestino, "D").Value = planillaFuente.Cells(filaIndiceFuente, "D").Value ' Rut usuario
                planillaDestino.Cells(filaIndiceDestino, "E").Value = planillaFuente.Cells(filaIndiceFuente, "E").Value
                planillaDestino.Cells(filaIndiceDestino, "F").Value = planillaFuente.Cells(filaIndiceFuente, "F").Value
                planillaDestino.Cells(filaIndiceDestino, "G").Value = planillaFuente.Cells(filaIndiceFuente, "G").Value
                planillaDestino.Cells(filaIndiceDestino, "H").Value = planillaFuente.Cells(filaIndiceFuente, "H").Value
                planillaDestino.Cells(filaIndiceDestino, "I").Value = fechaIndice
                fechaIndice = Application.Min(Application.EoMonth(fechaIndice, 0), fechaFin)
                planillaDestino.Cells(filaIndiceDestino, "J").Value = fechaIndice
                planillaDestino.Cells(filaIndiceDestino, "K").Value = planillaFuente.Cells(filaIndiceFuente, "K").Value
                planillaDestino.Cells(filaIndiceDestino, "L").Value = planillaFuente.Cells(filaIndiceFuente, "L").Value
                planillaDestino.Cells(filaIndiceDestino, "M").Value = planillaFuente.Cells(filaIndiceFuente, "M").Value
                planillaDestino.Cells(filaIndiceDestino, "N").Value = planillaFuente.Cells(filaIndiceFuente, "N").Value
                planillaDestino.Cells(filaIndiceDestino, "O").Value = planillaFuente.Cells(filaIndiceFuente, "O").Value
                planillaDestino.Cells(filaIndiceDestino, "P").Value = planillaFuente.Cells(filaIndiceFuente, "P").Value
                planillaDestino.Cells(filaIndiceDestino, "Q").Value = planillaFuente.Cells(filaIndiceFuente, "Q").Value
                planillaDestino.Cells(filaIndiceDestino, "R").Value = planillaFuente.Cells(filaIndiceFuente, "R").Value
            Next fechaIndice
        Next filaIndiceFuente
    
    ' LOS FORMATEOS SON ALFABÉTICOS, SE EJECUTAN DESDE A HASTA A+n, PARA INCLUIR UN NUEVO FORMATO ES NECESARIO DEJAR CADA LÍNEA DE FORMATO EN ORDEN ALFABÉTICO
    ' CREACIÓN COLUMNA PERPRO
    planillaDestino.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    planillaDestino.Range("A2:A" & filaIndiceDestino).Formula = "=CONCAT(YEAR(J2),IF(INT(MONTH(J2))<10,0,""""),MONTH(J2))" ' GENERADOR PERPRO

    ' CÁLCULO DE DÍAS TOTALES NO CORRIDOS, CONSIDERAR QUE LA FÓRMULA RESULTA NEGATIVO Y NO CUENTA EL DÍA DEL FIN DEL PERIÓDO
    planillaDestino.Columns("L:L").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    planillaDestino.Range("L2:L" & filaIndiceDestino).Formula = "=NETWORKDAYS(J2,K2,[libroFeriados.xlsm]hojaFeriados!$C$3:$C$38)"
    planillaDestino.Range("L2:L" & filaIndiceDestino).ClearFormats
    ' planillaDestino.Range("L2:L" & filaIndiceDestino).Formula = "=ABS(DAYS(I2,J2))+1"
    
    'planillaDestino.Columns("L").EntireColumn.Delete
    
    'TRANSFERIR CASOS ESPECIALES
    
    libroDestino.Sheets.Add after:=Sheets(Sheets.Count)
    
'    Dim libroFeriados As Workbook
'    Set libroFeriados = Workbooks.Open("libroFeriados.xlsm")
'    Dim viernesLista As Range
'    Set viernesLista = libroFeriados.Sheets(1).Range("$K$2:$K$53")
'    Dim iniLista As Range
'    Set iniLista = libroDestino.Sheets(1).Range("$J$2:$J")
    
    
    Dim r As Integer
    

    For r = filaIndiceDestino To 1 Step -1
        'If Cells(r, "L") = "0" Then
        If Cells(r, "J") = 0 Then
            'planillaDestino.Rows(r).EntireRow.Copy Destination:=libroDestino.Sheets(2).Rows(r)
            planillaDestino.Rows(r).EntireRow.Delete
        End If
    Next r
    
    libroFeriados.Close
    libroFuente.Close savechanges:=False
    libroDestino.Save
    libroDestino.Close
    
    ' ALERTA DE PROCESO FINALIZADO
    MsgBox "Ampliación finalizada, revisar archivo generado: libroDestino.xlsx", vbOKOnly
    
End Sub
