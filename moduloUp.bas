Attribute VB_Name = "moduloUp"
Option Explicit
Sub popularSubG()
    ' POPULAR BASE SUBG
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' BUSCAR ARCHIVO FUENTE Y DEFINICIÓN
    Dim libroFuenteDir As Office.FileDialog
    Dim libroFuenteSub As Workbook
    Set libroFuenteDir = Application.FileDialog(msoFileDialogFilePicker)
        With libroFuenteDir
            .AllowMultiSelect = False
            .Title = "SELECCIONE BASE FUENTE SUBG:"
            .Filters.Clear
            If .Show = True Then
                Set libroFuenteSub = Application.Workbooks.Open(libroFuenteDir.SelectedItems(1))
            Else
                MsgBox "Proceso cancelado, no se hicieron cambios.", vbOKOnly
                Exit Sub
            End If
        End With
        
    ' BUSCAR ARCHIVO DESTINO Y DEFINICIÓN
    Dim libroDestinoDir As Office.FileDialog
    Dim libroDestino As Workbook
    Set libroDestinoDir = Application.FileDialog(msoFileDialogFilePicker)
        With libroDestinoDir
            .AllowMultiSelect = False
            .Title = "SELECCIONE DESTINO:"
            .Filters.Clear
            If .Show = True Then
                Set libroDestino = Application.Workbooks.Open(libroDestinoDir.SelectedItems(1))
            Else
                MsgBox "Proceso cancelado, no se hicieron cambios.", vbOKOnly
                Exit Sub
            End If
        End With
    
    ' DECLARACIÓN PLANILLA ARCHIVO FUENTE n CORRESPONDE A LA POSICIÓN DE LA PLANILLA
    Dim planillaFuente As Worksheet
    Set planillaFuente = libroFuenteSub.Worksheets(1)
    
    ' DECLARACIÓN Y CREACIÓN PLANILLA DESTINO
    Dim planillaDestino As Worksheet
    Set planillaDestino = libroDestino.Worksheets(1)
    
    ' DEFINICIÓN DE ÚLTIMAS LÍNEAS EN AMBAS PLANILLAS
    Dim filaFuenteUltima As Long
    filaFuenteUltima = planillaFuente.Cells(planillaFuente.Rows.Count, "B").End(xlUp).Row
    Dim filaDestinoUltima As Long
    filaDestinoUltima = planillaFuente.Cells(planillaDestino.Rows.Count, "A").End(xlUp).Row
    
    ' DEFINICIÓN DE ÚLTIMAS LÍNEAS EN AMBAS PLANILLAS
    Dim filaIndiceFuente As Long
    filaIndiceFuente = 1
    
    Dim filaIndiceDestino As Long
    filaIndiceDestino = 1
    
    ' CREAR EN PLANILLAFUENTE COLUMNA QUE CONCATENE PERPRO Y IDHR
    planillaFuente.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    planillaFuente.Range("A2:A" & filaFuenteUltima).Formula = "=CONCAT($B2,$C2)"
    
    ' CREAR EN PLANILLA DESTINO COLUMNA QUE CONCATENE PERPRO Y IDHR
    planillaDestino.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    planillaDestino.Range("A2:A" & filaDestinoUltima).Formula = "=CONCAT($B2,$D2)"

    Dim dict As New Scripting.Dictionary
    
    ' ITERACIÓN PARA ALMACENAR CADA ID GENERADA
    For filaIndiceFuente = 2 To filaFuenteUltima
        dict.Add CStr(planillaFuente.Range("A" & filaIndiceFuente).Value), filaIndiceFuente ' PUNTERO INDICE DE FILA
    Next filaIndiceFuente
    
    MsgBox "DICCIONARIO COMPLETO", vbOKOnly
    
    For filaIndiceDestino = 2 To filaDestinoUltima
        ' SI UNA DE LAS ID'S GENERADAS COINCIDE CON LAS DEL ARCHIVO DESTINO, LA FILA DE FUENTE SE TRANSFIERE COMPLETA
        If dict.Exists(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)) Then
            'CELDAS SE TRANSFIEREN DESDE EL LIBRO FUENTE AL DESTINO EN LA PLANILLA CORRESPONDIENTE
            planillaDestino.Cells(filaIndiceDestino, "O").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "D").Value
            planillaDestino.Cells(filaIndiceDestino, "P").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "E").Value
            planillaDestino.Cells(filaIndiceDestino, "Q").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "F").Value
            planillaDestino.Cells(filaIndiceDestino, "R").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "G").Value
            planillaDestino.Cells(filaIndiceDestino, "S").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "H").Value
            planillaDestino.Cells(filaIndiceDestino, "T").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "I").Value
            planillaDestino.Cells(filaIndiceDestino, "U").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "J").Value
            planillaDestino.Cells(filaIndiceDestino, "V").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "K").Value
            planillaDestino.Cells(filaIndiceDestino, "W").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "L").Value
            planillaDestino.Cells(filaIndiceDestino, "X").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "M").Value
            planillaDestino.Cells(filaIndiceDestino, "Y").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "N").Value
            planillaDestino.Cells(filaIndiceDestino, "Z").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "O").Value
            planillaDestino.Cells(filaIndiceDestino, "AA").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "P").Value
            planillaDestino.Cells(filaIndiceDestino, "AB").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "Q").Value
            planillaDestino.Cells(filaIndiceDestino, "AC").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "R").Value
            planillaDestino.Cells(filaIndiceDestino, "AD").Value = planillaFuente.Cells(dict(CStr(planillaDestino.Range("A" & filaIndiceDestino).Value)), "S").Value
        End If
    Next filaIndiceDestino
    Set dict = Nothing
    
    ' COPIADOR DE ENCABEZADO
    libroFuenteSub.Sheets(1).Range("D1:S1").Copy Destination:=libroDestino.Sheets(1).Range("O1")
    libroDestino.Sheets(1).Range("A1:AE1").ClearFormats
    
    ' ELIMINACIÓN DE COLUMNAS EN BASE A REQUERIMIENTO
    planillaDestino.Columns("A").EntireColumn.Delete
    
    Application.StatusBar = ""
    
    ' PROPIEDADES DEL DOCUMENTO
    libroFuenteSub.Close savechanges:=False
    libroDestino.Close savechanges:=True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Proceso finalizado."
End Sub
