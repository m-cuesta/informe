Dim libro As Workbook
Dim fila2tpv
Dim final3
Dim penultimafila
Dim espacionabsolutorelativo
Dim ultimafila
Dim fila
Dim columna
Dim fila1rechazo
Dim fila2rechazo
Dim columna2
Dim fila2
Dim final2
Dim fila1tpv
Dim columna1_tpv_anterior
Dim filainicialalertarechazo
Dim filainiciocomerciosvalidacion
Dim filainicialalertavalidacion
Dim Altura_grafico_tpv
Dim contador_grupos
Dim contador_graficos
Dim cantidadhojas
Dim salir_por_fecha
Dim fecha_guia
Dim filainicialcrecimientotop10
Dim filainicialtablas_tpv_grande
Dim filainicialtablas_tpv_negativo
Dim entre_semana
Dim fin_Semana
Dim multiplicador_dias_no_laboral
Dim multiplicador_dias_laboral
Dim dias_mes_corridos
Dim filainicialtablas_comercios_activos
Dim fecha_mes_anterior













Sub todo()
Application.DisplayAlerts = False
Dim Cache As PivotCache

Dim Cache2 As PivotCache
Dim dinamica As PivotTable
Dim dinamica2 As PivotTable
Dim dinamica3 As PivotTable
Dim dinamica4 As PivotTable

Dim hoja1() As String
Dim hoja2() As String
Dim perso(1 To 2)
ReDim hoja1(1)
ReDim hoja2(1)



Call normalizacion_hojas_sub

  If salir_por_fecha = "si" Then
            'Salir si las fechas entre rechazo y tpv no coinciden
        Exit Sub
  End If
Call nombres_tpv("Hoja2")
Call nombres_tpv("Hoja5")


libro.Sheets("Hoja3").Select

Call tablas_dinamicas

Call macros_insertadas
  
 
  
  Cells(fila2tpv + 7, 2).Select
  ActiveWindow.FreezePanes = True
  
   ActiveSheet.Name = "Informe"
   
   
   'insertar procedimiento para crear tablas dinámicas de tpv anterior y actual(falta)_"tablas_dinamicas_tpv_comercios".
     
     Call tipo_dia_semana("hoja2")
     Call tipo_dia_semana("hoja5")
     
     libro.Names("tpv_anterior").Delete
     libro.Names("tpv_actual").Delete
     
     
     Call tabla_dinamica_comercios_tpv("hoja2")
     Call tabla_dinamica_comercios_tpv("hoja5")
     
     libro.Sheets("hoja2_td").Range("a1").CurrentRegion.Name = "tpv_actual"
     libro.Sheets("hoja5_td").Range("a1").CurrentRegion.Name = "tpv_anterior"
     
 
  Sheets("hoja5").Select
 
  Sheets("hoja5").Range(Cells(2, 1), Cells(Sheets("hoja5").Range("a1").CurrentRegion.Rows.Count, Sheets("hoja5").Range("a1").CurrentRegion.Columns.Count)).Copy Destination:=Sheets("Hoja2").Cells(Sheets("Hoja2").Range("a1").End(xlDown).Row + 1, 1)
 
 
  
  Sheets("hoja1").Delete
  Sheets("hoja4").Delete
  Sheets("hoja5").Delete
  
  Sheets("Hoja2").Select
  Sheets("Hoja2").Columns("G:G").Cut
  Sheets("Hoja2").Columns("B:B").Insert Shift:=xlToRight
  Sheets("Hoja2").Range(Columns(3), Columns(5)).Delete Shift:=xlToLeft
  Sheets("Hoja2").Range("a1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
  Sheets("Hoja2").Columns(buscar("entre_semana", "Hoja2")).Delete
  
  Sheets("hoja2").Name = "hoja1"
 
  
  'el procedimiento valores_tpv debería actualizar el tpv y el pronostico de tpv por comercio
  
  Call valores_tpv("hoja1")
  
  
  Sheets("hoja2_td").Delete
  Sheets("hoja5_td").Delete
  
  libro.Sheets.Add before:=Sheets(1)
  
  
  

   
 
libro.SaveAs "D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\informe_" & Replace(Date, "/", "-") & ".xlsm", FileFormat:=52
'libro.SaveAs "C:\Users\user\Documents\Javier\Pagos Online\informe jose\Abril 27\informe_" & Replace(Date, "/", "-") & ".xlsm", FileFormat:=52
'13.5
Workbooks("informe_" & Replace(Date, "/", "-") & ".xlsm").Activate
numeropais = Sheets("Hoja3").Cells(1, 3).End(xlDown).Row - 1
tipo = 2


contador_graficos = 1
cantidadhojas = 0

Sheets.Add before:=Worksheets(1)
Sheets(1).Name = "Resumen Graficos"
Sheets(1).Cells.Interior.Color = RGB(255, 255, 255)
Sheets(1).Cells.RowHeight = 20
For i = 1 To numeropais

    For j = 1 To tipo
         
         Sheets("Informe").Copy before:=Sheets("informe")
         'Sheets(1).Select
         Sheets(Sheets.Count - 1).Name = Sheets("hoja3").Cells(i + 1, 3) & "-" & Sheets("hoja3").Cells(j, 5)
         Sheets(Sheets.Count - 1).Cells(fila2tpv + 2, 5) = Sheets("hoja3").Cells(i + 1, 3)
         Sheets(Sheets.Count - 1).Cells(fila2tpv + 3, 5) = Sheets("hoja3").Cells(j, 5)
         'Sheets(Sheets.Count - 1).Cells.Locked = False
         'Sheets(Sheets.Count - 1).Range(Rows(fila2tpv + 2), Rows(fila2tpv + 3)).Locked = True
         'Sheets(Sheets.Count - 1).Range(Rows(fila2tpv + 1), Rows(300)).FormulaHidden = True
         'Sheets(Sheets.Count - 1).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingRows:=True, AllowUsingPivotTables:=True, Password:="1234"
         'Sheets(Sheets.Count - 1).EnableSelection = xlUnlockedCells
         
         
         tpv_acumulado_actual = Cells(final2 - 1, fila2rechazo - fila1rechazo + 3)
         tpv_acumulado_mes_anterior = Cells(final2 - 1, fila2rechazo - fila1rechazo + 4)
         
         
         'comportamiento tpv comercio
         
         
            For ma = 1 To 19
            
                    Sheets(Sheets.Count - 1).PivotTables("dinamica2_anterior").PivotFields("usuario_id").ClearAllFilters
                    On Error GoTo errordinamica:
                    Sheets(Sheets.Count - 1).PivotTables("dinamica2_anterior").PivotFields("usuario_id").CurrentPage = Cells(final3 + ma + 2, 2).Value
                
                    tpv_comercio_anterior = Cells(final2 - 1, fila2rechazo - fila1rechazo + 4)
                    tpv_comercio_actual = Cells(final3 + ma + 2, 3)
                    If tpv_comercio_actual / (fila2rechazo - fila1rechazo) < tpv_comercio_anterior / dias_mes_anterior Then '(fila2rechazo - fila1rechazo) corresponde a los dias transcurridos del mes
                        
                        Cells(final3 + ma + 2, 2).Interior.Color = RGB(255, 0, 0)
            
            
                    End If
                    Sheets(Sheets.Count - 1).PivotTables("dinamica2_anterior").PivotFields("usuario_id").ClearAllFilters
                           
volver:
                    On Error GoTo 0
            Next

         'fin comportamiento tpv comercio
         
         
         If (tpv_acumulado_mes_anterior) > (tpv_acumulado_actual) Then
         
            
            Sheets(Sheets.Count - 1).Tab.Color = RGB(255, 0, 0)
         
         End If
                          
                numero_hoja = Sheets.Count - 1
         
            Call grafico_tpv(numero_hoja)
             
             Sheets(Sheets.Count - 1).Cells(filainicialcrecimientotop10 + 1, 2).Select
             
            Call tablas_areas(Sheets("hoja3").Cells(i + 1, 3), Sheets("hoja3").Cells(j, 5))
            
            Call grafico_areas
            
            Sheets(Sheets.Count - 1).Cells(filainicialtablas_tpv_grande + 1, 2).Select
            
            Call tablas_tpv_grande(Sheets("hoja3").Cells(i + 1, 3), Sheets("hoja3").Cells(j, 5))
            
            Call grafico_tpv_grande

            Sheets(Sheets.Count - 1).Cells(filainicialtablas_tpv_negativo + 1, 2).Select
            
            Call tablas_tpv_negativo(Sheets("hoja3").Cells(i + 1, 3), Sheets("hoja3").Cells(j, 5))
            
            Call grafico_tpv_negativo
            
            
            Call tabla_comercios_activos(Sheets("hoja3").Cells(i + 1, 3), Sheets("hoja3").Cells(j, 5))
            
            Call grafico_comercios_activos

        
         
     'copiar y pegar  datos como valores
    Intersect(Range(Rows(fila2tpv + 4), Rows(filainicialalertarechazo - 4)), Range(Columns(1), Columns(fila2rechazo - fila1rechazo + 7))).Copy
    Intersect(Range(Rows(fila2tpv + 4), Rows(filainicialalertarechazo - 4)), Range(Columns(1), Columns(fila2rechazo - fila1rechazo + 7))).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
   ' copiar y pegar  datos como valores, 2 parte
    Intersect(Range(Rows(filainiciocomerciosvalidacion), Rows(filainicialalertavalidacion - 4)), Range(Columns(1), Columns(fila2rechazo - fila1rechazo + 4))).Copy
    Cells(filainiciocomerciosvalidacion, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         
     'eliminando columnas con tablas dinámicas
     
    
        Range(Columns(2), Columns(fila2rechazo - fila1rechazo + 4)).ColumnWidth = 20
      
      
      'centrar fechas
      With Rows(fila2tpv + 5)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
      End With
           
 
        
       
        
         'Emparejar ejes de los graficos
        Call tamaño_eje_graficas(Sheets(Sheets.Count - 1).Name, Sheets(Sheets.Count - 1).Name, 7, 8, "dos")
        Call tamaño_eje_graficas(Sheets(Sheets.Count - 1).Name, Sheets(Sheets.Count - 1).Name, 2, 5, "uno")
        Call tamaño_eje_graficas(Sheets(Sheets.Count - 1).Name, Sheets(Sheets.Count - 1).Name, 4, 6, "uno")
        
        
        'resumen graficos
         Call resumen_graficos(Sheets(Sheets.Count - 1).Name)
        
        'quitando lo no necesario y bloqueando la hoja
        Range(Columns(200), Columns(16384)).Delete
        Range(Rows(fila2tpv), Rows(1)).Delete
        Rows("1:" & fila2tpv).Insert Shift:=xlDown
        
        
        'revisar que pasa
        Cells(7 + fila2tpv, 2).Select
        ActiveWindow.FreezePanes = False
        
        ActiveWindow.FreezePanes = True
        
        
        'formato de las columnas pronostico y acumulado
        
            With Intersect(Range(Columns(fila2rechazo - fila1rechazo + 2), Columns(fila2rechazo - fila1rechazo + 3)), Range(Rows(fila2tpv + 6), Rows(final2 + 1)))
               .Interior.Color = RGB(255, 192, 0)
            End With
  
  
             'color del mes anterior
  
  
             With Intersect(Columns(fila2rechazo - fila1rechazo + 4), Range(Rows(fila2tpv + 6), Rows(final2 + 1)))
                .Interior.Color = RGB(220, 230, 241)
             End With
        
        
        Sheets(Sheets.Count - 1).Cells.Locked = False
        Sheets(Sheets.Count - 1).Range(Rows(1), Rows(fila2tpv + 3)).Locked = True
        Sheets(Sheets.Count - 1).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingRows:=True, AllowUsingPivotTables:=False, Password:="1234"
        Sheets(Sheets.Count - 1).EnableSelection = xlUnlockedCells
        
        
        'para arreglar altura de comentarios
        
         Set comentarios_hoja = Sheets(Sheets.Count - 1).Comments
            For Each comentario_ In comentarios_hoja
               comentario_.Shape.Height = 12
            Next
        Sheets(Sheets.Count - 1).Cells(final3 + 2, fila2rechazo - fila1rechazo + 4).Select
        
        
          
  
       
        
        
            
    Next
    
Next



   For i = 1 To Sheets.Count
 
     If Sheets(i).Name = "Hoja3" Or Sheets(i).Name = "Hoja7" Then
       Sheets(i).Visible = False
     End If
 
   Next
   

Sheets("Resumen Graficos").Move after:=Sheets(Sheets.Count)
Sheets("hoja1").Range("a1").AutoFilter
Sheets("hoja1").Name = "Resumen TPV"
Sheets("Resumen TPV").Move after:=Sheets(Sheets.Count)


libro.Worksheets("informe").Select
Intersect(Columns(1), Range(Rows(filainicialcrecimientotop10), Rows(1048576))) = ""

Sheets("Resumen Graficos").Select

'aca va el proceso de comercios con nuevas tx

Call comercios_nuevos

libro.Sheets("Resumen Graficos").Select





Workbooks("informe_" & Replace(Date, "/", "-") & ".xlsm").Save



'13.5





libro.Close

Workbooks("macro13.14.1.xlsm").Close

'Fin proceso principal

Exit Sub
ehoja2:
a = UBound(hoja2) + 1
ReDim Preserve hoja2(a)
hoja2(a) = Cells(i, 1)
Resume Next


ehoja1:

a = UBound(hoja1) + 1
ReDim Preserve hoja1(a)
hoja1(a) = Cells(i, 3)
Resume Next



errordinamica:

Resume volver:

End Sub


Function buscar(valor, hoja)
columnas = Sheets(hoja).Range("a1").End(xlToRight).Column
filas = Sheets(hoja).Range("a1").End(xlDown).Row


For i = 1 To columnas

    If Sheets(hoja).Cells(1, i) = valor Then
        a = i
        i = columnas
    End If
   
Next
buscar = a
End Function

Function diasemana(fecha)
a = Weekday(CDate(fecha), 2)

If a = 1 Then

    f = "Lunes"
    
     ElseIf a = 2 Then
     f = "Martes"
     ElseIf a = 3 Then
     f = "Miércoles"
     ElseIf a = 4 Then
     f = "Jueves"
     ElseIf a = 5 Then
     f = "Viernes"
     ElseIf a = 6 Then
     f = "Sábado"
     ElseIf a = 7 Then
     f = "Domingo"
 
 
 End If


diasemana = f

End Function


Sub Normalizar_hojas(hoja, fecha_)

Sheets(3).Select
Range("a1").Select
Sheets(3).Columns(1).Delete
ReDim arreglo_inicial(1 To 7)
ReDim perso(1 To 2)

arreglo_inicial(1) = "polv3"
arreglo_inicial(2) = "PA"
arreglo_inicial(3) = "MX"
arreglo_inicial(4) = "CO"
arreglo_inicial(5) = "PE"
arreglo_inicial(6) = "BR"
arreglo_inicial(7) = "AR"
'arreglo_inicial(8) = "prueba2"

perso(1) = "f"
perso(2) = "t"


filafinal = Sheets(hoja).Range("a1").End(xlDown).Row

Sheets(hoja).Columns(buscar("pais", hoja)).Copy Destination:=Sheets(3).Range("a1")
Sheets(3).Range("a1").RemoveDuplicates Columns:=1, Header:=xlYes
filahoja_3 = Sheets(3).Range("a60000").End(xlUp).Row
fecha = fecha_

pais = buscar("pais", hoja)
agregador = buscar("tiene_c_personalizados", hoja)
rango = buscar("rango", hoja)

iteracion_fecha = CInt(Format(fecha, "dd"))
s = 0
For i = 1 To UBound(arreglo_inicial)
    
    For j = 2 To filahoja_3
    
      If arreglo_inicial(i) = Sheets(3).Cells(j, 1) Then
        Exit For
      
      End If
    
       
    Next
    
    If j > filahoja_3 Then
    
      For m = 1 To UBound(perso)
      
         
         For m2 = 1 To iteracion_fecha
            
            Sheets(hoja).Cells((filafinal + m + m2 - 1) + s, pais) = arreglo_inicial(i)
            'Cells((ultima + i + j - 3) + s, rango).Format = "Text"
            Sheets(hoja).Cells((filafinal + m + m2 - 1) + s, rango).Formula = "'" & Replace(DateAdd("d", -m2 + 1, fecha), "/", "-")
            Sheets(hoja).Cells((filafinal + m + m2 - 1) + s, agregador) = perso(m)
            
         
         
         Next
         
          filafinal = Sheets(hoja).Cells(1, pais).End(xlDown).Row - 1
      Next
    
      's = m + m1 - 1
    filafinal = Sheets(hoja).Cells(1, pais).End(xlDown).Row
    
    End If
    
    

Next
    
    
    Sheets(3).Cells(1, 1) = "pais"
    Sheets(3).Cells(1, 3) = "pais"
    
    
    For i = 1 To UBound(arreglo_inicial)

        Sheets(3).Cells(i + 1, 1) = arreglo_inicial(i)
        Sheets(3).Cells(i + 1, 3) = arreglo_inicial(i)

    Next

        Sheets(3).Cells(i + 1, 1) = "todos"
        Sheets(3).Cells(i + 1, 3) = "todos"

End Sub








Sub normalizacion_hojas_sub()






'abrir archivos y luego copia y pega los datos
Set libro = Application.Workbooks.Add

libro.Sheets.Add after:=libro.Sheets(Sheets.Count)
libro.Sheets.Add after:=libro.Sheets(Sheets.Count)
libro.Sheets.Add after:=libro.Sheets(Sheets.Count)
libro.Sheets.Add after:=libro.Sheets(Sheets.Count)
libro.Sheets.Add after:=libro.Sheets(Sheets.Count)
libro.Sheets.Add after:=libro.Sheets(Sheets.Count)
libro.Sheets.Add after:=libro.Sheets(Sheets.Count)

Application.Workbooks.Open ("D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\rechazo.xlsx")
'Application.Workbooks.Open ("C:\Users\user\Documents\Javier\Pagos Online\informe jose\Abril 27\rechazo.xlsx")

fecha_rechazo = mayor_fecha(Sheets(1).Name, "rango")

Workbooks("rechazo.xlsx").Sheets(1).Cells.Copy Destination:=libro.Sheets(1).Cells

Application.Workbooks.Open ("D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\tpvagregador.xlsx")
'Application.Workbooks.Open ("C:\Users\user\Documents\Javier\Pagos Online\informe jose\Abril 27\tpvagregador.xlsx")

fecha_tpv = mayor_fecha(Sheets(1).Name, "rango")

Workbooks("tpvagregador.xlsx").Sheets(1).Cells.Copy Destination:=libro.Sheets(2).Cells

If fecha_rechazo <> fecha_tpv Then

    salir_por_fecha = "si"
    Exit Sub
    Else
    salir_por_fecha = "no"
    fecha_guia = fecha_rechazo
End If


'Abrir valores mes anterior

fecha_mes_anterior = DatePart("yyyy", DateAdd("m", -1, fecha_guia)) & "-" & DatePart("m", DateAdd("m", -1, fecha_guia)) & ".xlsx"
     Application.Workbooks.Open ("D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\rechazo_" & fecha_mes_anterior)

    Workbooks("rechazo_" & fecha_mes_anterior).Sheets(1).Cells.Copy Destination:=libro.Sheets(4).Cells



   Application.Workbooks.Open ("D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\tpvagregador_" & fecha_mes_anterior)

   Workbooks("tpvagregador_" & fecha_mes_anterior).Sheets(1).Cells.Copy Destination:=libro.Sheets(5).Cells


Workbooks("rechazo.xlsx").Close
Workbooks("tpvagregador.xlsx").Close
Workbooks("rechazo_" & fecha_mes_anterior).Close
Workbooks("tpvagregador_" & fecha_mes_anterior).Close


'fin mes anterior


libro.Activate
'normalizacion hojas

  For i = 1 To 5

        If i = 1 Or i = 2 Then
        
            Call Normalizar_hojas(Sheets(i).Name, fecha_guia)

        ElseIf i = 4 Or i = 5 Then

             Call Normalizar_hojas(Sheets(i).Name, Application.WorksheetFunction.EoMonth(CDate(DateAdd("m", -1, fecha_guia)), 0))

        
            
        End If
    Next
    
    
    Sheets(3).Range("E1") = "Agregador"
    Sheets(3).Range("E2") = "Gateway"
    Sheets(3).Range("E3") = "todos"
    Sheets(3).Range("f1") = "f"
    Sheets(3).Range("f2") = "t"
    Sheets(3).Range("f3") = "(All)"
    Sheets(3).Range("g1") = "Valores Absolutos"
    Sheets(3).Range("g2") = "Valores Relativos"
    
    
    
'fin normalización hojas

End Sub





Sub tablas_dinamicas()


Worksheets.Add after:=Sheets(Sheets.Count)
Range(Columns(1), Columns(1)).ColumnWidth = 80

'hoja1(la hoja del informe)

'tabla dinamica rechazos mes actual
Set Cache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=Sheets("hoja1").Range("A1").CurrentRegion)

Set dinamica = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache, _
        TableDestination:=Range("A3"))
dinamica.Name = "dinamica1"

With dinamica
        .ColumnGrand = False
        .RowGrand = False
        
        .CalculatedFields.Add "Validacion M", "=Validacion_MAF-RechazadasMAFsinValidar", True
        

             
                
        With .PivotFields("Transacciones_totales")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "Transacciones totales"
        End With
        
        With .PivotFields("Total_Aprobadas")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "Total Aprobadas"
        End With
        
        
         With .PivotFields("Rechazototal")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "Rechazo Total"
        End With
        
         With .PivotFields("Rechazo_Lista_Negra")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "    Rechazo Lista Negra"
        End With
        
        With .PivotFields("Rechazo_Reglas")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "    Rechazo Reglas"
        End With
        
        With .PivotFields("RechazadasMAFsinValidar")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "    Rechazo Reglas (comercios sin validación)"
        End With
        
       With .PivotFields("RechazadasAnalista")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "    Rechazadas Analista"
        End With
        
        With .PivotFields("RechazoBancario")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "    Rechazo Bancario"
        End With
        
          With .PivotFields("Detenidas")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "Detenidas "
        End With
        
           With .PivotFields("CanceladaValidacion")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "Cancelada Validacion "
        End With
        
        With .PivotFields("Intermitencias")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "Transacciones sin Estado Final "
        End With
        
       With .PivotFields("Validacion M")
         .Orientation = xlDataField
         .Function = xlSum
         .Caption = "Validadas MAF"
        End With
        
        '.PivotFields("codigo_comercio").NumberFormat = "General"
        .PivotFields("rango").Orientation = xlRowField
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField

  End With

Final = Cells(1048576, 1).End(xlUp).Row + 5

columna_inicial_rechazos_manterior = Cells(5, 16000).End(xlToLeft).Column + 5
'fin tabla dinamica rechazos mes actual



'inici tabla dinamica rechazasos mes anterior


Set Cache_1 = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=Sheets("hoja4").Range("A1").CurrentRegion)

Set dinamica_1 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache_1, _
        TableDestination:=Cells(3, columna_inicial_rechazos_manterior))
dinamica_1.Name = "dinamicarechazosantes"

With dinamica_1
        .ColumnGrand = False
        .RowGrand = False
        
        .CalculatedFields.Add "Validacion M", "=Validacion_MAF-RechazadasMAFsinValidar", True
        

             
                
        With .PivotFields("Transacciones_totales")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "Transacciones totales"
        End With
        
        With .PivotFields("Total_Aprobadas")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "Total Aprobadas"
        End With
        
        
         With .PivotFields("Rechazototal")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "Rechazo Total"
        End With
        
         With .PivotFields("Rechazo_Lista_Negra")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "    Rechazo Lista Negra"
        End With
        
        With .PivotFields("Rechazo_Reglas")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "    Rechazo Reglas"
        End With
        
        With .PivotFields("RechazadasMAFsinValidar")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "    Rechazo Reglas (comercios sin validación)"
        End With
        
       With .PivotFields("RechazadasAnalista")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "    Rechazadas Analista"
        End With
        
        With .PivotFields("RechazoBancario")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "    Rechazo Bancario"
        End With
        
          With .PivotFields("Detenidas")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "Detenidas "
        End With
        
           With .PivotFields("CanceladaValidacion")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "Cancelada Validacion "
        End With
        
        With .PivotFields("Intermitencias")
        .Orientation = xlDataField
        .Function = xlSum
        .Caption = "Transacciones sin Estado Final "
        End With
        
       With .PivotFields("Validacion M")
         .Orientation = xlDataField
         .Function = xlSum
         .Caption = "Validadas MAF"
        End With
        
        '.PivotFields("codigo_comercio").NumberFormat = "General"
        '.PivotFields("rango").Orientation = xlRowField
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField

  End With








'fin tabla dinamica rechazos mes anterior

'tabla dinamica tpv mes actual
Set Cache2 = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=Sheets("hoja2").Range("A1").CurrentRegion)

Set dinamica2 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache2, _
        TableDestination:=Cells(Final, 1))

dinamica2.Name = "dinamica2"







With dinamica2
        .ColumnGrand = False
        .RowGrand = False
        
        .PivotFields("rango").Orientation = xlRowField
        
        With .PivotFields("suma")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "TPV Millones"
         .NumberFormat = "$ #,##0,, "
         
        End With
         With .PivotFields("count")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "Cantidad Liberadas"
         
        End With
         .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField


 End With

'fin tpv actual



'inicio tpv mes anterior

Set Cache2_1 = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=Sheets("hoja5").Range("A1").CurrentRegion)

Set dinamica2_1 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache2_1, _
        TableDestination:=Cells(Final, columna_inicial_rechazos_manterior))

dinamica2_1.Name = "dinamica2_anterior"







With dinamica2_1
        .ColumnGrand = False
        .RowGrand = False
        
        '.PivotFields("rango").Orientation = xlRowField
        
        With .PivotFields("suma")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "TPV Millones"
         .NumberFormat = "$ #,##0,, "
         
        End With
         With .PivotFields("count")
        .Orientation = xlDataField
        .Function = xlSum
         .Caption = "Cantidad Liberadas"
         
        End With
         .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        '.PivotFields("usuario_id").Orientation = xlPageField


 End With


'fin tpv mes anterior





fila1rechazo = Range("a1").End(xlDown).End(xlDown).Row
columna1rechazo = Cells(fila1rechazo, 1).End(xlToRight).Column

columna1participaciontpv = columna1rechazo + 1000
columna2participaciontpv = columna1rechazo + 1002



Set dinamica3 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache2, _
        TableDestination:=Cells(1, columna1participaciontpv))



With dinamica3
        .ColumnGrand = False
        .RowGrand = False
    
        
        .PivotFields("usuario_id").Orientation = xlRowField
        
        With .PivotFields("suma")
         .Orientation = xlDataField
         .Function = xlSum
         .Caption = "TPV Millones"
         .NumberFormat = "$ #,##0,, "
         
        End With
        
         With .PivotFields("suma")
          .Orientation = xlDataField
          .Function = xlSum
          .Caption = "Participación Total"
          .Calculation = xlPercentOfColumn
         
        End With
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        '.PivotFields("usuario_id").PivotFilters.Add Type:=xlTopCount, DataField:=dinamica3.PivotFields("TPV Millones"), Value1:=20
        .PivotFields("usuario_id").AutoSort xlDescending, "TPV Millones", dinamica3.PivotColumnAxis.PivotLines(1), 1

End With

Cells(1048576, columna1participaciontpv).End(xlUp).CurrentRegion.Name = "tpv_actual"

Cells(5, columna1participaciontpv) = "Usuario id"
columna1participaciontx = columna2participaciontpv + 5



Set dinamica4 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache2, _
        TableDestination:=Cells(1, columna1participaciontx))
        
        
 With dinamica4
        .ColumnGrand = False
        .RowGrand = False
    
        
        .PivotFields("usuario_id").Orientation = xlRowField
        
        With .PivotFields("count")
         .Orientation = xlDataField
         .Function = xlSum
         .Caption = "Cantidad Transacciones"
         
         
        End With
        
         With .PivotFields("count")
          .Orientation = xlDataField
          .Function = xlSum
          .Caption = "Participación Total"
          .Calculation = xlPercentOfColumn
         
        End With
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        '.PivotFields("usuario_id").PivotFilters.Add Type:=xlTopCount, DataField:=dinamica4.PivotFields("Cantidad Transacciones"), Value1:=20
        .PivotFields("usuario_id").AutoSort xlDescending, "Cantidad Transacciones", dinamica4.PivotColumnAxis.PivotLines(1), 1

End With
     
  Cells(5, columna1participaciontx) = "Usuario id"
        
        
        
        'dinammica rechazos maf y total
        
        columna1rechazomaf = Cells(5, 16384).End(xlToLeft).Column + 6
        
        
        
        
    Set dinamica5 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache, _
        TableDestination:=Cells(1, columna1rechazomaf))
        
        dinamica5.CalculatedFields.Add "Porcentaje Rechazo MAF", "=(Rechazo_Lista_Negra +RechazadasMAFsinValidar +Rechazo_Reglas)/Transacciones_totales", True
        dinamica5.CalculatedFields.Add "Porcentaje Rechazo Total", "=Rechazototal/Transacciones_totales", True
    With dinamica5
        .ColumnGrand = False
        .RowGrand = False
    
        
        .PivotFields("rango").Orientation = xlRowField
        
        With .PivotFields("Porcentaje Rechazo MAF")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         '.Calculation = xlPercentOfColumn
         .Caption = "Porcentaje MAF"
         
         
        End With
        
         With .PivotFields("Porcentaje Rechazo Total")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         '.Calculation = xlPercentOfColumn
         .Caption = "Porcentaje Total"
         
         
        End With
        
        
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        

     End With
        
        Cells(4, columna1rechazomaf).CurrentRegion.Name = "rechazos_maf"
        
        
        'fin dinamica rechazos maf y total
        
        
        
        'dinammica comportamiento MAF
        
        columna1rechazototal = Cells(5, 16384).End(xlToLeft).Column + 6
        
        
        
        
    Set dinamica6 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache, _
        TableDestination:=Cells(1, columna1rechazototal))
        dinamica6.CalculatedFields.Add "Rechazo MAF total ", "=RechazadasMAFsinValidar+Rechazo_Lista_Negra+Rechazo_Reglas", True
        dinamica6.CalculatedFields.Add "Aprobacion MAF", "=Transacciones_totales-'Validacion M'-'Rechazo MAF total '", True
    With dinamica6
        .ColumnGrand = False
        .RowGrand = False
    
        
        
        
        With .PivotFields("Aprobacion MAF")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " #,##0"
         '.Calculation = xlPercentOfColumn
         .Caption = "Aprobar/Liberar"
         
        End With
        
        With .PivotFields("Rechazo MAF total ")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " #,##0"
         '.Calculation = xlPercentOfColumn
         .Caption = "Rechazar"
         
        End With
        
        With .PivotFields("Validacion M")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " #,##0"
         '.Calculation = xlPercentOfColumn
         .Caption = "Detener/Validar"
         
        End With
        
        
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        

     End With
        
        Cells(4, columna1rechazototal).CurrentRegion.Name = "rechazos_total"
        
        
        'fin dinamica rechazos totales
        
        
     'dinammica comercios problema rechazos maf
        
    columna1comerciorechazomaf = Cells(5, 16384).End(xlToLeft).Column + 6
        
        
        
        
    Set dinamica7 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache, _
        TableDestination:=Cells(1, columna1comerciorechazomaf))
        
       
    With dinamica7
         
         
        .ColumnGrand = False
        .RowGrand = False
    
        
        .PivotFields("codigo_comercio").Orientation = xlRowField
        
        With .PivotFields("Rechazo MAF total ")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " #,##0 "
         .Caption = "Cantidad Rechazo MAF"
         
         
        End With
        
        With .PivotFields("Rechazo MAF total ")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         .Calculation = xlPercentOfColumn
         .Caption = "Participación Total"
         
         
        End With
        
        
         With .PivotFields("Porcentaje Rechazo MAF")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         .Caption = "Rechazo Comercio"
      
         
         
        End With
        
        
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        .PivotFields("codigo_comercio").AutoSort xlDescending, "Cantidad Rechazo MAF", dinamica7.PivotColumnAxis.PivotLines(1), 1

     End With
        
        Cells(5, columna1comerciorechazomaf).CurrentRegion.Name = "comercios_rechazo_maf"
        
        
        'fin dinammica comercios problema rechazos maf
        
       

'ss
 'dinammica validacion maf
        
    columna1validacionmaf = Cells(5, 16384).End(xlToLeft).Column + 6
        
        
         
        
    Set dinamica8 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache, _
        TableDestination:=Cells(1, columna1validacionmaf))
        
    dinamica8.CalculatedFields.Add "Porcentaje Validacion ", "='Validacion M'/Transacciones_totales", True
    dinamica8.CalculatedFields.Add "Porcentaje Efectividad ", "= RechazadasAnalista /('Validacion M'-Detenidas-CanceladaValidacion)", True
    
    
    With dinamica8
         
         
        .ColumnGrand = False
        .RowGrand = False
    
        
        .PivotFields("rango").Orientation = xlRowField
        
        With .PivotFields("Porcentaje Validacion ")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         .Caption = "Porcentaje Validacion"
         
         
        End With
        
        With .PivotFields("Porcentaje Efectividad ")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         .Caption = "Efectividad"
         
         
        End With
        
        
        
        
        
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        

     End With
        
        Cells(5, columna1validacionmaf).CurrentRegion.Name = "validacion_maf"
        
        
        'fin dinammica comercios problema rechazos maf


'ss



'dinammica comercios problema validacion maf
        
    columna1comerciovalidacionmaf = Cells(5, 16384).End(xlToLeft).Column + 6
        
        
        
        
    Set dinamica9 = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache, _
        TableDestination:=Cells(1, columna1comerciovalidacionmaf))
        
       
    With dinamica9
         
         
        .ColumnGrand = False
        .RowGrand = False
    
        
        .PivotFields("codigo_comercio").Orientation = xlRowField
        
        With .PivotFields("Validacion M")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " #,##0 "
         .Caption = "Cantidad Validar MAF"
         
         
        End With
        
        With .PivotFields("Validacion M")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         .Calculation = xlPercentOfColumn
         .Caption = "Participación Total V"
         
         
        End With
        
        
         With .PivotFields("Porcentaje Validacion ")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         .Caption = "Validacion Comercio"
      
         
         
        End With
        
        
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        .PivotFields("codigo_comercio").AutoSort xlDescending, "Cantidad Validar MAF", dinamica9.PivotColumnAxis.PivotLines(1), 1

     End With
        
        Cells(5, columna1comerciovalidacionmaf).CurrentRegion.Name = "comercios_validacion_maf"
        
        
        'fin dinammica comercios problema validacion maf

'inicia dinamica rechazos maf mes anterior

        columna1rechazomaf_anterior = Cells(5, 16384).End(xlToLeft).Column + 6
        
        
        
        
    Set dinamica5_anterior = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache_1, _
        TableDestination:=Cells(1, columna1rechazomaf_anterior))
        
        dinamica5_anterior.CalculatedFields.Add "Porcentaje Rechazo MAF", "=(Rechazo_Lista_Negra +RechazadasMAFsinValidar +Rechazo_Reglas)/Transacciones_totales", True
        dinamica5_anterior.CalculatedFields.Add "Porcentaje Rechazo Total", "=Rechazototal/Transacciones_totales", True
    With dinamica5_anterior
        .ColumnGrand = False
        .RowGrand = False
    
        
        .PivotFields("rango").Orientation = xlRowField
        
        With .PivotFields("Porcentaje Rechazo MAF")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         '.Calculation = xlPercentOfColumn
         .Caption = "Porcentaje MAF"
         
         
        End With
        
         With .PivotFields("Porcentaje Rechazo Total")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         '.Calculation = xlPercentOfColumn
         .Caption = "Porcentaje Total"
         
         
        End With
        
        
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        

     End With
        
        Cells(4, columna1rechazomaf_anterior).CurrentRegion.Name = "rechazos_maf_anterior"
        

'fin dinamica rechazos maf mes anterior


'inicio dinamica validacion maf mes anterior
    columna1validacionmaf_anterior = Cells(5, 16384).End(xlToLeft).Column + 6
        
        
         
        
    Set dinamica8_anterior = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache_1, _
        TableDestination:=Cells(1, columna1validacionmaf_anterior))
        
    dinamica8_anterior.CalculatedFields.Add "Porcentaje Validacion ", "='Validacion M'/Transacciones_totales", True
    dinamica8_anterior.CalculatedFields.Add "Porcentaje Efectividad ", "= RechazadasAnalista /('Validacion M'-Detenidas-CanceladaValidacion)", True
    
    
    With dinamica8_anterior
         
         
        .ColumnGrand = False
        .RowGrand = False
    
        
        .PivotFields("rango").Orientation = xlRowField
        
        With .PivotFields("Porcentaje Validacion ")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         .Caption = "Porcentaje Validacion"
         
         
        End With
        
        With .PivotFields("Porcentaje Efectividad ")
         .Orientation = xlDataField
         .Function = xlSum
         .NumberFormat = " 0.0 %"
         .Caption = "Efectividad"
         
         
        End With
        
        
        
        
        
        .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        
        

     End With
        
        Cells(5, columna1validacionmaf_anterior).CurrentRegion.Name = "validacion_maf_anterior"



'inicio tpv comercios mes anterior

columna1_tpv_anterior = Cells(5, 16384).End(xlToLeft).Column + 6

Set dinamica_tpv_anterior = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache2_1, _
        TableDestination:=Cells(1, columna1_tpv_anterior))

dinamica_tpv_anterior.Name = "dinamica_tpv_anterior"







With dinamica_tpv_anterior
        .ColumnGrand = False
        .RowGrand = False
        
        '.PivotFields("rango").Orientation = xlRowField
        
        With .PivotFields("suma")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "TPV Millones"
            '.Calculation = xlPercentOfColumn
            .NumberFormat = "$ #,##0,,"
         
        End With
        
         .PivotFields("usuario_id").AutoSort xlDescending, "TPV Millones", dinamica_tpv_anterior.PivotColumnAxis.PivotLines(1), 1
        
        
         .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        .PivotFields("usuario_id").Orientation = xlRowField


 End With

Cells(1048576, columna1_tpv_anterior).End(xlUp).CurrentRegion.Name = "tpv_anterior"
'fin tpv comercios mes anterior



'dinamica tpv mes anterior por dia
columna1_tpv_anterior_dia = Cells(5, 16384).End(xlToLeft).Column + 6

Set dinamica_tpv_anterior_dia = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache2_1, _
        TableDestination:=Cells(1, columna1_tpv_anterior_dia))

dinamica_tpv_anterior_dia.Name = "dinamica_tpv_anterior_dia"







 With dinamica_tpv_anterior_dia
        .ColumnGrand = False
        .RowGrand = False
        
        .PivotFields("rango").Orientation = xlRowField
        
        With .PivotFields("suma")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "TPV Millones"
           
         
        End With
        
        
        
        
         .PivotFields("pais").Orientation = xlPageField
        .PivotFields("tiene_c_personalizados").Orientation = xlPageField
        


 End With

Cells(5, columna1_tpv_anterior_dia).CurrentRegion.Name = "tpv_anterior_dia"
'fin 'dinamica tpv mes anterior por dia




'fin dinamica validacion maf mes anterior


dinamica.HasAutoFormat = False
dinamica2.HasAutoFormat = False
dinamica3.HasAutoFormat = False
dinamica4.HasAutoFormat = False
dinamica_1.HasAutoFormat = False
dinamica2_1.HasAutoFormat = False



fila2rechazo = Cells(fila1rechazo, 1).End(xlDown).Row
fila1tpv = Cells(fila1rechazo, 1).End(xlDown).End(xlDown).End(xlDown).End(xlDown).Row + 8

columna1tpv = Cells(fila1tpv - 8, 1).End(xlToRight).Column
fila2tpv = Cells(1048576, 1).End(xlUp).Row + 8


Cells(fila2tpv + 5, 1).FormulaR1C1 = "=+TRANSPOSE(R[" & -(fila2tpv + 5 - fila1rechazo) & "]C[" & 0 & "]:R[" & -(fila2tpv + 5 - fila2rechazo) & "]C[" & columna1rechazo - 1 & "])"

Range(Cells(fila2tpv + 5, 1), Cells(fila2tpv + 5 + columna1rechazo - 1, fila2rechazo - fila1rechazo + 1)).FormulaArray = "=+TRANSPOSE(R[" & -(fila2tpv + 5 - fila1rechazo) & "]C[" & 0 & "]:R[" & -(fila2tpv + 5 - fila2rechazo) & "]C[" & columna1rechazo - 1 & "])"

'datos rechazo mes anterior

Cells(fila2tpv + 5 + 1, fila2rechazo - fila1rechazo + 4).FormulaArray = "=+transpose(" & Range(Cells(6, columna_inicial_rechazos_manterior).Address & ":" & Cells(6, columna_inicial_rechazos_manterior).End(xlToRight).Address).Address & ")"

Range(Cells(fila2tpv + 5 + 1, fila2rechazo - fila1rechazo + 4), Cells(fila2tpv + 5 + 1 + Cells(6, columna_inicial_rechazos_manterior).End(xlToRight).Column - columna_inicial_rechazos_manterior, fila2rechazo - fila1rechazo + 4)).FormulaArray = "=+transpose(" & Range(Cells(6, columna_inicial_rechazos_manterior).Address & ":" & Cells(6, columna_inicial_rechazos_manterior).End(xlToRight).Address).Address & ")"

Cells(fila2tpv + 5, fila2rechazo - fila1rechazo + 4) = "Mes Anterior"

'fin datos rechazo mes anterior


dias_mes_actual = Format(DateAdd("d", -1, "01/" & Format(DateAdd("m", 1, Cells(fila2tpv + 5, fila2rechazo - fila1rechazo + 1)), "mm/yyyy")), "d")
dias_mes_anterior = Format(DateAdd("d", -1, "01/" & Format(DateAdd("m", 0, fecha_guia), "mm/yyyy")), "d")
         


'Calculo variacion diaria promedio rechazos

Cells(fila2tpv + 5, fila2rechazo - fila1rechazo + 5) = "Variación"
For i = 1 To columna1rechazo - 1
   Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 5).Formula = "=+" & Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 3).Address & "/" & Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 4).Address & "-1"
   Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 5).NumberFormat = "0.00%"
   
   'formato condicional
    Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 5).Select
    Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 5).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 5).FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 5).FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 5).FormatConditions(1).StopIfTrue = False
   
Next


'fin variacion diaria rechazos

referencia = Cells(fila2tpv + 6, 2).Address


'valores absolutos hasta el mes
Cells(fila2tpv + 5, fila2rechazo - fila1rechazo + 2) = "M-t-d"
For i = 0 To columna1rechazo - 2
 Cells(fila2tpv + 6 + i, fila2rechazo - fila1rechazo + 2).Formula = "=+sum(offset(" & referencia & "," & i & ",0):offset(" & referencia & "," & i & "," & fila2rechazo - fila1rechazo - 1 & "))"


Next

Call vector_dias_laborales(fecha_guia)

'pronóstico de valores absolutos
Cells(fila2tpv + 5, fila2rechazo - fila1rechazo + 3) = "Forecast"
For i = 0 To columna1rechazo - 2
 
Cells(fila2tpv + 6 + i, fila2rechazo - fila1rechazo + 3).FormulaLocal = "=+redondear(sumaproducto(" & "desref(" & referencia & ";" & i & ";0):desref(" & referencia & ";" & i & ";" & fila2rechazo - fila1rechazo - 1 & ")" & ";" & entre_semana & ") *" & multiplicador_dias_laboral & ";0)+" & "redondear(sumaproducto(" & "desref(" & referencia & ";" & i & ";0):desref(" & referencia & ";" & i & ";" & fila2rechazo - fila1rechazo - 1 & ")" & ";" & fin_Semana & ") *" & multiplicador_dias_no_laboral & ";0)"

Next




' cálculo valor relativos

penultimafila = Cells(fila2tpv + 5, 1).End(xlDown).Row

'cells(penultimafila+3,1)="Etiquetas de fila"


Dim arreglo(1 To 12)
arreglo(1) = "Total Aprobadas"
arreglo(2) = "Rechazo Total"
arreglo(3) = "    Rechazo Lista Negra"
arreglo(4) = "    Rechazo Reglas"
arreglo(5) = "    Rechazo Reglas (comercios sin validación)"
arreglo(6) = "    Rechazadas Analista"
arreglo(7) = "    Rechazo Bancario"
arreglo(8) = "Detenidas"
arreglo(9) = "Cancelada Validacion"
arreglo(10) = "Transacciones sin Estado Final"
arreglo(11) = "Validadas MAF"
arreglo(12) = "Efectividad Reglas"




espacionabsolutorelativo = 5
For k = 0 To UBound(arreglo) - 1 'itera fila
For i = 0 To fila2rechazo - fila1rechazo + 3 'itera columna
    
    If i = 0 Then
    
    Cells(penultimafila + espacionabsolutorelativo + k, i + 1) = arreglo(k + 1)
    
       
    ElseIf k = UBound(arreglo) - 1 Then
       
       Cells(penultimafila + espacionabsolutorelativo + k, i + 1).FormulaLocal = "=+si(esnumero(" & "desref(" & referencia & ";" & k - 5 & ";" & i - 1 & ") / (desref(" & referencia & ";" & k & ";" & i - 1 & ")- desref(" & referencia & ";" & k - 2 & ";" & i - 1 & ")- desref(" & referencia & ";" & k - 3 & ";" & i - 1 & "))" & ");" & "desref(" & referencia & ";" & k - 5 & ";" & i - 1 & ") / (desref(" & referencia & ";" & k & ";" & i - 1 & ")- desref(" & referencia & ";" & k - 2 & ";" & i - 1 & ")- desref(" & referencia & ";" & k - 3 & ";" & i - 1 & "))" & ";0)"
        Cells(penultimafila + espacionabsolutorelativo + k, i + 1).NumberFormat = "#0.00%"
    
    Else
    
      Cells(penultimafila + espacionabsolutorelativo + k, i + 1).FormulaLocal = "=+si(esnumero(" & "desref(" & referencia & ";" & k + 1 & ";" & i - 1 & ") / desref(" & referencia & ";" & 0 & ";" & i - 1 & ")" & ");" & "desref(" & referencia & ";" & k + 1 & ";" & i - 1 & ") / desref(" & referencia & ";" & 0 & ";" & i - 1 & ")" & ";0)"
   
   
     Cells(penultimafila + espacionabsolutorelativo + k, i + 1).NumberFormat = "#0.00%"
    
    End If
   Next
Next



'variacion valores relativos


For i = 1 To columna1rechazo - 1



    Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 5).Formula = "=+" & Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 3).Address & "/" & Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 4).Address & "-1"
    
    Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 5).NumberFormat = "0.0%"
     
   'formato condicional
    'Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 5).Select
    Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 5).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 5).FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 5).FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Cells(penultimafila + espacionabsolutorelativo + i - 1, fila2rechazo - fila1rechazo + 5).FormatConditions(1).StopIfTrue = False



Next



'fin variacion valores relativos

'TPV



ultimafila = Cells(1048576, 1).End(xlUp).Row + 6
referencia3 = Cells(ultimafila + 1, 2).Address

Cells(ultimafila, 1).FormulaR1C1 = "=+transpose(OFFSET(" & "R" & Cells(fila1tpv, 1).Row & "C" & Cells(fila1tpv, 1).Column & ", 0,0):OFFSET(" & "R" & Cells(fila1tpv, 1).Row & "C" & Cells(fila1tpv, 1).Column & "," & fila2tpv - fila1tpv & "," & columna1tpv - 1 & "))"
Range(Cells(ultimafila, 1), Cells(ultimafila + columna1tpv - 1, fila2tpv - fila1tpv + 1)).FormulaArray = "=+transpose(OFFSET(" & "R" & Cells(fila1tpv - 8, 1).Row & "C" & Cells(fila1tpv - 8, 1).Column & ", 0,0):OFFSET(" & "R" & Cells(fila1tpv - 8, 1).Row & "C" & Cells(fila1tpv - 8, 1).Column & "," & fila2tpv - fila1tpv & "," & columna1tpv - 1 & "))"


'tpv del mes anterior

  Cells(ultimafila + 1, fila2rechazo - fila1rechazo + 4).FormulaArray = "=+transpose(" & Range(Cells(fila1tpv - 7, columna_inicial_rechazos_manterior).Address & ":" & Cells(fila1tpv - 7, columna_inicial_rechazos_manterior).End(xlToRight).Address).Address & ")"

  Range(Cells(ultimafila + 1, fila2rechazo - fila1rechazo + 4), Cells(ultimafila + 1 + Cells(fila1tpv - 7, columna_inicial_rechazos_manterior).End(xlToRight).Column - columna_inicial_rechazos_manterior, fila2rechazo - fila1rechazo + 4)).FormulaArray = "=+transpose(" & Range(Cells(fila1tpv - 7, columna_inicial_rechazos_manterior).Address & ":" & Cells(fila1tpv - 7, columna_inicial_rechazos_manterior).End(xlToRight).Address).Address & ")"




'fin tpv  del mes anterior


'Variacion tpv diario

 For i = 1 To 3
 
 Cells(ultimafila + i, fila2rechazo - fila1rechazo + 5).Formula = "=+" & Cells(ultimafila + i, fila2rechazo - fila1rechazo + 3).Address & "/" & Cells(ultimafila + i, fila2rechazo - fila1rechazo + 4).Address & "-1"
 
 'formato condicional
 
    'Cells(fila2tpv + 5 + i, fila2rechazo - fila1rechazo + 5).Select
    Cells(ultimafila + i, fila2rechazo - fila1rechazo + 5).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Cells(ultimafila + i, fila2rechazo - fila1rechazo + 5).FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Cells(ultimafila + i, fila2rechazo - fila1rechazo + 5).FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Cells(ultimafila + i, fila2rechazo - fila1rechazo + 5).FormatConditions(1).StopIfTrue = False
    

Next
'fin Variacion tpv diario



'valor agregado tpv

Cells(ultimafila + 1, fila2tpv - fila1tpv + 2) = "=+sum(offset(" & referencia3 & "," & 0 & ",0):offset(" & referencia3 & "," & 0 & "," & fila2rechazo - fila1rechazo - 1 & "))"
Cells(ultimafila + 2, fila2tpv - fila1tpv + 2) = "=+sum(offset(" & referencia3 & "," & 1 & ",0):offset(" & referencia3 & "," & 1 & "," & fila2rechazo - fila1rechazo - 1 & "))"

'pronosticos tpv

Cells(ultimafila + 1, fila2tpv - fila1tpv + 3).FormulaLocal = "=+redondear(sumaproducto(" & "desref(" & referencia3 & ";" & 0 & ";" & 0 & "):" & "desref(" & referencia3 & ";" & 0 & ";" & fila2rechazo - fila1rechazo - 1 & ")" & ";" & entre_semana & ") *" & multiplicador_dias_laboral & ";0)+" & "redondear(sumaproducto(" & "desref(" & referencia3 & ";" & 0 & ";" & 0 & "):" & "desref(" & referencia3 & ";" & 0 & ";" & fila2rechazo - fila1rechazo - 1 & ")" & ";" & fin_Semana & ") *" & multiplicador_dias_no_laboral & ";0)"
Cells(ultimafila + 2, fila2tpv - fila1tpv + 3).FormulaLocal = "=+redondear(sumaproducto(" & "desref(" & referencia3 & ";" & 1 & ";" & 0 & "):" & "desref(" & referencia3 & ";" & 1 & ";" & fila2rechazo - fila1rechazo - 1 & ")" & ";" & entre_semana & ") *" & multiplicador_dias_laboral & ";0)+" & "redondear(sumaproducto(" & "desref(" & referencia3 & ";" & 1 & ";" & 0 & "):" & "desref(" & referencia3 & ";" & 1 & ";" & fila2rechazo - fila1rechazo - 1 & ")" & ";" & fin_Semana & ") *" & multiplicador_dias_no_laboral & ";0)"


Rows(ultimafila + 1).NumberFormat = "$ #,##0,,.00 "

anchoboton = Columns(1).ColumnWidth + Columns(2).ColumnWidth + Columns(3).ColumnWidth + Columns(4).ColumnWidth + Columns(5).ColumnWidth
  
'ActiveSheet.Buttons.Add(anchoboton * 5.54142985342507 + 10, 15 * (fila2tpv + 1), 60.75, 20.25).Select
'ActiveSheet.Buttons(1).Caption = "Actualizar"


Cells(fila2tpv + 2, 4) = "País"
Cells(fila2tpv + 3, 4) = "Tipo"
Cells(fila2tpv + 2, 5) = "todos"
Cells(fila2tpv + 3, 5) = "todos"




Sheets("hoja3").Cells(1, 5) = "Agregador"
Sheets("hoja3").Cells(2, 5) = "Gateway"
Sheets("hoja3").Cells(3, 5) = "todos"
Sheets("hoja3").Cells(1, 6) = "f"
Sheets("hoja3").Cells(2, 6) = "t"
Sheets("hoja3").Cells(3, 6) = "(All)"
Sheets("hoja3").Cells(largo2 + 1, 1) = "todos"
Sheets("hoja3").Cells(1, 7) = "Valores Absolutos"
Sheets("hoja3").Cells(2, 7) = "Valores Relativos"

  
     Range("e" & fila2tpv + 2).Select
    
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
         xlBetween, Formula1:="=Hoja3!$a$2:$a$" & Sheets("hoja3").Range("a1").End(xlDown).Row
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

   Range("e" & fila2tpv + 3).Select
    
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Hoja3!$e$1:$e$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

   Rows("1:" & fila2tpv).EntireRow.Hidden = True

   With Cells
           .Interior.Color = RGB(20, 67, 172)
           .Font.ThemeColor = xlThemeColorDark1

   End With
   
   Range("a" & fila2tpv + 4) = "Comportamiento MAF Transacciones con tarjeta de crédito:"
    Range("a" & ultimafila - 1) = "TPV Transacciones todos los medios de pago"
    
    
    With Rows(fila2tpv + 4)
           .Interior.Color = RGB(255, 255, 255)
           .Font.Color = RGB(0, 0, 0)
           .Font.Bold = True
    
   End With
   With Rows(fila2tpv + 5)
           .Interior.Color = RGB(0, 0, 0)
           '.Font.Color = RGB(0, 0, 0)
           .Font.Bold = True
    
   End With
 
   With Cells(fila2tpv + 5, 1)
           '.Interior.Color = RGB(0, 0, 0)
           .Font.Color = RGB(0, 0, 0)
           .Font.Bold = True
    
   End With
   
 
  
 
 
 
   With Range(Rows(ultimafila - 1), Rows(ultimafila))
           .Interior.Color = RGB(255, 255, 255)
           .Font.Color = RGB(0, 0, 0)
           .Font.Bold = True
    
   End With
    With Rows(ultimafila + 3)
           .Interior.Color = RGB(0, 0, 0)
           '.Font.Color = RGB(0, 0, 0)
           .Font.Bold = True
    
   End With
   With Cells(ultimafila + 3, 1)
           '.Interior.Color = RGB(0, 0, 0)
           .Font.Color = RGB(0, 0, 0)
           .Font.Bold = True
    
   End With
   
   Range(Rows(fila2tpv + 6), Rows(penultimafila)).NumberFormat = "#,##0"
   Rows(ultimafila + 2).NumberFormat = "#,##0"
   
    With Range(Columns(2), Columns(100))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
     Range("d" & fila2tpv + 4).Select
    
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
         xlBetween, Formula1:="=Hoja3!$g$1:$g$" & Sheets("hoja3").Range("g1").End(xlDown).Row
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("d" & fila2tpv + 4) = "Valores Absolutos"
    
    Rows(penultimafila + 2 & ":" & ultimafila - espacionabsolutorelativo).EntireRow.Hidden = True
      
    Columns(1).ColumnWidth = 80
    
    final2 = Cells(1048576, 1).End(xlUp).Row + 1
    referencia2 = Cells(final2 - 2, 2).Address
    Cells(final2, 1) = "Ticket Promedio"
    
    
    
    
    'Calculo ticket promedio
    For i = 2 To fila2rechazo - fila1rechazo + 4
    
    Cells(final2, i).FormulaLocal = "=+si(esnumero(" & "desref(" & referencia2 & ";" & 0 & ";" & i - 2 & ") / desref(" & referencia2 & ";" & 1 & ";" & i - 2 & ")" & ");" & "desref(" & referencia2 & ";" & 0 & ";" & i - 2 & ") / desref(" & referencia2 & ";" & 1 & ";" & i - 2 & ")" & ";0)"
    
    Next
     Rows(final2).NumberFormat = "$ #,##0.00"
     
 
 final3 = final2 + 5 ' inicio de resumen comercios
 
 
 comilla = """"
 
 Cells(final3, 1) = "Top 20 Comercios por TPV y Transacciones"
 
 ' formato letras a la izquierda usuarios.
 Union(Range(Cells(final3 + 2, 2), Cells(final3 + 23, 2)), Range(Cells(final3 + 2, 6), Cells(final3 + 23, 6))).HorizontalAlignment = xlLeft
 
 Range(Rows(final3), Rows(final3 + 1)).Font.Bold = True
 
 ' participacion comercio por tpv
 For j = 1 To 20
   For i = 1 To 3
       
        
            Cells(final3 + j, i + 1).FormulaLocal = "=si(y(" & Cells(4 + j, columna1participaciontpv + i - 1).Address & " <>" & comilla & comilla & ";" & Cells(4 + j, columna1participaciontpv + i - 1).Address & " <>" & comilla & "(en blanco)" & comilla & ");" & Cells(4 + j, columna1participaciontpv + i - 1).Address & ";" & comilla & comilla & ")"
        
       
   Next
 Next
 
  
       'valores tpv de todos los demás comercios
       
       
       
suma_columna2 = "suma(" & Cells(final3 + 1, 3).Address & ":" & Cells(final3 + j - 1, 3).Address & ")" 'suma columna 2
suma_columna3 = "suma(" & Cells(final3 + 1, 4).Address & ":" & Cells(final3 + j - 1, 4).Address & ")"  'suma columna 3

cadena1 = "(1-" & suma_columna3 & ")/" & suma_columna3 & "*" & suma_columna2
cadena2 = "(1-" & suma_columna3 & ")"
       
        
            Cells(final3 + j, 1 + 1).FormulaLocal = "=si(y(" & Cells(4 + j, columna1participaciontpv + 1 - 1).Address & " <>" & comilla & comilla & ";" & Cells(4 + j, columna1participaciontpv + 1 - 1).Address & " <>" & comilla & "(en blanco)" & comilla & ");" & comilla & "Otros Comercios" & comilla & ";" & comilla & comilla & ")"
            Cells(final3 + j, 2 + 1).FormulaLocal = "=si(y(" & Cells(4 + j, columna1participaciontpv + 2 - 1).Address & " <>" & comilla & comilla & ";" & Cells(4 + j, columna1participaciontpv + 2 - 1).Address & " <>" & comilla & "(en blanco)" & comilla & ");" & cadena1 & ";" & comilla & comilla & ")"
            Cells(final3 + j, 3 + 1).FormulaLocal = "=si(y(" & Cells(4 + j, columna1participaciontpv + 3 - 1).Address & " <>" & comilla & comilla & ";" & Cells(4 + j, columna1participaciontpv + 3 - 1).Address & " <>" & comilla & "(en blanco)" & comilla & ");" & cadena2 & ";" & comilla & comilla & ")"
       
  
 
 
 
 
  Range(Cells(final3 + 1, 3), Cells(final3 + 21, 3)).NumberFormat = "$ #,##0.00,," '$ #,##000..
  Range(Cells(final3 + 1, 4), Cells(final3 + 21, 4)).NumberFormat = "#0.00%"
  
  'formato porcentajes variaciones promedio
  
  Intersect(Columns(fila2rechazo - fila1rechazo + 5), Range(Rows(1), Rows(ultimafila + 3))).NumberFormat = "0.0%"
 
 'participacion comercios por tx
  For j = 1 To 20
   For i = 1 To 3
       
        
            Cells(final3 + j, i + 5).FormulaLocal = "=si(y(" & Cells(4 + j, columna1participaciontx + i - 1).Address & " <>" & comilla & comilla & ";" & Cells(4 + j, columna1participaciontx + i - 1).Address & " <>" & comilla & "(en blanco)" & comilla & ");" & Cells(4 + j, columna1participaciontx + i - 1).Address & ";" & comilla & comilla & ")"
        
       
   Next
 Next
 
  Range(Cells(final3 + 1, 7), Cells(final3 + 20, 7)).NumberFormat = "#,##0"
  Range(Cells(final3 + 1, 8), Cells(final3 + 20, 8)).NumberFormat = "#0.00%"
     
ActiveSheet.Shapes.AddPicture "D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\PayU.png", True, True, 0, 0, 120, 45
'ActiveSheet.Shapes.AddPicture "C:\Users\user\Documents\Javier\Pagos Online\informe jose\Abril 27\PayU.png", True, True, 0, 0, 120, 45
'insertar dias de la semana
  
   'Rows(fila2tpv + 5).Insert Shift:=xlUp, CopyOrigin:=1
   Range(Cells(fila2tpv + 5, 1), Cells(Cells(1048576, 2).End(xlUp).Row, 1000)).Cut (Cells(fila2tpv + 6, 1))
   
   
   
   For i = 2 To fila2tpv - fila1tpv + 1
        Cells(fila2tpv + 5, i) = diasemana(Cells(fila2tpv + 6, i))
        
    
   Next
   
   'Graficas varias
   
   
     'grafica tpv
        'Dim grafico1 As Chart
        
         Cells(1048576, 2).End(xlUp).Select
         Set grafico1 = ActiveSheet.Shapes.AddChart
         grafico1.Chart.ChartType = xlPie
         grafico1.Chart.ChartTitle.Text = "Participación TPV"
         'grafico.Chart.SetSourceData Source:=Cells(filainicialgraficas, 2).CurrentRegion
         grafico1.Width = Columns(1).Width
         grafico1.Left = 1
         Altura_grafico_tpv = Range(Rows(1), Rows(Cells(1048576, 1).End(xlUp).Row + 2)).Height
         grafico1.Top = Altura_grafico_tpv
         grafico1.Height = Range(Rows(1), Rows(Cells(1048576, 1).End(xlUp).Row + 20)).Height - Range(Rows(1), Rows(Cells(1048576, 1).End(xlUp).Row + 2)).Height
         grafico1.Line.Visible = msoFalse
         grafico1.Chart.Legend.Font.Size = 6.8
         grafico1.Chart.Legend.Top = 8.575
         grafico1.Chart.Legend.Width = 170.78
         grafico1.Chart.Legend.Height = 256.285
         grafico1.Chart.Legend.Left = 249.663
         grafico1.Chart.ChartTitle.Left = 69.6
         
        
           
           
         filainicialgraficas = Cells(1048576, 2).End(xlUp).Row + 6 ' no incluye la grafica tpv
          
          
          
          'grafica rechazos maf y total
          Cells(filainicialgraficas, 1) = "Analisis Rechazos"
          Cells(filainicialgraficas, 1).Font.Bold = True
    
          
       ' Dim grafico2 As Chart
           
           
   
        Set grafico2 = ActiveSheet.Shapes.AddChart
        grafico2.Chart.ChartType = xlLine
        grafico2.Chart.SetSourceData Source:=Range("rechazos_maf")
        grafico2.Chart.HasTitle = True
        grafico2.Chart.ChartTitle.Text = "Rechazos MAF"
        grafico2.Left = 0
        grafico2.Width = Columns(1).Width
        grafico2.Height = Range(Rows(filainicialgraficas), Rows(filainicialgraficas + 15)).Height
        grafico2.Top = Range(Rows(1), Rows(filainicialgraficas + 1)).Height
        grafico2.Chart.ShowAllFieldButtons = False
        grafico2.Line.Visible = msoFalse
        grafico2.Chart.HasLegend = True
        grafico2.Chart.SetElement (msoElementLegendBottom)
        grafico2.Chart.Axes(xlCategory).TickLabels.Orientation = xlUpward
        grafico2.Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00%"
         grafico2.Chart.Axes(xlValue).MajorGridlines.Delete
        
        
        
        'grafico comportamientomaf
          Set grafico3 = ActiveSheet.Shapes.AddChart
        grafico3.Chart.ChartType = xlPie
        grafico3.Chart.SetSourceData Source:=Range("rechazos_total")
        grafico3.Chart.ChartTitle.Text = "Comportamiento MAF"
        grafico3.Left = Range(Columns(1), Columns(5)).Width + 50
        grafico3.Width = Columns(1).Width
        grafico3.Height = Range(Rows(filainicialgraficas), Rows(filainicialgraficas + 15)).Height
        grafico3.Top = Range(Rows(1), Rows(filainicialgraficas + 1)).Height
        grafico3.Chart.ShowAllFieldButtons = False
        grafico3.Line.Visible = msoFalse
        grafico3.Chart.HasLegend = True
        grafico3.Chart.SetElement (msoElementLegendBottom)
        grafico3.Chart.PlotBy = xlRows
        
        
        
        
        filainicioanalisisrechazo = filainicialgraficas + 22
                
        Cells(filainicioanalisisrechazo, 1) = "Comercios Objetivo Rechazo MAF"
        Cells(filainicioanalisisrechazo, 1).Font.Bold = True
        
        ' participacion comercio rechazo maf
 For j = 1 To 30
   For i = 1 To 4
       
        
            Cells(filainicioanalisisrechazo + j, i).FormulaLocal = "=si(y(" & Cells(4 + j, columna1comerciorechazomaf + i - 1).Address & " <>" & comilla & comilla & ";" & Cells(4 + j, columna1comerciorechazomaf + i - 1).Address & " <>" & comilla & "(en blanco)" & comilla & ");" & Cells(4 + j, columna1comerciorechazomaf + i - 1).Address & ";" & comilla & comilla & ")"
        If i > 2 Then
        
        Cells(filainicioanalisisrechazo + j, i).NumberFormat = "0.0 %"
        
        End If
       
   Next
 Next
        
        Cells(filainicioanalisisrechazo + 1, 1) = "Código Comercio"
        
        
        
        'ActiveSheet.Shapes.AddPicture "C:\Users\user\Documents\Javier\Pagos Online\informe jose\Abril 27\target.jpg", True, True, 10, Range(Rows(1), Rows(filainicioanalisisrechazo + 3)).Height, Columns(1).Width, Range(Rows(filainicioanalisisrechazo), Rows(filainicioanalisisrechazo + 16)).Height
        ActiveSheet.Shapes.AddPicture "D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\target.jpg", True, True, Range(Columns(1), Columns(5)).Width, Range(Rows(1), Rows(filainicioanalisisrechazo + 3)).Height, 200, Range(Rows(filainicioanalisisrechazo), Rows(filainicioanalisisrechazo + 16)).Height
        
        
        
        
        
        'alertas comercios rechazo
        
        filainicialalertarechazo = Cells(1048576, 2).End(xlUp).Row + 6
        Cells(filainicialalertarechazo, 1) = "Alerta Diaria Comercios Rechazo MAF"
        Cells(filainicialalertarechazo, 1).Font.Bold = True
        
            'dinammica alertas  rechazos maf
        
                
            'ActiveSheet.Shapes.AddPicture "C:\Users\user\Documents\Javier\Pagos Online\informe jose\Abril 27\warning.jpg", True, True, 30, Range(Rows(1), Rows(filainicialalertarechazo + 3)).Height, Columns(1).Width - 70, Range(Rows(filainicialalertarechazo), Rows(filainicialalertarechazo + 10)).Height
             ActiveSheet.Shapes.AddPicture "D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\warning.jpg", True, True, 30, Range(Rows(1), Rows(filainicialalertarechazo + 3)).Height, 200, Range(Rows(filainicialalertarechazo), Rows(filainicialalertarechazo + 10)).Height
            Set dinamica8 = ActiveSheet.PivotTables.Add( _
                PivotCache:=Cache, _
                TableDestination:=Cells(filainicialalertarechazo + 1, 2))
                
               
            With dinamica8
                 .CalculatedFields.Add "Alerta", "if(and('Porcentaje Rechazo MAF'>.25,'Rechazo MAF total '>5),1,0)", True
                 
                .ColumnGrand = False
                .RowGrand = False
            
                
                .PivotFields("codigo_comercio").Orientation = xlRowField
                
                With .PivotFields("Rechazo MAF total ")
                 .Orientation = xlDataField
                 .Function = xlSum
                 .NumberFormat = " #,##0 "
                 .Caption = "Cantidad Rechazo MAF"
                 
                 
                End With
                
                
                
                 With .PivotFields("Porcentaje Rechazo MAF")
                 .Orientation = xlDataField
                 .Function = xlSum
                 .NumberFormat = " 0.0 %"
                 .Caption = "Rechazo Comercio"
              
                 
                 
                End With
                
                
                 With .PivotFields("Alerta")
                 .Orientation = xlDataField
                 .Function = xlSum
                 .Caption = "Alertas"
              
                 
                 
                End With
                
                 dinamica8.AllowMultipleFilters = True
                .PivotFields("pais").Orientation = xlPageField
                .PivotFields("tiene_c_personalizados").Orientation = xlPageField
                .PivotFields("rango").Orientation = xlPageField
                .PivotFields("rango").CurrentPage = comillas & Cells(fila2tpv + 6, fila2tpv - fila1tpv + 1) & comillas '"25-04-2013"
                .PivotFields("codigo_comercio").AutoSort xlDescending, "Cantidad Rechazo MAF", dinamica8.PivotColumnAxis.PivotLines(1), 1
                 
                
                .PivotFields("codigo_comercio").PivotFilters.Add Type:=xlValueEquals, DataField:=dinamica8.PivotFields("Alertas"), Value1:=1
                
               .HasAutoFormat = False
               .DisplayFieldCaptions = False
             End With
             
             
             Range(Rows(filainicialalertarechazo - 1), Rows(filainicialalertarechazo - 3)).Hidden = True
                
                'Cells(5, columna1comerciorechazomaf).CurrentRegion.Name = "comercios_rechazo_maf"
        
        
        'fin alertas  rechazos maf
        
        
        
        
        'inicio graficas Validacion
        
        filainicioanalisisvalidacion = Application.WorksheetFunction.Max(filainicialalertarechazo + 21, Cells(1048576, 2).End(xlUp).Row + 7)
        
        
        Cells(filainicioanalisisvalidacion, 1) = "Analisis Validacion"
        Cells(filainicioanalisisvalidacion, 1).Font.Bold = True
        
        
        


    Set grafico4 = ActiveSheet.Shapes.AddChart
        grafico4.Chart.ChartType = xlLine
        grafico4.Chart.SetSourceData Source:=Range("validacion_maf")
        grafico4.Chart.HasTitle = True
        grafico4.Chart.ChartTitle.Text = "Validación"
        grafico4.Left = 0
        grafico4.Width = Columns(1).Width
        grafico4.Height = Range(Rows(filainicioanalisisvalidacion), Rows(filainicioanalisisvalidacion + 15)).Height
        grafico4.Top = Range(Rows(1), Rows(filainicioanalisisvalidacion + 1)).Height
        grafico4.Chart.ShowAllFieldButtons = False
        grafico4.Line.Visible = msoFalse
        grafico4.Chart.HasLegend = True
        grafico4.Chart.SetElement (msoElementLegendBottom)
        grafico4.Chart.Axes(xlCategory).TickLabels.Orientation = xlUpward
        grafico4.Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00%"
        grafico4.Chart.Axes(xlValue).MajorGridlines.Delete
        
        'fin graficas validacion
        
        
        filainiciocomerciosvalidacion = Cells(1048576, 1).End(xlUp).Row + 20
        Cells(filainiciocomerciosvalidacion, 1) = "Comercios Objetivo Validación"
        Cells(filainiciocomerciosvalidacion, 1).Font.Bold = True
        filainiciocomerciosvalidacion = filainiciocomerciosvalidacion + 1
        
        'inicio comercios problema  validacion
        
        
       For j = 1 To 30
          For i = 1 To 4
              
               
                   Cells(filainiciocomerciosvalidacion + j, i).FormulaLocal = "=si(y(" & Cells(4 + j, columna1comerciovalidacionmaf + i - 1).Address & " <>" & comilla & comilla & ";" & Cells(4 + j, columna1comerciovalidacionmaf + i - 1).Address & " <>" & comilla & "(en blanco)" & comilla & ");" & Cells(4 + j, columna1comerciovalidacionmaf + i - 1).Address & ";" & comilla & comilla & ")"
               If i > 2 Then
               
               Cells(filainiciocomerciosvalidacion + j, i).NumberFormat = "0.0 %"
               
               End If
              
          Next
        Next
        
        Cells(filainiciocomerciosvalidacion + 1, 1) = "Código Comercio"
        
        
        
        'ActiveSheet.Shapes.AddPicture "C:\Users\user\Documents\Javier\Pagos Online\informe jose\Abril 27\target.jpg", True, True, 10, Range(Rows(1), Rows(filainicioanalisisrechazo + 3)).Height, Columns(1).Width, Range(Rows(filainicioanalisisrechazo), Rows(filainicioanalisisrechazo + 16)).Height
        ActiveSheet.Shapes.AddPicture "D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\target.jpg", True, True, Range(Columns(1), Columns(5)).Width, Range(Rows(1), Rows(filainiciocomerciosvalidacion + 3)).Height, 200, Range(Rows(filainiciocomerciosvalidacion), Rows(filainiciocomerciosvalidacion + 16)).Height
        
        
        
        
        
        
        'fin comercios problema validacion
   'fin graficas varias
   
   
   'alertas comercios rechazo
        
        filainicialalertavalidacion = Cells(1048576, 2).End(xlUp).Row + 6
        Cells(filainicialalertavalidacion, 1) = "Alerta Diaria Comercios Validación MAF"
        Cells(filainicialalertavalidacion, 1).Font.Bold = True
        
            'dinammica alertas  validacion maf
        
                
            'ActiveSheet.Shapes.AddPicture "C:\Users\user\Documents\Javier\Pagos Online\informe jose\Abril 27\warning.jpg", True, True, 30, Range(Rows(1), Rows(filainicialalertarechazo + 3)).Height, Columns(1).Width - 70, Range(Rows(filainicialalertarechazo), Rows(filainicialalertarechazo + 10)).Height
             ActiveSheet.Shapes.AddPicture "D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\warning.jpg", True, True, 30, Range(Rows(1), Rows(filainicialalertavalidacion + 3)).Height, 200, Range(Rows(filainicialalertavalidacion), Rows(filainicialalertavalidacion + 10)).Height
            Set dinamica9 = ActiveSheet.PivotTables.Add( _
                PivotCache:=Cache, _
                TableDestination:=Cells(filainicialalertavalidacion + 1, 2))
                
               
            With dinamica9
                 .CalculatedFields.Add "AlertaV", "if(and('Porcentaje Validacion '>.5,'Validacion M'>5),1,0)", True
                 
                .ColumnGrand = False
                .RowGrand = False
            
                
                .PivotFields("codigo_comercio").Orientation = xlRowField
                
                With .PivotFields("Validacion M")
                 .Orientation = xlDataField
                 .Function = xlSum
                 .NumberFormat = " #,##0 "
                 .Caption = "Cantidad Validar MAF "
                 
                 
                End With
                
                
                
                 With .PivotFields("Porcentaje Validacion ")
                 .Orientation = xlDataField
                 .Function = xlSum
                 .NumberFormat = " 0.0 %"
                 .Caption = "Validación Comercio"
              
                 
                 
                End With
                
                
                 With .PivotFields("AlertaV")
                 .Orientation = xlDataField
                 .Function = xlSum
                 .Caption = "Alertas v"
              
                 
                 
                End With
                
                 dinamica9.AllowMultipleFilters = True
                .PivotFields("pais").Orientation = xlPageField
                .PivotFields("tiene_c_personalizados").Orientation = xlPageField
                .PivotFields("rango").Orientation = xlPageField
                .PivotFields("rango").CurrentPage = comillas & Cells(fila2tpv + 6, fila2tpv - fila1tpv + 1) & comillas '"25-04-2013"
                .PivotFields("codigo_comercio").AutoSort xlDescending, "Cantidad Validar MAF", dinamica9.PivotColumnAxis.PivotLines(1), 1
                 
                
                .PivotFields("codigo_comercio").PivotFilters.Add Type:=xlValueEquals, DataField:=dinamica9.PivotFields("Alertas v"), Value1:=1
                
               .HasAutoFormat = False
               .DisplayFieldCaptions = False
             End With
             
             
             Range(Rows(filainicialalertavalidacion - 1), Rows(filainicialalertavalidacion - 3)).Hidden = True
                
                'Cells(5, columna1comerciorechazomaf).CurrentRegion.Name = "comercios_rechazo_maf"
        
        
        'fin alertas  validacion maf
        
   
  filainicialgraficasanterior = Application.WorksheetFunction.Max(Cells(1048576, 2).End(xlUp).Row + 6, filainicialalertavalidacion + 16)
  
  
  Cells(filainicialgraficasanterior, 1) = "Comportamiento TPV(Millones)"
  Cells(filainicialgraficasanterior, 1).Font.Bold = True
   
   filainicialcrecimientotop10 = filainicialgraficasanterior + 22
   
  Cells(filainicialcrecimientotop10, 1) = "Crecimiento TOP 10"
  Cells(filainicialcrecimientotop10, 1).Font.Bold = True
  
  
 filainicialtablas_tpv_grande = filainicialcrecimientotop10 + 26
 
   Cells(filainicialtablas_tpv_grande, 1) = "TOP 10 Comercios con Mayor Crecimiento absoluto"
   Cells(filainicialtablas_tpv_grande, 1).Font.Bold = True
   
   
   
    filainicialtablas_tpv_negativo = filainicialtablas_tpv_grande + 26
 
   Cells(filainicialtablas_tpv_negativo, 1) = "TOP 10 Comercios con Mayor Disminución Absoluta"
   Cells(filainicialtablas_tpv_negativo, 1).Font.Bold = True
   
    filainicialtablas_comercios_activos = filainicialtablas_tpv_negativo + 26
 
   Cells(filainicialtablas_comercios_activos, 1) = "Comercios Activos por Pronóstico de TPV vs TPV Anterior"
   Cells(filainicialtablas_comercios_activos, 1).Font.Bold = True
   
   
   
   
   
   ' grafico rechazos antes
   
   
       Set grafico2_anterior = ActiveSheet.Shapes.AddChart
        grafico2_anterior.Chart.ChartType = xlLine
        grafico2_anterior.Chart.SetSourceData Source:=Range("rechazos_maf_anterior")
        grafico2_anterior.Chart.HasTitle = True
        grafico2_anterior.Chart.ChartTitle.Text = "Rechazos MAF Mes Anterior"
        grafico2_anterior.Left = Range(Columns(1), Columns(2)).Width - 15
        grafico2_anterior.Width = Columns(1).Width
        grafico2_anterior.Height = Range(Rows(filainicialgraficas), Rows(filainicialgraficas + 15)).Height
        grafico2_anterior.Top = Range(Rows(1), Rows(filainicialgraficas + 1)).Height
        grafico2_anterior.Chart.ShowAllFieldButtons = False
        grafico2_anterior.Line.Visible = msoFalse
        grafico2_anterior.Chart.HasLegend = True
        grafico2_anterior.Chart.SetElement (msoElementLegendBottom)
        grafico2_anterior.Chart.Axes(xlCategory).TickLabels.Orientation = xlUpward
        grafico2_anterior.Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00%"
        grafico2_anterior.Chart.Axes(xlValue).MajorGridlines.Delete
        
   
   ' grafico validacion  antes
   
   
       Set grafico4_anterior = ActiveSheet.Shapes.AddChart
        grafico4_anterior.Chart.ChartType = xlLine
        grafico4_anterior.Chart.SetSourceData Source:=Range("validacion_maf_anterior")
        grafico4_anterior.Chart.HasTitle = True
        grafico4_anterior.Chart.ChartTitle.Text = "Validacion Mes Anterior"
        grafico4_anterior.Left = Range(Columns(1), Columns(2)).Width - 15
        grafico4_anterior.Width = Columns(1).Width
        grafico4_anterior.Height = Range(Rows(filainicialgraficas), Rows(filainicialgraficas + 15)).Height
        grafico4_anterior.Top = Range(Rows(1), Rows(filainicioanalisisvalidacion + 1)).Height
        grafico4_anterior.Chart.ShowAllFieldButtons = False
        grafico4_anterior.Line.Visible = msoFalse
        grafico4_anterior.Chart.HasLegend = True
        grafico4_anterior.Chart.SetElement (msoElementLegendBottom)
        grafico4_anterior.Chart.Axes(xlCategory).TickLabels.Orientation = xlUpward
        grafico4_anterior.Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00%"
        grafico4_anterior.Chart.Axes(xlValue).MajorGridlines.Delete
        
   
         
   
'grafico tpv actual dias
Range(Cells(ultimafila + 1, 2), Cells(ultimafila + 2, fila2rechazo - fila1rechazo + 1)).Name = "tpv_actual_dia"




  Set grafico_tpv_actual_dia = ActiveSheet.Shapes.AddChart
        grafico_tpv_actual_dia.Chart.ChartType = xlLine
        grafico_tpv_actual_dia.Chart.SetSourceData Source:=Range("tpv_actual_dia"), PlotBy:=xlRows
        grafico_tpv_actual_dia.Chart.HasTitle = True
        grafico_tpv_actual_dia.Chart.ChartTitle.Text = "TPV Mes Actual"
        grafico_tpv_actual_dia.Left = 0
        grafico_tpv_actual_dia.Width = Columns(1).Width
        grafico_tpv_actual_dia.Height = Range(Rows(filainicialgraficas), Rows(filainicialgraficas + 15)).Height
        grafico_tpv_actual_dia.Top = Range(Rows(1), Rows(filainicialgraficasanterior + 1)).Height
        'grafico_tpv_actual_dia.Chart.ShowAllFieldButtons = False
        grafico_tpv_actual_dia.Line.Visible = msoFalse
        grafico_tpv_actual_dia.Chart.HasLegend = False
        'grafico_tpv_actual_dia.Chart.SetElement (msoElementLegendBottom)
        grafico_tpv_actual_dia.Chart.Axes(xlCategory).TickLabels.Orientation = xlUpward
        grafico_tpv_actual_dia.Chart.Axes(xlValue).TickLabels.NumberFormat = "$ #,###,,"
        grafico_tpv_actual_dia.Chart.Axes(xlValue).MajorGridlines.Delete


'fin grafico tpv actual dias


 ' grafico tpv dia antes
   
   
       Set grafico_tpv_anterior_dia = ActiveSheet.Shapes.AddChart
        grafico_tpv_anterior_dia.Chart.ChartType = xlLine
        grafico_tpv_anterior_dia.Chart.SetSourceData Source:=Range("tpv_anterior_dia"), PlotBy:=xlColumns
        grafico_tpv_anterior_dia.Chart.HasTitle = True
        grafico_tpv_anterior_dia.Chart.ChartTitle.Text = "TPV Mes Anterior"
        grafico_tpv_anterior_dia.Left = Range(Columns(1), Columns(2)).Width - 15
        grafico_tpv_anterior_dia.Width = Columns(1).Width
        grafico_tpv_anterior_dia.Height = Range(Rows(filainicialgraficas), Rows(filainicialgraficas + 15)).Height
        grafico_tpv_anterior_dia.Top = Range(Rows(1), Rows(filainicialgraficasanterior + 1)).Height
        grafico_tpv_anterior_dia.Chart.ShowAllFieldButtons = False
        grafico_tpv_anterior_dia.Line.Visible = msoFalse
        grafico_tpv_anterior_dia.Chart.HasLegend = False
        'grafico_tpv_anterior_dia.Chart.SetElement (msoElementLegendBottom)
        grafico_tpv_anterior_dia.Chart.Axes(xlCategory).TickLabels.Orientation = xlUpward
        grafico_tpv_anterior_dia.Chart.Axes(xlValue).TickLabels.NumberFormat = "$ #,###,,"
        grafico_tpv_anterior_dia.Chart.Axes(xlValue).MajorGridlines.Delete



'fin grafico tpv dia antes
   
     
   
   
   
   
   
'inicio grafico tpv actual vs tpv anterior

   
   
    Cells(ultimafila + 1, fila2rechazo - fila1rechazo + 4) = "Mes Anterior"
    Cells(ultimafila + 1, fila2rechazo - fila1rechazo + 3) = "Mes actual"
     
    Range(Cells(ultimafila + 1, fila2rechazo - fila1rechazo + 3), Cells(ultimafila + 2, fila2rechazo - fila1rechazo + 4)).Name = "tpvactual_vs_tpvanterior"


            'color de mes actual
             With Cells(ultimafila, fila2rechazo - fila1rechazo + 3).Font
                    .Color = -16727809
                    .TintAndShade = 0
             End With
            
            'color de mes anterior
             With Cells(ultimafila, fila2rechazo - fila1rechazo + 4).Font
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
             End With
               
               Set grafico_tpvactual_tpvanterior = ActiveSheet.Shapes.AddChart
                   grafico_tpvactual_tpvanterior.Chart.ChartType = xlColumnClustered
                   grafico_tpvactual_tpvanterior.Chart.SetSourceData Source:=Range("tpvactual_vs_tpvanterior"), PlotBy:=xlRows
                    grafico_tpvactual_tpvanterior.Chart.HasTitle = True
                   grafico_tpvactual_tpvanterior.Chart.ChartTitle.Text = "Pronóstico TPV vs Mes Anterior"
                   
                    
                    grafico_tpvactual_tpvanterior.Left = Range(Columns(1), Columns(7)).Width + 50
                    grafico_tpvactual_tpvanterior.Width = Columns(1).Width
                    grafico_tpvactual_tpvanterior.Height = Range(Rows(filainicialgraficas), Rows(filainicialgraficas + 15)).Height
                    grafico_tpvactual_tpvanterior.Top = Range(Rows(1), Rows(filainicialgraficasanterior + 1)).Height
                    grafico_tpvactual_tpvanterior.Line.Visible = msoFalse
                    grafico_tpvactual_tpvanterior.Chart.HasLegend = False
                    grafico_tpvactual_tpvanterior.Chart.Axes(xlValue).MajorGridlines.Delete
                    grafico_tpvactual_tpvanterior.Chart.Axes(xlValue).MinimumScale = 0
               
   
'fin grafico tpv actual vs tpv anterior
   
   
'formato final
ActiveWorkbook.ShowPivotTableFieldList = False
Cells.Interior.Color = RGB(255, 255, 255)
Cells.Font.Color = RGB(0, 0, 0)
Columns.ColumnWidth = 20
Columns(1).ColumnWidth = 80


  'formato de las columnas pronostico y acumulado
   
  With Intersect(Range(Columns(fila2rechazo - fila1rechazo + 2), Columns(fila2rechazo - fila1rechazo + 3)), Range(Rows(fila2tpv + 6), Rows(final2 + 1)))
      .Interior.Color = RGB(255, 192, 0)
  End With
  
  
  'color del mes anterior
  
  
   With Intersect(Columns(fila2rechazo - fila1rechazo + 4), Range(Rows(fila2tpv + 6), Rows(final2 + 1)))
      .Interior.Color = RGB(220, 230, 241)
  End With
  
  
  
  'ocultar columnas finales
  
  If fila2rechazo - fila1rechazo + 5 > 12 Then
     Range(Columns(fila2rechazo - fila1rechazo + 6), Columns(Cells(1048576, fila2rechazo - fila1rechazo + 6).End(xlToRight).Column)).Hidden = True
   Else
     Range(Columns(12), Columns(Cells(1048576, fila2rechazo - fila1rechazo + 7).End(xlToRight).Column)).Hidden = True
  End If
  grafico4_anterior.Width = Columns(1).Width
  grafico2_anterior.Width = Columns(1).Width
  grafico_tpvactual_tpvanterior.Width = Columns(1).Width
End Sub


Sub macros_insertadas()




'insertar macro dentro del mismo libro

Const DQUOTE = """"
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim VBComp1 As VBIDE.VBComponent
        Dim VBcomp2 As VBIDE.VBComponent
         
        Set VBProj = libro.VBProject
       
        Set VBComp1 = VBProj.VBComponents.Add(vbext_ct_StdModule)
        VBComp1.Name = "NewModule"
        Set VBcomp2 = VBProj.VBComponents("ThisWorkbook")
        
        Set VBComp = VBProj.VBComponents("hoja6")
        Set CodeMod = VBComp.CodeModule
        Set codemod1 = VBComp1.CodeModule
        Set codemod2 = VBcomp2.CodeModule
     With codemod2
         LineNum = .CountOfLines + 1
            .InsertLines LineNum, "Private Sub Workbook_Open()"
        LineNum = LineNum + 1
            .InsertLines LineNum, "   'Sheets(" & DQUOTE & "Informe" & DQUOTE & ").Cells(" & fila2tpv & " + 2, 5) = " & DQUOTE & "todos" & DQUOTE
      LineNum = LineNum + 1
            .InsertLines LineNum, "   End Sub"
         

     
     
     End With
     
    With codemod1
            LineNum = .CountOfLines + 1
            .InsertLines LineNum, "Sub Comercio"
            LineNum = LineNum + 1
            .InsertLines LineNum, "   For i=1 to Sheets.count"
            LineNum = LineNum + 1
        .InsertLines LineNum, "   If Sheets(i).PivotTables.Count <> 0 Then"
            LineNum = LineNum + 1
           .InsertLines LineNum, "      Sheets(i).Range(" & DQUOTE & " b1" & DQUOTE & ") = Sheets(" & DQUOTE & "indice" & DQUOTE & ").Range(" & DQUOTE & "i25" & DQUOTE & ").Value"
            LineNum = LineNum + 1
           .InsertLines LineNum, "      Sheets(i).Range(" & DQUOTE & "b1" & DQUOTE & ").NumberFormat = " & DQUOTE & "General" & DQUOTE & ""
            LineNum = LineNum + 1
        .InsertLines LineNum, "      End If"
            LineNum = LineNum + 1
            .InsertLines LineNum, "      Next"
            LineNum = LineNum + 1

           .InsertLines LineNum, "End Sub"
     End With

With CodeMod

LineNum = .CountOfLines + 1
.InsertLines LineNum, "Private Sub Worksheet_Change(ByVal Target As Excel.Range)"
 LineNum = LineNum + 1
.InsertLines LineNum, "If Target.Address = Cells( " & fila2tpv + 2 & ", 5).Address Then"
 LineNum = LineNum + 1
      
      .InsertLines LineNum, "For i = 1 To ActiveSheet.PivotTables.Count"
 LineNum = LineNum + 1
.InsertLines LineNum, "   ActiveSheet.PivotTables(i).PivotFields(" & DQUOTE & "pais" & DQUOTE & ").ClearAllFilters"
 LineNum = LineNum + 1
            
.InsertLines LineNum, "   If ActiveSheet.Cells(" & fila2tpv + 2 & ", 5) <> " & DQUOTE & "todos" & DQUOTE & "Then"
 LineNum = LineNum + 1
.InsertLines LineNum, "      ActiveSheet.PivotTables(i).PivotFields(" & DQUOTE & "pais" & DQUOTE & ").CurrentPage = ActiveSheet.Cells(" & fila2tpv + 2 & ", 5).Value"
 LineNum = LineNum + 1
.InsertLines LineNum, "      Else"
 LineNum = LineNum + 1
.InsertLines LineNum, "      ActiveSheet.PivotTables(i).PivotFields(" & DQUOTE & "pais" & DQUOTE & ").  CurrentPage =" & DQUOTE & "(All)" & DQUOTE
 LineNum = LineNum + 1
.InsertLines LineNum, "    End If"
 LineNum = LineNum + 1
 
   
.InsertLines LineNum, " Next"
 LineNum = LineNum + 1
 
  
 
  ' fin comentarios nombres de comercios
  

          
.InsertLines LineNum, "ElseIf Target.Address = Cells(" & fila2tpv + 3 & ", 5).Address Then"
 LineNum = LineNum + 1
           
.InsertLines LineNum, " For i = 1 To ActiveSheet.PivotTables.Count"
 LineNum = LineNum + 1
        
.InsertLines LineNum, "     ActiveSheet.PivotTables(i).PivotFields(" & DQUOTE & "tiene_c_personalizados" & DQUOTE & ").ClearAllFilters"
 LineNum = LineNum + 1
.InsertLines LineNum, "     If ActiveSheet.Cells(" & fila2tpv + 3 & ", 5) <>" & DQUOTE & "todos" & DQUOTE & "Then"
 LineNum = LineNum + 1
.InsertLines LineNum, "     ActiveSheet.PivotTables(i).PivotFields(" & DQUOTE & "tiene_c_personalizados" & DQUOTE & ").CurrentPage = Application.WorksheetFunction.VLookup(ActiveSheet.Cells(" & fila2tpv + 3 & ", 5), Sheets(" & DQUOTE & "hoja3" & DQUOTE & ").Range(" & DQUOTE & "e1:f3" & DQUOTE & "), 2, 0)"
 LineNum = LineNum + 1
            
.InsertLines LineNum, "     Else"
 LineNum = LineNum + 1
            
.InsertLines LineNum, "    ActiveSheet.PivotTables(i).PivotFields(" & DQUOTE & "tiene_c_personalizados" & DQUOTE & ").CurrentPage =" & DQUOTE & "(All)" & DQUOTE
 LineNum = LineNum + 1
.InsertLines LineNum, "     End If"
 LineNum = LineNum + 1
            
            
.InsertLines LineNum, " Next "
 LineNum = LineNum + 1
 
 
 'fin comentarios nombres de comercios
 
                
.InsertLines LineNum, "ElseIf Target.Address = Cells(" & fila2tpv + 4 & ", 4).Address Then"
 LineNum = LineNum + 1
       
.InsertLines LineNum, "For i = 1 To ActiveSheet.PivotTables.Count"
 LineNum = LineNum + 1
       
      .InsertLines LineNum, "     If Cells(" & fila2tpv + 4 & ", 4) =" & DQUOTE & "Valores Absolutos" & DQUOTE & "Then"
 LineNum = LineNum + 1
            
.InsertLines LineNum, "      Range(Rows(" & fila2tpv + 7 & "), Rows(" & penultimafila + espacionabsolutorelativo & ")).EntireRow.Hidden = False"
 LineNum = LineNum + 1

.InsertLines LineNum, "      Range(Rows(" & penultimafila + espacionabsolutorelativo + 1 & "), Rows(" & ultimafila - 1 & ")).EntireRow.Hidden = True"
 LineNum = LineNum + 1
            
.InsertLines LineNum, "    Else"
 LineNum = LineNum + 1
.InsertLines LineNum, "      Range(Rows(" & fila2tpv + 7 & "), Rows(" & penultimafila + espacionabsolutorelativo & ")).EntireRow.Hidden = True"
 LineNum = LineNum + 1

.InsertLines LineNum, "       Range(Rows(" & penultimafila + espacionabsolutorelativo + 1 & "), Rows(" & ultimafila - 1 & ")).EntireRow.Hidden = False"
 LineNum = LineNum + 1
            
            
.InsertLines LineNum, "     End If"
 LineNum = LineNum + 1
          
.InsertLines LineNum, "Next"
 LineNum = LineNum + 1
                  
.InsertLines LineNum, "End If"
 LineNum = LineNum + 1
.InsertLines LineNum, "End Sub"
 LineNum = LineNum + 1

'mostrar detalle tabla dinamica

 
 .InsertLines LineNum, "Private Sub Worksheet_beforeDoubleClick(ByVal Target As Range, Cancel as Boolean)"
 LineNum = LineNum + 1
 
 .InsertLines LineNum, "nombrehoja=activesheet.name"
 LineNum = LineNum + 1
 
.InsertLines LineNum, "fila=Target.Row"
 LineNum = LineNum + 1
.InsertLines LineNum, "columna=Target.column"
 LineNum = LineNum + 1
.InsertLines LineNum, "if fila <= " & penultimafila + 1 & " and fila>= " & fila2tpv + 5 + 1 & " and columna>=2 and columna <= " & fila2rechazo - fila1rechazo + 1 & " then"
 LineNum = LineNum + 1
.InsertLines LineNum, "fila2=columna+4"
 LineNum = LineNum + 1
.InsertLines LineNum, "columna2= 2"
 LineNum = LineNum + 1
 
 .InsertLines LineNum, "on error goto prueba:"
 LineNum = LineNum + 1

 
.InsertLines LineNum, "cells(fila2,columna2).showdetail=true"
 LineNum = LineNum + 1


.InsertLines LineNum, "end if"
 LineNum = LineNum + 1


.InsertLines LineNum, "if fila <= " & final2 & " and fila>= " & final2 - 1 & " and columna>=2 and columna<= " & fila2rechazo - fila1rechazo + 1 & " then"
 LineNum = LineNum + 1

.InsertLines LineNum, "fila2=columna+ " & fila1tpv - 9
 LineNum = LineNum + 1
.InsertLines LineNum, "columna2= 2"
 LineNum = LineNum + 1
 
 .InsertLines LineNum, "on error goto prueba:"
 LineNum = LineNum + 1

 
.InsertLines LineNum, "cells(fila2,columna2).showdetail=true"
 LineNum = LineNum + 1


.InsertLines LineNum, "end if"
 LineNum = LineNum + 1

'detalles tablas dinámicas acumulados hasta el mes rechazo

.InsertLines LineNum, "if fila <= " & penultimafila + 1 & " and fila>= " & fila2tpv + 5 + 1 & " and columna = " & fila2rechazo - fila1rechazo + 2 & " then"
 LineNum = LineNum + 1








.InsertLines LineNum, "sheets(nombrehoja).pivottables(" & DQUOTE & "dinamica1" & DQUOTE & ").pivotfields(" & DQUOTE & "rango" & DQUOTE & ").orientation=xlhidden"
 LineNum = LineNum + 1

.InsertLines LineNum, "on error goto prueba:"
 LineNum = LineNum + 1


.InsertLines LineNum, "Cells(6, 1).ShowDetail = True"
 LineNum = LineNum + 1


.InsertLines LineNum, "sheets(nombrehoja).pivottables(" & DQUOTE & "dinamica1" & DQUOTE & ").pivotfields(" & DQUOTE & "rango" & DQUOTE & ").orientation=xlrows"
 LineNum = LineNum + 1



.InsertLines LineNum, "end if"
 LineNum = LineNum + 1

' detalles tablas dinámicas acumulados hasta el mes tpv


.InsertLines LineNum, "if fila <= " & final2 & " and fila>= " & final2 - 1 & " and columna= " & fila2rechazo - fila1rechazo + 2 & " then"
 LineNum = LineNum + 1



.InsertLines LineNum, "sheets(nombrehoja).pivottables(" & DQUOTE & "dinamica2" & DQUOTE & ").pivotfields(" & DQUOTE & "rango" & DQUOTE & ").orientation=xlhidden"
 LineNum = LineNum + 1

.InsertLines LineNum, "on error goto prueba:"
 LineNum = LineNum + 1


.InsertLines LineNum, "Cells(" & fila1tpv - 7 & ", 1).ShowDetail = True"
 LineNum = LineNum + 1


.InsertLines LineNum, "sheets(nombrehoja).pivottables(" & DQUOTE & "dinamica2" & DQUOTE & ").pivotfields(" & DQUOTE & "rango" & DQUOTE & ").orientation=xlrows"
 LineNum = LineNum + 1


.InsertLines LineNum, "end if"
 LineNum = LineNum + 1


.InsertLines LineNum, "Exit Sub"
 LineNum = LineNum + 1


.InsertLines LineNum, "prueba:"
 LineNum = LineNum + 1



.InsertLines LineNum, "MsgBox " & DQUOTE & "No hay datos para mostrar" & DQUOTE
 LineNum = LineNum + 1


.InsertLines LineNum, "Resume Next"
 LineNum = LineNum + 1






.InsertLines LineNum, "end sub"
 LineNum = LineNum + 1






'hasta acá detalle table dinámica



End With
  VBComp.Activate
End Sub





Sub grafico_tpv(numero_hoja)

'Cells(1048576, columna1_tpv_anterior).End(xlUp).End(xlUp).Row
'dinamica_tpv_anterior.PivotFields ("usuario_id")
bandera = 0
contador_grupos = 1

 On Error Resume Next
 
  Sheets(numero_hoja).PivotTables("dinamica_tpv_anterior").PivotFields("usuario_id2").Orientation = xlHidden
  Range(Sheets(numero_hoja).Cells(5, columna1_tpv_anterior), Sheets(numero_hoja).Cells(5, columna1_tpv_anterior).End(xlDown)).Ungroup
    
 
 On Error GoTo 0
    
    
    
fila_tpv_comercios_anterior_1 = 5




If Sheets(numero_hoja).Cells(5, columna1_tpv_anterior).End(xlDown).Row <> 1048576 Then

    fila_tpv_comercios_anterior_2 = Sheets(numero_hoja).Cells(5, columna1_tpv_anterior).End(xlDown).Row
   
Else
fila_tpv_comercios_anterior_2 = 5

End If


If fila_tpv_comercios_anterior_2 - fila_tpv_comercios_anterior_1 > 19 Then


  Range(Sheets(numero_hoja).Cells(fila_tpv_comercios_anterior_1 + 19, columna1_tpv_anterior), Sheets(numero_hoja).Cells(fila_tpv_comercios_anterior_2, columna1_tpv_anterior)).Group
  Sheets(numero_hoja).PivotTables("dinamica_tpv_anterior").PivotFields("usuario_id2").ShowDetail = False
  Sheets(numero_hoja).PivotTables("dinamica_tpv_anterior").PivotFields("usuario_id2").AutoSort xlAscending, "usuario_id2"
  
  Sheets(numero_hoja).PivotTables("dinamica_tpv_anterior").PivotFields("usuario_id2").PivotItems(Sheets(numero_hoja).PivotTables("dinamica_tpv_anterior").PivotFields("usuario_id2").PivotItems.Count).Caption = "Otros Comercios"
    'contador_grupos = contador_grupos + 1
    'Sheets(numero_hoja).PivotTables("dinamica_tpv_anterior").PivotFields("usuario_id2").PivotItems("Otros Comercios").Position = Sheets(numero_hoja).PivotTables("dinamica_tpv_anterior").PivotFields("usuario_id2").PivotItems.Count

  Call grafico_temporal_sub
  
  
  Else
  
   Call grafico_temporal_sub
  
  
  
  

End If

End Sub

Sub grafico_temporal_sub()





Set grafico_temporal = ActiveSheet.Shapes.AddChart
         grafico_temporal.Chart.ChartType = xlPie
         grafico_temporal.Chart.ChartTitle.Text = "Participación TPV Mes Anterior"
         grafico_temporal.Chart.SetSourceData Source:=Cells(4, columna1_tpv_anterior).CurrentRegion
         grafico_temporal.Width = Columns(1).Width
         grafico_temporal.Left = Range(Columns(1), Columns(fila2rechazo - fila1rechazo + 4)).Width
         grafico_temporal.Top = Altura_grafico_tpv
         grafico_temporal.Height = Range(Rows(1), Rows(Cells(1048576, 1).End(xlUp).Row + 20)).Height - Range(Rows(1), Rows(Cells(1048576, 1).End(xlUp).Row + 2)).Height
         grafico_temporal.Line.Visible = msoFalse
         grafico_temporal.Chart.ShowAllFieldButtons = False
         grafico_temporal.Chart.Legend.Font.Size = 6.8
         grafico_temporal.Chart.Legend.Top = 8.575
         grafico_temporal.Chart.Legend.Width = 180.782
         grafico_temporal.Chart.Legend.Height = 256.28
         grafico_temporal.Chart.Legend.Left = 249.663
         grafico_temporal.Chart.ChartTitle.Left = 9.6
         





End Sub

Sub resumen_graficos(hoja)
  
'
Dim orden(1 To 13)

orden(1) = 11
orden(2) = 13
orden(3) = 12
orden(4) = 14
orden(5) = 9
orden(6) = 7
orden(7) = 8

orden(8) = 1
orden(9) = 10

orden(10) = 2
orden(11) = 5
orden(12) = 4
orden(13) = 6


'para insertar banderas
 
 pais_temporal = Mid(hoja, 1, Application.WorksheetFunction.Find("-", hoja) - 1)
 imagen = "D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\Banderas\" & pais_temporal & ".jpg"
  
 
 For i = 1 To 6
 Sheets("Resumen Graficos").Cells(cantidadhojas * 17 + 3, 1 + 15 * (i - 1)) = hoja
 Sheets("Resumen Graficos").Cells(cantidadhojas * 17 + 3, 1 + 15 * (i - 1)).Font.ThemeColor = xlThemeColorLight2
 Sheets("Resumen Graficos").Cells(cantidadhojas * 17 + 3, 1 + 15 * (i - 1)).Font.Size = 14
 Sheets("Resumen Graficos").Cells(cantidadhojas * 17 + 3, 1 + 15 * (i - 1)).Font.Bold = True
 
 
 alto = Range(Sheets("Resumen Graficos").Cells(1, 1), Sheets("Resumen Graficos").Cells(cantidadhojas * 17 + 1, 1)).Height
 izquierda = Range(Sheets("Resumen Graficos").Columns(1), Sheets("Resumen Graficos").Columns(15 * (i - 1) + 2)).Width
 
 If pais_temporal <> "todos" Then
    Sheets("Resumen Graficos").Shapes.AddPicture imagen, True, True, izquierda, alto, 60, 35
 End If
 
 Next
 
 
 
 

  For i = 1 To UBound(orden)
      
     Sheets(hoja).ChartObjects(orden(i)).Select
     Sheets(hoja).ChartObjects(orden(i)).Copy
     
     Sheets("Resumen Graficos").Select
     Sheets("Resumen Graficos").Paste
     
     Sheets("Resumen Graficos").ChartObjects(contador_graficos).Left = 10 + 420 * (i - 1)
     Sheets("Resumen Graficos").ChartObjects(contador_graficos).Top = Sheets("Resumen Graficos").Range(Rows(1), Rows(3)).Height + (Sheets("Resumen Graficos").Range(Rows(1), Rows(3)).Height + Sheets("Resumen Graficos").Range(Rows(1), Rows(14)).Height) * cantidadhojas
     Sheets("Resumen Graficos").ChartObjects(contador_graficos).Width = 420
     Sheets("Resumen Graficos").ChartObjects(contador_graficos).Height = Sheets("Resumen Graficos").Range(Rows(1), Rows(14)).Height
     
     
     contador_graficos = contador_graficos + 1
  Next

cantidadhojas = cantidadhojas + 1

Sheets(hoja).Select

End Sub








Sub tamaño_eje_graficas(hoja1, hoja2, indice_grafico1, indice_grafico2, uno_total)

'
'Sheets(3).ChartObjects(2).Select

maximo_grafico1 = Sheets(hoja1).ChartObjects(indice_grafico1).Chart.Axes(xlValue).MaximumScale

maximo_grafico2 = Sheets(hoja2).ChartObjects(indice_grafico2).Chart.Axes(xlValue).MaximumScale

If uno_total = "uno" Then
maximo = Application.WorksheetFunction.Min(Application.WorksheetFunction.Max(maximo_grafico1, maximo_grafico2), 1)
Else

maximo = Application.WorksheetFunction.Max(maximo_grafico1, maximo_grafico2)


End If


Sheets(hoja2).ChartObjects(indice_grafico2).Chart.Axes(xlValue).MaximumScale = maximo

Sheets(hoja1).ChartObjects(indice_grafico1).Chart.Axes(xlValue).MaximumScale = maximo



'MsgBox Sheets(3).ChartObjects(1).Name



End Sub



Function mayor_fecha(hoja, campo_fecha)
Dim rango As Range
Dim a As Range
valor = campo_fecha
'Set rango = Range("d2:d" & Cells(2, 4).End(xlDown).Row)
columna = buscar(valor, hoja)
Set rango = Range(Cells(2, columna), Cells(2, columna).End(xlDown))
i = 1
mayor = 0
For Each a In rango
    If mayor < CDate(a) Then
     mayor = CDate(a)
    End If

 i = i + 1
Next

mayor_fecha = mayor


End Function


Sub prueba()

MsgBox mayor_fecha("RapidMiner Data", "rango")

End Sub
Sub nombres_tpv(hoja)
' crea en la columna la columna concatenada del usuario_id y su respectivo nombre
Dim rango1 As Range

Dim rango2 As Range
Dim vector1()
Dim vector2()
libro.Sheets(hoja).Select


numero_filas_nombre_tpv = libro.Sheets(hoja).Range("a1").CurrentRegion.Rows.Count

Set rango1 = Range(Cells(2, buscar("usuario_id", hoja)), Cells(numero_filas_nombre_tpv, buscar("usuario_id", hoja)))
Set rango2 = Range(Cells(2, buscar("nombres", hoja)), Cells(numero_filas_nombre_tpv, buscar("nombres", hoja)))

ReDim vector1(1 To numero_filas_nombre_tpv - 1)
ReDim vector2(1 To numero_filas_nombre_tpv - 1)

vector1 = Application.WorksheetFunction.Transpose(rango1)
vector2 = Application.WorksheetFunction.Transpose(rango2)






For i = 1 To numero_filas_nombre_tpv - 1
   
   vector1(i) = Mid(vector1(i) & "-" & vector2(i), 1, 40)


Next


rango1 = Application.WorksheetFunction.Transpose(vector1)



End Sub



Sub valores_tpv(nombre_hoja)
'actualiza los valores de tpv_anterior y el pronostico de tpv

dias_mes = DatePart("d", Application.WorksheetFunction.EoMonth(fecha_guia, 0))
dias_pasados = DatePart("d", fecha_guia)

libro.Sheets(nombre_hoja).Select

Dim rango_temporal As Range
Dim rango_usuario As Range
filas_usuario = Range("a1").End(xlDown).Row
columnas = Range("a1").End(xlToRight).Column

Set rango_usuario = Range(Cells(2, 1), Cells(filas_usuario, 1))

ReDim vector_usuario(1 To (filas_usuario - 1))
ReDim vector_tpv_anterior(1 To (filas_usuario - 1))
ReDim vector_tpv_actual(1 To (filas_usuario - 1))
ReDim vector_diferencia(1 To (filas_usuario - 1))
ReDim vector_porcentaje_crecimiento(1 To (filas_usuario - 1))

vector_usuario = Application.WorksheetFunction.Transpose(rango_usuario)

For i = 1 To (filas_usuario - 1)


     On Error GoTo vtpval:
    
      vector_tpv_actual(i) = Application.WorksheetFunction.VLookup(vector_usuario(i), Range("tpv_actual"), 2, 0) * multiplicador_dias_no_laboral + Application.WorksheetFunction.VLookup(vector_usuario(i), Range("tpv_actual"), 3, 0) * multiplicador_dias_laboral
     On Error GoTo 0
     
     On Error GoTo vtpvar:
      vector_tpv_anterior(i) = Application.WorksheetFunction.VLookup(vector_usuario(i), Range("tpv_anterior"), 2, 0) + Application.WorksheetFunction.VLookup(vector_usuario(i), Range("tpv_anterior"), 3, 0)
     On Error GoTo 0
       
       vector_diferencia(i) = vector_tpv_actual(i) - vector_tpv_anterior(i)

    On Error GoTo vpc:
    vector_porcentaje_crecimiento(i) = vector_tpv_actual(i) / vector_tpv_anterior(i) - 1
    On Error GoTo 0
Next

Cells(1, columnas + 2) = "Tpv_Pronosticado"
Set rango_temporal = Range(Cells(2, columnas + 2), Cells(filas_usuario, columnas + 2))
   rango_temporal.NumberFormat = "$ #,##0.0,,"
 rango_temporal = Application.WorksheetFunction.Transpose(vector_tpv_actual)


Cells(1, columnas + 1) = "Tpv_Anterior"
Set rango_temporal = Range(Cells(2, columnas + 1), Cells(filas_usuario, columnas + 1))
 rango_temporal = Application.WorksheetFunction.Transpose(vector_tpv_anterior)
  rango_temporal.NumberFormat = "$ #,##0.0,,"

Cells(1, columnas + 3) = "Diferencia"
Set rango_temporal = Range(Cells(2, columnas + 3), Cells(filas_usuario, columnas + 3))
 rango_temporal = Application.WorksheetFunction.Transpose(vector_diferencia)
  rango_temporal.NumberFormat = "$ #,##0.0,,"
  
  Cells(1, columnas + 4) = "Porcentaje"
Set rango_temporal = Range(Cells(2, columnas + 4), Cells(filas_usuario, columnas + 4))
 rango_temporal = Application.WorksheetFunction.Transpose(vector_porcentaje_crecimiento)
  rango_temporal.NumberFormat = "0.0 %"
  
  
  
 columna_filtro = buscar("Tpv_Anterior", "hoja1")
   
   
  ' On Error Resume Next
    'ActiveSheet.AutoFilter.Sort.SortFields.Clear
  ' On Error GoTo 0
   
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Cells( _
        1, columna_filtro), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  
  
  

Exit Sub




vtpval:
vector_tpv_actual(i) = 0
Resume Next

vtpvar:
vector_tpv_anterior(i) = 0
Resume Next


vpc:
vector_porcentaje_crecimiento(i) = "Nuevo"
Resume Next







End Sub


Sub tablas_areas(pais, tipo)
    
    'libro.Sheets("Hoja7").Select
  
    
    Call filtro_pais_tipo(pais, tipo)
    Call ordernar_hoja_1("Tpv_Anterior", 2)
    
    
    Sheets("hoja1").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    ActiveSheet.Paste
    columnas_temporal = Cells(filainicialcrecimientotop10 + 1, 2).CurrentRegion.Columns.Count - 1
    filas_temporal = Cells(filainicialcrecimientotop10 + 1, 2).CurrentRegion.Rows.Count + filainicialcrecimientotop10 - 1
    
    tpv_total_mes_anterior = Application.WorksheetFunction.Sum(Range(Cells(filainicialcrecimientotop10 + 2, buscar("Tpv_Anterior", "hoja1") + 1), Cells(filas_temporal, buscar("Tpv_Anterior", "hoja1") + 1)))
    tpv_total_mes_actual = Application.WorksheetFunction.Sum(Range(Cells(filainicialcrecimientotop10 + 2, buscar("Tpv_Pronosticado", "hoja1") + 1), Cells(filas_temporal, buscar("Tpv_Pronosticado", "hoja1") + 1)))
    
    
    If Cells(filainicialcrecimientotop10 + 1, 2).CurrentRegion.Rows.Count - 2 > 10 Then
    
         Range(Cells(filainicialcrecimientotop10 + 2 + 9, 2), Cells(filas_temporal, 1 + columnas_temporal)).Clear
         
         tpv_total_mes_anterior_top10 = Application.WorksheetFunction.Sum(Range(Cells(filainicialcrecimientotop10 + 2, buscar("Tpv_Anterior", "hoja1") + 1), Cells(filas_temporal, buscar("Tpv_Anterior", "hoja1") + 1)))
         tpv_total_mes_actual_top10 = Application.WorksheetFunction.Sum(Range(Cells(filainicialcrecimientotop10 + 2, buscar("Tpv_Pronosticado", "hoja1") + 1), Cells(filas_temporal, buscar("Tpv_Pronosticado", "hoja1") + 1)))
         
         'valores
          Cells(filainicialcrecimientotop10 + 2 + 9, 2) = "Otros Comercios"
          Cells(filainicialcrecimientotop10 + 2 + 9, 3) = "Otros Comercios"
          Cells(filainicialcrecimientotop10 + 2 + 9, 4) = tipo
          Cells(filainicialcrecimientotop10 + 2 + 9, 5) = pais
          Cells(filainicialcrecimientotop10 + 2 + 9, 6) = tpv_total_mes_anterior - tpv_total_mes_anterior_top10
          Cells(filainicialcrecimientotop10 + 2 + 9, 6).NumberFormat = "$#,##0.0,,"
          Cells(filainicialcrecimientotop10 + 2 + 9, 7) = tpv_total_mes_actual - tpv_total_mes_actual_top10
          Cells(filainicialcrecimientotop10 + 2 + 9, 7).NumberFormat = "$#,##0.0,,"
          Cells(filainicialcrecimientotop10 + 2 + 9, 8) = Cells(filainicialcrecimientotop10 + 2 + 9, 7) - Cells(filainicialcrecimientotop10 + 2 + 9, 6)
          Cells(filainicialcrecimientotop10 + 2 + 9, 8).NumberFormat = "$#,##0.0,,"
          
           If Cells(filainicialcrecimientotop10 + 2 + 9, 6) - 1 <> 0 Then
           
             Cells(filainicialcrecimientotop10 + 2 + 9, 9) = Cells(filainicialcrecimientotop10 + 2 + 9, 7) / Cells(filainicialcrecimientotop10 + 2 + 9, 6) - 1
          End If
          Cells(filainicialcrecimientotop10 + 2 + 9, 9).NumberFormat = "0.0 %"
          
          
          
           
          
    End If
     'cambiando etiquetas
      Cells(filainicialcrecimientotop10 + 1, 2) = "Usuario ID"
          Cells(filainicialcrecimientotop10 + 1, 3) = "Nombres"
          Cells(filainicialcrecimientotop10 + 1, 4) = "Tiene Convenios"
          Cells(filainicialcrecimientotop10 + 1, 5) = "País"
          Cells(filainicialcrecimientotop10 + 1, 6) = "TPV Anterior"
          Cells(filainicialcrecimientotop10 + 1, 7) = "TPV Pronóstico"
          Cells(filainicialcrecimientotop10 + 1, 8) = "Diferencia"
          Cells(filainicialcrecimientotop10 + 1, 9) = "Porcentaje Crecimiento"
          
          
          
          
          
          
          
          'formatos
          ActiveSheet.Cells.Interior.ThemeColor = xlThemeColorDark1
          
              
          With Range(Cells(filainicialcrecimientotop10 + 1, 4), Cells(filainicialcrecimientotop10 + 2 + 9, 9))
          
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
          End With
          
          Range(Cells(filainicialcrecimientotop10 + 1, 2), Cells(filainicialcrecimientotop10 + 1, 9)).Font.Bold = True
          
          'Formato  de nombres con porcentajes
          For i = 1 To 10
          
          
            If Cells(filainicialcrecimientotop10 + 1 + i, 2) <> "" Then
            
                                
                                
                If tpv_total_mes_anterior <> 0 Then
                                
                  porcentaje_anterior = Cells(filainicialcrecimientotop10 + 1 + i, buscar("Tpv_Anterior", "hoja1") + 1) / tpv_total_mes_anterior
                  
                  Else
                  porcentaje_anterior = 0
                  
                End If
                
                 If tpv_total_mes_actual <> 0 Then
                  porcentaje_actual = Cells(filainicialcrecimientotop10 + 1 + i, buscar("Tpv_Pronosticado", "hoja1") + 1) / tpv_total_mes_actual
                  
                  Else
                  
                  porcentaje_actual = 0
                
                 End If
                
                
                cadena_adicional = " / " & Format(porcentaje_anterior, "0.0 %") & "-" & Format(porcentaje_actual, "0.0 %")
                
                
                Cells(filainicialcrecimientotop10 + 1 + i, buscar("usuario_id", "hoja1") + 1) = Cells(filainicialcrecimientotop10 + 1 + i, buscar("usuario_id", "hoja1") + 1) & cadena_adicional
                            
            
            End If
          
          
          Next
    
            'Fin Formato  de nombres con porcentajes

End Sub

Sub filtro_pais_tipo(pais, tipo)

'filtro para la hoja 1

  If tipo = "Agregador" Then
       tipo = "f"
      ElseIf tipo = "Gateway" Then
       tipo = "t"
     
    End If
    
    
   
    
    Sheets("Hoja1").Range("A1").AutoFilter
    
    If tipo <> "todos" Then
      Sheets("Hoja1").Range("a1").CurrentRegion.AutoFilter Field:=buscar("tiene_c_personalizados", "hoja1"), Criteria1:=tipo
      Else
      Sheets("Hoja1").Range("a1").CurrentRegion.AutoFilter Field:=buscar("tiene_c_personalizados", "hoja1")
    
    End If
    
    If pais <> "todos" Then
       Sheets("Hoja1").Range("a1").CurrentRegion.AutoFilter Field:=buscar("pais", "hoja1"), Criteria1:=pais
      Else
       Sheets("Hoja1").Range("a1").CurrentRegion.AutoFilter Field:=buscar("pais", "hoja1")
        
    End If

End Sub

Sub test23234()


Call tablas_areas("todos", "Agregador")


End Sub



Sub grafico_areas()


fila_temporal = Cells(filainicialcrecimientotop10 + 1, 2).End(xlDown).Row
If fila_temporal = 1048576 Then

    fila_temporal = Cells(filainicialcrecimientotop10 + 1, 2).Row

End If

Set grafico_areas_a = ActiveSheet.Shapes.AddChart

  grafico_areas_a.Chart.ChartType = xlAreaStacked
  
  
  
  grafico_areas_a.Chart.SetSourceData Source:=Union(Range(Cells(filainicialcrecimientotop10 + 1, 1 + buscar("usuario_id", "hoja1")), Cells(fila_temporal, 1 + buscar("usuario_id", "hoja1"))), Range(Cells(filainicialcrecimientotop10 + 1, 1 + buscar("Tpv_Anterior", "hoja1")), Cells(fila_temporal, buscar("Tpv_Pronosticado", "hoja1") + 1))), PlotBy:=xlRows
  
  grafico_areas_a.Left = 0
  grafico_areas_a.Top = Range(Cells(1, 1), Cells(filainicialcrecimientotop10, 1)).Height
  grafico_areas_a.Width = Columns(1).Width
  grafico_areas_a.Height = Range(Rows(filainicialcrecimientotop10), Rows(filainicialcrecimientotop10 + 22)).Height
  grafico_areas_a.Chart.Legend.Format.TextFrame2.TextRange.Font.Size = 7.5
 ' grafico_areas_a.Chart.PlotArea.Top = 11.45


  'grafico_areas_a.Chart.SetElement (msoElementLegendBottom)
  grafico_areas_a.Chart.Legend.Height = 250.863
  grafico_areas_a.Chart.Legend.Top = 42.171
  grafico_areas_a.Chart.Legend.Width = 220.091
  
  grafico_areas_a.Chart.Legend.Left = 270.019
  
  grafico_areas_a.Chart.PlotArea.Width = 290.818
  grafico_areas_a.Chart.PlotArea.Height = 295.838
  grafico_areas_a.Line.Visible = msoFalse
  grafico_areas_a.Chart.Axes(xlValue).MajorGridlines.Delete

   a = 5
  
  'On Error Resume Next
  'grafico_areas_a.Chart.ChartTitle.Delete
  'On Error GoTo 0
  
  grafico_areas_a.Chart.HasTitle = True
  grafico_areas_a.Chart.ChartTitle.Text = "Comportamiento Top 10 del Mes Anterior"

End Sub




Sub tablas_tpv_grande(pais, tipo)
    
    'libro.Sheets("Hoja7").Select
  
    
    Call filtro_pais_tipo(pais, tipo)
    Call ordernar_hoja_1("Diferencia", 2)
       
   Sheets("hoja1").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    ActiveSheet.Paste
    columnas_temporal = Cells(filainicialtablas_tpv_grande + 1, 2).CurrentRegion.Columns.Count - 1
    filas_temporal = Cells(filainicialtablas_tpv_grande + 1, 2).CurrentRegion.Rows.Count + filainicialtablas_tpv_grande - 1
    
    tpv_total_mes_anterior = Application.WorksheetFunction.Sum(Range(Cells(filainicialtablas_tpv_grande + 2, buscar("Tpv_Anterior", "hoja1") + 1), Cells(filas_temporal, buscar("Tpv_Anterior", "hoja1") + 1)))
    tpv_total_mes_actual = Application.WorksheetFunction.Sum(Range(Cells(filainicialtablas_tpv_grande + 2, buscar("Tpv_Pronosticado", "hoja1") + 1), Cells(filas_temporal, buscar("Tpv_Pronosticado", "hoja1") + 1)))
    
    
    If Cells(filainicialtablas_tpv_grande + 1, 2).CurrentRegion.Rows.Count - 2 > 10 Then
    
         Range(Cells(filainicialtablas_tpv_grande + 2 + 10, 2), Cells(filas_temporal, 1 + columnas_temporal)).Clear
                    
    End If
    
    
     'cambiando etiquetas
      Cells(filainicialtablas_tpv_grande + 1, 2) = "Usuario ID"
          Cells(filainicialtablas_tpv_grande + 1, 3) = "Nombres"
          Cells(filainicialtablas_tpv_grande + 1, 4) = "Tiene Convenios"
          Cells(filainicialtablas_tpv_grande + 1, 5) = "País"
          Cells(filainicialtablas_tpv_grande + 1, 6) = "TPV Anterior"
          Cells(filainicialtablas_tpv_grande + 1, 7) = "TPV Pronóstico"
          Cells(filainicialtablas_tpv_grande + 1, 8) = "Diferencia"
          Cells(filainicialtablas_tpv_grande + 1, 9) = "Porcentaje Crecimiento"
          
          
          
          
          
          
          
          'formatos
          ActiveSheet.Cells.Interior.ThemeColor = xlThemeColorDark1
          
              
          With Range(Cells(filainicialtablas_tpv_grande + 1, 4), Cells(filainicialtablas_tpv_grande + 2 + 9, 9))
          
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
          End With
          
          Range(Cells(filainicialtablas_tpv_grande + 1, 2), Cells(filainicialtablas_tpv_grande + 1, 9)).Font.Bold = True
    
        

    
    
End Sub



Sub ordernar_hoja_1(campo, tipo_orden)

'ascendente,tipo_orden=1
'descendente,tipo_orden =2

 libro.Sheets("Hoja1").Range("A1").Sort _
        Key1:=libro.Sheets("Hoja1").Columns(buscar(campo, "Hoja1")), Order1:=tipo_orden, _
        Header:=xlGuess
        

End Sub


Sub grafico_tpv_grande()


  Set grafico_areas_a = ActiveSheet.Shapes.AddChart
  
 
  
  
  grafico_areas_a.Chart.ChartType = xlBubble
  grafico_areas_a.Chart.HasTitle = True
  
  On Error Resume Next
  grafico_areas_a.Chart.SetElement (msoElementChartTitleCenteredOverlay)
  
  
  grafico_areas_a.Chart.Axes(xlValue).MajorGridlines.Delete
  grafico_areas_a.Chart.Legend.Delete
  grafico_areas_a.Chart.Axes(xlCategory).Delete
  grafico_areas_a.Chart.Axes(xlValue).Delete
  
  On Error Resume Next
        grafico_areas_a.Chart.ChartTitle.Text = "Comercios con Mayor Crecimiento Absoluto"
  On Error GoTo 0
  
  
  
  On Error Resume Next
        For i = grafico_areas_a.Chart.SeriesCollection.Count To 1 Step -1
              grafico_areas_a.Chart.SeriesCollection(i).Delete
        Next
  On Error GoTo 0



  fila = filainicialtablas_tpv_grande + 1
  For i = 1 To 10
    
    If Cells(fila + i, buscar("Diferencia", "hoja1") + 1) <> "" And Cells(fila + i, buscar("Diferencia", "hoja1") + 1) > 0 Then
          
            With grafico_areas_a.Chart
            
               
            
            .SeriesCollection.NewSeries
            .SeriesCollection(i).Name = Cells(fila + i, buscar("usuario_id", "hoja1") + 1)
            .SeriesCollection(i).XValues = 0
            .SeriesCollection(i).Values = 20 - (i * 2)
            .SeriesCollection(i).BubbleSizes = Cells(fila + i, buscar("Diferencia", "hoja1") + 1)
            
             If i = 1 Then
                   .SetElement (msoElementDataLabelLeft)
             End If
                     
            
            
            .SeriesCollection(i).Points(1).DataLabel.Text = Format(Cells(fila + i, buscar("Diferencia", "hoja1") + 1).Value, "$ #,##0,,")
             
            
            End With
    End If
    
    Next
    
    For i = 1 To grafico_areas_a.Chart.SeriesCollection.Count
             izquierda = grafico_areas_a.Chart.SeriesCollection(1).Points(1).DataLabel.Left + 100
             arriba = grafico_areas_a.Chart.SeriesCollection(i).Points(1).DataLabel.Top + 1
    
             grafico_areas_a.Chart.Shapes.AddTextbox(msoTextOrientationHorizontal, izquierda, arriba, 280, 13).TextFrame.Characters.Text = Cells(fila + i, buscar("usuario_id", "hoja1") + 1) & " / " & Format(Cells(fila + i, buscar("Porcentaje", "hoja1") + 1), "0.0 %")
      
    Next
 
 
  grafico_areas_a.Left = 0
  grafico_areas_a.Top = Range(Cells(1, 1), Cells(filainicialtablas_tpv_grande, 1)).Height
  grafico_areas_a.Width = Columns(1).Width
  grafico_areas_a.Height = Range(Rows(filainicialtablas_tpv_grande), Rows(filainicialtablas_tpv_grande + 22)).Height
  grafico_areas_a.Line.Visible = msoFalse

End Sub

'prueba otro grafico


Sub tablas_tpv_negativo(pais, tipo)
    
    'libro.Sheets("Hoja7").Select
  
    
    Call filtro_pais_tipo(pais, tipo)
    Call ordernar_hoja_1("Diferencia", 1)
       
   Sheets("hoja1").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    ActiveSheet.Paste
    columnas_temporal = Cells(filainicialtablas_tpv_negativo + 1, 2).CurrentRegion.Columns.Count - 1
    filas_temporal = Cells(filainicialtablas_tpv_negativo + 1, 2).CurrentRegion.Rows.Count + filainicialtablas_tpv_negativo - 1
    
    tpv_total_mes_anterior = Application.WorksheetFunction.Sum(Range(Cells(filainicialtablas_tpv_negativo + 2, buscar("Tpv_Anterior", "hoja1") + 1), Cells(filas_temporal, buscar("Tpv_Anterior", "hoja1") + 1)))
    tpv_total_mes_actual = Application.WorksheetFunction.Sum(Range(Cells(filainicialtablas_tpv_negativo + 2, buscar("Tpv_Pronosticado", "hoja1") + 1), Cells(filas_temporal, buscar("Tpv_Pronosticado", "hoja1") + 1)))
    
    
    If Cells(filainicialtablas_tpv_negativo + 1, 2).CurrentRegion.Rows.Count - 2 > 10 Then
    
         Range(Cells(filainicialtablas_tpv_negativo + 2 + 10, 2), Cells(filas_temporal, 1 + columnas_temporal)).Clear
                    
    End If
    
    
     'cambiando etiquetas
      Cells(filainicialtablas_tpv_negativo + 1, 2) = "Usuario ID"
          Cells(filainicialtablas_tpv_negativo + 1, 3) = "Nombres"
          Cells(filainicialtablas_tpv_negativo + 1, 4) = "Tiene Convenios"
          Cells(filainicialtablas_tpv_negativo + 1, 5) = "País"
          Cells(filainicialtablas_tpv_negativo + 1, 6) = "TPV Anterior"
          Cells(filainicialtablas_tpv_negativo + 1, 7) = "TPV Pronóstico"
          Cells(filainicialtablas_tpv_negativo + 1, 8) = "Diferencia"
          Cells(filainicialtablas_tpv_negativo + 1, 9) = "Porcentaje Crecimiento"
          
          
          
          
          
          
          
          'formatos
          ActiveSheet.Cells.Interior.ThemeColor = xlThemeColorDark1
          
              
          With Range(Cells(filainicialtablas_tpv_negativo + 1, 4), Cells(filainicialtablas_tpv_negativo + 2 + 9, 9))
          
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
          End With
          
          Range(Cells(filainicialtablas_tpv_negativo + 1, 2), Cells(filainicialtablas_tpv_negativo + 1, 9)).Font.Bold = True
    
        

    
    
End Sub



Sub grafico_tpv_negativo()


  Set grafico_areas_a = ActiveSheet.Shapes.AddChart
  
 
  
  
  grafico_areas_a.Chart.ChartType = xlBubble
  grafico_areas_a.Chart.HasTitle = True
  
  On Error Resume Next
  grafico_areas_a.Chart.SetElement (msoElementChartTitleCenteredOverlay)
  
  
  grafico_areas_a.Chart.Axes(xlValue).MajorGridlines.Delete
  grafico_areas_a.Chart.Legend.Delete
  grafico_areas_a.Chart.Axes(xlCategory).Delete
  grafico_areas_a.Chart.Axes(xlValue).Delete
  
  On Error Resume Next
        grafico_areas_a.Chart.ChartTitle.Text = "Comercios con Mayor Disminución Absoluta"
  On Error GoTo 0
  
  
  
  On Error Resume Next
        For i = grafico_areas_a.Chart.SeriesCollection.Count To 1 Step -1
              grafico_areas_a.Chart.SeriesCollection(i).Delete
        Next
  On Error GoTo 0



  fila = filainicialtablas_tpv_negativo + 1
  For i = 1 To 10
    
    If Cells(fila + i, buscar("Diferencia", "hoja1") + 1) <> "" And Cells(fila + i, buscar("Diferencia", "hoja1") + 1) < 0 Then
          
            With grafico_areas_a.Chart
            
               
            
            .SeriesCollection.NewSeries
            .SeriesCollection(i).Name = Cells(fila + i, buscar("usuario_id", "hoja1") + 1)
            .SeriesCollection(i).XValues = 0
            .SeriesCollection(i).Values = 20 - (i * 2)
            medida_burbuja = CStr(Round(-Cells(fila + i, buscar("Diferencia", "hoja1") + 1), 0))
            .SeriesCollection(i).BubbleSizes = "{" & medida_burbuja & "}"
            
            
             If i = 1 Then
                   .SetElement (msoElementDataLabelLeft)
             End If
                     
            
            
            .SeriesCollection(i).Points(1).DataLabel.Text = Format(Cells(fila + i, buscar("Diferencia", "hoja1") + 1).Value, "$ #,##0,,")
             
            
            End With
    End If
    
    Next
    
    For i = 1 To grafico_areas_a.Chart.SeriesCollection.Count
             izquierda = grafico_areas_a.Chart.SeriesCollection(1).Points(1).DataLabel.Left + 100
             arriba = grafico_areas_a.Chart.SeriesCollection(i).Points(1).DataLabel.Top + 1
    
             grafico_areas_a.Chart.Shapes.AddTextbox(msoTextOrientationHorizontal, izquierda, arriba, 280, 13).TextFrame.Characters.Text = Cells(fila + i, buscar("usuario_id", "hoja1") + 1) & " / " & Format(Cells(fila + i, buscar("Porcentaje", "hoja1") + 1), "0.0 %")
      
    Next
 
 
  grafico_areas_a.Left = 0
  grafico_areas_a.Top = Range(Cells(1, 1), Cells(filainicialtablas_tpv_negativo, 1)).Height
  grafico_areas_a.Width = Columns(1).Width
  grafico_areas_a.Height = Range(Rows(filainicialtablas_tpv_negativo), Rows(filainicialtablas_tpv_negativo + 22)).Height
  grafico_areas_a.Line.Visible = msoFalse

End Sub

Function cantidad_dias_laborales_todomes(fecha)
     temporal = 0
     dias_mes = Day(Application.WorksheetFunction.EoMonth(fecha, 0))
     primer_dia = Format(fecha, "yyyy/mm") & "/01"
      
      For i = 0 To dias_mes - 1
     
            If Weekday(DateAdd("d", i, primer_dia), 2) < 6 Then
            
                 temporal = temporal + 1
                 
                 
              
            End If
           
              


       Next
       
        cantidad_dias_laborales_todomes = temporal
        
End Function
Sub test()

Call vector_dias_laborales(Date)

End Sub



Function cantidad_dias_laborales_mescorrido(fecha)
     temporal = 0
     dias_mes = Day(fecha)
     primer_dia = Format(fecha, "yyyy/mm") & "/01"
      
      For i = 0 To dias_mes - 1
     
            If Weekday(DateAdd("d", i, primer_dia), 2) < 6 Then
            
                 temporal = temporal + 1
                 
                 
              
            End If
           
              


       Next
       
        cantidad_dias_laborales_mescorrido = temporal
        
End Function

Sub vector_dias_laborales(fecha)

     dias_mes = Day(Application.WorksheetFunction.EoMonth(fecha, 0))
     dias_mes_corridos = Day(fecha)
     
     
     
     If cantidad_dias_laborales_mescorrido(fecha) <> 0 Then
        multiplicador_dias_laboral = cantidad_dias_laborales_todomes(fecha) / cantidad_dias_laborales_mescorrido(fecha)
        Else
        multiplicador_dias_laboral = 0
        
        
    End If
     
     
     If dias_mes_corridos - cantidad_dias_laborales_mescorrido(fecha) <> 0 Then
     multiplicador_dias_no_laboral = (dias_mes - cantidad_dias_laborales_todomes(fecha)) / (dias_mes_corridos - cantidad_dias_laborales_mescorrido(fecha))
        Else
        multiplicador_dias_no_laboral = 0
     End If
     
     'El pronóstico cambia cuando no hay dias laborales y dias no laborales en el mes
     
     If cantidad_dias_laborales_mescorrido(fecha) = 0 Or (dias_mes_corridos - cantidad_dias_laborales_mescorrido(fecha)) = 0 Then
       
        
       multiplicador_dias_laboral = dias_mes / dias_mes_corridos
       multiplicador_dias_no_laboral = dias_mes / dias_mes_corridos
     
     
     
     End If
     
       'Fin del cambio para el pronóstico de los primeros siete días
     
     
     
     
     temporal = 0
     
     primer_dia = Format(fecha, "yyyy/mm") & "/01"
      entre_semana = "{"
      fin_Semana = "{"
      For i = 0 To dias_mes_corridos - 1
     
            If Weekday(DateAdd("d", i, primer_dia), 2) < 6 Then
            
                 entre_semana = entre_semana & "1\"
                 fin_Semana = fin_Semana & "0\"
                 
                 Else
                 entre_semana = entre_semana & "0\"
                 fin_Semana = fin_Semana & "1\"
                 
                 
                 
              
            End If
           
              


       Next
       
       entre_semana = Mid(entre_semana, 1, Len(entre_semana) - 1) & "}"
       fin_Semana = Mid(fin_Semana, 1, Len(fin_Semana) - 1) & "}"
       
        
End Sub

Sub tablas_dinamicas_tpv_comercios()








End Sub



Sub tipo_dia_semana(hoja)

' Establece el tipo de día de cada uno de los dias(lunes-viernes o sábado-domingo), Tambien debe establecer una fila dummy con tipo 0 y tipo 1

nombre_temporal = ActiveSheet.Name
libro.Sheets(hoja).Select
Dim rango_temporal As Range
Dim rango_temporal_2 As Range
'filas_temporal = libro.Sheets(hoja).range("a1").CurrentRegion.Rows.Count - 1
filas_temporal = Sheets(hoja).Range("a1").CurrentRegion.Rows.Count
columnas_temporal_2 = Sheets(hoja).Range("a1").CurrentRegion.Columns.Count + 1
columnas_temporal = buscar("rango", hoja)
ReDim arreglo_temporal(1 To filas_temporal)

'Set rango_temporal = libro.Sheets(hoja).Range(Cells(2, columnas_temporal), Cells(, columnas_temporal))
Set rango_temporal = libro.Sheets(hoja).Range(Cells(2, columnas_temporal), Cells(filas_temporal, columnas_temporal))
Set rango_temporal_2 = libro.Sheets(hoja).Range(Cells(2, columnas_temporal_2), Cells(filas_temporal, columnas_temporal_2))
arreglo_temporal = Application.WorksheetFunction.Transpose(rango_temporal)


For i = 1 To filas_temporal - 1
  
    If Application.WorksheetFunction.Weekday(CDate(arreglo_temporal(i)), 2) < 6 Then
        arreglo_temporal(i) = 1
        Else
        arreglo_temporal(i) = 0
      
    End If


Next

libro.Sheets(hoja).Cells(1, columnas_temporal_2) = "entre_semana"


'valores dummy
libro.Sheets(hoja).Cells(filas_temporal + 1, buscar("usuario_id", hoja)) = "Dummy1"
libro.Sheets(hoja).Cells(filas_temporal + 2, buscar("usuario_id", hoja)) = "Dummy1"
libro.Sheets(hoja).Cells(filas_temporal + 1, columnas_temporal_2) = "1"
libro.Sheets(hoja).Cells(filas_temporal + 2, columnas_temporal_2) = "0"


'fin valores dummy

rango_temporal_2 = Application.WorksheetFunction.Transpose(arreglo_temporal)

libro.Sheets(nombre_temporal).Select


End Sub


Sub tabla_dinamica_comercios_tpv(hoja)
nombre_temporal = ActiveSheet.Name
libro.Worksheets.Add

libro.ActiveSheet.Name = hoja & "_td"

Set Cache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=Sheets(hoja).Range("A2").CurrentRegion)

Set dinamica_otra = ActiveSheet.PivotTables.Add( _
        PivotCache:=Cache, _
        TableDestination:=Range("A1"))
        

With dinamica_otra
        .ColumnGrand = False
        .RowGrand = False
        
        
             
                
        .PivotFields("usuario_id").Orientation = xlRowField
        .PivotFields("entre_semana").Orientation = xlColumnField
        
         With .PivotFields("suma")
         .Orientation = xlDataField
         .Function = xlSum
        End With
   
        
     libro.Sheets(nombre_temporal).Select
        
End With




End Sub

Sub comercios_nuevos()


libro.Sheets.Add after:=libro.Sheets(libro.Sheets.Count)
libro.ActiveSheet.Name = "temporal_comercios_nuevos"
Application.Workbooks.Open ("D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\comercios_nuevos.xlsx")
libro.Activate
Workbooks("comercios_nuevos.xlsx").Sheets(1).Range("a1").CurrentRegion.Copy (libro.Sheets(libro.Sheets.Count).Range("a1"))

With libro.Sheets(libro.Sheets.Count).Range("a1").CurrentRegion
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

With libro.Sheets(libro.Sheets.Count).Columns(buscar("nombres", libro.Sheets(libro.Sheets.Count).Name))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

libro.Sheets(libro.Sheets.Count).Cells(1, buscar("nombres", libro.Sheets(libro.Sheets.Count).Name)).HorizontalAlignment = xlCenter


libro.Sheets(libro.Sheets.Count).Columns(buscar("suma", libro.Sheets(libro.Sheets.Count).Name)).NumberFormat = "$ #,##0.0,,"

libro.Sheets(libro.Sheets.Count).Cells(1, buscar("suma", libro.Sheets(libro.Sheets.Count).Name)) = "Valor(Millones)"
libro.Sheets(libro.Sheets.Count).Cells(1, buscar("nombres", libro.Sheets(libro.Sheets.Count).Name)) = "Nombre"
libro.Sheets(libro.Sheets.Count).Cells(1, buscar("usuario_id", libro.Sheets(libro.Sheets.Count).Name)) = "Usuario id"
libro.Sheets(libro.Sheets.Count).Cells(1, buscar("fecha_creacion_comercio", libro.Sheets(libro.Sheets.Count).Name)) = "Fecha Creación Comercio"
libro.Sheets(libro.Sheets.Count).Cells(1, buscar("count", libro.Sheets(libro.Sheets.Count).Name)) = "Cantidad"

Workbooks("comercios_nuevos.xlsx").Close


libro.Sheets.Add after:=libro.Sheets(libro.Sheets.Count)
libro.ActiveSheet.Name = "Comercios Nuevos"

libro.Sheets("Comercios Nuevos").Cells.RowHeight = 15

libro.Sheets("Comercios Nuevos").Columns(1).ColumnWidth = 20
libro.Sheets("Comercios Nuevos").Columns(2).ColumnWidth = 40
libro.Sheets("Comercios Nuevos").Columns(3).ColumnWidth = 30
libro.Sheets("Comercios Nuevos").Columns(4).ColumnWidth = 20


numeropais = Sheets("Hoja3").Cells(1, 3).End(xlDown).Row - 1
tipo = 2

fila_inicial = 5

For i = 1 To numeropais

    For j = 1 To tipo
    
        If Sheets("hoja3").Cells(i + 1, 3) <> "todos" And Sheets("hoja3").Cells(j, 5) <> "todos" Then
            
           Call filtro_pais_tipo_comercio_nuevo(Sheets("hoja3").Cells(i + 1, 3), Sheets("hoja3").Cells(j, 5))
        
    
           If Application.WorksheetFunction.Subtotal(103, Sheets("temporal_comercios_nuevos").Columns(1)) > 1 Then
                
                
                alto = Range(Sheets("Comercios Nuevos").Cells(1, 1), Sheets("Comercios Nuevos").Cells(fila_inicial - 2, 1)).Height
                izquierda = Sheets("Comercios Nuevos").Columns(1).Width
                Sheets("Comercios Nuevos").Cells(fila_inicial, 1) = Sheets("hoja3").Cells(i + 1, 3) & "-" & Sheets("hoja3").Cells(j, 5)
                Sheets("Comercios Nuevos").Cells(fila_inicial, 1).Font.ThemeColor = xlThemeColorLight2
                Sheets("Comercios Nuevos").Cells(fila_inicial, 1).Font.Size = 14
                Sheets("Comercios Nuevos").Cells(fila_inicial, 1).Font.Bold = True
                imagen = "D:\javier.cortes\Documents\Javier\Pagos Online\ultimo\Banderas\" & Sheets("hoja3").Cells(i + 1, 3) & ".jpg"
                Sheets("Comercios Nuevos").Shapes.AddPicture imagen, True, True, izquierda, alto, 60, 35
                Sheets("temporal_comercios_nuevos").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy (Sheets("Comercios Nuevos").Cells(fila_inicial + 2, 1))
                Intersect(Sheets("Comercios Nuevos").Cells(fila_inicial + 2, 1).CurrentRegion, Range(Sheets("Comercios Nuevos").Columns(6), Sheets("Comercios Nuevos").Columns(7))).Delete
                fila_inicial = Sheets("Comercios Nuevos").Cells(fila_inicial + 2, 1).CurrentRegion.Rows.Count + fila_inicial + 5
           
           End If
           
        End If
    
         Sheets("temporal_comercios_nuevos").Range("A1").AutoFilter
    
    Next
    
 Next

libro.Sheets("temporal_comercios_nuevos").Delete
libro.Sheets("Comercios Nuevos").Cells.Interior.Color = RGB(255, 255, 255)

libro.Sheets("Comercios Nuevos").Cells(2, 1).Value = "Comercios sin Transacciones Pagadas y Abonadas en Meses Anteriores"
libro.Sheets("Comercios Nuevos").Cells(2, 1).Font.Bold = True

libro.Sheets("Comercios Nuevos").Name = "TPV Comercios Nuevos"

End Sub




Sub filtro_pais_tipo_comercio_nuevo(pais, tipo)

'filtro para la hoja 1

  If tipo = "Agregador" Then
       tipo = "f"
      ElseIf tipo = "Gateway" Then
       tipo = "t"
     
    End If
    
    
   
    
    Sheets("temporal_comercios_nuevos").Range("A1").AutoFilter
    
    If tipo <> "todos" Then
      Sheets("temporal_comercios_nuevos").Range("a1").CurrentRegion.AutoFilter Field:=buscar("convenio_directo_cliente", "temporal_comercios_nuevos"), Criteria1:=tipo
      Else
      Sheets("temporal_comercios_nuevos").Range("a1").CurrentRegion.AutoFilter Field:=buscar("convenio_directo_cliente", "temporal_comercios_nuevos")
    
    End If
    
    If pais <> "todos" Then
       Sheets("temporal_comercios_nuevos").Range("a1").CurrentRegion.AutoFilter Field:=buscar("pais", "temporal_comercios_nuevos"), Criteria1:=pais
      Else
       Sheets("temporal_comercios_nuevos").Range("a1").CurrentRegion.AutoFilter Field:=buscar("pais", "temporal_comercios_nuevos")
        
    End If

End Sub



Sub filtro_valor(columna_campo, valor)

'filtro valor hoja "hoja1"
 libro.Sheets("hoja1").Range("a1").CurrentRegion.AutoFilter Field:=columna_campo, Criteria1:=">" & valor, Operator:=xlAnd



End Sub



Sub tabla_comercios_activos(pais, tipo)


libro.Sheets("hoja1").AutoFilterMode = False



valores_filtro = Array(0, 10000000, 100000000)
campos = Array("Tpv_Anterior", "Tpv_Pronosticado")
largo_valores = UBound(valores_filtro) + 1
largo_campos = UBound(campos) + 1

ReDim resultado_filtros_comercios_activos(1 To largo_campos, 1 To largo_valores)



Call filtro_pais_tipo(pais, tipo)

For c = 0 To UBound(campos)


    For vf = 0 To UBound(valores_filtro)
        
        Call filtro_valor(buscar(campos(c), "hoja1"), valores_filtro(vf))
        
        resultado_filtros_comercios_activos(c + 1, vf + 1) = Application.WorksheetFunction.Subtotal(103, libro.Sheets("hoja1").Columns(1)) - 1
        
        
        libro.Sheets("hoja1").Range("a1").CurrentRegion.AutoFilter Field:=buscar(campos(c), "hoja1")
         
  
    Next



Next


ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 + 1, 2 + 0) = "Anterior"
ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 + 2, 2 + 0) = "Actual"
ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 + 0, 2 + 1) = "> $0"
ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 + 0, 2 + 2) = "> $10"
ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 + 0, 2 + 3) = "> $100"
ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 - 1, 2 + 1) = "TPV Anterior/Pronostico TPV"

ActiveSheet.Range(Cells(filainicialtablas_comercios_activos + 2 - 1, 2 + 1), Cells(filainicialtablas_comercios_activos + 2 - 1, 2 + 3)).Merge
ActiveSheet.Range(Cells(filainicialtablas_comercios_activos + 2 - 1, 2 + 1), Cells(filainicialtablas_comercios_activos + 2 + 2, 2 + 3)).HorizontalAlignment = xlCenter
ActiveSheet.Range(Cells(filainicialtablas_comercios_activos + 2 - 1, 2 + 1), Cells(filainicialtablas_comercios_activos + 2 + 0, 2 + 3)).Font.Bold = True
 


For i = 1 To largo_campos
    For j = 1 To largo_valores
       ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 + i, 2 + j) = resultado_filtros_comercios_activos(i, j)
    Next
Next




End Sub


Sub test_adsfasdf()



Set libro = ActiveWorkbook



Call grafico_comercios_activos("CO", "f", libro)


End Sub



Sub grafico_comercios_activos()




  Set grafico_comercios_ac = ActiveSheet.Shapes.AddChart
  
 
  
  
  grafico_comercios_ac.Chart.ChartType = xlColumnClustered
  grafico_comercios_ac.Chart.HasTitle = True
  
  On Error Resume Next
  grafico_comercios_ac.Chart.SetElement (msoElementChartTitleCenteredOverlay)
  
  grafico_comercios_ac.Chart.ChartTitle.Text = "Comercios Activos"
  grafico_comercios_ac.Chart.SetSourceData Source:=Range(ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 - 1, 2 + 0), ActiveSheet.Cells(filainicialtablas_comercios_activos + 2 + 2, 2 + 3)), PlotBy:=xlRows
  
  grafico_comercios_ac.Left = 0
  grafico_comercios_ac.Width = Columns(1).Width
  grafico_comercios_ac.Height = Range(Rows(filainicialtablas_comercios_activos), Rows(filainicialtablas_comercios_activos + 16)).Height
  grafico_comercios_ac.Top = Range(Rows(1), Rows(filainicialtablas_comercios_activos + 1)).Height
  grafico_comercios_ac.Chart.Axes(xlValue).MajorGridlines.Delete
  grafico_comercios_ac.Line.Visible = msoFalse
  grafico_comercios_ac.Chart.SetElement (msoElementDataLabelOutSideEnd)

End Sub



