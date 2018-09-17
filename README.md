# Código VBA

Este repositorio tiene como objetivo recolectar código utilizado en soluciones basadas en MS Excel (vba) ya existentes, con el fin de agilizar el proceso de creación y respuesta, de modo que el código aquí presentado sirva como base de contrucción para nuevos desarrollos.

La primera versión de la estructura y clasificación será la siguiente:

    Título de Solución
    ¿Qué hace el código?
    Comentarios adicionales para su utilización
    Código

## Bucle Único (Single Loop)
Se utiliza para recorrer un rango unidimensional de celdas.
Debes agregar un **botón (control de formulario)** en la hoja de trabajo. Esto se logra desde _Insertar_ en la pestaña _Programador_ y nombrar el botón para luego agregar el código:
> En este caso, se agregará el valor "Excel la lleva" a la celda 1 hasta la 6.
```
Sub bucle()
    Dim i As Integer () '"Dim i%" es equivalente a definir i como integer
    
    For i = 1 To 6
        Cells(i, 1).Value = "Excel la lleva"
    Next i
End Sub
```

## Traspaso de Datos hoja principal -> hoja
Se utiliza para traspasar datos ingresados en la primera hoja, hacia otras fichas (con propios y distintos requerimientos de datos), al hacer click en el botón correspondiente.
Existen **botones (control de formulario)** correspondientes a cada ficha (hoja) a la cual se deseen traspasar los datos.

```
Public i&, j&, m&, r&, strSKU$
Private Const f1 = "Ficha Uno"
Private Const f2 = "Ficha Dos"

Public Sub TraspasoDatos()
    Application.ScreenUpdating = False
    On Error GoTo triggerDeError
    Dim btn As Excel.Button
    Set btn = Worksheets(ActiveSheet.Name).Buttons(Application.Caller)
    r = lasRows - 1
    Select Case btn.Caption
        Case f1: FichaUno r
        Case f2: FichaDos r
    End Select
    MsgBox "Ficha finalizada.", vbOKOnly, "Generador de Fichas"
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    Exit Sub
triggerDeError:
    MsgBox "Error inesperado. El proceso será interrumpido.", vbOKOnly + vbCritical, "Generador de Fichas"
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
End Sub

Private Sub FichaUno(r&)
    For j = 2 To r
        m = 1
        shtFichaUno.Cells(j, m) = shtNombreHojaMaestra.Range("A" & j + 1).Value2: m = m + 1
        shtFichaUno.Cells(j, m) = shtNombreHojaMaestra.Range("B" & j + 1).Value2: m = m + 1
        'esto dependiendo del número de columnas presentes en la hoja inicial.
        shtFichaUno.Cells(j, m) = shtNombreHojaMaestra.Range("AJ" & j + 1).Value2
    Next j
    shtFichaUno.Select
End Sub

Private Function lasRowss&()
    i = 1: While shtLOrealMaestra.Cells(i + 1, 1).Value <> "": i = i + 1: Wend: lasRows = i
End Function


Sub MetadataImagenes()
    shtNombreHojaMaestra.Range("AM3:AO" & lasRows).ClearContents
    Dim pic As Excel.Picture
    For Each pic In shtNombreHojaMaestra.Pictures
        pic.TopLeftCell.Offset(, 2).Value2 = pic.Name: pic.TopLeftCell.Offset(, 3).Value2 = pic.Height: pic.TopLeftCell.Offset(, 4).Value2 = pic.Width / 5
    Next pic
End Sub
```

