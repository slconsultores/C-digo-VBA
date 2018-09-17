# Código VBA

Este repositorio tiene como objetivo recolectar código utilizado en soluciones basadas en MS Excel (vba) ya existentes, con el fin de agilizar el proceso de creación y respuesta, de modo que el código aquí presentado sirva como base de contrucción para nuevos desarrollos.

La primera versión de la estructura y clasificación será la siguiente:

    Título de Solución
    ¿Qué hace el código?
    Comentarios adicionales para su utilización
    Código

Ej.

## Bucle Único (Single Loop)
Se utiliza para recorrer un rango unidimensional de celdas.
Debes agregar un **botón (control de formulario)** en la hoja de trabajo. Esto se logra desde _Insertar_ en la pestaña _Programador_ y nombrar el botón para luego agregar el código:
> En este caso, se agregará el valor "Excel la lleva" a la celda 1 hasta la 6.
```
Sub bucle()
    Dim i As Integer
    
    For i = 1 To 6
        Cells(i, 1).Value = "Excel la lleva"
    Next i
End Sub
```


