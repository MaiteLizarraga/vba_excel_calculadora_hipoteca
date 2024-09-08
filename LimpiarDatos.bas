Attribute VB_Name = "Módulo3"
Sub LimpiarDatos()

    ' Limpiar las celdas que contienen resultados
    ' No se limpian las celdas con los datos introducidos por el usuario
    Range("D3:F3").ClearContents
    Range("D6:H19").ClearContents

End Sub
