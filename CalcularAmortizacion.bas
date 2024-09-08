Attribute VB_Name = "Módulo2"
Sub CalcularAmortizacion()

    ' Definir variables de entrada
    Dim saldo_actual As Double
    Dim tipo_interes_anual As Double
    Dim tipo_interes_mensual As Double
    Dim cuota_mensual As Double
    Dim meses_hasta_revision As Integer
    Dim mes As Integer
    Dim parte_intereses As Double
    Dim parte_principal As Double
    
    ' Convertir el tipo de interés anual a mensual
    tipo_interes_anual = Range("B4").Value
    tipo_interes_mensual = (tipo_interes_anual / 100) / 12
    
    ' Inicializar variables de entrada
    saldo_actual = Range("B3").Value
    cuota_mensual = Range("D3").Value
    meses_hasta_revision = Range("B6").Value
    
    ' Inicializar el saldo pendiente
    Dim saldo_pendiente As Double
    saldo_pendiente = saldo_actual
    
    ' Definir dónde vamos a imprimir el resultado de esta macro
    Dim fila As Integer
    fila = 6    ' Empezar en la fila 6
    
    ' Bucle a través de los meses hasta el mes de revisión del tipo de interés
    For mes = 1 To meses_hasta_revision
    
        ' Calcular la parte de intereses
        parte_intereses = saldo_pendiente * tipo_interes_mensual
        
        ' Calcular la parte de principal
        parte_principal = cuota_mensual - parte_intereses
        
        ' Recalcular el saldo pendiente
        saldo_pendiente = saldo_pendiente - parte_principal
        
        ' Imprimir el desglose de la cuota mensual en celdas de Excel y aplicar formato número
        Cells(fila, 4).Value = mes
        Cells(fila, 5).Value = cuota_mensual: Cells(fila, 5).NumberFormat = "0.00"
        Cells(fila, 6).Value = parte_principal: Cells(fila, 6).NumberFormat = "0.00"
        Cells(fila, 7).Value = parte_intereses: Cells(fila, 7).NumberFormat = "0.00"
        Cells(fila, 8).Value = saldo_pendiente: Cells(fila, 8).NumberFormat = "0.00"
        
        ' Avanzar a la siguiente fila
        fila = fila + 1
        
    Next mes

End Sub
