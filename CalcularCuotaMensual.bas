Attribute VB_Name = "Módulo1"
Sub CalcularCuotaMensual()

    ' Datos iniciales
    Dim saldo_actual As Double
    Dim tipo_interes_anual As Double
    Dim meses_restantes As Integer
    Dim cuota_mensual As Double
    Dim tipo_interes_mensual As Double
    Dim temp As Double
    
    ' Asignación de datos desde celdas de Excel
    saldo_actual = Range("B3").Value
    tipo_interes_anual = Range("B4").Value
    meses_restantes = Range("B5").Value
    
    ' Convertir el tipo de interés anual a mensual
    tipo_interes_mensual = (tipo_interes_anual / 100) / 12
    
    ' Calcular la cuota mensual mediante la fórmula del sistema de amortización francés
    temp = (1 + tipo_interes_mensual) ^ meses_restantes
    cuota_mensual = saldo_actual * (tipo_interes_mensual * temp) / (temp - 1)
    
    ' Mostrar el resultado en una celda concreta
    Range("D3").Value = cuota_mensual
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Definir variable intereses
    Dim intereses As Double
    
    ' Calcular la parte de intereses de la cuota mensual
    intereses = saldo_actual * tipo_interes_mensual
    
    ' Mostrar el resultado en una celda concreta
    Range("F3").Value = intereses
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Definir variable principal
    Dim principal As Double
    
    ' Calcular la parte de principal de la cuota mensual
    principal = cuota_mensual - intereses
    
    ' Mostrar el resultado en una celda concreta
    Range("E3").Value = principal

End Sub
