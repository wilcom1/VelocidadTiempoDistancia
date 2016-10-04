Public Class VTD
    Public Function calculateVelocidad(Distancia As Double,
                                       Tiempo As Double)
        Dim Resultado As Double
        Resultado = Distancia / Tiempo
        Return Resultado
    End Function

    Public Function calculateDistancia(Velocidad As Double,
                                       Tiempo As Double)
        Dim Resultado As Double
        Resultado = Velocidad * Tiempo
        Return Resultado
    End Function
    'Función que calcula el tiempo y lo devuelve en minutos
    Public Function calculateTiempo(Velocidad As Double,
                                   Distancia As Double)
        Dim Resultado As Double
        Resultado = (Distancia / Velocidad) * 60
        Return Resultado
    End Function
End Class
