Public Class PickeoInfoDTO
    ' Representa el objeto "Hoja_detalle" que envías en el Map
    Public Property Hoja_detalle As DetalleHoja

    ' Representa la lista "Detalle" (los paquetes) que envías en el Map
    Public Property Detalle As List(Of PaquetePickeo)
End Class