Public Class EncabezadoPickeo
    Public Id As String

    Public Property Empresa As String
    Public Property Sucursal As String
    Public Property Numero_Picking As String
    Public Property Numero_NVV As String
    Public Property CodFuncionario_Pickea As String
    ' B4A envía Ticks (Long), en .NET usamos Int64
    Public Property Fecha_Picking As Int64
    Public Property Idmaeedo As Integer
End Class