' NOTA: si cambia aquí el nombre de interfaz "IService1", también debe actualizar la referencia a "IService1" en Web.config.
<ServiceContract()> _
Public Interface IService1

    <OperationContract()> _
    Function GetData(ByVal value As Integer) As String

    <OperationContract()> _
    Function GetDataUsingDataContract(ByVal composite As CompositeType) As CompositeType

    ' TAREAS PENDIENTES: agregue aquí sus operaciones de servicio

End Interface

' Utilice un contrato de datos, como se ilustra en el ejemplo siguiente, para agregar tipos compuestos a las operaciones de servicio.
<DataContract()> _
Public Class CompositeType

    Private boolValueField As Boolean
    Private stringValueField As String

    <DataMember()> _
    Public Property BoolValue() As Boolean
        Get
            Return Me.boolValueField
        End Get
        Set(ByVal value As Boolean)
            Me.boolValueField = value
        End Set
    End Property

    <DataMember()> _
    Public Property StringValue() As String
        Get
            Return Me.stringValueField
        End Get
        Set(ByVal value As String)
            Me.stringValueField = value
        End Set
    End Property

End Class
