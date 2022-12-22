' NOTA: si cambia aquí el nombre de clase "Service1", también debe actualizar la referencia a "Service1" tanto en Web.config como en el archivo .svc asociado.
Public Class Service1
    Implements IService1

    Public Sub New()
    End Sub

    Public Function GetData(value As Integer) As String Implements IService1.GetData
        Return String.Format("You entered: {0}", value)
    End Function

    Public Function GetDataUsingDataContract(composite As CompositeType) As CompositeType Implements IService1.GetDataUsingDataContract
        If composite.BoolValue Then
            composite.StringValue = (composite.StringValue & "Suffix")
        End If
        Return composite
    End Function

End Class
