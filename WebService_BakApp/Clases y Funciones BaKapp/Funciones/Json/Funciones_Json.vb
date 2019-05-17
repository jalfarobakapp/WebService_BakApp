Imports Newtonsoft.Json

Module Funciones_Json

    Public Function Fx_de_Json_a_Datatable(ByVal _Json As String) As DataTable

        _Json = Replace(_Json, """", "'")

        Dim dataSet As DataSet = JsonConvert.DeserializeObject(Of DataSet)(_Json)

    End Function



End Module
