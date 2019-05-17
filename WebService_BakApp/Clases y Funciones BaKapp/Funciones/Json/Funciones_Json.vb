Imports Newtonsoft.Json

Module Funciones_Json

    Public Function Fx_de_Json_a_Datatable(ByVal _Json As String) As DataTable
        ' Const _Json As String = "[{""Name"":""AAA"",""Age"":""22"",""Job"":""PPP""}," & "{""Name"":""BBB"",""Age"":""25"",""Job"":""QQQ""}," & "{""Name"":""CCC"",""Age"":""38"",""Job"":""RRR""}]"
        Dim table = JsonConvert.DeserializeObject(Of DataTable)(_Json)

        ' _Json = "{'Tabla 1': [{'id': 0,'elemento': 'elemento 0'},{'id': 1,'artículo': 'artículo 1'}]}"

        ' Dim Ds = JsonConvert.DeserializeObject(Of DataSet)(_Json)
        Dim dataSet As DataSet = JsonConvert.DeserializeObject(Of DataSet)(_Json)
        Return table

    End Function



End Module
