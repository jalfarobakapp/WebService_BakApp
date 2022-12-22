Public Module Variables

    Public Consulta_sql As String
    Public _Global_Cadena_Conexion_SQL_Server As String
    Public _Global_BaseBk As String

    Public Property Global_BaseBk As String
        Get

            Dim _Sql As Class_SQL
            _Sql = New Class_SQL

            Consulta_sql = "Select Top 1 *,NOKOCARAC+'.dbo.' As Global_BaseBk From TABCARAC Where KOTABLA = 'BAKAPP' And KOCARAC = 'BASE'"
            Dim _Row_Tabcarac = _Sql.Fx_Get_DataRow(Consulta_sql)

            _Global_BaseBk = _Row_Tabcarac.Item("Global_BaseBk")

            Return _Global_BaseBk
        End Get
        Set(value As String)
            _Global_BaseBk = value
        End Set
    End Property
End Module
