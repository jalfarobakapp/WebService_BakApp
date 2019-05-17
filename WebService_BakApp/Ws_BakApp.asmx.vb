Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports Newtonsoft.Json

Imports System.Web
Imports System.Web.Script.Services
Imports System.Web.Script.Serialization

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la siguiente línea.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://BakApp")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Ws_BakApp
    Inherits System.Web.Services.WebService

    Dim _Sql As Class_SQL '(_Global_Cadena_Conexion_SQL_Server)

    Public Sub New()
        _Global_Cadena_Conexion_SQL_Server = "data source = 192.168.0.75; initial catalog = RANDOM; user id = RANDOM; password = RANDOM"
    End Sub

    <WebMethod()> _
    Public Function Fx_Probar_Conexion_BD() As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Error As String = _Sql.Fx_Probar_Conexion
        Return _Error '"Hola a todos" 'http://localhost:34553
    End Function

    <WebMethod()> _
   Public Function Fx_Cadena_Conexion(ByVal Cadena_Conexion_SQL_Server As String) As String
        '_Global_Cadena_Conexion_SQL_Server = Cadena_Conexion_SQL_Server
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Error As String = _Sql.Fx_Probar_Conexion
        Return _Error
    End Function

    <WebMethod(True)> _
    Function Fx_GetDataSet(ByVal Consulta_Sql As String) As DataSet
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_Sql)
        Return _Ds
    End Function



    <WebMethod(True)> _
    Function Fx_Trae_Dato_String(ByVal _Tabla As String, _
                                 ByVal _Campo As String, _
                                 ByVal _Condicion As String) As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Dato As String = _Sql.Fx_Trae_Dato(_Tabla, _Campo, _Condicion, , False, "")
        Return _Dato
    End Function

    <WebMethod(True)> _
    Function Fx_Trae_Dato_Numero(ByVal _Tabla As String, _
                                 ByVal _Campo As String, _
                                 ByVal _Condicion As String) As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Dato As Double = _Sql.Fx_Trae_Dato(_Tabla, _Campo, _Condicion, , True, 0)
        Return _Dato
    End Function

    <WebMethod(True)> _
    Function Fx_Ej_consulta_IDU(ByVal Consulta_Sql As String) As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)

        If _Sql.Fx_Ej_consulta_IDU(Consulta_Sql) Then
            Return ""
        Else
            Return _Sql.Pro_Error
        End If
    End Function

    <WebMethod(True)> _
    Function Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(ByVal Consulta_Sql As String) As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)

        If _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_Sql) Then
            Return ""
        Else
            Return _Sql.Pro_Error
        End If
    End Function

    <WebMethod(True)> _
    Function Fx_Cuenta_Registros(ByVal _Tabla As String, _
                                 ByVal _Condicion As String) As Double
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Dato As Double = _Sql.Fx_Cuenta_Registros(_Tabla, _Condicion)
        Return _Dato
    End Function

    <WebMethod(True)> _
    Function Fx_Crear_Documento(ByVal _Global_BaseBk As String, _
                                ByVal _Funcionario As String, _
                                ByVal _Tido As String, _
                                ByVal _Nudo As String, _
                                ByVal _Es_ValeTransitorio As Boolean, _
                                ByVal _EsElectronico As Boolean, _
                                ByVal _Ds_Matriz_Documento As DataSet, _
                                ByVal _Es_Ajuste As Boolean) As String

        Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

        Dim _Idmaeedo As String
        _Idmaeedo = _New_Doc.Fx_Crear_Documento(_Tido, _
                                                    _Nudo, _
                                                    _Es_ValeTransitorio, _
                                                    _EsElectronico, _
                                                    _Ds_Matriz_Documento, _
                                                    _Es_Ajuste)

        Return _Idmaeedo

    End Function

    <WebMethod(True)> _
   Function Fx_Editar_Documento(ByVal _Global_BaseBk As String, _
                                ByVal _Idmaeedo_Dori As Integer, _
                                ByVal _Funcionario As String, _
                                ByVal _Ds_Matriz_Documento As DataSet) As Integer

        Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

        Dim _Idmaeedo As Integer
        _Idmaeedo = _New_Doc.Fx_Editar_Documento(_Idmaeedo_Dori, _Funcionario, _Ds_Matriz_Documento)

        Return _Idmaeedo

    End Function

    <WebMethod(True)> _
    Function Fx_Cambiar_Numeracion_Modalidad(ByVal _Tido As String, _
                                             ByVal _Nudo As String, _
                                             ByVal _Modalidad As String) As Double
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Dato As Double = Fx_Cambiar_Numeracion_Modalidad(_Tido, _Nudo, _Modalidad)  '_Sql.Fx_Cuenta_Registros(_Tabla, _Condicion)
        Return _Dato
    End Function

    Enum _Enum_Accion_EA
        Anular
        Eliminar
        Modificar
    End Enum

    <WebMethod(True)> _
   Function Fx_EliminarAnular_Doc(ByVal _Idmaeedo_Dori As Integer, _
                                  ByVal _Funcionario As String, _
                                  ByVal _Accion As _Enum_Accion_EA) As Boolean
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)

        Dim Cl_ClarDoc As New Clase_EliminarAnular_Documento

        If Cl_ClarDoc.Fx_EliminarAnular_Doc(_Idmaeedo_Dori, _
                                            _Funcionario, _
                                            _Accion, _
                                            False) Then
            Return True
        End If

    End Function

    <WebMethod(True)> _
   Function Fx_Traer_Numero_Documento(ByVal _Tido As String, _
                                      ByVal _NumeroDoc As String, _
                                      ByVal _Modalidad_Seleccionada As String, _
                                      ByVal _Empresa As String) As String
        Dim _NroDocumento As String = Traer_Numero_Documento(_Tido, _NumeroDoc, _Modalidad_Seleccionada, _Empresa)

        Return _NroDocumento
    End Function

    <WebMethod(True)> _
    Function Fx_Cadena_Conexion_SQL() As String
        Return System.Configuration.ConfigurationManager.ConnectionStrings("db_central").ToString()
    End Function

    <WebMethod(True)> _
    Function Fx_Conectado_Web_Service() As Boolean
        Return True
    End Function

    <WebMethod(True)> _
   Function Fx_Login_Usuario_Soap(ByVal _Clave As String) As DataSet

        Dim _Pw = Fx_TraeClaveRD(_Clave)

        Consulta_sql = "Select Top 1 KOFU,NOKOFU From TABFU Where PWFU = '" & _Pw & "'"
        _Sql = New Class_SQL

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Return _Ds

    End Function

#Region "JSON"

    <WebMethod(True)> _
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)> _
    Public Sub Sb_Login_Usuario_Json(ByVal _Clave As String)

        Dim js As New JavaScriptSerializer

        Dim _Pw = Fx_TraeClaveRD(_Clave)

        Consulta_sql = "Select Top 1 * From TABFU Where PWFU = '" & _Pw & "'"
        _Sql = New Class_SQL

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)> _
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)> _
    Public Sub Sb_Ds_Json(ByVal Key As String, ByVal _Consulta_Sql As String)

        Dim js As New JavaScriptSerializer

        _Sql = New Class_SQL
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)> _
    Public Sub Sb_Ds_Json_Prueba(ByVal Consulta_Sql As String)

        'Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = 670916" & vbCrLf & _
        '               "Select * From MAEDDO Where IDMAEEDO = 670916"

        Dim js As New JavaScriptSerializer

        _Sql = New Class_SQL
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_Sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)> _
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)> _
    Public Sub Sb_GetDataSet_Json(ByVal Consulta_Sql As String)

        Dim js As New JavaScriptSerializer

        _Sql = New Class_SQL
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_Sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)> _
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)> _
    Public Sub Sb_Buscar_Productos_Json(ByVal _Codigo As String, _
                                        ByVal _Descripcion As String)

        Dim js As New JavaScriptSerializer

        _Sql = New Class_SQL
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)> _
  Public Sub Sb_Json2Ds(ByVal _Json As String)

        If String.IsNullOrEmpty(_Json) Then
            _Json = "{'Tabla 1': [{""Name"":""AAA"",""Age"":""22"",""Job"":""PPP""}," &
                             "{""Name"":""BBB"",""Age"":""25"",""Job"":""QQQ""}," &
                             "{""Name"":""CCC"",""Age"":""38"",""Job"":""RRR""}]}"

        End If

        Dim _Tbl As DataTable = Fx_de_Json_a_Datatable(_Json)

    End Sub

#End Region





End Class