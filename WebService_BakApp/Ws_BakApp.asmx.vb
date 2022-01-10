Imports System.ComponentModel
Imports System.Web.Script.Serialization
Imports System.Web.Script.Services
Imports System.Web.Services
Imports Newtonsoft.Json

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la siguiente línea.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://BakApp")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class Ws_BakApp
    Inherits System.Web.Services.WebService

    Dim _Sql As Class_SQL
    Dim _Row_Tabcarac As DataRow

    Public Sub New()

        _Global_Cadena_Conexion_SQL_Server = "data source = 192.168.0.75; initial catalog = RANDOM; user id = RANDOM; password = RANDOM"

    End Sub

    <WebMethod()>
    Public Function Fx_Probar_Conexion_BD() As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Error As String = _Sql.Fx_Probar_Conexion
        Return _Error '"Hola a todos" 'http://localhost:34553
    End Function

    <WebMethod()>
    Public Function Fx_Cadena_Conexion(ByVal Cadena_Conexion_SQL_Server As String) As String
        '_Global_Cadena_Conexion_SQL_Server = Cadena_Conexion_SQL_Server
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Error As String = _Sql.Fx_Probar_Conexion
        Return _Error
    End Function

    <WebMethod(True)>
    Function Fx_GetDataSet(ByVal Consulta_Sql As String) As DataSet
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_Sql)
        Return _Ds
    End Function

    <WebMethod(True)>
    Function Fx_Trae_Dato_String(ByVal _Tabla As String,
                                 ByVal _Campo As String,
                                 ByVal _Condicion As String) As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Dato As String = _Sql.Fx_Trae_Dato(_Tabla, _Campo, _Condicion, , False, "")
        Return _Dato
    End Function

    <WebMethod(True)>
    Function Fx_Trae_Dato_Numero(ByVal _Tabla As String,
                                 ByVal _Campo As String,
                                 ByVal _Condicion As String) As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Dato As Double = _Sql.Fx_Trae_Dato(_Tabla, _Campo, _Condicion, , True, 0)
        Return _Dato
    End Function

    <WebMethod(True)>
    Function Fx_Ej_consulta_IDU(ByVal Consulta_Sql As String) As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)

        If _Sql.Fx_Ej_consulta_IDU(Consulta_Sql) Then
            Return ""
        Else
            Return _Sql.Pro_Error
        End If
    End Function

    <WebMethod(True)>
    Function Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(ByVal Consulta_Sql As String) As String
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)

        If _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_Sql) Then
            Return ""
        Else
            Return _Sql.Pro_Error
        End If
    End Function

    <WebMethod(True)>
    Function Fx_Cuenta_Registros(ByVal _Tabla As String,
                                 ByVal _Condicion As String) As Double
        _Sql = New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
        Dim _Dato As Double = _Sql.Fx_Cuenta_Registros(_Tabla, _Condicion)
        Return _Dato
    End Function

    <WebMethod(True)>
    Function Fx_Crear_Documento(ByVal _Global_BaseBk As String,
                                ByVal _Funcionario As String,
                                ByVal _Tido As String,
                                ByVal _Nudo As String,
                                ByVal _Es_ValeTransitorio As Boolean,
                                ByVal _EsElectronico As Boolean,
                                ByVal _Ds_Matriz_Documento As DataSet,
                                ByVal _Es_Ajuste As Boolean) As String

        Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

        Dim _Idmaeedo As String
        _Idmaeedo = _New_Doc.Fx_Crear_Documento(_Tido,
                                                    _Nudo,
                                                    _Es_ValeTransitorio,
                                                    _EsElectronico,
                                                    _Ds_Matriz_Documento,
                                                    _Es_Ajuste)

        Return _Idmaeedo

    End Function

    <WebMethod(True)>
    Function Fx_Editar_Documento(ByVal _Global_BaseBk As String,
                                ByVal _Idmaeedo_Dori As Integer,
                                ByVal _Funcionario As String,
                                ByVal _Ds_Matriz_Documento As DataSet) As Integer

        Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

        Dim _Idmaeedo As Integer
        _Idmaeedo = _New_Doc.Fx_Editar_Documento(_Idmaeedo_Dori, _Funcionario, _Ds_Matriz_Documento)

        Return _Idmaeedo

    End Function

    <WebMethod(True)>
    Function Fx_Cambiar_Numeracion_Modalidad(ByVal _Tido As String,
                                             ByVal _Nudo As String,
                                             ByVal _Modalidad As String) As Double
        _Sql = New Class_SQL
        Dim _Dato As Double = Fx_Cambiar_Numeracion_Modalidad(_Tido, _Nudo, _Modalidad)
        Return _Dato
    End Function

    Enum _Enum_Accion_EA
        Anular
        Eliminar
        Modificar
    End Enum

    <WebMethod(True)>
    Function Fx_EliminarAnular_Doc(ByVal _Idmaeedo_Dori As Integer,
                                  ByVal _Funcionario As String,
                                  ByVal _Accion As _Enum_Accion_EA) As Boolean
        _Sql = New Class_SQL

        Dim Cl_ClarDoc As New Clase_EliminarAnular_Documento

        If Cl_ClarDoc.Fx_EliminarAnular_Doc(_Idmaeedo_Dori,
                                            _Funcionario,
                                            _Accion,
                                            False) Then
            Return True
        End If

    End Function

    <WebMethod(True)>
    Function Fx_Traer_Numero_Documento(ByVal _Tido As String,
                                      ByVal _NumeroDoc As String,
                                      ByVal _Modalidad_Seleccionada As String,
                                      ByVal _Empresa As String) As String
        Dim _NroDocumento As String = Traer_Numero_Documento(_Tido, _NumeroDoc, _Modalidad_Seleccionada, _Empresa)

        Return _NroDocumento
    End Function

    <WebMethod(True)>
    Function Fx_Cadena_Conexion_SQL() As String
        Return System.Configuration.ConfigurationManager.ConnectionStrings("db_bakapp").ToString()
    End Function

    <WebMethod(True)>
    Function Fx_Conectado_Web_Service() As Boolean
        Return True
    End Function

    <WebMethod(True)>
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

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
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

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_GetModalidad_Gral(Global_BaseBk As String)

        Dim Consulta_Sql As String
        Consulta_Sql = "Select " & vbCrLf &
                       "Empresa, Pr_AutoPr_Crear_Codigo_Principal_Automatico, Pr_AutoPr_Correlativo_Por_Iniciales, Pr_AutoPr_Correlativo_General, " & vbCrLf &
                       "Pr_AutoPr_Tablas_Para_Iniciales_Cod_Automatico, Pr_AutoPr_Max_Cant_Caracteres_Del_Codigo, Pr_AutoPr_Ultimo_Codigo_Creado_Correlativo_General, " & vbCrLf &
                       "Pr_Desc_Producto_Solo_Mayusculas, Pr_Creacion_Exigir_Precio, Pr_Creacion_Exigir_Clasificacion_busqueda, Pr_Creacion_Exigir_Codigo_Alternativo, " & vbCrLf &
                       "Tbl_Ranking, Revisa_Taza_Cambio, Revisar_Taza_Solo_Mon_Extranjeras, Vnta_Dias_Venci_Coti, Vnta_TipoValor_Bruto_Neto, Vnta_EntidadXdefecto, " & vbCrLf &
                       "Vnta_SucEntXdefecto, Vnta_Producto_NoCreado, Vnta_Preguntar_Documento, SOC_CodTurno, SOC_Buscar_Producto, SOC_Aprueba_Solo_G1, " & vbCrLf &
                       "SOC_Aprueba_G1_y_G2, SOC_Prod_Crea_Solo_Marcas_Proveedor, SOC_Prod_Crea_Max_Carac_Nom, SOC_Valor_1ra_Aprobacion, SOC_Dias_Apela, " & vbCrLf &
                       "SOC_Tipo_Creacion_Producto_Normal_Matriz, Precio_Costos_Desde, Precios_Venta_Desde_Random, Precios_Venta_Desde_BakApp, " & vbCrLf &
                       "Vnta_Redondear_Dscto_Cero, Nodo_Raiz_Asociados, Vnta_Ofrecer_Otras_Bod_Stock_Insuficiente, Conservar_Responzable_Doc_Relacionado, " & vbCrLf &
                       "Preguntar_Si_Cambia_Responsable_Doc_Relacionado, ServTecnico_Empresa, ServTecnico_Sucursal, ServTecnico_Bodega" &
                       vbCrLf &
                       "Into #Paso" & vbCrLf &
                       vbCrLf &
                       "From " & Global_BaseBk & "Zw_Configuracion" & vbCrLf &
                       "Where Modalidad_General = 1" & vbCrLf &
                       vbCrLf &
                       "Select * From #Paso" & vbCrLf &
                       "Drop Table #Paso"

        Consulta_Sql = "Select top 1 * From TABPA"

        Dim js As New JavaScriptSerializer

        _Sql = New Class_SQL
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_Sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

        '"Where Modalidad = 'CAJA'--Modalidad = '  '" & vbCrLf &

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Buscar_Productos_Json(ByVal _Codigo As String,
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

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Traer_Productos_Json(Codigo As String,
                                       Empresa As String,
                                       Sucursal As String,
                                       Bodega As String,
                                       Lista As String,
                                       UnTrans As Integer,
                                       Koen As String)

        Consulta_sql = My.Resources.Recursos_Sql.SqlQuery_Traer_Producto
        Consulta_sql = Replace(Consulta_sql, "#Codigo#", Codigo)
        Consulta_sql = Replace(Consulta_sql, "#Empresa#", Empresa)
        Consulta_sql = Replace(Consulta_sql, "#Sucursal#", Sucursal)
        Consulta_sql = Replace(Consulta_sql, "#Bodega#", Bodega)
        Consulta_sql = Replace(Consulta_sql, "#Lista#", Lista)
        Consulta_sql = Replace(Consulta_sql, "#UnTrans#", UnTrans)


        _Sql = New Class_SQL
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim _PorIla As Double = _Sql.Fx_Trae_Dato("TABIM", "Isnull(Sum(POIM),0)", "KOIM In (SELECT KOIM FROM TABIMPR Where KOPR = '" & Codigo & "')")

        Consulta_sql = "SELECT Top 1 *,--PP01UD,PP02UD,DTMA01UD As DSCTOMAX,ECUACION,
                        (SELECT top 1 MELT FROM TABPP Where KOLT = '" & Lista & "') as MELT FROM TABPRE
                        Where KOLT = '" & Lista & "' And KOPR = '" & Codigo & "'"
        Dim _RowPrecios As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        Dim _Ecuacion As String

        If UnTrans = 1 Then
            _Ecuacion = _RowPrecios.Item("ECUACION")
        Else
            _Ecuacion = _RowPrecios.Item("ECUACION2")
        End If

        Dim _DescMaximo = Fx_Precio_Formula_Random(Empresa, Sucursal, _RowPrecios, "DTMA0" & UnTrans & "UD", "EDTMA0" & UnTrans & "UD", Nothing, True, Koen)

        'Dim _Campo_Precio
        'Dim _Campo_Ecuacion

        '_PrecioLinea = Fx_Precio_Formula_Random(_RowPrecios, _Campo_Precio, _Campo_Ecuacion, Nothing, True, _Koen)

        Dim _Precio As Double
        'Dim _StockBodega As Double

        Dim _PrecioListaUd1 As Double = Fx_Precio_Formula_Random(Empresa, Sucursal, _RowPrecios, "PP01UD", "ECUACION", Nothing, True, Koen)
        Dim _PrecioListaUd2 As Double = Fx_Precio_Formula_Random(Empresa, Sucursal, _RowPrecios, "PP02UD", "ECUACIONU2", Nothing, True, Koen)

        If UnTrans = 1 Then
            _Precio = _PrecioListaUd1
        Else
            _Precio = _PrecioListaUd2
        End If

        '_StockBodega = _Ds.Tables(0).Rows(0).Item("StockUd" & UnTrans)

        _Ds.Tables(0).Rows(0).Item("Ecuacion") = _Ecuacion.Trim
        _Ds.Tables(0).Rows(0).Item("DescMaximo") = _DescMaximo
        _Ds.Tables(0).Rows(0).Item("Precio") = _Precio
        _Ds.Tables(0).Rows(0).Item("PrecioListaUd1") = _PrecioListaUd1
        _Ds.Tables(0).Rows(0).Item("PrecioListaUd2") = _PrecioListaUd2

        _Ds.Tables(0).Rows(0).Item("PorIla") = _PorIla

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Traer_Entidad_Json(Koen As String, Suen As String)

        _Sql = New Class_SQL
        Dim _Ds As DataSet = Fx_Traer_Datos_Entidad(Koen, Suen)

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_SQlIDU(_Key As String, _Query As String)

        _Sql = New Class_SQL

        Consulta_sql = _Query
        Dim _Ds As DataSet

        If _Sql.Fx_Ej_consulta_IDU(Consulta_sql, False) Then
            Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'' As Error"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'" & Replace(_Sql.Pro_Error, "'", "''") & "' As Error"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Revisar_Stock_Fila(_Tido As String,
                                     _Empresa As String,
                                     _Sucursal As String,
                                     _Bodega As String,
                                     _Codigo As String,
                                     _Cantidad As Double,
                                     _UnTrans As Integer,
                                     _Tidopa As String)

        'Dim _Cantidad As Double

        'For Each _Row As DataRow In _TblDetalle.Rows

        '    Dim _Cod = _Row.Item("Codigo")
        '    Dim _Suc = _Row.Item("Sucursal")
        '    Dim _Bod = _Row.Item("Bodega")
        '    Dim _I = _Row.Item("Id")

        '    If _Cod = _Codigo And _Suc = _Sucursal And _Bod = _Bodega Then
        '        _Cantidad += _Row.Item("Cantidad")
        '    End If

        'Next
        _Sql = New Class_SQL

        Dim _Stock_Disponible As Double
        Dim _Revisar_Stock_Disponible As Boolean = True
        Dim _Campo_Formula_Stock = String.Empty

        Consulta_sql = "Select Top 1 * From TABTIDO Where TIDO = '" & _Tido & "'"
        Dim _RowTido As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If Not IsNothing(_RowTido) Then
            _Campo_Formula_Stock = NuloPorNro(_RowTido.Item("STOCK"), "")
        End If

        If _Tido = "NVV" Or _Tido = "RES" Or _Tido = "PRO" Or _Tido = "NVI" Then

            _Revisar_Stock_Disponible = Not (String.IsNullOrEmpty(_Campo_Formula_Stock))

        End If

        If _Revisar_Stock_Disponible Then

            _Stock_Disponible = Fx_Stock_Disponible(_Tido, _Empresa, _Sucursal, _Bodega, _Codigo, _UnTrans, "STFI" & _UnTrans)

            If _Tidopa = "NVV" And _Tido <> "NVV" Then

                If _Campo_Formula_Stock.Contains("-C") Then
                    _Stock_Disponible += _Cantidad
                End If

            End If

        Else
            _Stock_Disponible = 1 + _Cantidad
        End If

        'If _Revision_Remota Then
        '    _Stock_Disponible += _Cantidad
        'End If

        Dim _Stock As Double
        'Dim _Stock_Suficiente As Boolean

        _Stock = _Sql.Fx_Trae_Dato("MAEST", "STFI" & _UnTrans, "EMPRESA = '" & _Empresa &
                                   "' AND KOSU = '" & _Sucursal &
                                   "' AND KOBO = '" & _Bodega &
                                   "' AND KOPR = '" & _Codigo & "'", True)


        _Sql = New Class_SQL

        Consulta_sql = "Select " & _Stock_Disponible & " As Stock_Disponible," & _Stock & " As Stock_Fisico"
        Dim _Ds As DataSet
        _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

        Return

        'If _Tipr <> "SSN" Then

        '    Dim _Cantidad_Resul As Double = _Stock_Disponible - _Cantidad

        '    If _Stock_Disponible <= 0 Then
        '        _Stock_Suficiente = False
        '    Else
        '        If _Stock_Disponible - _Cantidad >= 0 Then
        '            _Stock_Suficiente = True
        '        End If
        '    End If

        '    'If Not _Stock_Suficiente Then

        '    '    Dim _CodFunAutoriza_Stock As String = _TblEncabezado.Rows(0).Item("Fun_Auto_Stock_Ins")

        '    '    'Código permiso vender sin stock "Bkp00015"

        '    '    If _Mostrar_Alerta Then

        '    '        'If _Ofrecer_Bodegas Then

        '    '        '    Dim _Vnta_Ofrecer_Otras_Bod_Stock_Insuficiente As Boolean = _Global_Row_Configuracion_Estacion.Item("Vnta_Ofrecer_Otras_Bod_Stock_Insuficiente")

        '    '        '    If _Vnta_Ofrecer_Otras_Bod_Stock_Insuficiente Then

        '    '        '        Consulta_sql = "Select Distinct EMPRESA+KOSU+KOBO As Cod,* From TABBO
        '    '        '                        Where EMPRESA+KOSU+KOBO
        '    '        '                        In (Select SUBSTRING(CodPermiso, 3, 10)
        '    '        '                            From " & _Global_BaseBk & "ZW_PermisosVsUsuarios
        '    '        '                                Where CodUsuario = '" & FUNCIONARIO & "' And 
        '    '        '                                CodPermiso In (Select CodPermiso From " & _Global_BaseBk & "ZW_Permisos Where CodFamilia = 'Bodega')) 
        '    '        '                                Or (EMPRESA = '" & ModEmpresa & "' And KOSU = '" & ModSucursal & "' And KOBO = '" & ModBodega & "')"

        '    '        '        Dim _Tbl_Bodegas As DataTable = _Sql.Fx_Get_Tablas(Consulta_sql)

        '    '        '        Dim _Filtro As String = Generar_Filtro_IN(_Tbl_Bodegas, "", "Cod", False, False, "'")

        '    '        '        _Filtro = "KOPR = '" & _Codigo & "' And EMPRESA+KOSU+KOBO In " & _Filtro

        '    '        '        Dim _Stock_Consolidado As Double = _Sql.Fx_Trae_Dato("MAEST", "Sum(STFI1)", _Filtro)

        '    '        '        If _Stock_Consolidado > 0 Then

        '    '        '            Dim _Row_Bodega As DataRow

        '    '        '            If Fx_Tiene_Permiso(Me, "Bkp00045") Then

        '    '        '                If Fr_Alerta_Stock.Visible Then
        '    '        '                    Fr_Alerta_Stock.Close()
        '    '        '                End If

        '    '        '                _Cantidad = NuloPorNro(_Fila.Cells("Cantidad").Value, 0)

        '    '        '                Dim _Es_Venta As Boolean = (_Tipo_Documento = csGlobales.Mod_Enum_Listados_Globales.Enum_Tipo_Documento.Venta)

        '    '        '                Dim Fm As New Frm_Formulario_Cantidad_Stock_X_Bodega(_Codigo, _Cantidad, _Sucursal, _Es_Venta, _Tido)
        '    '        '                Fm.ShowDialog(Me)
        '    '        '                _Row_Bodega = Fm.Row_Bodega
        '    '        '                Fm.Dispose()

        '    '        '                If Not (_Row_Bodega Is Nothing) Then

        '    '        '                    _Fila.Cells("Sucursal").Value = _Row_Bodega.Item("KOSU")
        '    '        '                    _Fila.Cells("Bodega").Value = _Row_Bodega.Item("KOBO")
        '    '        '                    Sb_Revisar_Stock_Fila(_Fila,,,, True, False)
        '    '        '                    Exit Sub

        '    '        '                Else

        '    '        '                    _Fila.Cells("Cantidad").Value = 0
        '    '        '                    Return

        '    '        '                End If

        '    '        '            End If

        '    '        '        End If

        '    '        '    End If

        '    '        'End If

        '    '        If CBool(_Cantidad) Then

        '    '            If Fx_Tiene_Permiso(Me, "Bkp00015", _CodFunAutoriza_Stock, False) Then

        '    '                MessageBoxEx.Show(Me, "¡Producto con Stock insuficiente!" & Environment.NewLine &
        '    '                              "Stock en Bodega  " & _Bodega & ": " & _Stock & Environment.NewLine &
        '    '                              "Cantidad vendida : " & _Cantidad & Environment.NewLine &
        '    '                              "Diferencia: " & _Stock - _Cantidad & " " & _UdTrans,
        '    '                              "Stock insuficiente", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, Me.TopMost)

        '    '            Else

        '    '                MessageBoxEx.Show(Me, "¡Producto con Stock insuficiente!" & Environment.NewLine &
        '    '                                  "Stock en Bodega  " & _Bodega & ": " & _Stock & Environment.NewLine &
        '    '                                  "Cantidad vendida : " & _Cantidad & Environment.NewLine &
        '    '                                  "Diferencia: " & _Stock - _Cantidad & " " & _UdTrans & Environment.NewLine & Environment.NewLine &
        '    '                                  "¡No permite hacer ventas sin autorización!",
        '    '                                  "Stock insuficiente", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, Me.TopMost)
        '    '            End If

        '    '        End If

        '    '    End If

        '    'End If

        'End If

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Sub Sb_Traer_Descuentos_Seteados_Desde_Lista(_Empresa As String,
                                                 _Sucursa As String,
                                                 _Codigo As String,
                                                 _CodLista As String,
                                                 _Prct As Boolean,
                                                 _Tict As String,
                                                 _PorIva As Double,
                                                 _PorIla As Double,
                                                 _Koen As String,
                                                 _ChkValoresNeto As Boolean)

        Dim _TblDscto As DataTable

        _Sql = New Class_SQL

        Consulta_sql = "Select Top 1 TABPRE.*,TABPP.MELT From TABPRE 
                        Inner Join TABPP On TABPP.KOLT = TABPRE.KOLT
                        Where TABPRE.KOLT = '" & _CodLista & "' And KOPR = '" & _Codigo & "'"

        Dim _Row_Tabpre As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)


        Consulta_sql = "Select Top 1 *,(SELECT top 1 MELT FROM TABPP Where KOLT = '" & _CodLista & "') As MELT From TABPRE
                        Where KOLT = '" & _CodLista & "' And KOPR = '" & _Codigo & "'"
        Dim _RowPrecios As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        Consulta_sql = "Declare @Campos Int

                        Set @Campos =(Select Count(*) From PNOMDIM Where DEPENDENCI = 'Por_tabpp' And NOMBRE = CODIGO And CODIGO <> 'PASO01P')

                        Select TOP 1 OPERA,Cast(SUBSTRING( OPERA,28,@Campos*2) As Varchar(200)) As Opera_Rev 
                        INTO #Paso
                        From TABPP Where OPERA LIKE 'XX%'

                         Update #Paso Set Opera_Rev = REPLACE(Opera_Rev,'  ','Dp,')
                         Update #Paso Set Opera_Rev = REPLACE(Opera_Rev,' 1','Dv,')
                         Update #Paso Set Opera_Rev = REPLACE(Opera_Rev,' 2','Rp,')
                         Update #Paso Set Opera_Rev = REPLACE(Opera_Rev,' 3','Rv,')

                        Select * From #Paso
                        Drop Table #Paso"

        Dim _Row_Opera As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If Not IsNothing(_Row_Opera) Then

            Dim _Opera = _Row_Opera.Item("Opera_Rev")
            Dim _Opera_Rev = Split(_Opera, ",")

            Consulta_sql = "Select Top 1 * From TABPRE Where KOLT = '" & _CodLista & "' And KOPR = '" & _Codigo & "'"
            Dim _TblTabpre As DataTable = _Sql.Fx_Get_Tablas(Consulta_sql)

            ' Asi es como actua el campo OPERA, este campo define como se comportaran los campos adicionales a partir del campo nro 29 en adelante

            '       28 - Ultimo Campo EDTMA02UD

            '        '' = Descuento Porcentaje
            '        1  = Descuento Valor
            '        2  = Recargo Porcentaje
            '        3  = Recargo Valor

            Dim _Campos_Adicionales = String.Empty
            Dim _j = 0
            Dim _a = 0

            Consulta_sql = "Select Cast('' As Varchar(20)) As Tcampo,Cast(0 As Float) As Dscto,Cast(0 As Float) As Valor" & Environment.NewLine &
                           "Where 1 < 0"
            _TblDscto = _Sql.Fx_Get_Tablas(Consulta_sql)

            For _i = 28 To _TblTabpre.Columns.Count - 1

                Dim _Columna As DataColumn = _TblTabpre.Columns(_i)
                Dim _Nombre_Columna As String = _Columna.ColumnName

                If Mid(_Nombre_Columna, 1, 5) <> "FORM_" Then

                    '_Campos_Adicionales += "[" & _Nombre_Columna & "] [CHAR] (121) Default ''," & Environment.NewLine

                    Dim _Valor As Double = NuloPorNro(_Row_Tabpre.Item(_Nombre_Columna), 0)
                    Dim _Valor_Fx As Double

                    'Dim _Koen = _TblEncabezado.Rows(0).Item("CodEntidad")

                    Dim _Campo_Ecuacion As String = "FORM_" & numero_(_i + 1, 3)

                    _Valor_Fx = Fx_Precio_Formula_Random(_Empresa, _Sucursa, _RowPrecios, _Nombre_Columna, _Campo_Ecuacion, Nothing, True, _Koen)

                    If _Valor = 0 Then _Valor = _Valor_Fx

                    'Dim _DocEn_Neto_Bruto = _TblEncabezado.Rows(0).Item("DocEn_Neto_Bruto")

                    If CBool(_Valor) Then

                        Dim _TCampo = _Opera_Rev(_j)
                        Dim _Dscto As Double
                        Dim _Incorporar_Dscto As Boolean

                        Select Case _TCampo
                            Case "Dp"
                                'Porcentaje
                                '_Array_Dsctos(_a, 0) = _TCampo
                                '_Array_Dsctos(_a, 1) = _Valor
                                '_Array_Dsctos(_a, 2) = 0
                                _Dscto = _Valor
                                _Valor = 0
                                _Incorporar_Dscto = True
                            Case "Dv" ', "Rv"
                                _Valor = _Valor '* -1
                                _Dscto = 0
                                _Incorporar_Dscto = True
                                '_Array_Dsctos(_a, 0) = _TCampo
                                '_Array_Dsctos(_a, 1) = 0
                                '_Array_Dsctos(_a, 2) = _Valor
                            Case "Rp"

                            Case "Rv"

                                'If _Prct And _Tict = "R" Then

                                Dim _Iva = _PorIva / 100
                                Dim _Ila = _PorIla / 100

                                Dim _Impuestos As Double = 1 + (_Iva + _Ila)

                                Dim _Melt = _Row_Tabpre.Item("MELT")

                                If _Melt = "B" Then

                                    If _ChkValoresNeto Then
                                        _Valor = Math.Round(_Valor / _Impuestos, 3)
                                    Else
                                        _Valor = _Valor * _Impuestos
                                    End If

                                Else

                                    If Not _ChkValoresNeto Then
                                        _Valor = Math.Round(_Valor * _Impuestos, 0)
                                    End If

                                End If

                                '_Fila.Cells("Recargo_Campo").Value = _Nombre_Columna
                                '_Fila.Cells("Recargo_Valor").Value = _Valor

                        End Select

                        If _Incorporar_Dscto Then

                            Dim NewFila As DataRow
                            NewFila = _TblDscto.NewRow
                            With NewFila
                                .Item("TCampo") = _TCampo
                                .Item("Dscto") = _Dscto
                                .Item("Valor") = _Valor
                                _TblDscto.Rows.Add(NewFila)
                            End With

                        End If

                    End If

                    _j += 1

                End If

                'ReDim Preserve _Array_Dsctos(1, 2)

            Next

            Dim _Ds As New DataSet

            If CBool(_TblDscto.Rows.Count) Then
                _Ds.Tables.Add(_TblDscto)
            Else
                Consulta_sql = "Select Cast('' As Varchar(20)) As Tcampo,Cast(0 As Float) As Dscto,Cast(0 As Float) As Valor"
                'Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'Sin Datos...' As Error"
                _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
            End If

            _Ds.Tables(0).TableName = "Table"

            Dim js As New JavaScriptSerializer

            Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
            Context.Response.ContentType = "application/json"
            Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
            Context.Response.Flush()

            Context.Response.End()

        End If

    End Sub

    <WebMethod(True)>
    Public Sub Sb_Json2Ds(ByVal _Json As String)

        If String.IsNullOrEmpty(_Json) Then
            _Json = "{'Tabla 1': [{""Name"":""AAA"",""Age"":""22"",""Job"":""PPP""}," &
                             "{""Name"":""BBB"",""Age"":""25"",""Job"":""QQQ""}," &
                             "{""Name"":""CCC"",""Age"":""38"",""Job"":""RRR""}]}"

        End If

        Dim _Tbl As DataTable = Fx_de_Json_a_Datatable(_Json)

    End Sub

    <WebMethod(True)>
    Public Sub Sb_Json_ImpBk(_Json As String, _NombreTabla As String)

        '_Json = Replace(_Json, """", "'")
        _Json = Replace(_Json, "\/", "/")
        _Json = _Json.Trim
        Dim _Json2 = Mid(_Json, 2, _Json.Length - 1)
        _Json2 = Mid(_Json2, 1, _Json2.Length - 1)

        Dim RutaArchivo As String = "D:\JsonB4Android\" & _NombreTabla & ".json"
        Dim Cuerpo = _Json

        Dim oSW As New System.IO.StreamWriter(RutaArchivo)

        oSW.WriteLine(Cuerpo)
        oSW.Close()

        Dim dataSet As DataSet = JsonConvert.DeserializeObject(Of DataSet)(_Json2)

        _Sql = New Class_SQL
        Dim _Ds As DataSet

        Dim _Existe = System.IO.File.Exists(RutaArchivo)

        If Not _Existe Then
            Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'' As Error"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'" & Replace(_Sql.Pro_Error, "'", "''") & "' As Error"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    '#End Region

    'Dim _Existe = System.IO.File.Exists(RutaArchivo)

    'If Not _Existe Then
    '    Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'' As Error"
    '    _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
    'Else
    '    Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'" & Replace(_Sql.Pro_Error, "'", "''") & "' As Error"
    '    _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
    'End If

    'Dim js As New JavaScriptSerializer

    'Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
    'Context.Response.ContentType = "application/json"
    'Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
    'Context.Response.Flush()

    'Context.Response.End()

    'End Sub


    <WebMethod(True)>
    Public Sub Sb_CreaDocumentoJsonBakapp(_EncabezadoJs As String, _DestalleJs As String, _DescuentosJs As String, _ObservacionesJs As String)

        _EncabezadoJs = _EncabezadoJs.Trim
        _DestalleJs = _DestalleJs.Trim
        _ObservacionesJs = _ObservacionesJs.Trim

        Dim _Ds_Matriz_Documentos As New Ds_Matriz_Documentos

        _Ds_Matriz_Documentos.Clear()
        _Ds_Matriz_Documentos = New Ds_Matriz_Documentos

        Fx_LlenarDatos(_Ds_Matriz_Documentos, _EncabezadoJs, "Encabezado_Doc")
        _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Post_Venta") = False
        _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Tipo_Documento") = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")

        Fx_LlenarDatos(_Ds_Matriz_Documentos, _DestalleJs, "Detalle_Doc")
        Fx_LlenarDatos(_Ds_Matriz_Documentos, _DescuentosJs, "Descuentos_Doc")
        Fx_LlenarDatos(_Ds_Matriz_Documentos, _ObservacionesJs, "Observaciones_Doc")

        Dim _Funcionario As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("CodFuncionario")
        _Global_BaseBk = "BAKAPP_VH.dbo."
        Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

        _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("NroDocumento") = "NVVXXXXXXX"

        Dim _Idmaeedo As String
        _Idmaeedo = _New_Doc.Fx_Crear_Documento_En_BakApp_Casi2("Bakapp4ndroid", _Ds_Matriz_Documentos, False, True, "B4A")

#Region "Insercion obsoleta"


        'Dim _Json = Mid(_EncabezadoJs, 2, _EncabezadoJs.Length - 1)
        '_Json = Mid(_Json, 1, _Json.Length - 1)

        'Dim _Ds As DataSet = JsonConvert.DeserializeObject(Of DataSet)(_Json)
        'Dim _RowEncabezado As DataRow = _Ds.Tables(0).Rows(0)

        'Dim NewFila As DataRow

        'With _RowEncabezado

        '    Dim _TblEncabezado As DataTable = _Ds_Matriz_Documentos.Tables("Encabezado_Doc")

        '    NewFila = _TblEncabezado.NewRow

        '    With NewFila

        '        .Item("Id_DocEnc") = 0
        '        .Item("Post_Venta") = False
        '        .Item("Tipo_Documento") = _RowEncabezado.Item("TipoDoc")
        '        .Item("Modalidad") = _RowEncabezado.Item("Modalidad")
        '        .Item("Empresa") = _RowEncabezado.Item("Empresa")
        '        .Item("Sucursal") = _RowEncabezado.Item("Sucursal")
        '        .Item("TipoDoc") = _RowEncabezado.Item("TipoDoc")
        '        .Item("SubTido") = _RowEncabezado.Item("Subtido")
        '        .Item("NroDocumento") = _RowEncabezado.Item("NroDocumento")
        '        .Item("CodEntidad") = _RowEncabezado.Item("CodEntidad")
        '        .Item("CodSucEntidad") = _RowEncabezado.Item("CodSucEntidad")
        '        .Item("CodSucEntidadFisica") = _RowEncabezado.Item("CodSucEntidadFisica")
        '        .Item("CodSucEntidadFisica") = _RowEncabezado.Item("CodSucEntidadFisica")
        '        .Item("ListaPrecios") = _RowEncabezado.Item("ListaPrecios")
        '        .Item("CodFuncionario") = _RowEncabezado.Item("CodFuncionario")
        '        .Item("NomFuncionario") = _RowEncabezado.Item("NomFuncionario")
        '        .Item("Es_Electronico") = CBool(_RowEncabezado.Item("Es_Electronico"))

        '        .Item("FechaEmision") = Fx_FechaStr2Datetime(_RowEncabezado.Item("FechaEmision"))
        '        .Item("Fecha_1er_Vencimiento") = Fx_FechaStr2Datetime(_RowEncabezado.Item("Fecha_1er_Vencimiento"))
        '        .Item("FechaUltVencimiento") = Fx_FechaStr2Datetime(_RowEncabezado.Item("FechaUltVencimiento"))
        '        .Item("FechaRecepcion") = Fx_FechaStr2Datetime(_RowEncabezado.Item("FechaRecepcion"))
        '        .Item("FechaMaxRecepcion") = .Item("FechaRecepcion")

        '        .Item("Cuotas") = _RowEncabezado.Item("Cuotas")

        '        .Item("Dias_1er_Vencimiento") = _RowEncabezado.Item("Dias_1er_Vencimiento")
        '        .Item("Dias_Vencimiento") = _RowEncabezado.Item("Dias_Vencimiento")
        '        .Item("Moneda_Doc") = _RowEncabezado.Item("Moneda_Doc")
        '        .Item("Valor_Dolar") = _RowEncabezado.Item("Valor_Dolar")
        '        .Item("DocEn_Neto_Bruto") = _RowEncabezado.Item("DocEn_Neto_Bruto")
        '        .Item("TipoMoneda") = _RowEncabezado.Item("TipoMoneda")
        '        .Item("Tasadorig_Doc") = _RowEncabezado.Item("Tasadorig_Doc")
        '        .Item("Centro_Costo") = _RowEncabezado.Item("Centro_Costo")
        '        .Item("Contacto_Ent") = _RowEncabezado.Item("Contacto_Ent")
        '        .Item("TotalNetoDoc") = _RowEncabezado.Item("TotalNetoDoc")
        '        .Item("TotalIvaDoc") = _RowEncabezado.Item("TotalIvaDoc")
        '        .Item("TotalIlaDoc") = _RowEncabezado.Item("TotalIlaDoc")
        '        .Item("TotalBrutoDoc") = _RowEncabezado.Item("TotalBrutoDoc")
        '        .Item("CantTotal") = _RowEncabezado.Item("CantTotal")
        '        .Item("CantDesp") = _RowEncabezado.Item("CantDesp")
        '        .Item("Vizado") = CBool(_RowEncabezado.Item("Vizado"))
        '        .Item("Idmaeedo_Origen") = 0
        '        .Item("CodUsuario_Permiso_Dscto") = String.Empty
        '        .Item("Fun_Auto_Deuda_Ven") = _RowEncabezado.Item("Fun_Auto_Deuda_Ven")
        '        .Item("Fun_Auto_Stock_Ins") = _RowEncabezado.Item("Fun_Auto_Stock_Ins")
        '        .Item("Fun_Auto_Cupo_Exe") = _RowEncabezado.Item("Fun_Auto_Cupo_Exe")
        '        .Item("Fun_Auto_Fecha_Des") = _RowEncabezado.Item("Fun_Auto_Fecha_Des")

        '        .Item("Stand_by") = False
        '        .Item("Libro") = _RowEncabezado.Item("Libro")
        '        .Item("Fecha_Tributaria") = Fx_FechaStr2Datetime(_RowEncabezado.Item("Fecha_Tributaria"))

        '        .Item("Reserva_NroOCC") = False
        '        .Item("Reserva_Idmaeedo") = 0

        '        .Item("Bodega_Destino") = _RowEncabezado.Item("Bodega_Destino")
        '        .Item("TotalKilos") = 0

        '        .Item("MinKgDesp") = 0
        '        .Item("MinNetoDesp") = 0

        '        _TblEncabezado.Rows.Add(NewFila)

        '    End With

        'End With

        '_Json = Mid(_ObservacionesJs, 2, _ObservacionesJs.Length - 1)
        '_Json = Mid(_Json, 1, _Json.Length - 1)

        '_Ds = JsonConvert.DeserializeObject(Of DataSet)(_Json)
        'Dim _RowObservaciones As DataRow = _Ds.Tables(0).Rows(0)

        'Dim _TblObservaciones As DataTable = _Ds_Matriz_Documentos.Tables("Observaciones_Doc")
        'NewFila = _TblObservaciones.NewRow
        'With NewFila
        '    .Item("Observaciones") = _RowObservaciones.Item("Observaciones")
        '    .Item("Forma_pago") = _RowObservaciones.Item("Forma_pago")
        '    .Item("Orden_compra") = _RowObservaciones.Item("Orden_compra")
        '    .Item("Placa") = _RowObservaciones.Item("Placa")
        '    .Item("CodRetirador") = _RowObservaciones.Item("CodRetirador")

        '    For i = 1 To 35
        '        .Item("Obs" & i) = _RowObservaciones.Item("Obs" & i)
        '    Next

        '    _Ds_Matriz_Documentos.Tables("Observaciones_Doc").Rows.Add(NewFila)
        'End With

#End Region

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Existe = System.IO.File.Exists(True)

        If Not _Existe Then
            Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'' As Error"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'" & Replace(_Sql.Pro_Error, "'", "''") & "' As Error"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    ''' <summary>
    ''' Convierte la fecha desde un string en datetime
    ''' </summary>
    ''' <param name="_Fecha"></param>
    ''' <returns></returns>
    Function Fx_FechaStr2Datetime(_Fecha As String) As DateTime

        Dim _Fecha_Datetime As DateTime

        _Fecha_Datetime = DateTime.ParseExact(_Fecha, "MM/dd/yyyy", Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None)

        Return _Fecha_Datetime

    End Function

    Function Fx_LlenarDatos(ByRef _Ds_Matriz_Documentos As DataSet, _Json As String, _NomTabla As String)

        Dim _Log = String.Empty

        _Json = Mid(_Json, 2, _Json.Length - 1)
        _Json = Mid(_Json, 1, _Json.Length - 1)

        Dim _Ds As DataSet = JsonConvert.DeserializeObject(Of DataSet)(_Json)
        'Dim _Row As DataRow = _Ds.Tables(0).Rows(0)

        Dim NewFila As DataRow

        Dim _Tbl As DataTable = _Ds_Matriz_Documentos.Tables(_NomTabla)

        For Each _Row As DataRow In _Ds.Tables(0).Rows

            NewFila = _Tbl.NewRow

            With NewFila

                Dim name(_Tbl.Columns.Count) As String
                Dim i As Integer = 0
                For Each column As DataColumn In _Tbl.Columns
                    name(i) = column.ColumnName
                    Dim _NomColumna As String = column.ColumnName
                    Try
                        If column.DataType.Name = "DateTime" Then
                            .Item(_NomColumna) = Fx_FechaStr2Datetime(_Row.Item(_NomColumna))
                        Else
                            If column.DataType.Name = "Boolean" Then
                                .Item(_NomColumna) = CBool(_Row.Item(_NomColumna))
                            Else
                                .Item(_NomColumna) = _Row.Item(_NomColumna)
                            End If
                        End If
                    Catch ex As Exception
                        _Log += ex.Message & vbCrLf
                    End Try
                    i += 1
                Next

                If _NomTabla = "Detalle_Doc" Or _NomTabla = "Descuentos_Doc" Then
                    .Item("Id") = _Row.Item("Id_DocDet")
                End If

                _Tbl.Rows.Add(NewFila)

            End With

        Next

    End Function

#End Region

End Class