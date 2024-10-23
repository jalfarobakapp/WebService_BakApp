Imports System.ComponentModel
Imports System.IO
Imports System.Security.Claims
Imports System.Web.Script.Serialization
Imports System.Web.Script.Services
Imports System.Web.Services
Imports DevComponents.DotNetBar
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la siguiente línea.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://BakApp")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class Ws_BakApp
    Inherits System.Web.Services.WebService

    Dim _Sql As Class_SQL
    Dim _Row_Tabcarac As DataRow

    Dim _Version As String = My.Application.Info.Version.ToString

    Public Sub New()

        _Global_Cadena_Conexion_SQL_Server = "data source = 192.168.0.75; initial catalog = RANDOM; user id = RANDOM; password = RANDOM"

    End Sub

    <WebMethod()>
    Public Function Fx_Probar_Conexion_BD() As String
        _Sql = New Class_SQL
        Dim _Error As String = _Sql.Fx_Probar_Conexion
        Return _Error '"Hola a todos" 'http://localhost:34553
    End Function

    <WebMethod()>
    Public Function Fx_Cadena_Conexion(Cadena_Conexion_SQL_Server As String) As String
        _Sql = New Class_SQL
        Dim _Error As String = _Sql.Fx_Probar_Conexion
        Return _Error
    End Function

    <WebMethod(True)>
    Function Fx_GetDataSet(Consulta_Sql As String) As DataSet
        _Sql = New Class_SQL
        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_Sql)
        Return _Ds
    End Function

    <WebMethod(True)>
    Function Fx_Trae_Dato_String(_Tabla As String,
                                 _Campo As String,
                                 _Condicion As String) As String
        _Sql = New Class_SQL
        Dim _Dato As String = _Sql.Fx_Trae_Dato(_Tabla, _Campo, _Condicion, , False, "")
        Return _Dato
    End Function

    <WebMethod(True)>
    Function Fx_Trae_Dato_Numero(_Tabla As String,
                                 _Campo As String,
                                 _Condicion As String) As String
        _Sql = New Class_SQL
        Dim _Dato As Double = _Sql.Fx_Trae_Dato(_Tabla, _Campo, _Condicion, , True, 0)
        Return _Dato
    End Function

    <WebMethod(True)>
    Function Fx_Ej_consulta_IDU(Consulta_Sql As String) As String
        _Sql = New Class_SQL

        If _Sql.Fx_Ej_consulta_IDU(Consulta_Sql) Then
            Return ""
        Else
            Return _Sql.Pro_Error
        End If
    End Function

    <WebMethod(True)>
    Function Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_Sql As String) As String
        _Sql = New Class_SQL

        If _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_Sql) Then
            Return ""
        Else
            Return _Sql.Pro_Error
        End If
    End Function

    <WebMethod(True)>
    Function Fx_Cuenta_Registros(_Tabla As String,
                                 _Condicion As String) As Double
        _Sql = New Class_SQL
        Dim _Dato As Double = _Sql.Fx_Cuenta_Registros(_Tabla, _Condicion)
        Return _Dato
    End Function

    '<WebMethod(True)>
    'Function Fx_Crear_Documento(_Global_BaseBk As String,
    '                            _Funcionario As String,
    '                            _Tido As String,
    '                            _Nudo As String,
    '                            _Es_ValeTransitorio As Boolean,
    '                            _EsElectronico As Boolean,
    '                            _Ds_Matriz_Documento As DataSet,
    '                            _Es_Ajuste As Boolean) As String

    '    Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

    '    Dim _Idmaeedo As String
    '    _Idmaeedo = _New_Doc.Fx_Crear_Documento(_Tido,
    '                                                _Nudo,
    '                                                _Es_ValeTransitorio,
    '                                                _EsElectronico,
    '                                                _Ds_Matriz_Documento,
    '                                                _Es_Ajuste)

    '    Return _Idmaeedo

    'End Function

    '<WebMethod(True)>
    'Function Fx_Editar_Documento(_Global_BaseBk As String,
    '                            _Idmaeedo_Dori As Integer,
    '                            _Funcionario As String,
    '                            _Ds_Matriz_Documento As DataSet) As Integer

    '    Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

    '    Dim _Idmaeedo As Integer
    '    _Idmaeedo = _New_Doc.Fx_Editar_Documento(_Idmaeedo_Dori, _Funcionario, _Ds_Matriz_Documento)

    '    Return _Idmaeedo

    'End Function

    <WebMethod(True)>
    Function Fx_Cambiar_Numeracion_Modalidad(_Tido As String,
                                             _Nudo As String,
                                             _Modalidad As String) As Double
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
    Function Fx_EliminarAnular_Doc(_Idmaeedo_Dori As Integer,
                                  _Funcionario As String,
                                  _Accion As _Enum_Accion_EA) As Boolean
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
    Function Fx_Traer_Numero_Documento(_Tido As String,
                                      _NumeroDoc As String,
                                      _Modalidad_Seleccionada As String,
                                      _Empresa As String) As String
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
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Version()

        _Sql = New Class_SQL
        Dim Consulta_sql As String

        Dim _Ds As DataSet


        Consulta_sql = "Select '" & _Version & "' As Version"
        _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Function Fx_Login_Usuario_Soap(_Clave As String) As DataSet

        Dim _Pw = Fx_TraeClaveRD(_Clave)

        Consulta_sql = "Select Top 1 KOFU,NOKOFU From TABFU Where PWFU = '" & _Pw & "'"
        _Sql = New Class_SQL

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Return _Ds

    End Function

#Region "JSON"

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Login_Usuario_Json(_Clave As String)

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

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Ds_Json(Key As String, _Consulta_Sql As String)

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
    Public Sub Sb_Ds_Json_Prueba(Consulta_Sql As String)

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
    Public Sub Sb_GetDataSet_Json(Consulta_Sql As String)

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
    Public Sub Sb_Buscar_Productos_Json(_Codigo As String,
                                        _Descripcion As String)

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

        _Sql = New Class_SQL
        Dim Consulta_sql As String

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'").ToString.Trim & ".dbo."

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Try

            Consulta_sql = "Select Top 1 * From TABCODAL Where KOPRAL = '" & Codigo & "' And KOEN = ''"
            Dim _RowTablcodal As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If Not IsNothing(_RowTablcodal) Then
                Codigo = _RowTablcodal.Item("KOPR")
            Else
                Dim _Kopr = _Sql.Fx_Trae_Dato("MAEPR", "KOPR", "KOPRTE = '" & Codigo & "'")
                If Not String.IsNullOrEmpty(_Kopr) Then
                    Codigo = _Kopr
                End If
            End If


            Consulta_sql = My.Resources.Recursos_Sql.SqlQuery_Traer_Producto
            Consulta_sql = Replace(Consulta_sql, "#Codigo#", Codigo)
            Consulta_sql = Replace(Consulta_sql, "#Empresa#", Empresa)
            Consulta_sql = Replace(Consulta_sql, "#Sucursal#", Sucursal)
            Consulta_sql = Replace(Consulta_sql, "#Bodega#", Bodega)
            Consulta_sql = Replace(Consulta_sql, "#Lista#", Lista)
            Consulta_sql = Replace(Consulta_sql, "#UnTrans#", UnTrans)


            _Sql = New Class_SQL
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

            Dim _PorIva As Double = _Ds.Tables(0).Rows(0).Item("PorIva")
            Dim _PorIla As Double = _Sql.Fx_Trae_Dato("TABIM", "Isnull(Sum(POIM),0)", "KOIM In (SELECT KOIM FROM TABIMPR Where KOPR = '" & Codigo & "')")

            Consulta_sql = "SELECT Top 1 *,--PP01UD,PP02UD,DTMA01UD As DSCTOMAX,ECUACION,
                        (SELECT top 1 MELT FROM TABPP Where KOLT = '" & Lista & "') As MELT FROM TABPRE
                        Where KOLT = '" & Lista & "' And KOPR = '" & Codigo & "'"
            Dim _RowPrecios As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_RowPrecios) Then
                Throw New System.Exception("Producto no asignado a la lista de precios [" & Lista & "]")
            End If

            Dim _Reg As Boolean = CBool(_Sql.Fx_Cuenta_Registros("TABBOPR",
                                                           "EMPRESA = '" & Empresa & "' And " &
                                                           "KOSU = '" & Sucursal & "' And " &
                                                           "KOBO = '" & Bodega & "' And " &
                                                           "KOPR = '" & Codigo & "'"))

            If Not _Reg Then
                Throw New System.Exception("Producto no asignado a la bodega [" & Bodega & "]")
            End If

            Dim _Ecuacion As String

            If UnTrans = 1 Then
                _Ecuacion = NuloPorNro(_RowPrecios.Item("ECUACION"), "")
            Else
                _Ecuacion = NuloPorNro(_RowPrecios.Item("ECUACION2"), "")
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

            Dim _Iva = _PorIva / 100
            Dim _Ila = _PorIla / 100

            Dim _Impuestos As Double = 1 + (_Iva + _Ila)

            Dim _PrecioNetoUdLista As Double
            Dim _PrecioBrutoUdLista As Double

            If _RowPrecios.Item("MELT") = "N" Then
                _PrecioNetoUdLista = _Precio
                _PrecioBrutoUdLista = Math.Round(_Precio * _Impuestos, 0)
            Else
                _PrecioBrutoUdLista = _Precio
                _PrecioNetoUdLista = Math.Round(_Precio / _Impuestos, 5)
            End If

            _Ds.Tables(0).Rows(0).Item("Ecuacion") = _Ecuacion.Trim
            _Ds.Tables(0).Rows(0).Item("DescMaximo") = _DescMaximo
            _Ds.Tables(0).Rows(0).Item("Precio") = _Precio
            _Ds.Tables(0).Rows(0).Item("PrecioListaUd1") = _PrecioListaUd1
            _Ds.Tables(0).Rows(0).Item("PrecioListaUd2") = _PrecioListaUd2

            _Ds.Tables(0).Rows(0).Item("PrecioNetoUdLista") = _PrecioNetoUdLista
            _Ds.Tables(0).Rows(0).Item("PrecioBrutoUdLista") = _PrecioBrutoUdLista

            _Ds.Tables(0).Rows(0).Item("PorIla") = _PorIla
            '_Ds.Tables(0).Rows(0).Item("CodLista") = Lista

            ' ESPECIAL SOLO PARA VILLAR HNOS.
            If _Sql.Fx_Existe_Tabla("@WMS_GATEWAY_STOCK") Then

                Dim _Stock As Double

                If Sucursal.Trim = "01" And Bodega.Trim = "01" Then _Stock = _Sql.Fx_Trae_Dato("[@WMS_GATEWAY_STOCK]", "STOCK_ALAMEDA", "SKU = '" & Codigo & "'", True)
                If Sucursal.Trim = "02" And Bodega.Trim = "02" Then _Stock = _Sql.Fx_Trae_Dato("[@WMS_GATEWAY_STOCK]", "STOCK_ENEA", "SKU = '" & Codigo & "'", True)

                _Ds.Tables(0).Rows(0).Item("StockBodega") = _Stock

            End If


        Catch ex As Exception

            Consulta_sql = "Select 'Error_" & Replace(ex.Message, "'", "''") & "' As Codigo,'" & _Version & "' As Version"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        End Try

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
    Public Sub Sb_Traer_Concepto_Json(_Concepto As String,
                                      _Empresa As String,
                                      _Sucursal As String,
                                      _Bodega As String,
                                      _Lista As String,
                                      _Koen As String)

        _Sql = New Class_SQL
        Dim _Ds As DataSet

        Consulta_sql = "Select * From TABCT Where KOCT = '" & _Concepto & "'"
        Dim _RowConcepto As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If Not IsNothing(_RowConcepto) Then

            Dim _Descripcion As String = _RowConcepto.Item("NOKOCT").ToString.Trim
            Dim _Iva As Integer = _RowConcepto.Item("POIVCT")
            Dim _Tict As String = _RowConcepto.Item("TICT")

            Dim _Valor_Concepto As Double

            If Not String.IsNullOrEmpty(_Koen) Then
                _Valor_Concepto = FX_Traer_Valor_Concepto(_Empresa, _RowConcepto, _Koen)
            End If

            If Not CBool(_Valor_Concepto) Then
                _Valor_Concepto = _RowConcepto.Item("POCT")
            End If

            Consulta_sql = "Select  Cast(0 As Int) As 'Id_DocEnc',
		                    '" & _Empresa & "' As Empresa,
		                    '" & _Sucursal & "' As Sucursal,
		                    '" & _Bodega & "' As Bodega,
		                    '" & _Concepto & "' As Codigo,
		                    '" & _Descripcion & "' As Descripcion,
		                    1 As 'UnTrans',
	                        '' As 'UdTrans',
		                    1 As 'Rtu',
		                    '' As 'Ud01PR',
		                    '' As 'Ud02PR',
		                    " & _Iva & " As 'PorIva',
		                    Cast(0 As Float) As 'PorIla',
		                    0 As 'StockBodega',
		                    '" & _Lista & "' As 'CodLista',
		                    Cast(1 as Bit) As 'Prct',
		                    '" & _Tict & "' As 'Tict',
		                    '' As 'Tipr',
		                    Cast(0 As Float) As 'Precio',
		                    Cast(0 As Float) As 'PrecioListaUd1',
		                    Cast(0 As Float) As 'PrecioListaUd2',
		                    " & De_Num_a_Tx_01(_Valor_Concepto, False, 5) & " As 'DescuentoPorc',
		                    Cast(0 As Float) As 'DescMaximo',
		                    Cast('' As Varchar(242)) As 'Ecuacion',
		                    0 As 'PmLinea',
		                    0 As 'PmSucLinea',
		                    0 As 'PmIFRS',
		                    '' As 'UbicacionBod',
		                    '' As 'Moneda',
		                    '' As 'Tipo_Moneda',
		                    Cast(0 As Float) As 'Tipo_Cambio'"

            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

            Dim a = Consulta_sql

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

        _Sql = New Class_SQL
        Dim Consulta_sql As String

        Dim _Stock_Disponible As Double
        Dim _Revisar_Stock_Disponible As Boolean = True
        Dim _Campo_Formula_Stock = String.Empty

        Consulta_sql = "Select Top 1 * From TABTIDO Where TIDO = '" & _Tido & "'"
        Dim _RowTido As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If Not IsNothing(_RowTido) Then
            _Campo_Formula_Stock = NuloPorNro(_RowTido.Item("STOCK"), "")
        End If

        If _Tido = "NVV" Or _Tido = "RES" Or _Tido = "PRO" Or _Tido = "NVI" Then

            _Revisar_Stock_Disponible = True

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

        Dim _Stock As Double

        _Stock = _Sql.Fx_Trae_Dato("MAEST", "STFI" & _UnTrans, "EMPRESA = '" & _Empresa &
                                   "' AND KOSU = '" & _Sucursal &
                                   "' AND KOBO = '" & _Bodega &
                                   "' AND KOPR = '" & _Codigo & "'", True)

        'CONFIGURACION ESPECIAL PARA VILLAR HERMANOS
        If _Sql.Fx_Existe_Tabla("@WMS_GATEWAY_STOCK") Then

            If _Sucursal.Trim = "01" And _Bodega.Trim = "01" Then _Stock = _Sql.Fx_Trae_Dato("[@WMS_GATEWAY_STOCK]", "STOCK_ALAMEDA", "SKU = '" & _Codigo & "'", True)
            If _Sucursal.Trim = "02" And _Bodega.Trim = "02" Then _Stock = _Sql.Fx_Trae_Dato("[@WMS_GATEWAY_STOCK]", "STOCK_ENEA", "SKU = '" & _Codigo & "'", True)

            _Stock_Disponible = _Stock

        End If

        _Sql = New Class_SQL

        Consulta_sql = "Select " & De_Num_a_Tx_01(_Stock_Disponible, False, 5) & " As Stock_Disponible," & De_Num_a_Tx_01(_Stock, False, 5) & " As Stock_Fisico"
        Dim _Ds As DataSet
        _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

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
            Dim _TblTabpre As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

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
            _TblDscto = _Sql.Fx_Get_DataTable(Consulta_sql)

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
    Public Sub Sb_Json2Ds(_Json As String)

        If String.IsNullOrEmpty(_Json) Then
            _Json = "{'Tabla 1': [{""Name"":""AAA"",""Age"":""22"",""Job"":""PPP""}," &
                             "{""Name"":""BBB"",""Age"":""25"",""Job"":""QQQ""}," &
                             "{""Name"":""CCC"",""Age"":""38"",""Job"":""RRR""}]}"

        End If

        Dim _Tbl As DataTable = Fx_de_Json_a_Datatable(_Json)

    End Sub

    <WebMethod(True)>
    Public Sub Sb_Json_ImpBk(_Json As String, _NombreTabla As String)

        _Sql = New Class_SQL
        Dim Consulta_sql As String

        Dim _Ruta As String = String.Empty ' = "D:\JsonB4Android\"
        _Ruta = System.Configuration.ConfigurationManager.AppSettings("Ruta_Tmp").ToString

        Dim _Existe_Ruta As Boolean = True
        Dim _Existe_Archivo As Boolean

        If String.IsNullOrEmpty(_Ruta) Then
            _Ruta = "Falta la consiguración de la carpeta de archivos temporales en [Web.config]" & vbCrLf &
                    "<appSettings>   <add key=""Ruta_Tmp"" value=""C:\JsonB4Android""/>   </appSettings>"
            _Existe_Ruta = False
        Else
            If Not Directory.Exists(_Ruta) Then
                _Ruta = "No existe el directorio: [" & _Ruta & "] en el servidor"
                _Existe_Ruta = False
            End If
        End If

        If _Existe_Ruta Then
            'Fx_Grabar_JsonArchivo(_Json, _Ruta, _NombreTabla)
            _Existe_Archivo = Fx_Grabar_JsonArchivo(_Json, _Ruta, _NombreTabla)
        End If

        Dim _Ds As DataSet
        Dim _Error As String = _Ruta

        If Not _Existe_Archivo Then
            Consulta_sql = "Select Cast(0 as Bit) As Respuesta,'" & _Error & "' As Error"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(1 as Bit) As Respuesta,'Archivo creado:" & _Error & "' As Error"
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
    Public Sub Sb_RevCarpetaTmp()

        _Sql = New Class_SQL
        Dim Consulta_sql As String

        Dim _Ruta As String = String.Empty ' = "D:\JsonB4Android\"
        _Ruta = System.Configuration.ConfigurationManager.AppSettings("Ruta_Tmp").ToString

        Dim _Existe_Ruta As Boolean = True

        If String.IsNullOrEmpty(_Ruta) Then
            _Ruta = "Falta la configuración de la carpeta de archivos temporales en [Web.config]" & vbCrLf &
                    "<appSettings>" & vbCrLf & "<add key=""Ruta_Tmp"" value=""""/>" & vbCrLf & "</appSettings>"
            _Existe_Ruta = False
        Else
            If Not Directory.Exists(_Ruta) Then
                _Ruta = "No existe el directorio de la carpete para archivos temporales en el servidor, revise el archivo [Web.config]" & vbCrLf &
                        "<appSettings>" & vbCrLf & "<add key=""Ruta_Tmp"" value=""" & _Ruta & """/>" & vbCrLf & "</appSettings>"
                _Existe_Ruta = False
            End If
        End If

        Dim _Ds As DataSet
        Dim _Error As String = _Ruta

        Consulta_sql = "Select Cast(" & Convert.ToInt32(_Existe_Ruta) & " as Bit) As ExisteRuta,'" & _Error & "' As Error,'" & _Version & "' As Version"
        _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_CreaDocumentoJsonBakapp(_EncabezadoJs As String,
                                          _DestalleJs As String,
                                          _DescuentosJs As String,
                                          _ObservacionesJs As String,
                                          _Id_Estacion As Integer)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error As String
        Dim _Idmaeedo As Integer
        Dim _Tido As String
        Dim _Nudo As String

        Dim _Row_EstacionBk As DataRow

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'").ToString.Trim & ".dbo."

        Try

            'Dim _Ruta = "D:\JsonB4Android\"

            'Fx_Grabar_JsonArchivo(_EncabezadoJs, _Ruta, "EncabezadoJs")
            'Fx_Grabar_JsonArchivo(_DestalleJs, _Ruta, "DestalleJs")
            'Fx_Grabar_JsonArchivo(_DescuentosJs, _Ruta, "DescuentosJs")
            'Fx_Grabar_JsonArchivo(_ObservacionesJs, _Ruta, "ObservacionesJs")

            'Throw New System.Exception("Archivo New creados...")

            'Return

            Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_EstacionesBkp Where Id = " & _Id_Estacion
            _Row_EstacionBk = _Sql.Fx_Get_DataRow(Consulta_sql)

            Dim _Ds_Matriz_Documentos As New Ds_Matriz_Documentos

            _Ds_Matriz_Documentos.Clear()
            _Ds_Matriz_Documentos = New Ds_Matriz_Documentos

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _EncabezadoJs, "Encabezado_Doc")

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Post_Venta") = False
            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Tipo_Documento") = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _DestalleJs, "Detalle_Doc")
            If Not String.IsNullOrEmpty(_DescuentosJs) Then Fx_LlenarDatos(_Ds_Matriz_Documentos, _DescuentosJs, "Descuentos_Doc")
            Fx_LlenarDatos(_Ds_Matriz_Documentos, _ObservacionesJs, "Observaciones_Doc")

            Dim _Funcionario As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("CodFuncionario")
            '_Global_BaseBk = "BAKAPP_VH.dbo."
            Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

            _New_Doc.NombreEquipo = _Row_EstacionBk.Item("NombreEquipo")
            _New_Doc.TipoEstacion = _Row_EstacionBk.Item("TipoEstacion")

            Dim _Modalidad As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Modalidad")
            Dim _Empresa As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Empresa")

            _Tido = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")
            _Nudo = Traer_Numero_Documento2(_Tido, _Empresa, _Modalidad)

            If _Nudo = "_Error" Then
                Throw New System.Exception("Problemas al obtener la numeración del documento." & vbCrLf &
                         "Informe esta situación al administrador del sistema.")
            End If

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("NroDocumento") = _Nudo

            _Idmaeedo = _New_Doc.Fx_Crear_Documento2(_Tido, _Nudo, False, False, _Ds_Matriz_Documentos)
            '_Idmaeedo = _New_Doc.Fx_Crear_Documento_En_BakApp_Casi2("Bakapp4ndroid", _Ds_Matriz_Documentos, False, True, "B4A")

            _Error = _New_Doc.Error

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If CBool(_Idmaeedo) Then

            Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _Idmaeedo
            Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            _Tido = _Row_Documento.Item("TIDO")
            _Nudo = _Row_Documento.Item("NUDO")

            Consulta_sql = "Select " & _Idmaeedo & " As Idmaeedo,'" & _Tido & "' As Tido,'" & _Nudo & "' As 'Nudo',Cast(1 as Bit) As Respuesta,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)

        Else
            Consulta_sql = "Select 0 As Idmaeedo,Cast(1 as Bit) As Respuesta,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_CreaDocumentoJsonBakapp2(_EncabezadoJs As String,
                                           _DestalleJs As String,
                                           _DescuentosJs As String,
                                           _ObservacionesJs As String,
                                           _DespachoSimpleJs As String,
                                           _Id_Estacion As Integer)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error As String
        Dim _Idmaeedo As Integer
        Dim _Tido As String
        Dim _Nudo As String

        Dim _Ds As DataSet
        Dim _Row_DespachoSimple As DataRow
        Dim _Row_EstacionBk As DataRow

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'").ToString.Trim & ".dbo."

        Try

            Dim _Ruta = "D:\JsonB4Android\"

            'Fx_Grabar_JsonArchivo(_EncabezadoJs, _Ruta, "EncabezadoJs")
            'Fx_Grabar_JsonArchivo(_DestalleJs, _Ruta, "DestalleJs")
            'Fx_Grabar_JsonArchivo(_DescuentosJs, _Ruta, "DescuentosJs")
            'Fx_Grabar_JsonArchivo(_ObservacionesJs, _Ruta, "ObservacionesJs")
            'Fx_Grabar_JsonArchivo(_DespachoSimpleJs, _Ruta, "DespachoSimpleJs")

            'Throw New System.Exception("Archivo New creados...")

            'Return

            Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_EstacionesBkp Where Id = " & _Id_Estacion
            _Row_EstacionBk = _Sql.Fx_Get_DataRow(Consulta_sql)

            Dim _Ds_Matriz_Documentos As New Ds_Matriz_Documentos

            _Ds_Matriz_Documentos.Clear()
            _Ds_Matriz_Documentos = New Ds_Matriz_Documentos

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _EncabezadoJs, "Encabezado_Doc")

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("FechaEmision") = FechaDelServidor()
            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Post_Venta") = False
            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Tipo_Documento") = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _DestalleJs, "Detalle_Doc")
            If Not String.IsNullOrEmpty(_DescuentosJs) Then Fx_LlenarDatos(_Ds_Matriz_Documentos, _DescuentosJs, "Descuentos_Doc")
            Fx_LlenarDatos(_Ds_Matriz_Documentos, _ObservacionesJs, "Observaciones_Doc")

            Dim _Json = _DespachoSimpleJs

            _Json = Mid(_Json, 2, _Json.Length - 1)
            _Json = Mid(_Json, 1, _Json.Length - 1)

            _Ds = JsonConvert.DeserializeObject(Of DataSet)(_Json)
            _Row_DespachoSimple = _Ds.Tables(0).Rows(0)

            Dim _Funcionario As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("CodFuncionario")

            Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

            _New_Doc.NombreEquipo = _Row_EstacionBk.Item("NombreEquipo")
            _New_Doc.TipoEstacion = _Row_EstacionBk.Item("TipoEstacion")

            Dim _Modalidad As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Modalidad")
            Dim _Empresa As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Empresa")

            _Tido = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")
            _Nudo = Traer_Numero_Documento2(_Tido, _Empresa, _Modalidad)

            If _Nudo = "_Error" Then
                Throw New System.Exception("Problemas al obtener la numeración del documento." & vbCrLf &
                         "Informe esta situación al administrador del sistema.")
            End If

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("NroDocumento") = _Nudo

            _Idmaeedo = _New_Doc.Fx_Crear_Documento2(_Tido, _Nudo, False, False, _Ds_Matriz_Documentos)

            _Error = _New_Doc.Error

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If CBool(_Idmaeedo) Then

            Dim _CodTipoDespacho As Integer = _Row_DespachoSimple.Item("CodTipoDespacho")
            Dim _TipoDespacho As String = _Row_DespachoSimple.Item("TipoDespacho")

            Dim _CodTipoPagoDesp As String = _Row_DespachoSimple.Item("CodTipoPagoDesp")
            Dim _TipoPagoDesp As String = _Row_DespachoSimple.Item("TipoPagoDesp")

            Dim _DireccionDesp As String = _Row_DespachoSimple.Item("DireccionDesp")
            Dim _TransporteDesp As String = _Row_DespachoSimple.Item("TransporteDesp")
            Dim _ObservacionesDesp As String = _Row_DespachoSimple.Item("ObservacionesDesp")

            Dim _CodDocDestino As String = _Row_DespachoSimple.Item("CodDocDestino")
            Dim _DocDestino As String = _Row_DespachoSimple.Item("DocDestino")

            Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Despacho_Simple (Idmaeedo,CodTipoDespacho,TipoDespacho,CodTipoPagoDesp,TipoPagoDesp," &
                           "DireccionDesp,TransporteDesp,ObservacionesDesp,CodDocDestino,DocDestino) Values " &
                           "(" & _Idmaeedo & "," & _CodTipoDespacho & ",'" & _TipoDespacho & "'," & _CodTipoPagoDesp & ",'" & _TipoPagoDesp &
                           "','" & _DireccionDesp & "','" & _TransporteDesp & "','" & _ObservacionesDesp & "','" & _CodDocDestino & "','" & _DocDestino & "')"
            If _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_sql) Then

                If _Tido = "NVV" Then

                    If _Sql.Fx_Existe_Tabla("@WMS_GATEWAY_ANEXO_PEDIDOS") Then

                        Consulta_sql = "Insert Into [@WMS_GATEWAY_ANEXO_PEDIDOS] (IDMAEEDO,TIPO_DESPACHO,FORMA_PAGO,DOCUMENTO_DESTINO) Values " &
                                   "(" & _Idmaeedo & "," & _CodTipoDespacho & ",'" & _CodTipoPagoDesp & "','" & _CodDocDestino & "')"
                        _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_sql)

                    End If

                End If

            End If

            Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _Idmaeedo
            Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            _Tido = _Row_Documento.Item("TIDO")
            _Nudo = _Row_Documento.Item("NUDO")

            Consulta_sql = "Select " & _Idmaeedo & " As Idmaeedo,'" & _Tido & "' As Tido,'" & _Nudo & "' As 'Nudo',Cast(1 as Bit) As Respuesta,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)

        Else
            Consulta_sql = "Select 0 As Idmaeedo,Cast(1 as Bit) As Respuesta,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_EditarDocumentoJsonBakapp(_OldIdmaeedo As Integer,
                                            _Cod_Func_Eliminador As String,
                                            _Global_BaseBk As String,
                                            _EncabezadoJs As String,
                                            _DestalleJs As String,
                                            _DescuentosJs As String,
                                            _ObservacionesJs As String,
                                            _Cambiar_NroDocumento As Boolean,
                                            _Id_Estacion As Integer)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error As String
        Dim _NewIdmaeedo As Integer
        Dim _Tido As String
        Dim _Nudo As String
        Dim _Old_Nudo As String

        Dim _Row_OldMaeedo As DataRow
        Dim _Row_EstacionBk As DataRow

        Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _OldIdmaeedo
        _Row_OldMaeedo = _Sql.Fx_Get_DataRow(Consulta_sql)

        _Tido = _Row_OldMaeedo.Item("TIDO")
        _Nudo = _Row_OldMaeedo.Item("NUDO")
        _Old_Nudo = _Nudo

        Try

            Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_EstacionesBkp Where Id = " & _Id_Estacion
            _Row_EstacionBk = _Sql.Fx_Get_DataRow(Consulta_sql)

            Dim _Ds_Matriz_Documentos As New Ds_Matriz_Documentos

            _Ds_Matriz_Documentos.Clear()
            _Ds_Matriz_Documentos = New Ds_Matriz_Documentos

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _EncabezadoJs, "Encabezado_Doc")

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("FechaEmision") = FechaDelServidor()
            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Post_Venta") = False
            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Tipo_Documento") = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _DestalleJs, "Detalle_Doc")
            If Not String.IsNullOrEmpty(_DescuentosJs) Then Fx_LlenarDatos(_Ds_Matriz_Documentos, _DescuentosJs, "Descuentos_Doc")
            Fx_LlenarDatos(_Ds_Matriz_Documentos, _ObservacionesJs, "Observaciones_Doc")

            Dim _Funcionario As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("CodFuncionario")
            '_Global_BaseBk = "BAKAPP_VH.dbo."
            Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

            _New_Doc.NombreEquipo = _Row_EstacionBk.Item("NombreEquipo")
            _New_Doc.TipoEstacion = _Row_EstacionBk.Item("TipoEstacion")

            Dim _Modalidad As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Modalidad")
            Dim _Empresa As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Empresa")

            If _Cambiar_NroDocumento Then
                '_Tido = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")
                _Nudo = Traer_Numero_Documento2(_Tido, _Empresa, _Modalidad)
            End If

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("NroDocumento") = _Nudo

            _NewIdmaeedo = _New_Doc.Fx_Crear_Documento2(_Tido, _Nudo, False, False, _Ds_Matriz_Documentos)

            If Not _Cambiar_NroDocumento Then _Nudo = _Old_Nudo

            _Error = _New_Doc.Error

            If CBool(_NewIdmaeedo) Then

                Dim _Class_E As New Clase_EliminarAnular_Documento

                Dim _Eliminado As Boolean = _Class_E.Fx_EliminarAnular_Doc(_OldIdmaeedo,
                                                                           _Cod_Func_Eliminador,
                                                                           Clase_EliminarAnular_Documento._Accion_EA.Modificar,
                                                                           False)

                If _Eliminado Then

                    Consulta_sql = "Update MAEEDO Set NUDO = '" & _Nudo & "' Where IDMAEEDO = " & _NewIdmaeedo & vbCrLf &
                                   "Update MAEDDO Set NUDO = '" & _Nudo & "' Where IDMAEEDO = " & _NewIdmaeedo
                    _Sql.Fx_Ej_consulta_IDU(Consulta_sql)

                Else

                    _Class_E.Fx_EliminarAnular_Doc(_NewIdmaeedo,
                                                   _Cod_Func_Eliminador,
                                                   Clase_EliminarAnular_Documento._Accion_EA.Modificar,
                                                   False)
                    _NewIdmaeedo = 0

                End If

            End If

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If CBool(_NewIdmaeedo) Then

            Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _NewIdmaeedo
            Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            _Tido = _Row_Documento.Item("TIDO")
            _Nudo = _Row_Documento.Item("NUDO")

            Consulta_sql = "Select " & _NewIdmaeedo & " As Idmaeedo,'" & _Tido & "' As Tido,'" & _Nudo & "' As 'Nudo',Cast(1 as Bit) As Respuesta,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select 0 As Idmaeedo,Cast(1 as Bit) As Respuesta,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_EditarDocumentoJsonBakapp2(_OldIdmaeedo As Integer,
                                            _Cod_Func_Eliminador As String,
                                            _Global_BaseBk As String,
                                            _EncabezadoJs As String,
                                            _DestalleJs As String,
                                            _DescuentosJs As String,
                                            _ObservacionesJs As String,
                                            _Cambiar_NroDocumento As Boolean,
                                            _DespachoSimpleJs As String,
                                            _Id_Estacion As Integer)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error As String
        Dim _NewIdmaeedo As Integer
        Dim _Tido As String
        Dim _Nudo As String
        Dim _Old_Nudo As String

        Dim _Row_OldMaeedo As DataRow
        Dim _Row_EstacionBk As DataRow

        Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _OldIdmaeedo
        _Row_OldMaeedo = _Sql.Fx_Get_DataRow(Consulta_sql)

        _Tido = _Row_OldMaeedo.Item("TIDO")
        _Nudo = _Row_OldMaeedo.Item("NUDO")
        _Old_Nudo = _Nudo

        Dim _Ds As DataSet
        Dim _Row_DespachoSimple As DataRow

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_EstacionesBkp Where Id = " & _Id_Estacion
        _Row_EstacionBk = _Sql.Fx_Get_DataRow(Consulta_sql)


        Try

            Dim _Ds_Matriz_Documentos As New Ds_Matriz_Documentos

            _Ds_Matriz_Documentos.Clear()
            _Ds_Matriz_Documentos = New Ds_Matriz_Documentos

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _EncabezadoJs, "Encabezado_Doc")

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Post_Venta") = False
            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Tipo_Documento") = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _DestalleJs, "Detalle_Doc")
            If Not String.IsNullOrEmpty(_DescuentosJs) Then Fx_LlenarDatos(_Ds_Matriz_Documentos, _DescuentosJs, "Descuentos_Doc")
            Fx_LlenarDatos(_Ds_Matriz_Documentos, _ObservacionesJs, "Observaciones_Doc")

            Dim _Funcionario As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("CodFuncionario")
            '_Global_BaseBk = "BAKAPP_VH.dbo."
            Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

            _New_Doc.NombreEquipo = _Row_EstacionBk.Item("NombreEquipo")
            _New_Doc.TipoEstacion = _Row_EstacionBk.Item("TipoEstacion")

            Dim _Modalidad As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Modalidad")
            Dim _Empresa As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Empresa")

            If _Cambiar_NroDocumento Then
                '_Tido = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")
                _Nudo = Traer_Numero_Documento2(_Tido, _Empresa, _Modalidad)
            End If

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("NroDocumento") = _Nudo


            Dim _Json = _DespachoSimpleJs

            _Json = Mid(_Json, 2, _Json.Length - 1)
            _Json = Mid(_Json, 1, _Json.Length - 1)

            _Ds = JsonConvert.DeserializeObject(Of DataSet)(_Json)
            _Row_DespachoSimple = _Ds.Tables(0).Rows(0)


            _NewIdmaeedo = _New_Doc.Fx_Crear_Documento2(_Tido, _Nudo, False, False, _Ds_Matriz_Documentos,,,,,, True)

            If Not _Cambiar_NroDocumento Then _Nudo = _Old_Nudo

            _Error = _New_Doc.Error

            If CBool(_NewIdmaeedo) Then

                Dim _Class_E As New Clase_EliminarAnular_Documento

                Dim _Eliminado As Boolean = _Class_E.Fx_EliminarAnular_Doc(_OldIdmaeedo,
                                                                           _Cod_Func_Eliminador,
                                                                           Clase_EliminarAnular_Documento._Accion_EA.Modificar,
                                                                           False)

                If _Eliminado Then

                    Consulta_sql = "Update MAEEDO Set NUDO = '" & _Nudo & "' Where IDMAEEDO = " & _NewIdmaeedo & vbCrLf &
                                   "Update MAEDDO Set NUDO = '" & _Nudo & "' Where IDMAEEDO = " & _NewIdmaeedo
                    _Sql.Fx_Ej_consulta_IDU(Consulta_sql)

                Else

                    _Class_E.Fx_EliminarAnular_Doc(_NewIdmaeedo,
                                                   _Cod_Func_Eliminador,
                                                   Clase_EliminarAnular_Documento._Accion_EA.Modificar,
                                                   False)
                    _NewIdmaeedo = 0

                End If

            End If

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If CBool(_NewIdmaeedo) Then

            Dim _CodTipoDespacho As Integer = _Row_DespachoSimple.Item("CodTipoDespacho")
            Dim _TipoDespacho As String = _Row_DespachoSimple.Item("TipoDespacho")

            Dim _CodTipoPagoDesp As String = _Row_DespachoSimple.Item("CodTipoPagoDesp")
            Dim _TipoPagoDesp As String = _Row_DespachoSimple.Item("TipoPagoDesp")

            Dim _DireccionDesp As String = _Row_DespachoSimple.Item("DireccionDesp")
            Dim _TransporteDesp As String = _Row_DespachoSimple.Item("TransporteDesp")
            Dim _ObservacionesDesp As String = _Row_DespachoSimple.Item("ObservacionesDesp")

            Dim _CodDocDestino As String = _Row_DespachoSimple.Item("CodDocDestino")
            Dim _DocDestino As String = _Row_DespachoSimple.Item("DocDestino")

            Consulta_sql = "Delete " & _Global_BaseBk & "Zw_Despacho_Simple Where Idmaeedo = " & _OldIdmaeedo & vbCrLf &
                           "Insert Into " & _Global_BaseBk & "Zw_Despacho_Simple (Idmaeedo,CodTipoDespacho,TipoDespacho,CodTipoPagoDesp,TipoPagoDesp," &
                           "DireccionDesp,TransporteDesp,ObservacionesDesp,CodDocDestino,DocDestino) Values " &
                           "(" & _NewIdmaeedo & "," & _CodTipoDespacho & ",'" & _TipoDespacho & "'," & _CodTipoPagoDesp & ",'" & _TipoPagoDesp &
                           "','" & _DireccionDesp & "','" & _TransporteDesp & "','" & _ObservacionesDesp & "','" & _CodDocDestino & "','" & _DocDestino & "')"
            If _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_sql) Then

                If _Tido = "NVV" Then

                    If _Sql.Fx_Existe_Tabla("@WMS_GATEWAY_ANEXO_PEDIDOS") Then

                        Consulta_sql = "Insert Into [@WMS_GATEWAY_ANEXO_PEDIDOS] (IDMAEEDO,TIPO_DESPACHO,FORMA_PAGO,DOCUMENTO_DESTINO) Values " &
                                   "(" & _NewIdmaeedo & "," & _CodTipoDespacho & ",'" & _CodTipoPagoDesp & "','" & _CodDocDestino & "')"
                        _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_sql)

                    End If

                End If

            End If

            Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _NewIdmaeedo
            Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            _Tido = _Row_Documento.Item("TIDO")
            _Nudo = _Row_Documento.Item("NUDO")

            Consulta_sql = "Select " & _NewIdmaeedo & " As Idmaeedo,'" & _Tido & "' As Tido,'" & _Nudo & "' As 'Nudo',Cast(1 as Bit) As Respuesta,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select 0 As Idmaeedo,Cast(1 as Bit) As Respuesta,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Function Sb_CreaDocumentoJson2XmlBakapp(_EncabezadoJs As String, _DestalleJs As String, _DescuentosJs As String, _ObservacionesJs As String) As DataSet

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error As String
        Dim _Idmaeedo As Integer
        Dim _Tido As String
        Dim _Nudo As String

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'").ToString.Trim & ".dbo."

        Try

            Dim _Ds_Matriz_Documentos As New Ds_Matriz_Documentos

            _Ds_Matriz_Documentos.Clear()
            _Ds_Matriz_Documentos = New Ds_Matriz_Documentos

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _EncabezadoJs, "Encabezado_Doc")

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Post_Venta") = False
            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Tipo_Documento") = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _DestalleJs, "Detalle_Doc")
            If Not String.IsNullOrEmpty(_DescuentosJs) Then Fx_LlenarDatos(_Ds_Matriz_Documentos, _DescuentosJs, "Descuentos_Doc")
            Fx_LlenarDatos(_Ds_Matriz_Documentos, _ObservacionesJs, "Observaciones_Doc")

            Dim _Funcionario As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("CodFuncionario")
            '_Global_BaseBk = "BAKAPP_VH.dbo."
            Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

            Dim _Modalidad As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Modalidad")
            Dim _Empresa As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Empresa")

            _Tido = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")
            _Nudo = Traer_Numero_Documento2(_Tido, _Empresa, _Modalidad)

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("NroDocumento") = _Nudo

            _Idmaeedo = _New_Doc.Fx_Crear_Documento2(_Tido, _Nudo, False, False, _Ds_Matriz_Documentos)

            _Error = _New_Doc.Error

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If CBool(_Idmaeedo) Then

            Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _Idmaeedo
            Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            _Tido = _Row_Documento.Item("TIDO")
            _Nudo = _Row_Documento.Item("NUDO")

            Consulta_sql = "Select " & _Idmaeedo & " As Idmaeedo,'" & _Tido & "' As Tido,'" & _Nudo & "' As 'Nudo',Cast(1 as Bit) As Respuesta,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select 0 As Idmaeedo,Cast(1 as Bit) As Respuesta,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Return _Ds2

    End Function

    <WebMethod(True)>
    Function Sb_CreaDocumentoJson2StrBakapp(_EncabezadoJs As String, _DestalleJs As String, _DescuentosJs As String, _ObservacionesJs As String) As String

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error As String
        Dim _Idmaeedo As Integer
        Dim _Tido As String
        Dim _Nudo As String

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'").ToString.Trim & ".dbo."

        Try

            Dim _Ds_Matriz_Documentos As New Ds_Matriz_Documentos

            _Ds_Matriz_Documentos.Clear()
            _Ds_Matriz_Documentos = New Ds_Matriz_Documentos

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _EncabezadoJs, "Encabezado_Doc")

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Post_Venta") = False
            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Tipo_Documento") = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _DestalleJs, "Detalle_Doc")
            If Not String.IsNullOrEmpty(_DescuentosJs) Then Fx_LlenarDatos(_Ds_Matriz_Documentos, _DescuentosJs, "Descuentos_Doc")
            Fx_LlenarDatos(_Ds_Matriz_Documentos, _ObservacionesJs, "Observaciones_Doc")

            Dim _Funcionario As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("CodFuncionario")
            '_Global_BaseBk = "BAKAPP_VH.dbo."
            Dim _New_Doc As New Clase_Crear_Documento(_Global_BaseBk, _Funcionario)

            Dim _Modalidad As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Modalidad")
            Dim _Empresa As String = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("Empresa")

            _Tido = _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("TipoDoc")
            _Nudo = Traer_Numero_Documento2(_Tido, _Empresa, _Modalidad)

            _Ds_Matriz_Documentos.Tables("Encabezado_Doc").Rows(0).Item("NroDocumento") = _Nudo

            _Idmaeedo = _New_Doc.Fx_Crear_Documento2(_Tido, _Nudo, False, False, _Ds_Matriz_Documentos)

            _Error = _New_Doc.Error

        Catch ex As Exception
            _Error = ex.Message
        End Try

        'If CBool(_Idmaeedo) Then

        '    Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _Idmaeedo
        '    Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        '    _Tido = _Row_Documento.Item("TIDO")
        '    _Nudo = _Row_Documento.Item("NUDO")

        '    Consulta_sql = "Select " & _Idmaeedo & " As Idmaeedo,'" & _Tido & "' As Tido,'" & _Nudo & "' As 'Nudo',Cast(1 as Bit) As Respuesta,'" & _Version & "' As Version"
        '    _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        'Else
        '    Consulta_sql = "Select 0 As Idmaeedo,Cast(1 as Bit) As Respuesta,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
        '    _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        'End If

        Return _Idmaeedo

    End Function
    Private Sub Fx_Limpia_Json(ByRef _Json As String)
        _Json = Replace(_Json, "\/", "/")
        _Json = _Json.Trim
    End Sub

    <WebMethod(True)>
    Public Sub Sb_Usar_Dscto_Poswii(_Clave As String, _Kofu As String, _Eliminar As Boolean)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Consulta_sql = "Select Top 1 KOFU,CLAVE,ESTADO,DESCUENTO,USUARIO,FECHA " & vbCrLf &
                       "From [@POSWI_USER_ADMIN]" & vbCrLf &
                       "Where KOFU = '" & _Kofu & "' And CLAVE = '" & _Clave & "' And CAST(FECHA AS DATE) = CAST(Getdate() AS DATE)"

        Consulta_sql = "Select ID,CLAVE,FECHA,DESCUENTO,KOFU,ESTADO,DIR_IP" & vbCrLf &
                       "From [@POSWI_DESCUENTO_CABECERA]" & vbCrLf &
                       "Where KOFU = '" & _Kofu & "' And CLAVE = '" & _Clave & "' And CAST(FECHA AS DATE) = CAST(Getdate() AS DATE)"

        Dim _Row_Permiso As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If Not IsNothing(_Row_Permiso) Then

            Dim _Id As Integer = _Row_Permiso.Item("ID")
            Dim _Estado As Integer = _Row_Permiso.Item("ESTADO")

            Consulta_sql = "Select Cast(1 As Bit) As Existe,Cast(" & _Estado & " As Bit) As Otorgado," & _Row_Permiso.Item("DESCUENTO") & " As Descuento"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)

            If _Eliminar Then
                'Consulta_sql = "Delete [@POSWI_USER_ADMIN]" & vbCrLf &
                '               "Where KOFU = '" & _Kofu & "' And CLAVE = '" & _Clave & "' And CAST(FECHA AS DATE) = CAST(Getdate() AS DATE)"
                Consulta_sql = "Delete [@POSWI_DESCUENTO_CABECERA]" & vbCrLf &
                               "Where ID = " & _Id
            Else
                'Consulta_sql = "Update [@POSWI_USER_ADMIN] Set ESTADO = 1" & vbCrLf &
                '               "Where KOFU = '" & _Kofu & "' And ESTADO = 0 And CLAVE = '" & _Clave & "' And CAST(FECHA AS DATE) = CAST(Getdate() AS DATE)"
                Consulta_sql = "Update [@POSWI_DESCUENTO_CABECERA] Set ESTADO = 1" & vbCrLf &
                               "Where ID = " & _Id
            End If
            _Sql.Fx_Ej_consulta_IDU(Consulta_sql)

        Else
            Consulta_sql = "Select Cast(0 As Bit) As Existe,Cast(0 As Bit) As Otorgado,0 As Descuento"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_Usar_Clave_DocDespSimple_Poswii(_Clave As String, _Koen As String, _Eliminar As Boolean)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Consulta_sql = "Select * From [@WMS_GATEWAY_DOCUMENTOS_CLAVES]" & vbCrLf &
                       "Where KOEN = '" & _Koen & "' And CLAVE = '" & _Clave & "' -- And ESTADO = 0"

        Dim _Row_Permiso As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If Not IsNothing(_Row_Permiso) Then

            Dim _Id As Integer = _Row_Permiso.Item("ID")
            Dim _Estado As Integer = _Row_Permiso.Item("ESTADO")

            Consulta_sql = "Select " & _Id & " As Id,Cast(1 As Bit) As Existe,Cast(" & _Estado & " As Bit) As Otorgado"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)

            If Not CBool(_Estado) Then

                If _Eliminar Then
                    Consulta_sql = "Delete [@WMS_GATEWAY_DOCUMENTOS_CLAVES] Where ID = " & _Id
                Else
                    Consulta_sql = "Update [@WMS_GATEWAY_DOCUMENTOS_CLAVES] Set ESTADO = 1,FECHA_USO=GETDATE() Where ID = " & _Id
                End If
                _Sql.Fx_Ej_consulta_IDU(Consulta_sql)

            End If

        Else
            Consulta_sql = "Select 0 As Id,Cast(0 As Bit) As Existe,Cast(0 As Bit) As Otorgado"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    Function Fx_Grabar_JsonArchivo(_Json As String, ByRef _Ruta As String, _NombreTabla As String) As Boolean

        _Json = Replace(_Json, "\/", "/")
        _Json = _Json.Trim
        Dim _Json2 = Mid(_Json, 2, _Json.Length - 1)
        _Json2 = Mid(_Json2, 1, _Json2.Length - 1)

        Dim RutaArchivo As String = _Ruta & "\" & _NombreTabla & ".json"
        Dim Cuerpo = _Json

        Dim oSW As New System.IO.StreamWriter(RutaArchivo)

        oSW.WriteLine(Cuerpo)
        oSW.Close()

        Dim _Existe As Boolean = System.IO.File.Exists(RutaArchivo)

        If _Existe Then
            _Ruta = RutaArchivo
        End If

        Return _Existe

    End Function

    ''' <summary>
    ''' Convierte la fecha desde un string en datetime
    ''' </summary>
    ''' <param name="_Fecha"></param>
    ''' <returns></returns>
    Function Fx_FechaStr2Datetime(_Fecha As String) As DateTime

        Dim _Fecha_Datetime As DateTime
        Dim _VolverArevisar = False

        _Fecha = Replace(_Fecha, "/", "-")
        Try
            _Fecha_Datetime = DateTime.ParseExact(_Fecha, "dd-MM-yyyy", Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None)
        Catch ex As Exception
            _VolverArevisar = True
        End Try

        If _VolverArevisar Then
            _Fecha_Datetime = DateTime.ParseExact(_Fecha, "MM-dd-yyyy", Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None)
        End If

        Return _Fecha_Datetime

    End Function

    Function Fx_LlenarDatos(ByRef _Ds_Matriz_Documentos As DataSet, _Json As String, _NomTabla As String)

        Dim _Log = String.Empty

        _Json = Mid(_Json, 2, _Json.Length - 1)
        _Json = Mid(_Json, 1, _Json.Length - 1)

        Dim _Ds As DataSet = JsonConvert.DeserializeObject(Of DataSet)(_Json)

        Dim NewFila As DataRow

        Dim _Tbl As DataTable = _Ds_Matriz_Documentos.Tables(_NomTabla)

        For Each _Row As DataRow In _Ds.Tables(0).Rows

            NewFila = _Tbl.NewRow

            With NewFila

                Dim name(_Tbl.Columns.Count) As String
                Dim i As Integer = 0
                For Each column As DataColumn In _Tbl.Columns
                    name(i) = column.ColumnName
                    If column.ColumnName = "Fecha_Tributaria" Then
                        Dim t = 0
                    End If
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

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Traer_Descuento_Global_X_Cliente(_Global_BaseBk As String, _Koen As String, _Suen As String)

        _Sql = New Class_SQL
        Dim Consulta_sql As String

        Dim _Ds As DataSet

        Try

            Consulta_sql = "Select * From PDIMCLI Where CODIGO = '" & _Koen & "'"
            Dim _Row_Pdimcli As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_Row_Pdimcli) Then
                Throw New System.Exception("No tiene registos en PDIMCLI el cliente")
            End If

            Dim _Campo_Dscto As String = _Sql.Fx_Trae_Dato(_Global_BaseBk & "Zw_TablaDeCaracterizaciones", "CodigoTabla", "Tabla = 'PDIMEN_Poswi'")

            If String.IsNullOrEmpty(_Campo_Dscto) Then
                Throw New System.Exception("Falta el campo en la tabla Zw_TablaDeCaracterizaciones, Tabla: PDIMEN_Poswi, Campo...")
            End If

            Dim _Porc_Dscto As Double = _Row_Pdimcli.Item(_Campo_Dscto)
            Dim _TieneDsctoEspecial As Boolean = (_Porc_Dscto > 0)

            Consulta_sql = "Select Cast(" & Convert.ToInt32(_TieneDsctoEspecial) & " As Bit) As TieneDsctoEspecial," &
                            _Porc_Dscto & " As Descuento,'' As Error"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        Catch ex As Exception

            Consulta_sql = "Select Cast(0 As Bit) As TieneDsctoEspecial,0 As Descuento,'" & ex.Message & "' As Error"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        End Try

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_EnviarCorreoBakapp(_Global_BaseBk As String,
                                     _Empresa As String,
                                     _Modalidad As String,
                                     _CodFuncionario As String,
                                     _Idmaeedo As Integer,
                                     _Para As String,
                                     _Cc As String)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error = String.Empty

        Dim _Tido As String
        Dim _Nudo As String
        Dim _Enviar_Correo As Boolean
        Dim _Id_Correo As Integer
        Dim _Nombre_Correo As String
        Dim _Asunto As String
        Dim _CuerpoMensaje As String
        Dim _NombreFormato_Correo As String

        Dim _Row_Documento As DataRow
        Dim _Row_Funcionario As DataRow
        Dim _Row_ConfModDocumento As DataRow
        Dim _Row_Correo As DataRow

        Try

            Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _Idmaeedo
            _Row_Documento = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_Row_Documento) Then
                Throw New System.Exception("Documento no encontrado")
            End If

            _Tido = _Row_Documento.Item("TIDO")
            _Nudo = _Row_Documento.Item("NUDO")

            Consulta_sql = "Select Us.*,Cr.Contrasena,Cr.Host,Cr.Puerto,Cr.SSL From " & _Global_BaseBk & "Zw_Usuarios Us
                            Inner Join " & _Global_BaseBk & "Zw_Correos_Cuentas Cr On Us.Email = Cr.Nombre_Usuario
                            Where Us.CodFuncionario = '" & _CodFuncionario & "'"
            _Row_Funcionario = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_Row_Funcionario) Then
                Throw New System.Exception("Falta la configuración del correo del funcionario" & vbCrLf &
                                           "Revise el Email en Zw_Usuarios Vs la conf. en Zw_Correos_Cuentas" & vbCrLf &
                                           "Avise de esta situación al administrador del sistema")
            End If

            Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Configuracion_Formatos_X_Modalidad" & vbCrLf &
                           "Where Empresa = '" & _Empresa & "' And Modalidad = '" & _Modalidad & "' And TipoDoc = '" & _Tido & "'"
            _Row_ConfModDocumento = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_Row_ConfModDocumento) Then
                Throw New System.Exception("Falta la configuración de la modalidad (" & _Modalidad & ")" & vbCrLf &
                                           "Avise de esta situación al administrador del sistema")
            End If

            _Enviar_Correo = _Row_ConfModDocumento.Item("Enviar_Correo")
            _Id_Correo = _Row_ConfModDocumento.Item("Id_Correo")
            _NombreFormato_Correo = _Row_ConfModDocumento.Item("NombreFormato_Correo")

            Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Correos Where Id = " & _Id_Correo
            _Row_Correo = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_Row_Correo) Then
                Throw New System.Exception("Falta el correo para el documento " & _Tido & " en la configuración de la modalidad (" & _Modalidad & ")" & vbCrLf &
                                           "Avise de esta situación al administrador del sistema")
            End If

            _Nombre_Correo = _Row_Correo.Item("Nombre_Correo")

            _Asunto = _Row_Correo.Item("Asunto")
            _CuerpoMensaje = _Row_Correo.Item("CuerpoMensaje")

            Dim _Email As String
            Dim _Host As String
            Dim _Puerto As Integer
            Dim _SSL As Boolean

            _Email = _Row_Funcionario.Item("Email")
            _Host = _Row_Funcionario.Item("Host")
            _Puerto = _Row_Funcionario.Item("Puerto")
            _SSL = _Row_Funcionario.Item("SSL")


            _CuerpoMensaje = Replace(_CuerpoMensaje, "&lt;", "<")
            _CuerpoMensaje = Replace(_CuerpoMensaje, "&gt;", ">")
            _CuerpoMensaje = Replace(_CuerpoMensaje, "&quot;", """")

            _CuerpoMensaje = Replace(_CuerpoMensaje, "'", "''")

            If _Enviar_Correo Then

                Dim _Fecha = "Getdate()"
                Dim _Adjuntar_Documento As Boolean = Not String.IsNullOrEmpty(_NombreFormato_Correo)

                'If _Enviar_al_otro_dia Then
                '    _Fecha = "DATEADD(D,1,Getdate())"
                'End If

                Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Demonio_Doc_Emitidos_Aviso_Correo (Id_Correo,Nombre_Correo,CodFuncionario,Asunto," &
                                "Para,Cc,Idmaeedo,Tido,Nudo,NombreFormato,Enviar,Mensaje,Fecha,Adjuntar_Documento,Doc_Adjuntos,Adjuntar_DTE,Id_Dte)" &
                                vbCrLf &
                                "Values (" & _Id_Correo & ",'" & _Nombre_Correo & "','" & _CodFuncionario & "','" & _Asunto & "','" & _Para & "','" & _Cc &
                                "'," & _Idmaeedo & ",'" & _Tido & "','" & _Nudo & "','" & _NombreFormato_Correo & "',1,'" & _CuerpoMensaje & "'," & _Fecha &
                                "," & Convert.ToInt32(_Adjuntar_Documento) & ",'',0,0)"

                _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_sql)

                _Error = _Sql.Pro_Error

                If Not String.IsNullOrEmpty(_Error) Then
                    Throw New System.Exception(_Error)
                End If

            End If

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If String.IsNullOrEmpty(_Error) Then
            Consulta_sql = "Select Cast(1 As Bit) As Enviado,'Ok' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(0 as Bit) As Enviado,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_EnviarImprimirBakapp(_Global_BaseBk2 As String,
                                       _Empresa As String,
                                       _Modalidad As String,
                                       _CodFuncionario As String,
                                       _Idmaeedo As Integer)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error = String.Empty

        Try

            _Global_BaseBk = _Global_BaseBk2

            Dim _Cl_Imprimir As New Cl_Enviar_Impresion_Diablito(_Empresa, _CodFuncionario)
            _Cl_Imprimir.SoloEnviarDocDeSucursalDelDiablito = False

            If Not _Cl_Imprimir.Fx_Enviar_Impresion_Al_Diablito(_Modalidad, _Idmaeedo) Then
                Throw New System.Exception(_Cl_Imprimir.Error)
            End If

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If String.IsNullOrEmpty(_Error) Then
            Consulta_sql = "Select Cast(1 As Bit) As Enviado,'Ok' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(0 as Bit) As Enviado,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_Traer_Documento(_Global_BaseBk2 As String, _Tido As String, _Nudo As String)

        _Sql = New Class_SQL
        Dim Consulta_sql As String
        Dim _Ds2 As DataSet

        Dim _Idmaeedo As Integer
        Dim _Error = String.Empty

        Try

            _Idmaeedo = _Sql.Fx_Trae_Dato("MAEEDO", "IDMAEEDO", "TIDO = '" & _Tido & "' And NUDO = '" & _Nudo & "'", True, False)

            If _Idmaeedo = 0 Then
                _Error = "No existe el documento " & _Tido & "-" & _Nudo
                Throw New System.Exception(_Error)
            End If

            Consulta_sql = "Insert Into MAEEDOOB (IDMAEEDO,OBDO,OCDO) Values (" & _Idmaeedo & ",'','')"
            _Sql.Fx_Ej_consulta_IDU(Consulta_sql, False)

            Consulta_sql = "Select Edo.*,Obs.OBDO,Obs.OCDO From MAEEDO Edo" & vbCrLf &
                           "Left Join MAEEDOOB Obs On Edo.IDMAEEDO = Obs.IDMAEEDO" & vbCrLf &
                           "Where Edo.IDMAEEDO = " & _Idmaeedo
            Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_Row_Documento) Then
                _Error = "No existe el documento con IDMAEEDO = " & _Idmaeedo
                Throw New System.Exception(_Error)
            End If

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If String.IsNullOrEmpty(_Error) Then
            Consulta_sql = "Select Edo.*,NOKOEN,EMAIL,EMAILCOMER,Obs.OBDO,Obs.OCDO,Cast(1 As Bit) As Enviado,'Ok' As Error,'" & _Version & "' As Version From MAEEDO Edo" & vbCrLf &
                           "Left Join MAEEDOOB Obs On Edo.IDMAEEDO = Obs.IDMAEEDO" & vbCrLf &
                           "Left Join MAEEN On KOEN = ENDO And SUEN = SUENDO" & vbCrLf &
                           "Where Edo.IDMAEEDO = " & _Idmaeedo
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(0 as Bit) As Enviado,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_Traer_Documento2(_Global_BaseBk2 As String, _Idmaeedo As Integer)

        _Sql = New Class_SQL
        Dim Consulta_sql As String
        Dim _Ds2 As DataSet

        Dim _Error = String.Empty

        Try

            Dim _New_Idmaeedo = _Sql.Fx_Trae_Dato("MAEEDO", "IDMAEEDO", "IDMAEEDO = " & _Idmaeedo, True, False)

            If _New_Idmaeedo = 0 Then
                _Error = "No existe el documento IDAMEEDO: " & _Idmaeedo
                Throw New System.Exception(_Error)
            End If

            Consulta_sql = "Insert Into MAEEDOOB (IDMAEEDO,OBDO,OCDO) Values (" & _Idmaeedo & ",'','')"
            _Sql.Fx_Ej_consulta_IDU(Consulta_sql, False)

            Consulta_sql = "Select Edo.*,Obs.OBDO,Obs.OCDO From MAEEDO Edo" & vbCrLf &
                           "Left Join MAEEDOOB Obs On Edo.IDMAEEDO = Obs.IDMAEEDO" & vbCrLf &
                           "Where Edo.IDMAEEDO = " & _Idmaeedo
            Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_Row_Documento) Then
                _Error = "No existe el documento con IDMAEEDO = " & _Idmaeedo
                Throw New System.Exception(_Error)
            End If

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If String.IsNullOrEmpty(_Error) Then
            Consulta_sql = "Select Edo.*,NOKOEN,EMAIL,EMAILCOMER,Obs.OBDO,Obs.OCDO,Cast(1 As Bit) As Enviado,'Ok' As Error,'" & _Version & "' As Version From MAEEDO Edo" & vbCrLf &
                           "Left Join MAEEDOOB Obs On Edo.IDMAEEDO = Obs.IDMAEEDO" & vbCrLf &
                           "Left Join MAEEN On KOEN = ENDO And SUEN = SUENDO" & vbCrLf &
                           "Where Edo.IDMAEEDO = " & _Idmaeedo
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(0 as Bit) As Enviado,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_Actualizar_Observaciones_Documento(_Idmaeedo As Integer,
                                                     _Observaciones As String,
                                                     _Orden_De_Compra As String)

        _Sql = New Class_SQL
        Dim Consulta_sql As String
        Dim _Ds2 As DataSet

        Dim _Error = String.Empty

        Try

            If _Idmaeedo = 0 Then
                _Error = "Falta Nro IDMAEEDO"
                Throw New System.Exception(_Error)
            End If

            Consulta_sql = "Update MAEEDOOB Set OBDO = '" & _Observaciones & "',OCDO = '" & _Orden_De_Compra & "' Where IDMAEEDO = " & _Idmaeedo
            _Sql.Fx_Ej_consulta_IDU(Consulta_sql, False)

            If Not String.IsNullOrEmpty(_Sql.Pro_Error) Then
                _Error = _Sql.Pro_Error
                Throw New System.Exception(_Error)
            End If

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If String.IsNullOrEmpty(_Error) Then
            Consulta_sql = "Select Cast(1 as Bit) As Actualizado,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(0 as Bit) As Actualizado,'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    Public Sub Sb_RevisarStockEnDetalle(_Tido As String, _DestalleJs As String, _Global_BaseBk As String)

        _Sql = New Class_SQL
        Dim _Ds2 As DataSet

        Dim _Error As String = String.Empty
        Dim _Stock_Insuficiente As Boolean = False

        Try

            'Dim _Ruta = "D:\JsonB4Android\"

            'Fx_Grabar_JsonArchivo(_EncabezadoJs, _Ruta, "EncabezadoJs")
            'Fx_Grabar_JsonArchivo(_DestalleJs, _Ruta, "DestalleJs")
            'Fx_Grabar_JsonArchivo(_DescuentosJs, _Ruta, "DescuentosJs")
            'Fx_Grabar_JsonArchivo(_ObservacionesJs, _Ruta, "ObservacionesJs")

            'Throw New System.Exception("Archivo New creados...")

            'Return

            Dim _Ds_Matriz_Documentos As New Ds_Matriz_Documentos

            _Ds_Matriz_Documentos.Clear()
            _Ds_Matriz_Documentos = New Ds_Matriz_Documentos

            Fx_LlenarDatos(_Ds_Matriz_Documentos, _DestalleJs, "Detalle_Doc")

            Dim _Tbl_Detalle As DataTable = _Ds_Matriz_Documentos.Tables("Detalle_Doc")

            For Each _Fila As DataRow In _Tbl_Detalle.Rows

                Dim _Empresa As String = _Fila.Item("Empresa")
                Dim _Sucursal As String = _Fila.Item("Sucursal")
                Dim _Bodega As String = _Fila.Item("Bodega")
                Dim _Codigo As String = _Fila.Item("Codigo")
                'Dim _Tidopa As String = _Fila.Item("")
                Dim _UnTrans As Integer = _Fila.Item("UnTrans")
                Dim _Cantidad As Double = _Fila.Item("Cantidad")

                Dim _Stock_Disponible = Fx_Stock_Disponible(_Tido, _Empresa, _Sucursal, _Bodega, _Codigo, _UnTrans, "STFI" & _UnTrans)

                'If _Tidopa = "NVV" And _Tido <> "NVV" Then

                '    If _Campo_Formula_Stock.Contains("-C") Then
                '        _Stock_Disponible += _Cantidad
                '    End If

                'End If

                If _Stock_Disponible - _Cantidad < 0 Then
                    _Stock_Insuficiente = True
                    Exit For
                End If

            Next

        Catch ex As Exception
            _Error = ex.Message
        End Try

        Consulta_sql = "Select Cast(" & Convert.ToInt32(_Stock_Insuficiente) & " As Bit) As Stock_Insuficiente," &
                       "'" & Replace(_Error, "'", "''") & "' As Error,'" & _Version & "' As Version"
        _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()



    End Sub

    <WebMethod(True)>
    Public Sub Sb_RevisarDocVsListaPrecio(_Idmaeedo As Integer, _Vnta_Dias_Venci_Coti As Integer)

        _Sql = New Class_SQL
        Dim Consulta_sql As String
        Dim _Ds2 As DataSet

        Dim _Error = String.Empty
        Dim _Respuesta = String.Empty
        Dim _RowMaeedo_Origen As DataRow
        Dim _Permitir = True
        Dim _HayDifPrecios = False
        Dim _Permiso = "Pbk00010"

        Try

            Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _Idmaeedo
            _RowMaeedo_Origen = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_RowMaeedo_Origen) Then
                Throw New System.Exception("No se pudo encontrar el documento ID: " & _Idmaeedo)
            End If

            Consulta_sql = My.Resources.Recursos_Sql.Revisar_Socumentos_VS_Lista_de_Precios
            Consulta_sql = Replace(Consulta_sql, "#Idmaeedo#", _Idmaeedo)

            Dim _Row_Diferencia As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            Dim _Diferencia As Double = _Row_Diferencia.Item("Diferencia")
            Dim _Total_New As Double = _Row_Diferencia.Item("New_VABRLI")
            Dim _Total_Old As Double = _Row_Diferencia.Item("VABRLI")

            Dim _FechaDoc As Date = _RowMaeedo_Origen.Item("FEEMDO")
            Dim _Dias_Coti As Integer = DateDiff(DateInterval.Day, _FechaDoc, Now.Date)
            Dim _Vencida As Boolean

            Dim _Dias_Venci_Coti As Integer = _Vnta_Dias_Venci_Coti '_Global_Row_Configuracion_General.Item("Vnta_Dias_Venci_Coti")

            If _Dias_Venci_Coti > 0 Then
                If _Dias_Venci_Coti < _Dias_Coti Then
                    _Vencida = True
                End If
            End If

            If _Vencida Then

                If _Diferencia < 101 And _Diferencia > -101 Then
                    _Diferencia = 0
                End If

                If CBool(_Diferencia) Then

                    _HayDifPrecios = True

                    If _Diferencia > 0 Then

                        _Respuesta = "El documento de origen tiene más de " & _Dias_Venci_Coti & " días." &
                                     vbCrLf &
                                     "Total Original: " & FormatCurrency(_Total_Old, 0) & vbCrLf &
                                     "Total Actual: " & FormatCurrency(_Total_New, 0) & vbCrLf &
                                     "Diferencia: " & FormatCurrency(_Diferencia, 0)
                        _Permitir = False

                        '(SI) Mantiene los precios originales del documento (requiere permiso)" & vbCrLf &
                        '(NO) Mantiene los precios actuales de la lista de precios"

                        'Requiere el permiso TienePermiso("Pbk00010")

                    Else

                        _Respuesta = "El documento de origen tiene más de " & _Dias_Venci_Coti & " días." &
                                      vbCrLf &
                                      "Total Original: " & FormatCurrency(_Total_Old, 0) & vbCrLf &
                                      "Total Actual: " & FormatCurrency(_Total_New, 0) & vbCrLf &
                                      "Diferencia: " & FormatCurrency(_Diferencia, 0)

                        _Permitir = True

                    End If

                End If

            End If

            'Throw New System.Exception(_Error)

        Catch ex As Exception
            _Error = ex.Message
        End Try

        If Not String.IsNullOrEmpty(_Error) Then
            _Permitir = False
            _HayDifPrecios = False
            _Respuesta = String.Empty
        End If

        Consulta_sql = "Select " &
                       "Cast(" & Convert.ToInt32(_Permitir) & " as Bit) As Permitir," &
                       "Cast(" & Convert.ToInt32(_HayDifPrecios) & " as Bit) As HayDifPrecios," &
                       "'" & Replace(_Error, "'", "''") & "' As Error," &
                       "'" & _Permiso & "' As Permiso," &
                       "'" & _Respuesta & "' As Respuesta,'" & _Version & "' As Version"
        _Ds2 = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim js As New JavaScriptSerializer

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds2, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_FormatoModalidad(_Empresa As String, _Modalidad As String, _Tido As String)

        Dim js As New JavaScriptSerializer

        _Sql = New Class_SQL

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'").ToString.Trim & ".dbo."

        Consulta_sql = "Select Top 1 Isnull(TIDO,'') As Tido,Rtrim(Ltrim(Isnull(NOTIDO,''))) As Notido,Isnull(FrmMod.Modalidad,'') As Modalidad," & vbCrLf &
                        "Isnull(FrmFx.NombreFormato,'') As NombreFomato,Isnull(FrmMod.NombreFormato,'') As NombreFomatoEnMod,Isnull(FrmFx.NroLineasXpag,0) As NroLineasXpag," & vbCrLf &
                        "CAST((Case When FrmFx.NombreFormato Is null Then 0 Else 1 End) As bit) As TieneFormato,Cast(1 As Bit) As EsCorrecto" & vbCrLf &
                        "From " & _Global_BaseBk & "Zw_Configuracion_Formatos_X_Modalidad FrmMod" & vbCrLf &
                        "Left Join " & _Global_BaseBk & "Zw_Format_01 FrmFx On FrmFx.TipoDoc = FrmMod.TipoDoc And " &
                        "FrmFx.NombreFormato = FrmMod.NombreFormato" & vbCrLf &
                        "Left Join TABTIDO On TIDO = FrmMod.TipoDoc" & vbCrLf &
                        "Where Empresa = '" & _Empresa & "' And Modalidad = '" & _Modalidad & "' And FrmMod.TipoDoc = '" & _Tido & "'"

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim _TieneFormato As Boolean

        Try
            If Not CBool(_Ds.Tables(0).Rows.Count) Then
                Throw New System.Exception("No existe formato o documento para Empresa: [" & _Empresa & "], Modalidad: [" & _Modalidad & "], Tido: [" & _Tido & "]")
            End If
            _TieneFormato = _Ds.Tables(0).Rows(0).Item("TieneFormato")
        Catch ex As Exception
            Consulta_sql = "Select Cast(0 as Bit) As EsCorrecto,'" & Replace(ex.Message, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        End Try

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Token_Generar(_Key As String)

        _Sql = New Class_SQL

        If _Sql.BaseConectada Then

            Dim js As New JavaScriptSerializer

            Dim _MinutosExpiracion As Integer = 15 ' Duración del token en minutos

            Dim _TokenGenerator As New TokenGenerator()

            Dim _Token As String = _TokenGenerator.GenerateToken(_Key, _MinutosExpiracion)

            Consulta_sql = "Select '" & _Token & "' As Token,Cast(1 As Bit) As EsCorrecto,'" & _Key & "' As Observacion"

        Else
            Consulta_sql = "Select '' As Token,Cast(0 As Bit) As EsCorrecto,'Error de conexión a la base de datos' As Observacion"
        End If

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Token_Validar(_Token As String, _Key As String)

        _Sql = New Class_SQL

        Dim _Estado As String
        Dim _EsCorrecto As Integer

        If _Sql.BaseConectada Then

            Dim js As New JavaScriptSerializer

            'Dim _MinutosExpiracion As Integer = 1 ' Duración del token en minutos

            Dim _TokenGenerator As New TokenGenerator()

            Dim principal As ClaimsPrincipal = _TokenGenerator.ValidateToken(_Token)

            If principal IsNot Nothing Then

                If principal.Identity.Name.ToString.Trim = _Key Then
                    _Estado = "Token válido para el dispositivo: " & principal.Identity.Name
                    _EsCorrecto = 1
                Else
                    _Estado = "Token no corresponde al dispositivo: " & _Key
                    _EsCorrecto = 0
                End If
            Else
                _Estado = "Token inválido o expirado."
                _EsCorrecto = 0
            End If

            Consulta_sql = "Select '" & _Estado & "' As Estado,Cast(" & _EsCorrecto & " As Bit) As EsCorrecto,'' As Observacion"

        Else
            Consulta_sql = "Select 'Error de conexión' As Estado,Cast(0 As Bit) As EsCorrecto,'Error de conexión a la base de datos' As Observacion"
        End If

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_ImprimirEnPDF2Bit(_Tido As String, _Nudo As String)

        Dim js As New JavaScriptSerializer

        _Sql = New Class_SQL

        Dim _NombreFormato As String = "Tamaño carta Alameda PDA"
        Dim _Path As String = "D:\PdfCreator impresiones"
        Dim _Idmaeedo As Integer = _Sql.Fx_Trae_Dato("MAEEDO", "IDMAEEDO", "TIDO = '" & _Tido & "' And NUDO = '" & _Nudo & "'", True)
        Dim _RutEmpresa As String = "81756300-7"

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'").ToString.Trim & ".dbo."


        Dim _Pdf_Adjunto As New Clas_PDF_Crear_Documento(_Idmaeedo,
                                                         _Tido,
                                                         _NombreFormato,
                                                         _Tido & "-" & _Nudo,
                                                         _Path, _Tido & "-" & _Nudo,
                                                         False,
                                                         _Global_BaseBk,
                                                         _RutEmpresa)

        _Pdf_Adjunto.Sb_Crear_PDF("", False, _Pdf_Adjunto.Pro_Nombre_Archivo)

        Dim _Error_Pdf = _Pdf_Adjunto.Pro_Error
        Dim _Existe_File = System.IO.File.Exists(_Pdf_Adjunto.Pro_Full_Path_Archivo_PDF & "\" & _Pdf_Adjunto.Pro_Nombre_Archivo & ".pdf")

        If String.IsNullOrEmpty(_Error_Pdf) Then

            '_Pdf_Adjunto.Sb_Abrir_Archivo()

            'If _Imprimir_Cedible Then
            '    _Pdf_Adjunto.Pro_Nombre_Archivo = _Pdf_Adjunto.Pro_Nombre_Archivo & "_Cedible"
            '    _Existe_File = System.IO.File.Exists(_Pdf_Adjunto.Pro_Full_Path_Archivo_PDF & "\" & _Pdf_Adjunto.Pro_Nombre_Archivo & ".pdf")
            '    _Pdf_Adjunto.Sb_Abrir_Archivo()
            'End If


        Else

            'MessageBoxEx.Show(Me, _Error_Pdf, "Problemas al crear el archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop)

        End If


        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Dim _TieneFormato As Boolean

        Try
            If Not CBool(_Ds.Tables(0).Rows.Count) Then
                Throw New System.Exception("No existe formato o documento, Tido: [" & _Tido & " - " & _Nudo & "]")
            End If
            _TieneFormato = _Ds.Tables(0).Rows(0).Item("TieneFormato")
        Catch ex As Exception
            Consulta_sql = "Select Cast(0 as Bit) As EsCorrecto,'" & Replace(ex.Message, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        End Try

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

#Region "Inventario"

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Inv_IngresarHoja(_Inv_Hoja As String, _Ls_Inv_Hoja_Detalle As String)

        Dim js As New JavaScriptSerializer

        _Sql = New Class_SQL

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'",, False).ToString.Trim & ".dbo."

        Dim Json = _Inv_Hoja
        Dim _Zw_Inv_Hoja = JsonConvert.DeserializeObject(Of Zw_Inv_Hoja)(Json)
        Json = _Ls_Inv_Hoja_Detalle
        Dim _Zw_Ls_Inv_Hoja_Detalle As List(Of Zw_Inv_Hoja_Detalle) = JsonConvert.DeserializeObject(Of List(Of Zw_Inv_Hoja_Detalle))(Json)

        Dim Cl_Conteo As New Cl_Conteo
        Cl_Conteo.Zw_Inv_Hoja = _Zw_Inv_Hoja
        Cl_Conteo.Ls_Zw_Inv_Hoja_Detalle = _Zw_Ls_Inv_Hoja_Detalle

        Dim _Mensaje As LsValiciones.Mensajes

        _Mensaje = Cl_Conteo.Fx_Grabar_Nueva_Hoja()

        Dim _Ds As DataSet

        If _Mensaje.EsCorrecto Then
            Consulta_sql = "Select Cast(1 as Bit) As EsCorrecto,'' As Error,'" & _Version & "' As Version," & _Mensaje.Id & " As Id,'" & Cl_Conteo.Zw_Inv_Hoja.Nro_Hoja & "' As Numero"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        Else
            Consulta_sql = "Select Cast(0 as Bit) As EsCorrecto,'" & Replace(_Mensaje.Mensaje, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        End If

        Try
            If Not CBool(_Ds.Tables(0).Rows.Count) Then
                'Throw New System.Exception("No existe formato o documento, Tido: [" & _Tido & " - " & _Nudo & "]")
            End If
        Catch ex As Exception
            Consulta_sql = "Select Cast(0 as Bit) As EsCorrecto,'" & Replace(ex.Message, "'", "''") & "' As Error,'" & _Version & "' As Version"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)
        End Try

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Dim _Respuesta = Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None)
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()

        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Inv_TraerProductoInventarioComentario(_IdInventario As Integer,
                                              _Empresa As String,
                                              _Sucursal As String,
                                              _Bodega As String,
                                              _Tipo As String,
                                              _Codigo As String)

        _Sql = New Class_SQL

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'",, False).ToString.Trim & ".dbo."

        Dim js As New JavaScriptSerializer
        Dim donde As String = ""
        Dim condicion As String = ""
        If _Tipo = "Principal" Then
            donde = "KOPR"
            condicion = "MP." & donde & " Like '" & _Codigo & "'"

        ElseIf _Tipo = "Tecnico" Then
            donde = "KOPRTE"
            condicion = "MP." & donde & " Like '" & _Codigo & "'"

        ElseIf _Tipo = "Rapido" Then
            donde = "KOPRRA"
            condicion = "MP." & donde & " Like '" & _Codigo & "'"

        ElseIf _Tipo = "Descripcion" Then
            donde = "NOKOPR"
            condicion = "MP." & donde & " Like '%" & _Codigo & "%'"

        End If

        Consulta_sql = "Select MP.KOPR as Principal ,MP.KOPRRA as Rapido, MP.KOPRTE as Tecnico,RLUD As Rtu,UD01PR As Ud1,UD02PR As Ud2,NOKOPR as Descripcion,Ft.StFisicoUd1 as StFisicoUd1 , Ft.StFisicoUd2 as StFisicoUd2,
MP.FMPR as SuperFamilia ,TABFM.NOKOFM as NombreSuper, 
MP.PFPR as Familia ,TABPF.NOKOPF as NombreFamilia, 
MP.HFPR as SubFamilia, TABHF.NOKOHF as NombreSub, MP.MRPR ,NOKOMR As 'MARCA',Cast(0 As float) As PrecioListaUd1,Cast(0 As float) As PrecioListaUd2
FROM MAEPR MP 
--INNER JOIN MAEST ST ON MP.KOPR = ST.KOPR
RIGHT JOIN TABFM ON  MP.FMPR = TABFM.KOFM
RIGHT JOIN TABPF ON  MP.FMPR = TABPF.KOFM AND MP.PFPR = TABPF.KOPF
RIGHT JOIN TABHF ON  MP.FMPR = TABHF.KOFM AND MP.PFPR = TABHF.KOPF AND  MP.HFPR = TABHF.KOHF
RIGHT JOIN TABMR On MP.MRPR = TABMR.KOMR
Left Join " & _Global_BaseBk & "Zw_Inv_FotoInventario Ft On Ft.Codigo = MP.KOPR
WHERE Ft.IdInventario = " & _IdInventario & " And Ft.Empresa = '" & _Empresa & "' AND Ft.Sucursal = '" & _Sucursal & "' AND Ft.Bodega = '" & _Bodega & "' AND " & condicion

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Try

            Dim _Lista As String = "01P"

            Consulta_sql = "SELECT Top 1 *,--PP01UD,PP02UD,DTMA01UD As DSCTOMAX,ECUACION,
                            (SELECT top 1 MELT FROM TABPP Where KOLT = '" & _Lista & "') As MELT FROM TABPRE
                            Where KOLT = '" & _Lista & "' And KOPR = '" & _Codigo & "'"
            Dim _RowPrecios As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            Dim _PrecioListaUd1 As Double = Fx_Precio_Formula_Random(_Empresa, _Sucursal, _RowPrecios, "PP01UD", "ECUACION", Nothing, True, "")
            Dim _PrecioListaUd2 As Double = Fx_Precio_Formula_Random(_Empresa, _Sucursal, _RowPrecios, "PP02UD", "ECUACIONU2", Nothing, True, "")

            _Ds.Tables(0).Rows(0).Item("PrecioListaUd1") = _PrecioListaUd1
            _Ds.Tables(0).Rows(0).Item("PrecioListaUd2") = _PrecioListaUd2

        Catch ex As Exception

            Consulta_sql = "Select 'Error_" & Replace(ex.Message, "'", "''") & "' As Codigo,'" & _Version & "' As Version"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        End Try

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Inv_TraerProductoInventario(_IdInventario As Integer,
                                              _Empresa As String,
                                              _Sucursal As String,
                                              _Bodega As String,
                                              _Tipo As String,
                                              _Codigo As String)

        _Sql = New Class_SQL

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'",, False).ToString.Trim & ".dbo."

        Dim js As New JavaScriptSerializer
        Dim donde As String = ""

        If _Tipo = "Principal" Then
            donde = "KOPR"
        ElseIf _Tipo = "Tecnico" Then
            donde = "KOPRTE"
        ElseIf _Tipo = "Rapido" Then
            donde = "KOPRRA"
        End If

        Consulta_sql = "Select MP.KOPR as Principal ,MP.KOPRRA as Rapido, MP.KOPRTE as Tecnico,RLUD As Rtu,UD01PR As Ud1,UD02PR As Ud2,NOKOPR as Descripcion," & vbCrLf &
                       "Isnull(Ft.StFisicoUd1,0) as StFisicoUd1, Isnull(Ft.StFisicoUd2,0) as StFisicoUd2," & vbCrLf &
                       "Isnull(MP.FMPR, '') as SuperFamilia ,Isnull(TABFM.NOKOFM,'') as NombreSuper," & vbCrLf &
                       "Isnull(MP.PFPR, '') as Familia ,Isnull(TABPF.NOKOPF,'') as NombreFamilia," & vbCrLf &
                       "Isnull(MP.HFPR, '') as SubFamilia, Isnull(TABHF.NOKOHF, '') as NombreSub, MP.MRPR ,Isnull(NOKOMR,'') As MARCA," & vbCrLf &
                       "Cast(0 As float) As PrecioListaUd1,Cast(0 As float) As PrecioListaUd2,Cast('' As Varchar(20)) As 'FechaUlt'" & vbCrLf &
                       "FROM MAEPR MP" & vbCrLf &
                       "Left Join TABFM ON  MP.FMPR = TABFM.KOFM" & vbCrLf &
                       "Left Join TABPF ON  MP.FMPR = TABPF.KOFM AND MP.PFPR = TABPF.KOPF" & vbCrLf &
                       "Left Join TABHF ON  MP.FMPR = TABHF.KOFM AND MP.PFPR = TABHF.KOPF AND  MP.HFPR = TABHF.KOHF" & vbCrLf &
                       "Left Join TABMR On MP.MRPR = TABMR.KOMR" & vbCrLf &
                       "Left Join " & _Global_BaseBk & "Zw_Inv_FotoInventario Ft On " &
                       "Ft.Codigo = MP.KOPR And Ft.IdInventario = " & _IdInventario &
                       " And Ft.Empresa = '" & _Empresa & "' And Ft.Sucursal = '" & _Sucursal & "' AND Ft.Bodega = '" & _Bodega & "'" & vbCrLf &
                       "Where MP." & donde & " = '" & _Codigo & "'"

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Try

            Dim _Lista As String = "01P"

            Consulta_sql = "SELECT Top 1 *,--PP01UD,PP02UD,DTMA01UD As DSCTOMAX,ECUACION,
                            (SELECT top 1 MELT FROM TABPP Where KOLT = '" & _Lista & "') As MELT FROM TABPRE
                            Where KOLT = '" & _Lista & "' And KOPR = '" & _Codigo & "'"
            Dim _RowPrecios As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            Dim _PrecioListaUd1 As Double = Fx_Precio_Formula_Random(_Empresa, _Sucursal, _RowPrecios, "PP01UD", "ECUACION", Nothing, True, "")
            Dim _PrecioListaUd2 As Double = Fx_Precio_Formula_Random(_Empresa, _Sucursal, _RowPrecios, "PP02UD", "ECUACIONU2", Nothing, True, "")

            _Ds.Tables(0).Rows(0).Item("PrecioListaUd1") = _PrecioListaUd1
            _Ds.Tables(0).Rows(0).Item("PrecioListaUd2") = _PrecioListaUd2

            Consulta_sql = "Select Top 1 TIDO,NUDO,FEEMLI From MAEDDO Where TIDO = 'FCC' And KOPRCT = '" & _Codigo & "' Order By IDMAEDDO Desc"
            Dim _RowUltComp As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If Not IsNothing(_RowUltComp) Then
                _Ds.Tables(0).Rows(0).Item("FechaUlt") = FormatDateTime(_RowUltComp.Item("FEEMLI"), DateFormat.ShortDate)
            End If

        Catch ex As Exception

            Consulta_sql = "Select 'Error_" & Replace(ex.Message, "'", "''") & "' As Codigo,'" & _Version & "' As Version"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        End Try

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Inv_TraerInfoProductoXbodega(_Empresa As String,
                                               _Sucursal As String,
                                               _Bodega As String,
                                               _Tipo As String,
                                               _Codigo As String)

        _Sql = New Class_SQL

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'",, False).ToString.Trim & ".dbo."

        Dim js As New JavaScriptSerializer
        Dim donde As String = ""

        If _Tipo = "Principal" Then
            donde = "KOPR"
        ElseIf _Tipo = "Tecnico" Then
            donde = "KOPRTE"
        ElseIf _Tipo = "Rapido" Then
            donde = "KOPRRA"
        End If

        Consulta_sql = "Select MP.KOPR As Principal,MP.KOPRRA As Rapido, MP.KOPRTE As Tecnico,RLUD As Rtu,UD01PR As Ud1,UD02PR As Ud2,NOKOPR as Descripcion," & vbCrLf &
                       "Isnull(ST.STFI1,0) as StFisicoUd1,Isnull(ST.STFI2,0) as StFisicoUd2," & vbCrLf &
                       "Isnull(MP.FMPR, '') as SuperFamilia ,Isnull(TABFM.NOKOFM,'') as NombreSuper," & vbCrLf &
                       "Isnull(MP.PFPR, '') as Familia ,Isnull(TABPF.NOKOPF,'') as NombreFamilia," & vbCrLf &
                       "Isnull(MP.HFPR, '') as SubFamilia, Isnull(TABHF.NOKOHF, '') as NombreSub, MP.MRPR ,Isnull(NOKOMR,'') As MARCA," &
                       "Cast(0 As float) As PrecioListaUd1,Cast(0 As float) As PrecioListaUd2,Cast('' As Varchar(20)) As 'FechaUlt'" & vbCrLf &
                       "From MAEPR MP" & vbCrLf &
                       "Left Join MAEST ST On MP.KOPR = ST.KOPR And ST.EMPRESA = '" & _Empresa & "' AND ST.KOSU = '" & _Sucursal & "' AND ST.KOBO = '" & _Bodega & "'" & vbCrLf &
                       "Left Join TABFM On MP.FMPR = TABFM.KOFM" & vbCrLf &
                       "Left Join TABPF ON  MP.FMPR = TABPF.KOFM AND MP.PFPR = TABPF.KOPF" & vbCrLf &
                       "Left Join TABHF ON  MP.FMPR = TABHF.KOFM AND MP.PFPR = TABHF.KOPF AND  MP.HFPR = TABHF.KOHF" & vbCrLf &
                       "Left Join TABMR On MP.MRPR = TABMR.KOMR" & vbCrLf &
                       "Where MP." & donde & " = '" & _Codigo & "'"

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Try

            Dim _Lista As String = "01P"

            Consulta_sql = "SELECT Top 1 *,--PP01UD,PP02UD,DTMA01UD As DSCTOMAX,ECUACION,
                            (SELECT top 1 MELT FROM TABPP Where KOLT = '" & _Lista & "') As MELT FROM TABPRE
                            Where KOLT = '" & _Lista & "' And KOPR = '" & _Codigo & "'"
            Dim _RowPrecios As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            Dim _PrecioListaUd1 As Double = Fx_Precio_Formula_Random(_Empresa, _Sucursal, _RowPrecios, "PP01UD", "ECUACION", Nothing, True, "")
            Dim _PrecioListaUd2 As Double = Fx_Precio_Formula_Random(_Empresa, _Sucursal, _RowPrecios, "PP02UD", "ECUACIONU2", Nothing, True, "")

            _Ds.Tables(0).Rows(0).Item("PrecioListaUd1") = _PrecioListaUd1
            _Ds.Tables(0).Rows(0).Item("PrecioListaUd2") = _PrecioListaUd2

            Consulta_sql = "Select Top 1 TIDO,NUDO,FEEMLI From MAEDDO Where TIDO = 'FCC' And KOPRCT = '" & _Codigo & "' Order By IDMAEDDO Desc"
            Dim _RowUltComp As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If Not IsNothing(_RowUltComp) Then
                _Ds.Tables(0).Rows(0).Item("FechaUlt") = FormatDateTime(_RowUltComp.Item("FEEMLI"), DateFormat.ShortDate)
            End If

        Catch ex As Exception

            Consulta_sql = "Select 'Error_" & Replace(ex.Message, "'", "''") & "' As Codigo,'" & _Version & "' As Version"
            _Ds = _Sql.Fx_Get_DataSet(Consulta_sql)

        End Try

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Inv_BuscarSector(Sector As String, IdInv As String)

        _Sql = New Class_SQL

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'",, False).ToString.Trim & ".dbo."

        Dim js As New JavaScriptSerializer

        If IsNumeric(Sector) Then
            Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Inv_Sector Where IdInventario = " & IdInv & " And Id = " & Sector
        Else
            Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Inv_Sector Where IdInventario = " & IdInv & " And Sector = '" & Sector & "'"
        End If

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Inv_BuscarInventario(Inventario As String)

        _Sql = New Class_SQL

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'",, False).ToString.Trim & ".dbo."

        Dim js As New JavaScriptSerializer

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Inv_Inventario Where Id = '" & Inventario & "'"

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=True, XmlSerializeString:=False)>
    Public Sub Sb_Inv_BuscarContador(Rut As String, Rut2 As String)

        _Sql = New Class_SQL

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'",, False).ToString.Trim & ".dbo."

        Dim js As New JavaScriptSerializer

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Inv_Contador Where Activo = '1' and Rut  != '" & Rut & "' and Rut  != '" & Rut2 & "'"

        Dim _Ds As DataSet = _Sql.Fx_Get_DataSet(Consulta_sql)

        Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
        Context.Response.ContentType = "application/json"
        Context.Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(_Ds, Newtonsoft.Json.Formatting.None))
        Context.Response.Flush()
        Context.Response.End()

    End Sub

    <WebMethod(True)>
    <Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=False, XmlSerializeString:=False)>
    Public Sub JS_ProcesarHojas(InventarioJson As String)
        Try
            ' Convertir el JSON en un JObject
            Dim inventarioJObject As JObject = JObject.Parse(InventarioJson)

            ' Obtener el objeto 'Hoja'
            Dim hoja As JObject = inventarioJObject("Hoja")

            ' Obtener el valor de 'Nro_Hoja'
            Dim nroHoja As String = hoja("Nro_Hoja").ToString()

            ' Crear un nuevo diccionario para la respuesta
            Dim MAPaux As New Dictionary(Of String, Object)
            MAPaux.Add("Numero_Hoja", nroHoja)

            ' Configurar la respuesta HTTP
            Context.Response.Cache.SetExpires(DateTime.Now.AddHours(-1))
            Context.Response.ContentType = "application/json"
            Context.Response.Write(JsonConvert.SerializeObject(MAPaux, Formatting.None))
            Context.Response.Flush()
            Context.Response.End()
        Catch ex As Exception
            ' Manejar cualquier error y devolver una respuesta adecuada
            Context.Response.StatusCode = 500
            Context.Response.ContentType = "application/json"
            Dim errorResponse As New Dictionary(Of String, String)
            errorResponse.Add("error", ex.Message)
            Context.Response.Write(JsonConvert.SerializeObject(errorResponse, Formatting.None))
            Context.Response.Flush()
            Context.Response.End()
        End Try
    End Sub

#End Region

#End Region


End Class