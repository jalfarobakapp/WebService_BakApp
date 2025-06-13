Public Class Class_Imprimir_Barras

    Dim _Sql As New Class_SQL
    Dim Consulta_sql As String

    Dim _Error As String

#Region "VARIABLES DE IMPRESION"

    Dim _Tido
    Dim _Nudo
    Dim _CodEntidad
    Dim _CodSucEntidad
    Dim _RazonSocial

    Dim _Codigo_principal
    Dim _Codigo_tecnico
    Dim _Codigo_rapido
    Dim _Codigo_Alternativo = String.Empty
    Dim _Descripcion
    Dim _Descripcion_Corta
    Dim _Desc0125 As String
    Dim _Desc2650 As String

    Dim _Desc_0135 As String
    Dim _Desc_0235 As String

    Dim _Ubicacion
    Dim _Precio_ud1
    Dim _Precio_ud2
    Dim _PrecioNetoXRtu
    Dim _PrecioBrutoXRtu
    Dim _Rtu
    Dim _Ud1
    Dim _Ud2

    Dim _Marca_Pr
    Dim _Nodim1
    Dim _Nodim2
    Dim _Nodim3
    Dim _PrecioLc1
    Dim _FechaProgramada_Futuro As Date

    Dim _PU01_Neto, _PU02_Neto As Double
    Dim _PU01_Bruto, _PU02_Bruto As Double

    Dim _Cantidad

    Dim _Wms_Ubic_BakApp
    Dim _Wms_Ubicacion_Codigo
    Dim _Wms_Ubicacion_Nombre
    Dim _Wms_Ubicacion_Columna
    Dim _Wms_Ubicacion_Fila
    Dim _Wms_Sector_Codigo
    Dim _Wms_Sector_Nombre
    Dim _Wms_Mapa_Nombre

    Dim _Stock_Minimo_Ubic
    Dim _Stock_Maximo_Ubic

    Dim _Tidopa As String
    Dim _Nudopa As String

    Public Sub New()
        _Sql = New Class_SQL
    End Sub

    Public Property [Error] As String
        Get
            Return _Error
        End Get
        Set(value As String)
            _Error = value
        End Set
    End Property

#End Region

#Region "IMPRIMIR CODIGO"
    Function Fx_Imprimir_Etiquea_Producto(_NombreEtiqueta As String,
                                          _Codigo As String,
                                          _CodLista As String,
                                          _Empresa As String,
                                          _Sucursal As String,
                                          _Bodega As String,
                                          _CodAlternativo As String) As LsValiciones.Mensajes

        Dim _Mensaje As New LsValiciones.Mensajes

        Try

            Dim _RowProducto As DataRow = Fx_DatosProducto(_Codigo,
                                                          _CodLista,
                                                          _Empresa,
                                                          _Sucursal,
                                                          _Bodega)

            If IsNothing(_RowProducto) Then
                Throw New System.Exception("No se encontró el producto")
            End If

            _Codigo_principal = _Codigo
            _Codigo_tecnico = _RowProducto.Item("KOPRTE")
            _Codigo_rapido = _RowProducto.Item("KOPRRA")
            _Descripcion = _RowProducto.Item("NOKOPR").ToString.Trim
            _Descripcion_Corta = _RowProducto.Item("NOKOPRRA").ToString.Trim

            If Not String.IsNullOrEmpty(_CodAlternativo) Then
                _Codigo_Alternativo = _CodAlternativo.ToString.Trim
            Else
                _Codigo_Alternativo = _RowProducto.Item("Codigo_Alternativo").ToString.Trim
            End If

            _Ud1 = _RowProducto.Item("UD01PR").ToString.Trim
            _Ud2 = _RowProducto.Item("UD02PR").ToString.Trim

            _Marca_Pr = _RowProducto.Item("Marca").ToString.Trim

            _Ubicacion = _RowProducto.Item("Ubic_Random")

            _Precio_ud1 = _RowProducto.Item("Precio_ud1")
            _Precio_ud2 = _RowProducto.Item("Precio_ud2")

            _PU01_Neto = _RowProducto.Item("PU01_Neto")
            _PU02_Neto = _RowProducto.Item("PU02_Neto")
            _PU01_Bruto = _RowProducto.Item("PU01_Bruto")
            _PU02_Bruto = _RowProducto.Item("PU02_Bruto")

            _Rtu = _RowProducto.Item("RLUD")
            _PrecioNetoXRtu = _RowProducto.Item("PrecioNetoXRtu")
            _PrecioBrutoXRtu = _RowProducto.Item("PrecioBrutoXRtu")

            _Stock_Minimo_Ubic = _RowProducto.Item("Stock_Minimo_Ubic")
            _Stock_Maximo_Ubic = _RowProducto.Item("Stock_Maximo_Ubic")

            _Descripcion = Replace(_Descripcion, Chr(34), "")
            _Desc0125 = Mid(_Descripcion, 1, 25)
            _Desc2650 = Mid(_Descripcion, 26, 50)

            Dim _Descri25_Aju = Fx_AjustarTexto(_Descripcion, 35)

            Dim _Desc_Aju = Split(_Descri25_Aju, vbCrLf, 2)

            If _Desc_Aju.Length > 1 Then 'AndAlso _Desc_Aju(1).ToString.Replace(vbCrLf, " ").ToString.Length <= 25 Then

                _Desc_0135 = _Desc_Aju(0)
                _Desc_0235 = _Desc_Aju(1)

            Else

                _Desc_0135 = _Desc_Aju(0)
                _Desc_0235 = String.Empty

            End If

            _Nodim1 = _RowProducto.Item("NODIM1").ToString.Trim
            _Nodim2 = _RowProducto.Item("NODIM2").ToString.Trim
            _Nodim3 = _RowProducto.Item("NODIM3").ToString.Trim

            Dim _Dim1, _Dim2 As Double

            If IsNumeric(_Nodim1) Then
                _Dim1 = Val(_Nodim1)
            End If

            If IsNumeric(_Nodim2) Then
                _Dim2 = Val(_Nodim2)
            End If

            _FechaProgramada_Futuro = _RowProducto.Item("FechaProgramada")

            If _Dim1 = 0 Then _Dim1 = 1
            If _Dim2 = 0 Then _Dim2 = 1

            _PrecioLc1 = (_Precio_ud1 / _Dim1) * _Dim2

            Dim _RowEtiqueta As DataRow = Fx_TraeEtiqueta(_NombreEtiqueta)

            If IsNothing(_RowEtiqueta) Then
                Throw New System.Exception("No se encontró la etiqueta: " & _NombreEtiqueta)
            End If

            Dim _Texto = _RowEtiqueta.Item("FUNCION")

            Fx_Imprimir_Etiqueta(_Texto)

            _Mensaje.EsCorrecto = True
            _Mensaje.Detalle = "Etiqueta impresa correctamente"
            _Mensaje.Mensaje = "Etiqueta impresa correctamente"
            _Mensaje.Tag = Fx_Imprimir_Etiqueta(_Texto)

            If String.IsNullOrEmpty(_Mensaje.Tag) Then
                Throw New System.Exception("La etiqueta esta en blanco")
            End If

        Catch ex As Exception
            _Mensaje.EsCorrecto = False
            _Mensaje.Detalle = "Error al imprimir etiqueta: " & ex.Message
            _Mensaje.Mensaje = ex.Message
            _Mensaje.Tag = String.Empty
            _Error = ex.Message
        End Try

        Return _Mensaje

    End Function

    Function Fx_Producto_Ubicaciones(_Codigo As String,
                                     _Lista As String,
                                     _Empresa As String,
                                     _Sucursal As String,
                                     _Bodega As String) As DataTable

        Consulta_sql = "Select MAEPR.*,Isnull((Select top 1 DATOSUBIC From TABBOPR
                        Where EMPRESA = '" & _Empresa & "' AND KOSU = '" & _Sucursal & "' AND KOBO = '" & _Bodega & "' And KOPR = '" & _Codigo & "'),'') As 'Ubic_Random',
                        Isnull((Select Top 1 PP01UD From TABPRE Where KOLT = '" & _Lista & "' And KOPR = '" & _Codigo & "'),0) As Precio_ud1,
                        Isnull((Select Top 1 PP02UD From TABPRE Where KOLT = '" & _Lista & "' And KOPR = '" & _Codigo & "'),0) As Precio_ud2,
                        Isnull((Select top 1 PM From MAEPREM Where EMPRESA = '" & _Empresa & "' And KOPR = '" & _Codigo & "'),0) As 'PM',
                        Isnull((Select top 1 PPUL01 From MAEPREM Where EMPRESA = '" & _Empresa & "' And KOPR = '" & _Codigo & "'),0) As 'PU01',
                        Isnull((Select top 1 PPUL02 From MAEPREM Where EMPRESA = '" & _Empresa & "' And KOPR = '" & _Codigo & "'),0) As 'PU02',
                        Isnull((Select top 1 KOPRAL From TABCODAL Where KOEN = '' And KOPR = '" & _Codigo & "'),'') As Codigo_Alternativo,
                        Isnull((Select Top 1 NOKOMR From TABMR Where KOMR = MRPR),'') As Marca,
                        Cast(0 As Float) As PU01_Neto,Cast(0 As Float) As PU02_Neto,Cast(0 As Float) As PU01_Bruto,Cast(0 As Float) As PU02_Bruto,
                        Mapa.Nombre_Mapa,isnull(Sector.Nombre_Sector,'') As Nombre_Sector,Isnull(Ubic.Descripcion_Ubic,'') As Descripcion_Ubic,
                        Isnull(Ubic.Columna,'') As Columna,Isnull(Ubic.NomColumna,'') As NomColumna,
                        Isnull(Ubic.Fila,'') As Fila,Isnull(Ubic.Alto,0) As Alto,Isnull(Ubic.Ancho,0) As Ancho,Isnull(Ubic.Peso_Max,0) As Peso_Max,
                        PrUbic.*
                        From MAEPR 
                        Inner Join " & _Global_BaseBk & "Zw_Prod_Ubicacion PrUbic On PrUbic.Empresa = '01' And PrUbic.Sucursal = '" & _Empresa & "' And PrUbic.Codigo = KOPR
                        Left Join " & _Global_BaseBk & "Zw_WMS_Ubicaciones_Bodega Ubic On Ubic.Empresa = PrUbic.Empresa And Ubic.Sucursal = PrUbic.Sucursal And Ubic.Codigo_Ubic = PrUbic.Codigo_Ubic
	                    Left Join " & _Global_BaseBk & "Zw_WMS_Ubicaciones_Mapa_Enc Mapa On Mapa.Id_Mapa = PrUbic.Id_Mapa
		                Left Join " & _Global_BaseBk & "Zw_WMS_Ubicaciones_Mapa_Det Sector On Sector.Id_Mapa = PrUbic.Id_Mapa And Sector.Codigo_Sector = PrUbic.Codigo_Sector

                        Where KOPR = '" & _Codigo & "'
                        Order By Primaria Desc"

        Fx_Producto_Ubicaciones = _Sql.Fx_Get_DataTable(Consulta_sql)

    End Function

#End Region

#Region "PROCEDIMIENTOS PRIVADOS"

#Region "TRAER ETIQUETA"
    Private Function Fx_TraeEtiqueta(_NombreEtiqueta As String) As DataRow

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Tbl_DisenoBarras Where NombreEtiqueta = '" & _NombreEtiqueta & "'"
        Dim _Tbl As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Dim _Row As DataRow

        If CBool(_Tbl.Rows.Count) Then
            _Row = _Tbl.Rows(0)
        End If

        Return _Row

    End Function
#End Region

#Region "IMPRIMIR EL ARCHIVO"

    Private Function Fx_Imprimir_Etiqueta(_Texto As String) As String

        Dim _TextoOri As String = _Texto
        Dim _Fecha_impresion As Date = Now

        _Error = String.Empty

        'Try

        _Texto = Replace(_Texto, "<CODIGO_PR>", _Codigo_principal.ToString.Trim)
        _Texto = Replace(_Texto, "<CODIGO_TC>", _Codigo_tecnico.ToString.Trim)
        _Texto = Replace(_Texto, "<CODIGO_RA>", _Codigo_rapido.ToString.Trim)
        _Texto = Replace(_Texto, "<CODIGO_ALT>", _Codigo_Alternativo.ToString.Trim)

        _Texto = Replace(_Texto, "<UD1_PR>", _Ud1.ToString.Trim)
        _Texto = Replace(_Texto, "<UD2_PR>", _Ud2.ToString.Trim)

        Dim _Descripcion_cortamr As String

        If Not String.IsNullOrEmpty(_Marca_Pr) Then
            _Descripcion_cortamr = _Descripcion_Corta.ToString.Replace(_Marca_Pr, "").Trim
        End If

        _Texto = Replace(_Texto, "<DESCRIPCION_CORTASMR>", _Descripcion_cortamr)

        _Texto = Replace(_Texto, "<DESCRIPCION_PR>", _Descripcion)
        _Texto = Replace(_Texto, "<DESCRIPCION_CORTA>", _Descripcion_Corta)

        Dim _Desc0125mr As String
        Dim _Desc2650mr As String

        If Not String.IsNullOrEmpty(_Marca_Pr) Then
            _Desc0125mr = _Desc0125.ToString.Replace(_Marca_Pr, "").Trim
            _Desc2650mr = _Desc2650.ToString.Replace(_Marca_Pr, "").Trim
        End If

        _Texto = Replace(_Texto, "<DESCRIPCION_1-25MR>", _Desc0125mr)
        _Texto = Replace(_Texto, "<DESCRIPCION_26-50MR>", _Desc2650mr)

        _Texto = Replace(_Texto, "<DESCRIPCION_1-25>", _Desc0125)
        _Texto = Replace(_Texto, "<DESCRIPCION_26-50>", _Desc2650)

        _Texto = Replace(_Texto, "<DESCRIPCION_1-35>", _Desc_0135)
        _Texto = Replace(_Texto, "<DESCRIPCION_2-35>", _Desc_0235)

        _Texto = Replace(_Texto, "<UBICACION_PR>", _Ubicacion)
        _Texto = Replace(_Texto, "<UBICACION>", _Ubicacion)
        _Texto = Replace(_Texto, "<MARCA_PR>", _Marca_Pr)
        _Texto = Replace(_Texto, "<NODIM1>", _Nodim1)
        _Texto = Replace(_Texto, "<NODIM2>", _Nodim2)
        _Texto = Replace(_Texto, "<NODIM3>", _Nodim3)
        _Texto = Replace(_Texto, "<FECHAPROGRFUTURO>", Format(_FechaProgramada_Futuro, "dd-MM-yyyy"))
        _Texto = Replace(_Texto, "<RTU>", _Rtu)

        '_Texto = Replace(_Texto, "<UBIC_BAKAPP>", _Ubic_BakApp)
        _Texto = Replace(_Texto, "<WMS_UBIC_COLUMNA>", _Wms_Ubicacion_Columna)
        _Texto = Replace(_Texto, "<WMS_UBIC_FILA>", _Wms_Ubicacion_Fila)
        _Texto = Replace(_Texto, "<WMS_UBIC_CODIGO>", _Wms_Ubicacion_Codigo)
        _Texto = Replace(_Texto, "<WMS_UBIC_DESCR>", _Wms_Ubicacion_Nombre)

        _Texto = Replace(_Texto, "<WMS_SECTOR_CODIGO> ", _Wms_Sector_Codigo)
        _Texto = Replace(_Texto, "<WMS_SECTOR_DESC>", _Wms_Sector_Nombre)

        _Texto = Replace(_Texto, "<WMS_MAPA_NOMBRE> ", _Wms_Mapa_Nombre)

        _Texto = Replace(_Texto, "<STOCK_MIN_UBIC>", _Stock_Minimo_Ubic)
        _Texto = Replace(_Texto, "<STOCK_MAX_UBIC>", _Stock_Maximo_Ubic)

        _Texto = Replace(_Texto, "<PRECIO_UD1>", _Precio_ud1)
        _Texto = Replace(_Texto, "<PRECIO_UD2>", _Precio_ud2)

        Dim _vPrecioNetoXRtu As String
        Dim _vPrecioBrutoXRtu As String

        Try
            _vPrecioNetoXRtu = Fx_Formato_Numerico(_PrecioNetoXRtu, "9", False)
        Catch ex As Exception
            _vPrecioNetoXRtu = "?"
        End Try

        Try
            _vPrecioBrutoXRtu = Fx_Formato_Numerico(_PrecioBrutoXRtu, "9", False)
        Catch ex As Exception
            _vPrecioBrutoXRtu = "?"
        End Try

        If _Texto.Contains("PBRUTOUD1X6") Or _Texto.Contains("PBRUTOUD1XMULTIPLO") Or _Texto.Contains("PBRUTOUD2XMULTIPLO") Then

            Consulta_sql = "Select Top 1 * From TABCODAL Where KOPRAL = '" & _Codigo_Alternativo & "' And KOPR = '" & _Codigo_principal & "'"
            Dim _Row_Kopral As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If Not IsNothing(_Row_Kopral) Then

                Dim _Unimulti As Integer = 1
                Dim _Multiplo As Integer = _Row_Kopral.Item("MULTIPLO")

                Dim _PrecioMulti As Double

                If _Texto.Contains("PBRUTOUD2XUNIMULTI2") Then
                    _Unimulti = 2
                End If

                If _Unimulti = 1 Then : _PrecioMulti = _PU01_Bruto : Else : _PrecioMulti = _PU02_Bruto : End If

                _PrecioBrutoXRtu = _Multiplo * _PrecioMulti

                _vPrecioBrutoXRtu = Fx_Formato_Numerico(_PrecioBrutoXRtu, "9", False)

                _Texto = Replace(_Texto, "<PBRUTOUD1X6>", _vPrecioBrutoXRtu)

            End If

        End If

        'If _Texto.Contains("PBRUTOUD1X6") Then
        '    _PrecioBrutoXRtu = 6 * _PU01_Bruto
        '    _vPrecioBrutoXRtu = Fx_Formato_Numerico(_PrecioBrutoXRtu, "9", False)
        '    '_Texto = Replace(_Texto, "<PNETOXRTU_UD1>", _vPrecioNetoXRtu)
        '    _Texto = Replace(_Texto, "<PBRUTOUD1X6>", _vPrecioBrutoXRtu)
        'End If

        Dim _St_PU01_Neto As String = Fx_Formato_Numerico(_PU01_Neto, "9", False)
        Dim _St_PU02_Neto As String = Fx_Formato_Numerico(_PU02_Neto, "9", False)
        Dim _St_PU01_Bruto As String = Fx_Formato_Numerico(_PU01_Bruto, "9", False)
        Dim _St_PU02_Bruto As String = Fx_Formato_Numerico(_PU02_Bruto, "9", False)

        Dim _St_PU01_Neto_d1 As String = Fx_Formato_Numerico(_PU01_Neto, "9,9", False)
        Dim _St_PU02_Neto_d1 As String = Fx_Formato_Numerico(_PU02_Neto, "9,9", False)
        Dim _St_PU01_Bruto_d1 As String = Fx_Formato_Numerico(_PU01_Bruto, "9,9", False)
        Dim _St_PU02_Bruto_d1 As String = Fx_Formato_Numerico(_PU02_Bruto, "9,9", False)

        Dim _St_PU01_Neto_d2 As String = Fx_Formato_Numerico(_PU01_Neto, "9,99", False)
        Dim _St_PU02_Neto_d2 As String = Fx_Formato_Numerico(_PU02_Neto, "9,99", False)
        Dim _St_PU01_Bruto_d2 As String = Fx_Formato_Numerico(_PU01_Bruto, "9,99", False)
        Dim _St_PU02_Bruto_d2 As String = Fx_Formato_Numerico(_PU02_Bruto, "9,99", False)

        Dim _St_PU01_Neto2 As String = Fx_Formato_Numerico(_PU01_Neto, "99.999", False)
        Dim _St_PU02_Neto2 As String = Fx_Formato_Numerico(_PU02_Neto, "99.999", False)
        Dim _St_PU01_Bruto2 As String = Fx_Formato_Numerico(_PU01_Bruto, "99.999", False)
        Dim _St_PU02_Bruto2 As String = Fx_Formato_Numerico(_PU02_Bruto, "99.999", False)

        Dim _St_PU01_Neto3 As String = Fx_Formato_Numerico(_PU01_Neto, "999.999", False)
        Dim _St_PU02_Neto3 As String = Fx_Formato_Numerico(_PU02_Neto, "999.999", False)
        Dim _St_PU01_Bruto3 As String = Fx_Formato_Numerico(_PU01_Bruto, "999.999", False)
        Dim _St_PU02_Bruto3 As String = Fx_Formato_Numerico(_PU02_Bruto, "999.999", False)

        Dim _Lc_PrecioEspLC1 As String
        Dim _Lc_PrecioEspLC2 As String

        Try
            _Lc_PrecioEspLC1 = Fx_Formato_Numerico(_PrecioLc1, "9", False)
        Catch ex As Exception
            _Lc_PrecioEspLC1 = "?"
        End Try

        Try
            _Lc_PrecioEspLC2 = Fx_Formato_Numerico(_PrecioLc1, "9.999.999", False)
        Catch ex As Exception
            _Lc_PrecioEspLC2 = "?"
        End Try


        Dim _St_PU01_Neto4 As String = Fx_Formato_Numerico(_PU01_Neto, "9.999.999", False)
        Dim _St_PU02_Neto4 As String = Fx_Formato_Numerico(_PU02_Neto, "9.999.999", False)
        Dim _St_PU01_Bruto4 As String = Fx_Formato_Numerico(_PU01_Bruto, "9.999.999", False)
        Dim _St_PU02_Bruto4 As String = Fx_Formato_Numerico(_PU02_Bruto, "9.999.999", False)

        _Texto = Replace(_Texto, "<PNETO_UD1>", _St_PU01_Neto)
        _Texto = Replace(_Texto, "<PNETO_UD2>", _St_PU02_Neto)

        ' Neto reserva 4 espacios a la derecha
        _Texto = Replace(_Texto, "<PNETO_UD1_2>", _St_PU01_Neto2)
        _Texto = Replace(_Texto, "<PNETO_UD2_2>", _St_PU02_Neto2)
        ' Neto reserva 5 espacios a la derecha
        _Texto = Replace(_Texto, "<PNETO_UD1_3>", _St_PU01_Neto3)
        _Texto = Replace(_Texto, "<PNETO_UD2_3>", _St_PU02_Neto3)
        ' Neto reserva 6 espacios a la derecha para millones
        _Texto = Replace(_Texto, "<PNETO_UD1_4>", _St_PU01_Neto4)
        _Texto = Replace(_Texto, "<PNETO_UD2_4>", _St_PU02_Neto4)

        _Texto = Replace(_Texto, "<PBRUTO_UD1>", _St_PU01_Bruto)
        _Texto = Replace(_Texto, "<PBRUTO_UD2>", _St_PU02_Bruto)

        _Texto = Replace(_Texto, "<PBRUTO_UD1_2>", _St_PU01_Bruto2)
        _Texto = Replace(_Texto, "<PBRUTO_UD2_2>", _St_PU02_Bruto2)

        _Texto = Replace(_Texto, "<PBRUTO_UD1_3>", _St_PU01_Bruto3)
        _Texto = Replace(_Texto, "<PBRUTO_UD2_3>", _St_PU02_Bruto3)

        _Texto = Replace(_Texto, "<PBRUTO_UD1_4>", _St_PU01_Bruto4)
        _Texto = Replace(_Texto, "<PBRUTO_UD2_4>", _St_PU02_Bruto4)

        'Precios especiales La Colchaguina
        _Texto = Replace(_Texto, "<PBRUTO_LCESP1>", _Lc_PrecioEspLC1)
        _Texto = Replace(_Texto, "<PBRUTO_LCESP2>", _Lc_PrecioEspLC2)

        _Texto = Replace(_Texto, "<FECHA_IMPRESION>", _Fecha_impresion)
        _Texto = Replace(_Texto, "<FECHA_IMPRESION2>", _Fecha_impresion.ToShortDateString)
        _Texto = Replace(_Texto, "<FECHA_IMPRESION3>", Format(_Fecha_impresion, "dd-MM-yyyy"))

        _Texto = Replace(_Texto, "<TIDO>", _Tido)
        _Texto = Replace(_Texto, "<NUDO>", _Nudo)

        _Texto = Replace(_Texto, "<TIDOPA>", _Tidopa)
        _Texto = Replace(_Texto, "<NUDOPA>", _Nudopa)

        Dim _PU01_BrutoCtdo As String
        Dim _PU02_BrutoCtdo As String

        _PU01_BrutoCtdo = Fx_FormatearValorCentrado(_PU01_Bruto, 12)
        _PU02_BrutoCtdo = Fx_FormatearValorCentrado(_PU02_Bruto, 12)

        _Texto = Replace(_Texto, "<PBRUTO_UD1_CENT>", _PU01_BrutoCtdo)
        _Texto = Replace(_Texto, "<PBRUTO_UD2_CENT>", _PU02_BrutoCtdo)

        Dim _Nudopa_Sc As String

        Try
            _Nudopa_Sc = CInt(_Nudopa)
        Catch ex As Exception
            _Nudopa_Sc = _Nudopa
        End Try

        _Texto = Replace(_Texto, "<NUDOPA_SC>", _Nudopa_Sc.Trim)

        Return _Texto

    End Function

    Function Fx_FormatearValorCentrado(valor As String, Optional largo As Integer = 12) As String
        ' Intenta convertir a número y dar formato con puntos como separador de miles
        Dim valorNumerico As Decimal
        If Decimal.TryParse(valor, valorNumerico) Then
            valor = valorNumerico.ToString("#,##0") ' Ej: 9.999.999
        End If

        ' Agrega el símbolo $
        Dim valorFormateado As String = "$ " & valor

        ' Si el resultado es mayor que el largo, recorta
        If valorFormateado.Length > largo Then
            valorFormateado = valorFormateado.Substring(0, largo)
        End If

        ' Centrado visual: calcula espacios a la izquierda y derecha
        Dim espaciosTotales As Integer = largo - valorFormateado.Length
        Dim espaciosIzquierda As Integer = espaciosTotales \ 2
        Dim espaciosDerecha As Integer = espaciosTotales - espaciosIzquierda

        ' Devuelve el valor con espacios a ambos lados
        Return New String(" "c, espaciosIzquierda) & valorFormateado & New String(" "c, espaciosDerecha)
    End Function

#End Region

#Region "TRAER DATOS DEL PRODUCTO"

    Private Function Fx_DatosProducto(_Codigo As String,
                                      _CodLista As String,
                                      _Empresa As String,
                                      _Sucursal As String,
                                      _Bodega As String,
                                      Optional _CodEntidad As String = "",
                                      Optional _Codigo_Ubic As String = "",
                                      Optional _ImprimirDesdePrecioFuturo As Boolean = False,
                                      Optional _Id_PrecioFuturo As Integer = 0) As DataRow

        If String.IsNullOrEmpty(_Codigo_Ubic) Then

            Consulta_sql = "Select Top 1 * From " & _Global_BaseBk & "Zw_Prod_Ubicacion Where Codigo = '" & _Codigo & "' And Primaria = 1"

        Else

            Consulta_sql = "Select Top 1 * From " & _Global_BaseBk & "Zw_Prod_Ubicacion 
                            Where Empresa = '" & _Empresa & "' And Sucursal = '" & _Sucursal & "'" & Space(1) &
                            "And Bodega = '" & _Bodega & "' And Codigo_Ubic = '" & _Codigo_Ubic & "' And Codigo = '" & _Codigo & "'"

        End If

        Dim _TblUbicacion As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Dim _Ubic_BakApp As String
        Dim _Stock_Minimo_Ubic As Double
        Dim _Stock_Maximo_Ubic As Double

        If CBool(_TblUbicacion.Rows.Count) Then

            _Ubic_BakApp = _TblUbicacion.Rows(0).Item("Codigo_Ubic")
            _Stock_Minimo_Ubic = _TblUbicacion.Rows(0).Item("Stock_Minimo_Ubic")
            _Stock_Maximo_Ubic = _TblUbicacion.Rows(0).Item("Stock_Maximo_Ubic")

        End If

        Consulta_sql = "Select TOP 1 *,Isnull((Select top 1 DATOSUBIC From TABBOPR" & vbCrLf &
                       "Where EMPRESA = '" & _Empresa & "' AND KOSU = '" & _Sucursal &
                       "' AND KOBO = '" & _Bodega & "' And KOPR = '" & _Codigo & "'),'') As 'Ubic_Random'," & vbCrLf &
                       "'" & _Ubic_BakApp & "' As 'Ubic_BakApp'," & vbCrLf &
                       "Cast(" & De_Num_a_Tx_01(_Stock_Minimo_Ubic, False, 5) & " As Float) As 'Stock_Minimo_Ubic'," & vbCrLf &
                       "Cast(" & De_Num_a_Tx_01(_Stock_Maximo_Ubic, False, 5) & " As Float) As 'Stock_Maximo_Ubic'," & vbCrLf &
                       "Isnull((Select Top 1 PP01UD From TABPRE Where KOLT = '" & _CodLista & "' And KOPR = '" & _Codigo & "'),0) As Precio_ud1," & vbCrLf &
                       "Isnull((Select Top 1 PP02UD From TABPRE Where KOLT = '" & _CodLista & "' And KOPR = '" & _Codigo & "'),0) As Precio_ud2," & vbCrLf &
                       "Cast(0 As Float) As 'PrecioNetoXRtu',Cast(0 As Float) As 'PrecioBrutoXRtu'," & vbCrLf &
                       "Isnull((Select top 1 PM From MAEPREM Where EMPRESA = '" & _Empresa & "' And KOPR = '" & _Codigo & "'),0) As 'PM'," & vbCrLf &
                       "Isnull((Select top 1 PPUL01 From MAEPREM Where EMPRESA = '" & _Empresa & "' And KOPR = '" & _Codigo & "'),0) As 'PU01'," & vbCrLf &
                       "Isnull((Select top 1 PPUL02 From MAEPREM Where EMPRESA = '" & _Empresa & "' And KOPR = '" & _Codigo & "'),0) As 'PU02'," & vbCrLf &
                       "Isnull((Select top 1 KOPRAL From TABCODAL Where KOEN = '" & _CodEntidad & "' And KOPR = '" & _Codigo & "'),'') As Codigo_Alternativo," & vbCrLf &
                       "Isnull((Select Top 1 NOKOMR From TABMR Where KOMR = MRPR),'') As Marca," & vbCrLf &
                       "Cast(0 As Float) As PU01_Neto,Cast(0 As Float) As PU02_Neto,Cast(0 As Float) As PU01_Bruto,Cast(0 As Float) As PU02_Bruto,Getdate() As FechaProgramada" & vbCrLf &
                       "From MAEPR Where KOPR = '" & _Codigo & "'"

        Dim _Tbl As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        If CBool(_Tbl.Rows.Count) Then

            Dim _RowProducto As DataRow = _Tbl.Rows(0)

            Sb_Incorporar_Precios(_Empresa, _RowProducto, _CodLista, _ImprimirDesdePrecioFuturo, _Id_PrecioFuturo)

            Return _RowProducto

        Else
            Return Nothing
        End If

    End Function

    Sub Sb_Incorporar_Precios(_Empresa As String,
                              ByRef _RowProducto As DataRow,
                              _CodLista As String,
                              _ImprimirDesdePrecioFuturo As Boolean,
                              _Id_PrecioFuturo As Integer)

        'Dim _RowProducto As DataRow = _Tbl.Rows(0)
        Dim _Codigo As String = _RowProducto.Item("KOPR")

        Consulta_sql = "Select Isnull(Sum(POIM),0) As Impuesto From TABIM" & Space(1) &
                       "Where KOIM In (Select KOIM From TABIMPR Where KOPR = '" & _Codigo & "')"
        Dim _RowImpuestos As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        Dim _PorIva As Double = _RowProducto.Item("POIVPR")
        Dim _PorIla As Double = _RowImpuestos.Item("Impuesto")

        Dim _Iva = _PorIva / 100
        Dim _Ila = _PorIla / 100

        Dim _Impuestos As Double = 1 + (_Iva + _Ila)


        Consulta_sql = "Select Top 1 *,(Select top 1 MELT From TABPP Where KOLT = '" & _CodLista & "') As MELT From TABPRE
                            Where KOLT = '" & _CodLista & "' And KOPR = '" & _Codigo & "'"
        Dim _RowPrecios As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        Dim _Ecuacion As String
        Dim _Ecuacionu2 As String

        Dim _PrecioListaUd1 As Double
        Dim _PrecioListaUd2 As Double

        Dim _Melt As String = _RowPrecios.Item("MELT")

        If Not IsNothing(_RowPrecios) Then

            _Ecuacion = NuloPorNro(_RowPrecios.Item("ECUACION").ToString.Trim, "")
            _Ecuacionu2 = NuloPorNro(_RowPrecios.Item("ECUACIONU2").ToString.Trim, "")

            '_PrecioListaUd1 = Fx_Funcion_Ecuacion_Random(Nothing, _CodEntidad, _Ecuacion, _Codigo, 1, _RowPrecios, 0, 0, 0)
            '_PrecioListaUd2 = Fx_Funcion_Ecuacion_Random(Nothing, _CodEntidad, _Ecuacionu2, _Codigo, 2, _RowPrecios, 0, 0, 0)

            _PrecioListaUd1 = Fx_Precio_Formula_Random(_Empresa, _CodEntidad, _RowPrecios, "PP01UD", "ECUACION", Nothing, True, "")
            _PrecioListaUd2 = Fx_Precio_Formula_Random(_Empresa, _CodEntidad, _RowPrecios, "PP02UD", "ECUACIONU2", Nothing, True, "")

            If _PrecioListaUd1 = 0 Then _PrecioListaUd1 = NuloPorNro(_RowPrecios.Item("PP01UD"), 0)
            If _PrecioListaUd2 = 0 Then _PrecioListaUd2 = NuloPorNro(_RowPrecios.Item("PP02UD"), 0)

        End If

        If _ImprimirDesdePrecioFuturo Then

            If CBool(_Id_PrecioFuturo) Then
                Consulta_sql = "Select Top 1 LEnc.Codigo, NombreProgramacion, FechaCreacion, FechaProgramada, Funcionario, Activo,LDet.*" & vbCrLf &
                           "From " & _Global_BaseBk & "Zw_ListaLC_Programadas LEnc" & vbCrLf &
                           "Inner Join " & _Global_BaseBk & "Zw_ListaLC_Programadas_Detalles LDet On LEnc.Id = LDet.Id_Enc" & vbCrLf &
                           "Where LDet.Id = " & _Id_PrecioFuturo
            Else
                Consulta_sql = "Select Top 1 LEnc.Codigo, NombreProgramacion, FechaCreacion, FechaProgramada, Funcionario, Activo,LDet.*" & vbCrLf &
                           "From " & _Global_BaseBk & "Zw_ListaLC_Programadas LEnc" & vbCrLf &
                           "Inner Join " & _Global_BaseBk & "Zw_ListaLC_Programadas_Detalles LDet On LEnc.Id = LDet.Id_Enc" & vbCrLf &
                           "Where LEnc.Codigo = '" & _Codigo & "' And LDet.Lista = '" & _CodLista & "' Order by LEnc.Id"
            End If


            Dim _Row_PrecioFuturo As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If Not IsNothing(_Row_PrecioFuturo) Then

                _PrecioListaUd1 = _Row_PrecioFuturo.Item("PrecioUd1")
                _PrecioListaUd2 = _Row_PrecioFuturo.Item("PrecioUd2")

                _RowProducto.Item("Precio_ud1") = _PrecioListaUd1
                _RowProducto.Item("Precio_ud2") = _PrecioListaUd2
                _RowProducto.Item("FechaProgramada") = _Row_PrecioFuturo.Item("FechaProgramada")

            End If

        End If


        Dim _PU01_Neto, _PU02_Neto As Double
        Dim _PU01_Bruto, _PU02_Bruto As Double

        If _Melt = "N" Then
            _PU01_Neto = _PrecioListaUd1
            _PU02_Neto = _PrecioListaUd2
            _PU01_Bruto = Math.Round(_PU01_Neto * _Impuestos, 0)
            _PU02_Bruto = Math.Round(_PU02_Neto * _Impuestos, 0)
        End If

        If _Melt = "B" Then
            _PU01_Bruto = _PrecioListaUd1
            _PU02_Bruto = _PrecioListaUd2
            _PU01_Neto = Math.Round(_PU01_Bruto / _Impuestos, 2)
            _PU02_Neto = Math.Round(_PU02_Bruto / _Impuestos, 2)
        End If

        _RowProducto.Item("PU01_Neto") = _PU01_Neto
        _RowProducto.Item("PU02_Neto") = _PU02_Neto

        _RowProducto.Item("PU01_Bruto") = _PU01_Bruto
        _RowProducto.Item("PU02_Bruto") = _PU02_Bruto

        _RowProducto.Item("PrecioNetoXRtu") = _RowProducto.Item("RLUD") * _PU01_Neto
        _RowProducto.Item("PrecioBrutoXRtu") = _RowProducto.Item("RLUD") * _PU01_Bruto

    End Sub

#End Region

#Region "TRAER DATOS DE UBICACION"

    Private Function Fx_Datos_Ubicacion(_Empresa As String,
                                        _Sucursal As String,
                                        _Bodega As String,
                                        _Id_Mapa As Integer,
                                        _Codigo_Sector As String,
                                        _Codigo_Ubic As String) As DataRow

        Consulta_sql = "Select Mapa.Nombre_Mapa,Sector.Nombre_Sector,Ubic.*
                        From " & _Global_BaseBk & "Zw_WMS_Ubicaciones_Bodega Ubic
                        Left Join " & _Global_BaseBk & "Zw_WMS_Ubicaciones_Mapa_Enc Mapa On Mapa.Id_Mapa = Ubic.Id_Mapa
                        Left Join " & _Global_BaseBk & "Zw_WMS_Ubicaciones_Mapa_Det Sector On Sector.Id_Mapa = Ubic.Id_Mapa And Sector.Codigo_Sector = Ubic.Codigo_Sector
                        Where Ubic.Empresa = '" & _Empresa & "' And Ubic.Sucursal = '" & _Sucursal & "' And Ubic.Bodega = '" & _Bodega &
                        "' And Ubic.Id_Mapa = " & _Id_Mapa & " And Ubic.Codigo_Sector = '" & _Codigo_Sector & "' And Ubic.Codigo_Ubic = '" & _Codigo_Ubic & "'"

        Dim _RowUbicacion As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        Return _RowUbicacion

    End Function

#End Region

#End Region

End Class

