Imports System.Reflection.Assembly
Imports DevComponents.DotNetBar
Imports System.IO
Imports System.Security.Cryptography
Imports System.Drawing.Printing
Imports System.Drawing


Public Module Funciones_Especiales_BakApp

    Dim Consulta_sql As String

    Function Fx_Cambiar_Numeracion_Modalidad(_Tido As String,
                                             _Empresa As String,
                                             _Modalidad As String) As Boolean

        Dim _Sql As New Class_SQL()

        Dim _Consulta_sql = "Select Top 1 " & _Tido & " From CONFIEST Where MODALIDAD = '" & _Modalidad & "'"
        Dim _Tbl As DataTable = _Sql.Fx_Get_DataTable(_Consulta_sql) 'get_Tablas(_Consulta_sql, cn1)

        Dim _Nudo_Modalidad As String

        _Consulta_sql = String.Empty

        If CBool(_Tbl.Rows.Count) Then

            _Nudo_Modalidad = Trim(_Tbl.Rows(0).Item(_Tido))

            If String.IsNullOrEmpty(_Nudo_Modalidad) Then
                If Fx_Cambiar_Numeracion_Modalidad(_Tido, _Empresa, "  ") Then
                    _Consulta_sql = String.Empty
                End If
            ElseIf _Nudo_Modalidad = "0000000000" Then
                _Consulta_sql = String.Empty
            Else

                Dim Continua As Boolean = True

                If Not String.IsNullOrEmpty(Trim(_Nudo_Modalidad)) Then

                    Dim _ProxNumero As String = Fx_Proximo_NroDocumento(_Nudo_Modalidad, 10)

                    _Consulta_sql = "UPDATE CONFIEST SET " & _Tido & " = '" & _ProxNumero & "'" & vbCrLf &
                                    "WHERE EMPRESA = '" & _Empresa & "' AND  MODALIDAD = '" & _Modalidad & "'"

                End If

            End If

        End If

        If Not String.IsNullOrEmpty(_Consulta_sql) Then
            Return _Sql.Fx_Ej_consulta_IDU(_Consulta_sql)
        End If

    End Function

    Public Function Fx_Licencia(_RutEmpresa As String, ByRef _Mensaje As String) As Boolean

        Dim _Sql As New Class_SQL()

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Licencia Where Rut = '" & _RutEmpresa & "'"
        Dim _TblLicencia As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        If CBool(_TblLicencia.Rows.Count) Then

            With _TblLicencia.Rows(0)

                Dim _Rut = .Item("Rut")
                Dim _Fecha_caduca As Date = .Item("Fecha_caduca")
                Dim _Cant_licencias = .Item("Cant_licencias")

                Dim _Llave1 As String = .Item("Llave1")
                Dim _Llave2 As String = .Item("Llave2")
                Dim _Llave3 As String = .Item("Llave3")
                Dim _Llave4 As String = .Item("Llave4")

                Dim _LLaves = Fx_Genera_Licencia_BakApp(_Rut, _Fecha_caduca, _Cant_licencias, "b4s1ng4")

                Dim _Llave1_R = _LLaves(0)
                Dim _Llave2_R = _LLaves(1)
                Dim _Llave3_R = _LLaves(2)
                Dim _Llave4_R = _LLaves(3) '= Encripta_md5(_Llave1 & _Llave2 & _Llave3)

                Dim _Dias = DateDiff(DateInterval.Day, FechaDelServidor, _Fecha_caduca)

                If _Dias > 0 Then

                    If _Dias < 10 Then
                        _Mensaje = "Faltan " & _Dias & " días para que caduque la licencia"
                        Return True
                    End If

                Else
                    _Mensaje = "Sistema sin licencias de uso"
                    Return False
                End If

                If _Llave1_R <> _Llave1 Or
                   _Llave2_R <> _Llave2 Or
                   _Llave3_R <> _Llave3 Or
                   _Llave4_R <> _Llave4 Then

                    _Mensaje = "Licencia corrupta, error en base de datos"
                    Return False
                Else
                    Return True
                End If

            End With

        Else

            _Mensaje = "No existe llave para el uso del sistema"
            Return False

        End If

    End Function

    Function Fx_Genera_Licencia_BakApp(_RutEmpresa As String,
                                        _FechaCaduca As Date,
                                        _CantLicencias As Integer,
                                        _Palabra_X As String) As String()

        Dim _Llave1, _Llave2, _Llave3, _Llave4 As String

        _Llave1 = Encripta_md5(Trim(_RutEmpresa) & _Palabra_X)
        _Llave2 = Encripta_md5(Format(_FechaCaduca, "yyyyMMdd"))
        _Llave3 = Encripta_md5(_CantLicencias & _Palabra_X)
        _Llave4 = Encripta_md5(_Llave1 & _Llave2 & _Llave3 & _Palabra_X)

        Dim Licencia(3) As String

        Licencia(0) = _Llave1
        Licencia(1) = _Llave2
        Licencia(2) = _Llave3
        Licencia(3) = _Llave4

        Return Licencia

    End Function

    Private Function Encripta_md5(TextoAEncriptar As String) As String
        Dim vlo_MD5 As New MD5CryptoServiceProvider
        Dim vlby_Byte(), vlby_Hash() As Byte
        Dim vls_TextoEncriptado As String = ""

        'Convierte texto a encriptar a Bytes
        vlby_Byte = System.Text.Encoding.UTF8.GetBytes(TextoAEncriptar)

        'Aplicación del algoritmo hash
        vlby_Hash = vlo_MD5.ComputeHash(vlby_Byte)

        'Convierte la matriz de byte en una cadena
        For Each vlby_Aux As Byte In vlby_Hash
            vls_TextoEncriptado += vlby_Aux.ToString("x2")
        Next

        'Retorno de función
        Return vls_TextoEncriptado
    End Function

    Public Function Fx_Tipo_Grab_Modalidad(_TipoDoc As String,
                                           _NrNumeroDoco As String) As String


        Dim Continua As Boolean = True

        If String.IsNullOrEmpty(Trim(_NrNumeroDoco)) Then
            Return "EnBlanco"
        ElseIf _NrNumeroDoco = "0000000000" Then
            Return "Puros_Ceros"
        Else
            Return "Con_Numeración"
        End If

    End Function

    Function _Dev_HoraGrab(Hora As String)

        Dim _HH, _MM, _SS As Double
        Dim _Horagrab As Integer

        _HH = Mid(Hora, 1, 2)
        _MM = Mid(Hora, 4, 2)
        _SS = Mid(Hora, 7, 2)

        _Horagrab = Math.Round((_HH * 3600) + (_MM * 60) + _SS, 0)

        Return _Horagrab

    End Function

    Function Fx_Decodifica_Horagrab(_Horagrab As Integer) As String

        Dim _Hora = Math.Floor(_Horagrab * 1.0 / 3600)
        Dim _Minutos = Math.Floor((_Horagrab * 1.0 / 3600 - _Hora) * 60)
        Dim _Segundos = Math.Round(((_Horagrab * 1.0 / 3600 - _Hora) * 60 - _Minutos) * 60, 0)

        Fx_Decodifica_Horagrab = _Hora & ":" & _Minutos

    End Function

    Function Fx_Fecha_y_Hora_del_Servidor() As DateTime

        Dim _Sql As New Class_SQL()

        Consulta_sql = "select getdate() As Fecha"
        Dim _Row As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)
        Fx_Fecha_y_Hora_del_Servidor = _Row.Item("Fecha")

    End Function

    Function Fx_Trae_Permiso_Bk(_CodUsuario As String,
                                _CodPremiso As String) As DataTable

        Dim _Sql As New Class_SQL()

        Dim _Tbl As DataTable

        Consulta_sql = "Select * From " & _Global_BaseBk & "ZW_PermisosVsUsuarios" & vbCrLf &
                       "Where CodPermiso = '" & _CodPremiso & "' And CodUsuario = '" & _CodUsuario & "'"

        _Tbl = _Sql.Fx_Get_DataTable(Consulta_sql)

        Return _Tbl

    End Function

    Function Fx_Traer_Datos_Entidad(_CodEntidad As String, _SucEntidad As String) As DataSet

        Dim _Sql As New Class_SQL()

        Dim _Ds_Entidad As DataSet

        Consulta_sql = My.Resources.Recursos_Sql.SqlQuery_Datos_Entidad
        Consulta_sql = Replace(Consulta_sql, "#CodEntidad#", _CodEntidad)
        Consulta_sql = Replace(Consulta_sql, "#SucEntidad#", _SucEntidad)

        _Ds_Entidad = _Sql.Fx_Get_DataSet(Consulta_sql)

        If CBool(_Ds_Entidad.Tables(0).Rows.Count) Then
            Dim _Rut As String = _Ds_Entidad.Tables(0).Rows(0).Item("RTEN").ToString.Trim
            If _Rut.Contains("-") Then
                Dim _Rt = Split(_Rut, "-")
                _Rut = _Rt(0)
            End If
            _Rut = FormatNumber(_Rut, 0) & "-" & RutDigito(_Rut)
            _Ds_Entidad.Tables(0).Rows(0).Item("Rut") = _Rut
        End If

        Return _Ds_Entidad

    End Function

    Function TraeClaveRD(Texto As String) As String

        Dim valorAscii As Integer
        Dim PassEncriptado, Letra As String
        Dim CadenaRD As Long


        Texto = Trim(Texto)
        For x = 1 To Len(Texto)
            Letra = Mid(Texto, x, 1)
            valorAscii = Asc(Letra)
            'txtAscii.Text = valor.ToString

            If x = 1 Then
                CadenaRD = (17225 + valorAscii) * 1
            ElseIf x = 2 Then
                CadenaRD = (1847 + valorAscii) * 8
            ElseIf x = 3 Then
                CadenaRD = (1217 + valorAscii) * 27
            ElseIf x = 4 Then
                CadenaRD = (237 + valorAscii) * 64
            ElseIf x = 5 Then
                CadenaRD = (201 + valorAscii) * 125
            End If

            PassEncriptado = PassEncriptado & CadenaRD
            CadenaRD = 0
        Next

        Return PassEncriptado

    End Function

    Function Fx_Stock_Disponible(_Tido As String,
                                 _Empresa As String,
                                 _Sucursal As String,
                                 _Bodega As String,
                                 _Codigo As String,
                                 _Ud As Integer,
                                 _Campo As String) As Double


        Dim _Sql As New Class_SQL()

        Consulta_sql = "Select Top 1 * From TABTIDO Where TIDO = '" & _Tido & "'"
        Dim _RowTido As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        Dim _Campo_Formula_Stock As String

        Try
            _Campo_Formula_Stock = _RowTido.Item("STOCK")
        Catch ex As Exception
            _Campo_Formula_Stock = String.Empty
        End Try

        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "A", "[A]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "B", "[B]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "C", "[C]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "D", "[D]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "E", "[E]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "F", "[F]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "G", "[G]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "H", "[H]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "I", "[I]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "J", "[J]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "K", "[K]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "L", "[L]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "M", "[M]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "N", "[N]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "O", "[O]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "P", "[P]")
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "Q", "[Q]")


        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[A]", "STFI" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[B]", "STDV" & _Ud)

        'Comprometido
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[C]", "(STOCNV" & _Ud & "+Isnull(StComp" & _Ud & ",0))")

        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[D]", "STDV" & _Ud & "C")

        'Pedido
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[E]", "(STOCNV" & _Ud & "C+Isnull(StPedi" & _Ud & ",0))")

        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[F]", "DESPNOFAC" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[G]", "RECENOFAC" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[H]", "PRESALCLI" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[I]", "PRESDEPRO" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[J]", "CONSALCLI" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[K]", "CONDESPRO" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[L]", "DEVENGNCV" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[M]", "DEVENFNCC" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[N]", "DEVSINNCV" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[O]", "DEVSINNCC" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[P]", "STENFAB" & _Ud)
        _Campo_Formula_Stock = Replace(_Campo_Formula_Stock, "[Q]", "STREQFAB" & _Ud)

        If String.IsNullOrEmpty(_Campo_Formula_Stock) Then
            _Campo_Formula_Stock = _Campo
        End If

        Consulta_sql = "Select " & _Campo_Formula_Stock & " As Stock_Disponible" & vbCrLf &
                       "From MAEST" & vbCrLf &
                       "Left Join " & Global_BaseBk & "Zw_Prod_Stock On EMPRESA = Empresa And KOSU = Sucursal And KOBO = Bodega And KOPR = Codigo" & vbCrLf &
                       "Where" & vbCrLf &
                       "EMPRESA = '" & _Empresa & "' And KOSU = '" & _Sucursal & "'" & Space(1) &
                       "And KOBO = '" & _Bodega & "' And KOPR = '" & _Codigo & "'"

        Dim _RowStock As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If Not (_RowStock Is Nothing) Then
            Fx_Stock_Disponible = _RowStock.Item("Stock_Disponible")
        End If

    End Function

    Public Function Fx_Suma_cantidades(CAMPO As String,
                                       TABLA As String,
                                       Optional condicion As String = "",
                                       Optional Decimales As Integer = 0) As Double

        Dim _Sql As New Class_SQL()

        Dim Suma As Double

        If condicion <> "" Then
            condicion = " Where " & condicion
        End If

        Consulta_sql = "SELECT ROUND(SUM(" & CAMPO & ")," & Decimales & ") AS CAMPO FROM " & TABLA & condicion & ""
        Dim _Tbl As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Dim cuenta As Long = _Tbl.Rows.Count
        Dim dr As DataRow = _Tbl.Rows(0)

        Suma = NuloPorNro(dr("CAMPO"), 0)

        Return Suma

        If CBool(cuenta) Then
            If IsDBNull(dr("CAMPO")) Then
                Return 0
            Else
                Return RTrim(dr("CAMPO"))
            End If
        Else
            Return 0
        End If

    End Function

    Function Fx_Solo_Enteros(_Cantidad As Double,
                             _Divisible As String) As Boolean

        Dim _Sql As New Class_SQL()

        Dim _Cant_Tiene_Decimales As Boolean
        Dim _Campo_Divisible As String

        If CBool(_Cantidad) Then

            If IsNumeric(_Cantidad) Then
                If CInt(_Cantidad) = _Cantidad Then
                    ' es entero
                    _Cant_Tiene_Decimales = False
                Else
                    ' es decimal
                    _Cant_Tiene_Decimales = True
                End If
            End If

            Dim _Solo_Enteros_Ud1, _Solo_Enteros_Ud2 As Boolean

            If _Cant_Tiene_Decimales Then
                If _Divisible = "0" Or _Divisible = "N" Then
                    Fx_Solo_Enteros = True
                End If
            End If

        Else
            Return True
        End If

    End Function

    Function Fx_Precio_Formula(_CPrecio As _Campo_Precio,
                           _RowPrecio As DataRow)

        Dim _Sql As New Class_SQL()

        Dim _Ej_Fx_documento As Boolean = _RowPrecio.Item("Ej_Fx_documento")

        Dim _Lista = _RowPrecio.Item("Lista")
        Dim _Codigo = _RowPrecio.Item("Codigo")

        Dim _Rtu = De_Num_a_Tx_01(_RowPrecio.Item("Rtu"), False, 5)
        Dim _Precio As Double '= _RowPrecio.Item("Precio")
        'im _Precio2 As Double = _RowPrecio.Item("Precio2")

        Dim _Formula '= Split(_RowPrecio.Item("Formula"), "#")

        If _CPrecio = _Campo_Precio.Precio_Ud1 Then
            _Formula = Split(_RowPrecio.Item("Formula"), "#")
            _Precio = _RowPrecio.Item("Precio")
        Else
            _Formula = Split(_RowPrecio.Item("Formula2"), "#")
            _Precio = _RowPrecio.Item("Precio2")
        End If

        'Dim _Formula2 = Split(_RowPrecio.Item("Formula2"), "#")

        Dim _Flete = De_Num_a_Tx_01(_RowPrecio.Item("Flete"), False, 5)
        Dim _Margen = De_Num_a_Tx_01(_RowPrecio.Item("Margen"), False, 5)
        Dim _Margen_Adicional = De_Num_a_Tx_01(_RowPrecio.Item("Margen_Adicional"), False, 5)
        Dim _Margen2 = De_Num_a_Tx_01(_RowPrecio.Item("Margen2"), False, 5)
        Dim _Margen_Adicional2 = De_Num_a_Tx_01(_RowPrecio.Item("Margen_Adicional2"), False, 5)
        Dim _Costo = De_Num_a_Tx_01(_RowPrecio.Item("Costo"), False, 5)
        Dim _Costo2 = De_Num_a_Tx_01(_RowPrecio.Item("Costo2"), False, 5)

        Dim _Pm = De_Num_a_Tx_01(_RowPrecio.Item("Pm"), False, 5)
        Dim _UC_Ud1 = De_Num_a_Tx_01(_RowPrecio.Item("UC_Ud1"), False, 5)
        Dim _UC_Ud2 = De_Num_a_Tx_01(_RowPrecio.Item("UC_Ud2"), False, 5)

        Dim _Fx1, _Fx2, _Redondeo
        Dim _New_Precio

        If _Ej_Fx_documento Then
            If CBool(_Formula.Length) Then
                _Fx1 = _Formula(0)

                If _Formula.Length > 1 Then
                    _Redondeo = _Formula(1)
                Else
                    _Redondeo = 0
                End If

                _Fx1 = Replace(_Fx1, "Flete", _Flete)
                '_Fx1 = Replace(_Fx1, "Ila", 1)
                '_Fx1 = Replace(_Fx1, "Iva", 1)
                _Fx1 = Replace(_Fx1, "Costo", _Costo)
                _Fx1 = Replace(_Fx1, "Costo2", _Costo2)
                _Fx1 = Replace(_Fx1, "Rtu", _Rtu)
                _Fx1 = Replace(_Fx1, "Pm", _Pm)
                _Fx1 = Replace(_Fx1, "UC_Ud1", _UC_Ud1)
                _Fx1 = Replace(_Fx1, "UC_Ud2", _UC_Ud2)
                _Fx1 = Replace(_Fx1, "Margen", _Margen)
                _Fx1 = Replace(_Fx1, "Margen_Adicional", _Margen_Adicional)
                _Fx1 = Replace(_Fx1, "Margen2", _Margen2)
                _Fx1 = Replace(_Fx1, "Margen_Adicional2", _Margen_Adicional2)

                _Fx1 = Replace(_Fx1, ",", ".")

                'Dim _Sql As New Class_SQL(Cadena_ConexionSQL_Server)

                Consulta_sql = "Select " & _Fx1 & " As Valor"
                Dim _RowPr As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

                _Precio = _RowPr.Item("Valor") '_Sql.Fx_Trae_Dato(, _Fx1, "", "") '_
                '   "Lista = '" & _Lista & "' And Codigo = '" & _Codigo & "'")

                _New_Precio = Fx_Redondear_Precio(_Precio, _Redondeo)

            End If
        Else
            _New_Precio = _Precio
        End If

        Return _New_Precio

    End Function

    Function Fx_Precio_Formula_Random(_Empresa As String,
                                      _Sucursal As String,
                                      _RowPrecio As DataRow,
                                      _Campo_Precio As String,
                                      _Campo_Ecuacion As String,
                                      _RowCostos_PM As DataRow,
                                      _Aplicar_Formula_Dinamica As Boolean,
                                      _Koen As String)

        Dim _Sql As New Class_SQL()

        If (_RowPrecio Is Nothing) Then
            Return 0
        End If

        Dim _Lista = _RowPrecio.Item("KOLT")
        Dim _Codigo = _RowPrecio.Item("KOPR")

        Dim _Rtu = De_Num_a_Tx_01(_RowPrecio.Item("RLUD"), False, 5)
        Dim _Precio As Double

        Dim _Formula

        Dim _Ecuacion As String = NuloPorNro(_RowPrecio.Item(_Campo_Ecuacion), "").ToString.Trim

        Dim _Ejecutar_Ecuacion = False

        If _Aplicar_Formula_Dinamica Then

            If Not String.IsNullOrEmpty(_Ecuacion) Then
                _Ejecutar_Ecuacion = (_Ecuacion = LCase(_Ecuacion))
            End If

        Else

            _Ejecutar_Ecuacion = True

        End If


        _Precio = _RowPrecio.Item(_Campo_Precio)
        Dim _New_Precio As Double

        If _Ejecutar_Ecuacion Then

            _Formula = Split(_Ecuacion, "#")

            Dim _Pp01ud = De_Num_a_Tx_01(_RowPrecio.Item("PP01UD"), False, 5)
            Dim _Pp02ud = De_Num_a_Tx_01(_RowPrecio.Item("PP02UD"), False, 5)

            Dim _Mg01ud = De_Num_a_Tx_01(_RowPrecio.Item("MG01UD"), False, 5)
            Dim _Mg02ud = De_Num_a_Tx_01(_RowPrecio.Item("MG02UD"), False, 5)

            Dim _Dtma01ud = De_Num_a_Tx_01(_RowPrecio.Item("DTMA01UD"), False, 5)
            Dim _Dtma02ud = De_Num_a_Tx_01(_RowPrecio.Item("DTMA02UD"), False, 5)

            Dim _Pm As String = 0
            Dim _Ppul01 As String = 0
            Dim _Ppul02 As String = 0
            Dim _Pmsuc As String = 0

            If (_RowCostos_PM Is Nothing) Then
                'aqui puede ser 

                Consulta_sql = "Select Top 1 PM,PM AS PM01,PPUL01,PPUL02,Isnull(Round(PMSUC,5),0) As PMSUC
                            From MAEPREM EM
                            Left Join MAEPMSUC SUC On EM.EMPRESA = SUC.EMPRESA AND SUC.KOSU = '" & _Sucursal & "' AND EM.KOPR = SUC.KOPR
                            Where EM.EMPRESA = '" & _Empresa & "' And EM.KOPR = '" & _Codigo & "'"

                _RowCostos_PM = _Sql.Fx_Get_DataRow(Consulta_sql)

            End If

            If Not (_RowCostos_PM Is Nothing) Then

                _Pm = Math.Round(_RowCostos_PM.Item("PM01"), 5)
                _Ppul01 = Math.Round(_RowCostos_PM.Item("PPUL01"), 5)
                _Ppul02 = Math.Round(_RowCostos_PM.Item("PPUL02"), 5)
                _Pmsuc = Math.Round(_RowCostos_PM.Item("PMSUC"), 5)

            End If

            Dim _Fx1, _Redondeo


            _Fx1 = UCase(_Formula(0))

            If _Formula.Length > 1 Then
                _Redondeo = Trim(_Formula(1))
            Else
                _Redondeo = 0
            End If

            If String.IsNullOrEmpty(_Fx1) Then

                _New_Precio = 0

            Else

                If _Fx1.ToString.Contains("<") Then
                    _Fx1 = Fx_Traer_Campo_Desde_Otra_Lista(_Empresa, _Sucursal, _Codigo, _Fx1, _Koen)
                End If

                _Fx1 = Replace(_Fx1, "RLUD", _Rtu)

                _Fx1 = Replace(_Fx1, "PMSUC", _Pmsuc)
                _Fx1 = Replace(_Fx1, "PM", _Pm)
                _Fx1 = Replace(_Fx1, "PPUL01", _Ppul01)
                _Fx1 = Replace(_Fx1, "PPUL02", _Ppul02)
                _Fx1 = Replace(_Fx1, "PP01UD", _Pp01ud)
                _Fx1 = Replace(_Fx1, "PP02UD", _Pp02ud)

                _Fx1 = Replace(_Fx1, "MG01UD", _Mg01ud)
                _Fx1 = Replace(_Fx1, "MG02UD", _Mg02ud)

                _Fx1 = Replace(_Fx1, "DTMA01UD", _Dtma01ud)
                _Fx1 = Replace(_Fx1, "DTMA02UD", _Dtma02ud)

                _Fx1 = Replace(_Fx1, ",", ".")

                _Fx1 = UCase(_Fx1)

                'aqui esta el error
                Sb_Buscar_Valor_En_Dimensiones(_Empresa, _Fx1, _Codigo, _Koen)


                Consulta_sql = "Select " & _Fx1 & " As Valor"
                Dim _RowPr As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

                _Precio = _RowPr.Item("Valor")

                _Redondeo = Fx_Redondeo_Random(_Redondeo)

            End If

            _New_Precio = Fx_Redondear_Precio(_Precio, _Redondeo)

        Else
            _New_Precio = _Precio
        End If

        _New_Precio = Math.Round(_New_Precio, 2)

        Return _New_Precio

    End Function

    Function Fx_Traer_Campo_Desde_Otra_Lista(_Empresa As String, _Sucursal As String, _Codigo As String, _Ecuacion As String, _Koen As String) As String

        Dim _Sql As New Class_SQL()

        Dim _Ecuacion_Original As String = _Ecuacion

        Dim _Ecuaciones = Split(_Ecuacion, ">")
        Dim _Listas() As String
        Dim _Filtro_Listas As String

        Dim _Cont = 0

        For i = 0 To _Ecuaciones.Length - 1

            Dim _Lt = _Ecuaciones(i)

            If _Lt.Contains("<") Then
                _Lt = Replace(_Lt, "<", "")
                ReDim Preserve _Listas(_Cont)
                _Listas(_Cont) = _Lt
                _Cont += 1
            End If

        Next

        _Filtro_Listas = Generar_Filtro_IN_Arreglo(_Listas, False)

        Consulta_sql = "Select * From TABPP Where KOLT In (" & _Filtro_Listas & ")"
        Dim _Tbl_Listas As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Dim _Campo As String

        For Each _FLista As DataRow In _Tbl_Listas.Rows

            Dim _Kolt As String = _FLista.Item("KOLT")
            _Campo = "<" & _Kolt & ">"

            If _Ecuacion.Contains(_Campo) Then

                Consulta_sql = "Select * From TABPRE Where KOLT = '" & _Kolt & "' And KOPR = '" & _Codigo & "'"
                Dim _RowPrecio As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

                Dim _Contador = 0

                Consulta_sql = "Select COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME = 'TABPRE'"
                Dim _Tbl_Campos_Tabpre As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

                For Each _FColumnas As DataRow In _Tbl_Campos_Tabpre.Rows

                    Dim _Columna As String = _FColumnas.Item("COLUMN_NAME").ToString.Trim
                    Dim _Campo_Lista As String = _Campo & _Columna
                    'Dim _Resultado As String

                    Dim _Campo_Precio As String = _Columna
                    Dim _Campo_Ecacion As String = String.Empty

                    If _Ecuacion.Contains(_Campo_Lista) Then

                        Select Case _Campo_Precio
                            Case "PP01UD"
                                _Campo_Ecacion = "ECUACION"
                            Case "PP02UD"
                                _Campo_Ecacion = "ECUACIONU2"
                            Case "MG01UD"
                                _Campo_Ecacion = "EMG01UD"
                            Case "MG02UD"
                                _Campo_Ecacion = "EMG01UD"
                            Case "DTMA01UD"
                                _Campo_Ecacion = "EDTMA01UD"
                            Case "DTMA02UD"
                                _Campo_Ecacion = "DTMA02UD"
                            Case Else
                                If _Contador = 28 Then
                                    _Campo_Ecacion = _Tbl_Campos_Tabpre.Rows(_Contador + 1).Item("COLUMN_NAME")
                                End If
                        End Select

                        Dim _Valor = Fx_Precio_Formula_Random(_Empresa, _Sucursal, _RowPrecio, _Campo_Precio, _Campo_Ecacion, Nothing, False, _Koen)

                        _Ecuacion = Replace(_Ecuacion, _Campo_Lista, LCase(_Valor))

                        If _Ecuacion <> _Ecuacion_Original Then
                            Return _Ecuacion
                        End If

                    End If

                    _Contador += 1

                Next

            End If

        Next

        Return _Ecuacion

    End Function

    Sub Sb_Buscar_Valor_En_Dimensiones(_Empresa As String, ByRef _Fx1 As String, _Codigo As String, _Koen As String)

        Dim _Sql As New Class_SQL()

        Dim _Contiene_Campos As Boolean

        For Each _Caracter As String In _Fx1.ToString
            _Contiene_Campos = CBool(InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ", _Caracter))
            If _Contiene_Campos Then
                Exit For
            End If
        Next

        If Not _Contiene_Campos Then
            Return
        End If

        Consulta_sql = "Select * From PNOMDIM Where DEPENDENCI = 'Valor_propio'"
        Dim _Tbl_Dimension_Valor_Propio As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        For Each _FDim_Vp As DataRow In _Tbl_Dimension_Valor_Propio.Rows

            Dim _Dimension = _FDim_Vp.Item("CODIGO").ToString.Trim
            Dim _Valor = De_Num_a_Tx_01(_FDim_Vp.Item("VALOR"), False, 5)

            _Fx1 = Replace(_Fx1, _Dimension, _Valor)

            For Each _Caracter As String In _Fx1.ToString
                _Contiene_Campos = CBool(InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ", _Caracter))
                If _Contiene_Campos Then
                    Exit For
                End If
            Next

            If Not _Contiene_Campos Then
                Return
            End If

        Next

        Consulta_sql = "Select * From PNOMDIM Where DEPENDENCI = 'Por_producto'"
        Dim _Tbl_Dimension_Por_Producto As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        For Each _FDim_Vp As DataRow In _Tbl_Dimension_Por_Producto.Rows

            Dim _Dimension = _FDim_Vp.Item("CODIGO").ToString.Trim
            Dim _Valor_Dim As Double = _Sql.Fx_Trae_Dato("PDIMEN", _Dimension, "EMPRESA = '" & _Empresa & "' And CODIGO = '" & _Codigo & "'", True, False)
            Dim _Valor = De_Num_a_Tx_01(_Valor_Dim, False, 5)

            _Fx1 = Replace(_Fx1, _Dimension, _Valor)

            For Each _Caracter As String In _Fx1.ToString
                _Contiene_Campos = CBool(InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ", _Caracter))
                If _Contiene_Campos Then
                    Exit For
                End If
            Next

            If Not _Contiene_Campos Then
                Return
            End If

        Next

        Consulta_sql = "Select * From PNOMDIM 
                        Inner Join INFORMATION_SCHEMA.COLUMNS On PNOMDIM.CODIGO = COLUMN_NAME And DATA_TYPE In ('float','int')
                        Where DEPENDENCI = 'Por_entidad' And TABLE_NAME = 'PDIMCLI'"
        Dim _Tbl_Dimension_Por_Entidad As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        For Each _FDim_Vp As DataRow In _Tbl_Dimension_Por_Entidad.Rows

            Dim _Dimension = _FDim_Vp.Item("CODIGO").ToString.Trim
            Dim _Valor_Dim As Double = _Sql.Fx_Trae_Dato("PDIMCLI", _Dimension, "CODIGO = '" & _Koen & "'", True, False)
            Dim _Valor = De_Num_a_Tx_01(_Valor_Dim, False, 5)

            _Fx1 = Replace(_Fx1, _Dimension, _Valor)

            For Each _Caracter As String In _Fx1.ToString
                _Contiene_Campos = CBool(InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ", _Caracter))
                If _Contiene_Campos Then
                    Exit For
                End If
            Next

            If Not _Contiene_Campos Then
                Return
            End If

        Next

    End Sub

    Function FX_Traer_Valor_Concepto(_Empresa As String, _RowConcepto As DataRow, _Koen As String) As Double

        Dim _Sql As New Class_SQL()

        Dim _Codigo As String = _RowConcepto.Item("KOCT")
        Dim _Poct As Double = _RowConcepto.Item("POCT")
        Dim _Ecuct As String = _RowConcepto.Item("ECUCT")

        Dim _Pp01ud = De_Num_a_Tx_01(_RowConcepto.Item("POCT"), False, 5)

        Dim _Formula = Split(_Ecuct, "#")
        Dim _Fx1, _Redondeo

        _Fx1 = UCase(_Formula(0))

        If _Formula.Length > 1 Then
            _Redondeo = Trim(_Formula(1))
        Else
            _Redondeo = 0
        End If

        Sb_Buscar_Valor_En_Dimensiones(_Empresa, _Fx1, _Codigo, _Koen)

        If String.IsNullOrEmpty(_Fx1.ToString.Trim) Then
            _Fx1 = 0
        End If

        Dim _Precio As Double

        Consulta_sql = "Select " & _Fx1 & " As Valor"
        Dim _RowPr As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        _Precio = _RowPr.Item("Valor")

        _Redondeo = Fx_Redondeo_Random(_Redondeo)

        _Precio = Fx_Redondear_Precio(_Precio, _Redondeo)

        Return _Precio

    End Function

    Function Fx_Redondeo_Random(_Redondeo As String) As Redondeo

        Select Case _Redondeo
            Case 1
                Return Redondeo.Redondear_2_decimales
            Case 2
                Return Redondeo.Redondear_1_decimales
            Case 3
                Return Redondeo.Redondear_0_decimales
            Case 4
                Return Redondeo.Redondear_con_multiplo_de_5
            Case 5
                Return Redondeo.Redondear_a_la_decena_mas_proxima
            Case 6
                Return Redondeo.Redondear_con_multiplo_de_5
            Case 7
                Return Redondeo.No_aplicar_redondeo
            Case 8
                Return Redondeo.Redondear_a_la_centena_mas_proxima
            Case 9
                Return Redondeo.Redondear_990
            Case "E"
                Return Redondeo.Redondear_al_entero_superior
            Case "F"
                Return Redondeo.Redondear_4_decimales
            Case "T"
                Return Redondeo.Redondear_3_decimales
            Case Else
                Return Redondeo.No_aplicar_redondeo
        End Select

    End Function

    Enum Redondeo
        No_aplicar_redondeo
        Redondear_0_decimales
        Redondear_1_decimales
        Redondear_2_decimales
        Redondear_3_decimales
        Redondear_4_decimales
        Redondear_5_decimales
        Redondear_a_la_decena_mas_proxima
        Redondear_a_la_centena_mas_proxima
        Redondear_con_multiplo_de_5
        Redondear_990
        Redondear_al_entero_superior
    End Enum

    Function Fx_Redondear_Precio(_Precio As Double, _Redondedo As Redondeo) As Double

        Dim _Valor As String = De_Num_a_Tx_01(_Precio, False, 5)

        Dim _Redondear, _UltNro, _Digito As Integer
        Dim _Red As Boolean

        Select Case _Redondedo

            Case Redondeo.No_aplicar_redondeo

                Return _Precio

            Case Redondeo.Redondear_0_decimales

                _Precio = Math.Round(_Precio, 0)

            Case Redondeo.Redondear_1_decimales

                '_Precio = Math.Round(_Precio, 1, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 1)

            Case Redondeo.Redondear_2_decimales

                '_Precio = Math.Round(_Precio, 2, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 2)

            Case Redondeo.Redondear_3_decimales

                '_Precio = Math.Round(_Precio, 3, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 3)

            Case Redondeo.Redondear_4_decimales

                '_Precio = Math.Round(_Precio, 4, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 4)

            Case Redondeo.Redondear_5_decimales

                '_Precio = Math.Round(_Precio, 5, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 5)


            Case Redondeo.Redondear_a_la_decena_mas_proxima

                _Precio = Math.Round(_Precio, 0)
                Dim _Decena = Split(_Precio, ".")
                Dim _Len = Len(_Decena(0))
                Dim _Ult_Dig = Mid(_Decena(0), _Len, 1)

                If _Ult_Dig = 0 And _Decena.Length = 1 Then
                    Return _Precio
                End If

                _Redondear = 10
                _UltNro = 1
                _Red = True

            Case Redondeo.Redondear_a_la_centena_mas_proxima

                _Redondear = 100
                _UltNro = 2
                _Red = True

            Case Redondeo.Redondear_con_multiplo_de_5

                _Redondear = 5
                _UltNro = 1
                _Red = True

            Case Redondeo.Redondear_990

                _Redondear = 990
                _UltNro = 3
                _Red = True

            Case Redondeo.Redondear_al_entero_superior

                _Valor = De_Num_a_Tx_01(_Precio, False, 1)
                _Digito = CInt(Right(_Valor, 1))
                _Precio = Math.Round(_Precio, 0)
                If _Digito > 0 And _Digito < 5 Then
                    _Precio += 1
                End If

            Case Else

        End Select

        If _Red Then

            _Valor = De_Num_a_Tx_01(_Precio, True, 0)

            _Digito = CInt(Right(_Valor, _UltNro))
            _Precio = (_Valor - _Digito) + _Redondear
        End If

        Return _Precio

    End Function

End Module

Public Module Modulo_Precios_Costos

    Dim _Sql As New Class_SQL()
    Dim Consulta_sql As String

    Enum _Campo_Precio
        Precio_Ud1
        Precio_Ud2
    End Enum

    'Function Fx_Precio_Formula_Random(_RowPrecio As DataRow,
    '                                  _Campo_Precio As String,
    '                                  _Campo_Ecuacion As String,
    '                                  _RowCostos_PM As DataRow,
    '                                  _Aplicar_Formula_Dinamica As Boolean,
    '                                  _ModEmpresa As String,
    '                                  _ModSucursal As String)


    '    If (_RowPrecio Is Nothing) Then
    '        Return 0
    '    End If

    '    Dim _Lista = _RowPrecio.Item("KOLT")
    '    Dim _Codigo = _RowPrecio.Item("KOPR")

    '    Dim _Rtu = De_Num_a_Tx_01(_RowPrecio.Item("RLUD"), False, 5)
    '    Dim _Precio As Double

    '    Dim _Formula

    '    Dim _Ecuacion As String = NuloPorNro(_RowPrecio.Item(_Campo_Ecuacion), "").ToString.Trim

    '    Dim _Ejecutar_Ecuacion = False

    '    If _Aplicar_Formula_Dinamica Then

    '        If Not String.IsNullOrEmpty(_Ecuacion) Then
    '            _Ejecutar_Ecuacion = (_Ecuacion = LCase(_Ecuacion))
    '        End If

    '    Else

    '        _Ejecutar_Ecuacion = True

    '    End If


    '    _Precio = _RowPrecio.Item(_Campo_Precio)
    '    Dim _New_Precio As Double

    '    If _Ejecutar_Ecuacion Then ' CBool(_Formula.Length) And Not String.IsNullOrEmpty(_Ecuacion) Then

    '        _Formula = Split(_Ecuacion, "#")

    '        Dim _Pp01ud = De_Num_a_Tx_01(_RowPrecio.Item("PP01UD"), False, 5)
    '        Dim _Pp02ud = De_Num_a_Tx_01(_RowPrecio.Item("PP02UD"), False, 5)

    '        Dim _Mg01ud = De_Num_a_Tx_01(_RowPrecio.Item("MG01UD"), False, 5)
    '        Dim _Mg02ud = De_Num_a_Tx_01(_RowPrecio.Item("MG02UD"), False, 5)

    '        Dim _Pm As String = 0
    '        Dim _Ppul01 As String = 0
    '        Dim _Ppul02 As String = 0
    '        Dim _Pmsuc As String = 0

    '        If (_RowCostos_PM Is Nothing) Then

    '            Consulta_sql = "Select Top 1 PM,PM AS PM01,PPUL01,PPUL02,Isnull(Round(PMSUC,5),0) As PMSUC
    '                        From MAEPREM EM
    '                        Left Join MAEPMSUC SUC On EM.EMPRESA = SUC.EMPRESA AND SUC.KOSU = '" & _ModSucursal & "' AND EM.KOPR = SUC.KOPR
    '                        Where EM.EMPRESA = '" & _ModEmpresa & "' And EM.KOPR = '" & _Codigo & "'"

    '            _RowCostos_PM = _Sql.Fx_Get_DataRow(Consulta_sql)

    '        End If

    '        If Not (_RowCostos_PM Is Nothing) Then

    '            _Pm = Math.Round(_RowCostos_PM.Item("PM01"), 5)
    '            _Ppul01 = Math.Round(_RowCostos_PM.Item("PPUL01"), 5)
    '            _Ppul02 = Math.Round(_RowCostos_PM.Item("PPUL02"), 5)
    '            _Pmsuc = Math.Round(_RowCostos_PM.Item("PMSUC"), 5)

    '        End If

    '        Dim _Fx1, _Fx2, _Redondeo


    '        _Fx1 = UCase(_Formula(0))

    '        If _Formula.Length > 1 Then
    '            _Redondeo = Trim(_Formula(1))
    '        Else
    '            _Redondeo = 0
    '        End If

    '        _Fx1 = Fx_Traer_Campo_Desde_Lista(_Codigo, _Fx1)

    '        _Fx1 = Replace(_Fx1, "RLUD", _Rtu)
    '        '_Fx1 = Replace(_Fx1, "PM01", _Pm)

    '        _Fx1 = Replace(_Fx1, "PMSUC", _Pmsuc)
    '        _Fx1 = Replace(_Fx1, "PM", _Pm)
    '        _Fx1 = Replace(_Fx1, "PPUL01", _Ppul01)
    '        _Fx1 = Replace(_Fx1, "PPUL02", _Ppul02)
    '        _Fx1 = Replace(_Fx1, "PP01UD", _Pp01ud)
    '        _Fx1 = Replace(_Fx1, "PP02UD", _Pp02ud)

    '        _Fx1 = Replace(_Fx1, "MG01UD", _Mg01ud)
    '        _Fx1 = Replace(_Fx1, "MG02UD", _Mg02ud)

    '        _Fx1 = Replace(_Fx1, ",", ".")

    '        _Fx1 = UCase(_Fx1)

    '        Consulta_sql = "Select * From PNOMDIM Where DEPENDENCI = 'Valor_propio'"
    '        Dim _Tbl_Dimension_Valor_Propio As DataTable = _Sql.Fx_Get_Tablas(Consulta_sql)

    '        For Each _FDim_Vp As DataRow In _Tbl_Dimension_Valor_Propio.Rows

    '            Dim _Dimension = _FDim_Vp.Item("CODIGO").ToString.Trim
    '            Dim _Valor = De_Num_a_Tx_01(_FDim_Vp.Item("VALOR"), False, 5)

    '            _Fx1 = Replace(_Fx1, _Dimension, _Valor)

    '        Next

    '        Consulta_sql = "Select * From PNOMDIM Where DEPENDENCI = 'Por_producto'"
    '        Dim _Tbl_Dimension_Por_Producto As DataTable = _Sql.Fx_Get_Tablas(Consulta_sql)

    '        For Each _FDim_Vp As DataRow In _Tbl_Dimension_Por_Producto.Rows

    '            Dim _Dimension = _FDim_Vp.Item("CODIGO").ToString.Trim
    '            Dim _Valor_Dim As Double = _Sql.Fx_Trae_Dato("PDIMEN", _Dimension, "EMPRESA = '" & _ModEmpresa & "' And CODIGO = '" & _Codigo & "'", True)
    '            Dim _Valor = De_Num_a_Tx_01(_Valor_Dim, False, 5)

    '            _Fx1 = Replace(_Fx1, _Dimension, _Valor)

    '        Next


    '        Consulta_sql = "Select " & _Fx1 & " As Valor"
    '        Dim _RowPr As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

    '        _Precio = _RowPr.Item("Valor")

    '        _Redondeo = Fx_Redondeo_Random(_Redondeo)
    '        _New_Precio = Fx_Redondear_Precio(_Precio, _Redondeo)

    '    Else
    '        _New_Precio = _Precio
    '    End If

    '    _New_Precio = Math.Round(_New_Precio, 2)

    '    Return _New_Precio

    'End Function

    Function Fx_Traer_Campo_Desde_Lista(_Codigo As String, _Ecuacion As String) As String

        Consulta_sql = "Select * From TABPP"
        Dim _Tbl_Listas As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Consulta_sql = "Select * From PNOMDIM Where CODIGO <> ''"
        Dim _Tbl_Dimensiones As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Consulta_sql = "Select COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME = 'TABPRE'"
        Dim _Tbl_Campos_Tabpre As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Dim _Campo As String

        For Each _FLista As DataRow In _Tbl_Listas.Rows

            Dim _Kolt As String = _FLista.Item("KOLT")
            _Campo = "<" & _Kolt & ">"

            For Each _FColumnas As DataRow In _Tbl_Campos_Tabpre.Rows

                Dim _Columna As String = _FColumnas.Item("COLUMN_NAME")
                Dim _Campo_Lista As String = _Campo & _Columna.Trim

                If _Ecuacion.Contains(_Campo_Lista) Then

                    Dim _Resultado = "(Select " & _Columna & " From TABPRE Where KOLT = '" & _Kolt & "' And KOPR = '" & _Codigo & "')"
                    _Ecuacion = Replace(_Ecuacion, _Campo_Lista, LCase(_Resultado))

                End If

            Next

        Next

        Return _Ecuacion

    End Function

    Function Fx_Redondeo_Random(_Redondeo As String) As Redondeo

        Select Case _Redondeo
            Case 1
                Return Redondeo.Redondear_2_decimales
            Case 2
                Return Redondeo.Redondear_1_decimales
            Case 3
                Return Redondeo.Redondear_0_decimales
            Case 4
                Return Redondeo.Redondear_con_multiplo_de_5
            Case 5
                Return Redondeo.Redondear_a_la_decena_mas_proxima
            Case 6
                Return Redondeo.Redondear_con_multiplo_de_5
            Case 7
                Return Redondeo.No_aplicar_redondeo
            Case 8
                Return Redondeo.Redondear_a_la_centena_mas_proxima
            Case 9
                Return Redondeo.Redondear_990
            Case "E"
                Return Redondeo.Redondear_al_entero_superior
            Case "F"
                Return Redondeo.Redondear_4_decimales
            Case "T"
                Return Redondeo.Redondear_3_decimales
            Case Else
                Return Redondeo.No_aplicar_redondeo
        End Select

    End Function

    Enum Redondeo
        No_aplicar_redondeo
        Redondear_0_decimales
        Redondear_1_decimales
        Redondear_2_decimales
        Redondear_3_decimales
        Redondear_4_decimales
        Redondear_5_decimales
        Redondear_a_la_decena_mas_proxima
        Redondear_a_la_centena_mas_proxima
        Redondear_con_multiplo_de_5
        Redondear_990
        Redondear_al_entero_superior
    End Enum

    Function Fx_Redondear_Precio(_Precio As Double, _Redondedo As Redondeo) As Double

        Dim _Valor As String = De_Num_a_Tx_01(_Precio, False, 5)

        Dim _Redondear, _UltNro, _Digito As Integer
        Dim _Red As Boolean

        Select Case _Redondedo

            Case Redondeo.No_aplicar_redondeo

                Return _Precio

            Case Redondeo.Redondear_0_decimales

                _Precio = Math.Round(_Precio, 0)

            Case Redondeo.Redondear_1_decimales

                '_Precio = Math.Round(_Precio, 1, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 1)

            Case Redondeo.Redondear_2_decimales

                '_Precio = Math.Round(_Precio, 2, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 2)

            Case Redondeo.Redondear_3_decimales

                '_Precio = Math.Round(_Precio, 3, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 3)

            Case Redondeo.Redondear_4_decimales

                '_Precio = Math.Round(_Precio, 4, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 4)

            Case Redondeo.Redondear_5_decimales

                '_Precio = Math.Round(_Precio, 5, MidpointRounding.ToEven)
                _Precio = De_Txt_a_Num_01(_Valor, 5)


            Case Redondeo.Redondear_a_la_decena_mas_proxima

                _Precio = Math.Round(_Precio, 0)
                Dim _Decena = Split(_Precio, ".")
                Dim _Len = Len(_Decena(0))
                Dim _Ult_Dig = Mid(_Decena(0), _Len, 1)

                If _Ult_Dig = 0 And _Decena.Length = 1 Then
                    Return _Precio
                End If

                _Redondear = 10
                _UltNro = 1
                _Red = True

            Case Redondeo.Redondear_a_la_centena_mas_proxima

                _Redondear = 100
                _UltNro = 2
                _Red = True

            Case Redondeo.Redondear_con_multiplo_de_5

                _Redondear = 5
                _UltNro = 1
                _Red = True

            Case Redondeo.Redondear_990

                _Redondear = 990
                _UltNro = 3
                _Red = True

            Case Redondeo.Redondear_al_entero_superior

                _Valor = De_Num_a_Tx_01(_Precio, False, 1)
                _Digito = CInt(Right(_Valor, 1))
                _Precio = Math.Round(_Precio, 0)
                If _Digito > 0 And _Digito < 5 Then
                    _Precio += 1
                End If

            Case Else

        End Select

        If _Red Then

            _Valor = De_Num_a_Tx_01(_Precio, True, 0)

            _Digito = CInt(Right(_Valor, _UltNro))
            _Precio = (_Valor - _Digito) + _Redondear
        End If

        Return _Precio

    End Function

    Function Fx_Formato_Numerico(_Valor As String,
                                 _Formato As String,
                                 _Es_Descuento As Boolean,
                                 Optional _Moneda_Str As String = "$") As String

        Dim _Cant_Caracteres As Integer = Len(_Formato)
        Dim _Moneda As Boolean

        Dim _Precio_Val As Double
        Dim _Es_Prcentaje As Boolean

        Dim _Decimales = 0

        If _Formato.Contains("%") Then
            _Es_Prcentaje = True
            _Cant_Caracteres -= 1
            _Formato = Replace(_Formato, "%", "")
        End If

        If _Formato.Contains(",") Then
            Dim _Dec = Split(_Formato, ",", 2)
            _Decimales = Len(_Dec(1))
        End If

        Dim _FORMAT As String

        If Not _Valor.Contains(",") Then

            _Decimales = 0
            Dim _Frmt = Split(_Formato, ",", 2)
            _Formato = _Frmt(0)

        End If

        Dim _Relleno As String = Mid(_Formato, 1, 1)

        If IsNumeric(_Relleno) Then
            _Relleno = " "
        Else
            _Formato = Replace(_Formato, _Relleno, "")
        End If

        If _Relleno = "$" Then

            _Moneda = True
            _Relleno = " "
            _Cant_Caracteres -= _Moneda_Str.Length '1

        End If

        _Precio_Val = De_Txt_a_Num_01(_Valor, _Decimales)
        _Precio_Val = Math.Round(_Precio_Val, _Decimales)

        If _Es_Descuento Then
            _Precio_Val = _Precio_Val * -1
        End If

        Dim _Precio = FormatNumber(_Precio_Val, _Decimales)

        ' Alinear a la derecha si el formato contiene 9
        If _Formato.Contains(9) Then 'Mid(_Formato, 1, 1) <> 8 Then
            _Precio = Rellenar(_Precio, _Cant_Caracteres, _Relleno, False)
        End If

        If _Moneda Then
            _Valor = _Moneda_Str & _Precio
        Else
            If _Es_Prcentaje Then
                _Valor = _Precio & "%"
            Else
                _Valor = _Precio
            End If
        End If

        Fx_Formato_Numerico = _Valor

    End Function

    Enum Enum_Tipo_Lista
        Compra
        Venta
    End Enum

End Module

Public Module Colores_Bakapp

    Enum Enum_Colores_Bakapp
        Rojo
        Celeste
        Naranjo
        Verde_Claro
        Amarillo
        Fuxia
        Azul_Petroleo
        Verde_Pistacho
        Crema
        Vino
        Azul_Oscuro
        Beig
        Verde
        Marron
        Gris
        Azul_Bakapp
        Gris_Oscuro
    End Enum

    Function Fx_Color_Bakapp(_Color_Bakapp As Enum_Colores_Bakapp) As Color

        Dim _Color As Color

        Select Case _Color_Bakapp
            Case Enum_Colores_Bakapp.Rojo
                _Color = ColorTranslator.FromHtml("#DC0000")'("#E2404C")
            Case Enum_Colores_Bakapp.Celeste
                _Color = ColorTranslator.FromHtml("#5BBDD9")
            Case Enum_Colores_Bakapp.Naranjo
                _Color = ColorTranslator.FromHtml("#E67D46")
            Case Enum_Colores_Bakapp.Verde_Claro
                _Color = ColorTranslator.FromHtml("#65C38B")
            Case Enum_Colores_Bakapp.Amarillo
                _Color = ColorTranslator.FromHtml("#F4B545")
            Case Enum_Colores_Bakapp.Fuxia
                _Color = ColorTranslator.FromHtml("#D15D7B")
            Case Enum_Colores_Bakapp.Azul_Petroleo
                _Color = ColorTranslator.FromHtml("#3D838C")
            Case Enum_Colores_Bakapp.Verde_Pistacho
                _Color = ColorTranslator.FromHtml("#A4B45D")
            Case Enum_Colores_Bakapp.Crema
                _Color = ColorTranslator.FromHtml("#F7BA8F")
            Case Enum_Colores_Bakapp.Vino
                _Color = ColorTranslator.FromHtml("#801812")
            Case Enum_Colores_Bakapp.Azul_Oscuro
                _Color = ColorTranslator.FromHtml("#263068")
            Case Enum_Colores_Bakapp.Beig
                _Color = ColorTranslator.FromHtml("#CCB59A")
            Case Enum_Colores_Bakapp.Verde
                _Color = ColorTranslator.FromHtml("#4B7F51")
            Case Enum_Colores_Bakapp.Marron
                _Color = ColorTranslator.FromHtml("#6E5746")
            Case Enum_Colores_Bakapp.Gris
                _Color = ColorTranslator.FromHtml("#9C9B9B")
            Case Enum_Colores_Bakapp.Azul_Bakapp
                _Color = ColorTranslator.FromHtml("#349FCE")
            Case Enum_Colores_Bakapp.Gris_Oscuro
                _Color = ColorTranslator.FromHtml("#777E91")
        End Select

        Return _Color

    End Function

    Function Fx_Revisar_Expiracion_Folio_SII(_Empresa As String, _Tido As String, _Folio As String) As Boolean

        Dim _Sql As New Class_SQL()

        If _Tido = "GDP" Or _Tido = "GDD" Or _Tido = "GTI" Then
            _Tido = "GDV"
        End If

        Dim _Td = Fx_Tipo_DTE_VS_TIDO(_Tido)

        Consulta_sql = "Select Top 1 * From FFOLIOS With ( NOLOCK )" & vbCrLf &
                       "Where Cast(RNG_D AS INT)<=" & Val(_Folio) & "  And Cast(RNG_H AS INT)>=" & Val(_Folio) &
                       " And TD='" & _Td & "' And EMPRESA='" & _Empresa & "' "

        Dim _Row_Folios As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If IsNothing(_Row_Folios) Then

            'If Not IsNothing(_Formulario) Then

            '    'MessageBoxEx.Show(_Formulario, "el folio del documento electrónico no está  autorizado por el SII: " & _Folio & vbCrLf & vbCrLf &
            '    '                  "INFORME ESTA SITUACION AL ADMINISTRADOR DEL SISTEMA POR FAVOR", "Validación Modalidad: " & Modalidad,
            '    '                  MessageBoxButtons.OK, MessageBoxIcon.Stop)

            'End If

        Else

            'Dim _Hasta = _Row_Folios.Item("RNG_H")
            'Dim _Folios_Restantes = _Hasta - CInt(_Folio)

            Dim _Fa As DateTime = FormatDateTime(CDate(_Row_Folios.Item("FA")), DateFormat.ShortDate)
            Dim _Fecha_Servisor As DateTime = FormatDateTime(FechaDelServidor(), DateFormat.ShortDate)

            Dim _Meses As Integer = 6

            If _Sql.Fx_Existe_Tabla("FDTECONF") Then

                Try
                    _Meses = _Sql.Fx_Trae_Dato("FDTECONF", "VALOR", "CAMPO = 'sii.meses.expiran.folios' And ACTIVO=1 And EMPRESA = '" & _Empresa & "'")
                Catch ex As Exception
                    If _Tido = "BLV" Then
                        _Meses = 24
                    ElseIf _Tido = "GDV" Then
                        _Meses = 12
                    End If
                End Try

            End If

            Dim _Meses_Dif As Double = DateDiff(DateInterval.Month, _Fa, _Fecha_Servisor)
            Dim _Dias_Dif As Integer = DateDiff(DateInterval.Day, _Fa, _Fecha_Servisor)

            _Meses_Dif = Math.Round(_Dias_Dif / 31, 2)

            If _Meses_Dif > _Meses Then

                'If Not IsNothing(_Formulario) Then

                '    MessageBoxEx.Show(_Formulario, "Este folio " & _Folio & " tiene mas de (" & _Meses & ") meses desde su fecha de creación" & vbCrLf &
                '              "en el SII y su configuración indica que podría estar vencido." & vbCrLf &
                '              "Si usted insite en el envío, este documento podria ser rechazado." & vbCrLf & vbCrLf &
                '              "INFORME ESTA SITUACION AL ADMINISTRADOR DEL SISTEMA POR FAVOR", "Validación Modalidad: " & Modalidad, MessageBoxButtons.OK, MessageBoxIcon.Stop)

                'End If

            Else

                Return True

            End If

        End If

    End Function

    Function Fx_Tipo_DTE_VS_TIDO(_Tido As String) As Integer

        Select Case _Tido
            Case "FCV"
                Return 33
            Case "BLV", "BSV"
                Return 39
            Case "GDV", "GDP", "GTI", "GDD"
                Return 52
            Case "NCV"
                Return 61
            Case "OCC"
                Return 801
            Case Else
                Return 0
        End Select

        'Return "FACTURA" 33
        'Return "FACTURA EXENTA" 34
        'Return "GUIA DE DESPACHO" 52
        'Return "FACTURA DE COMPRA" 46
        'Return "NOTA DE DEBITO" 56
        'Return "NOTA DE CREDITO" 61
        'Return "ORDEN DE COMPRA" 801

    End Function

    Function Fx_Caracter_Raro_Quitar(ByRef _Texto As String)

        _Texto = Replace(_Texto, "&", "&amp;")
        _Texto = Replace(_Texto, "<", "&lt;")
        _Texto = Replace(_Texto, ">", "&gt;")
        _Texto = Replace(_Texto, "'", "&apos;")
        _Texto = Replace(_Texto, """", "&quot;")
        _Texto = Replace(_Texto, "´", "")
        _Texto = Replace(_Texto, "°", "")
        _Texto = Replace(_Texto, "º", "")
        _Texto = Replace(_Texto, "ñ", "n")
        _Texto = Replace(_Texto, "Ñ", "N")

        _Texto = Replace(_Texto, "á", "a")
        _Texto = Replace(_Texto, "é", "e")
        _Texto = Replace(_Texto, "í", "i")
        _Texto = Replace(_Texto, "ó", "o")
        _Texto = Replace(_Texto, "ú", "u")

        _Texto = Replace(_Texto, "Á", "A")
        _Texto = Replace(_Texto, "É", "E")
        _Texto = Replace(_Texto, "Í", "I")
        _Texto = Replace(_Texto, "Ó", "O")
        _Texto = Replace(_Texto, "Ú", "U")

        _Texto = Replace(_Texto, "ü", "u")
        _Texto = Replace(_Texto, "Ü", "U")

        _Texto = Replace(_Texto, vbCrLf, "")
        _Texto = Replace(_Texto, " ", "")
        _Texto = Replace(_Texto, "ª", "")

        If Not String.IsNullOrEmpty(_Texto) Then
            For i = 1 To _Texto.Length
                Dim Letra As String = Mid(_Texto, i, 1)
                Dim codeInt = Asc(Letra)
                If (codeInt >= 0 And codeInt <= 31) Or (codeInt >= 127 And codeInt <= 255) Then
                    _Texto = Replace(_Texto, Letra, " ")
                End If
            Next
        End If

        If IsNothing(_Texto) Then
            _Texto = String.Empty
        End If

        _Texto = _Texto.Trim

    End Function

End Module

Namespace LsValiciones

    Public Class Mensajes

        Public Property EsCorrecto As Boolean
        Public Property Id As String
        Public Property Fecha As DateTime
        Public Property Detalle As String = String.Empty
        Public Property Mensaje As String = String.Empty
        Public Property Resultado As String = String.Empty
        Public Property Tag As Object
        Public Property UsarImagen As Boolean
        Public Property NombreImagen As String = String.Empty
        Public Property Icono As Object
        Public Property Cancelado As Boolean
        Public Property MostrarMensaje As Boolean = True
        Public Property Cerrar As Boolean
        Public Property ErrorDeConexionSQL As Boolean

    End Class

    Public Class Columnas

        Public Property Nombre As String
        Public Property Descripcion As String
        Public Property Ancho As Integer

    End Class


End Namespace


