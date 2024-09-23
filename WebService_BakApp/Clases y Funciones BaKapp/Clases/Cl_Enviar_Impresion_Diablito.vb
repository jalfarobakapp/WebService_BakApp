
Public Class Cl_Enviar_Impresion_Diablito

    Dim _Sql As New Class_SQL
    Dim Consulta_sql As String

    Dim _Tbl_Conf_Impresion_Normal As DataTable
    Dim _Tbl_Conf_Impresion_Vale_Transitorio As DataTable
    Dim _Tbl_Conf_Impresion_Picking As DataTable

    Dim _Empresa As String
    Dim _CodFuncionario As String

    Dim _Error As String
    Dim _SoloEnviarDocDeSucursalDelDiablito As Boolean

    Enum Enum_Tipo
        Normal
        Vale_Transitorio
        Picking
    End Enum

    Public Property [Error] As String
        Get
            Return _Error
        End Get
        Set(value As String)
            _Error = value
        End Set
    End Property

    Public Property SoloEnviarDocDeSucursalDelDiablito As Boolean
        Get
            Return _SoloEnviarDocDeSucursalDelDiablito
        End Get
        Set(value As Boolean)
            _SoloEnviarDocDeSucursalDelDiablito = value
        End Set
    End Property

    Public Sub New(_Empresa As String, _CodFuncionario As String)
        Me._Empresa = _Empresa
        Me._CodFuncionario = _CodFuncionario
    End Sub

    Function Fx_Trae_Tbl_Configuracion_Estaciones_Impresion(_Modalidad As String, _Tido As String, _Tipo As Enum_Tipo) As DataTable

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Usuarios_Impresion 
                        Where CodFuncionario = '" & _CodFuncionario & "' And Empresa = '" & _Empresa & "' And " &
                       "Modalidad = '" & _Modalidad & "' And Tido = '" & _Tido & "' And Tipo = '" & _Tipo.ToString & "' And Activo = 1"

        Dim _Tbl As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Return _Tbl

    End Function

    Function Fx_Trae_Tbl_Configuracion_Estaciones_Impresion_Todas_Modalidades(_Tido As String, _Tipo As Enum_Tipo) As DataTable

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Usuarios_Impresion 
                        Where CodFuncionario = '" & _CodFuncionario & "' And Empresa = '" & _Empresa & "' And " &
                       "Imp_Todas_Modalidades = 1 And Tido = '" & _Tido & "' And Tipo = '" & _Tipo.ToString & "' And Activo = 1 And Modalidad = ''"

        Dim _Tbl As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        Return _Tbl

    End Function

    Function Fx_Enviar_Impresion_Al_Diablito(_Modalidad As String,
                                             _Idmaeedo As Integer) As Boolean

        _Error = String.Empty

        Try

            Consulta_sql = "Select * From MAEEDO Where IDMAEEDO = " & _Idmaeedo
            Dim _Row_Documento As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            If IsNothing(_Row_Documento) Then
                _Error = "No se encontro el documento con el indice: " & _Idmaeedo
                Return False
            End If

            Dim _Tido = _Row_Documento.Item("TIDO")
            Dim _Nudo = _Row_Documento.Item("NUDO")
            Dim _Kofudo = _Row_Documento.Item("KOFUDO")
            Dim _Nudonodefi = _Row_Documento.Item("NUDONODEFI")

            _Tbl_Conf_Impresion_Vale_Transitorio = Fx_Trae_Tbl_Configuracion_Estaciones_Impresion_Todas_Modalidades(_Tido, Enum_Tipo.Vale_Transitorio)
            If Not CBool(_Tbl_Conf_Impresion_Vale_Transitorio.Rows.Count) Then
                _Tbl_Conf_Impresion_Vale_Transitorio = Fx_Trae_Tbl_Configuracion_Estaciones_Impresion(_Modalidad, _Tido, Enum_Tipo.Vale_Transitorio)
            End If

            _Tbl_Conf_Impresion_Normal = Fx_Trae_Tbl_Configuracion_Estaciones_Impresion_Todas_Modalidades(_Tido, Enum_Tipo.Normal)
            If Not CBool(_Tbl_Conf_Impresion_Normal.Rows.Count) Then
                _Tbl_Conf_Impresion_Normal = Fx_Trae_Tbl_Configuracion_Estaciones_Impresion(_Modalidad, _Tido, Enum_Tipo.Normal)
            End If

            _Tbl_Conf_Impresion_Picking = Fx_Trae_Tbl_Configuracion_Estaciones_Impresion_Todas_Modalidades(_Tido, Enum_Tipo.Picking)
            If Not CBool(_Tbl_Conf_Impresion_Picking.Rows.Count) Then
                _Tbl_Conf_Impresion_Picking = Fx_Trae_Tbl_Configuracion_Estaciones_Impresion(_Modalidad, _Tido, Enum_Tipo.Picking)
            End If

            'Dim _NombreFormato As String
            'Dim _NombreEquipo_Imprime As String
            'Dim _Nro_Copias As Integer
            'Dim _Impresora As String

            Consulta_sql = String.Empty

            If _Nudonodefi Then

                If Not CBool(_Tbl_Conf_Impresion_Vale_Transitorio.Rows.Count) Then
                    _Error += " - No existe configuración para la salida de impresión de vales transitorios para este tipo de documento en el diablito de impresiones automáticas" & vbCrLf
                End If

                For Each _Row As DataRow In _Tbl_Conf_Impresion_Vale_Transitorio.Rows

                    Dim _NombreEquipo_Imprime = _Row.Item("NombreEquipo_Imprime")

                    Dim _Query = "Select ESUCURSAL,EBODEGA 
                                    From CONFIEST
                                    Inner Join " & _Global_BaseBk & "Zw_EstacionesBkp On Modalidad_X_Defecto = MODALIDAD
                                    Where NombreEquipo = '" & _NombreEquipo_Imprime & "'"
                    Dim _Row_Eim As DataRow = _Sql.Fx_Get_DataRow(_Query)

                    If IsNothing(_Row_Eim) Then
                        Consulta_sql = "Falta la modalidad en la Estacion del Equipo: " & _NombreEquipo_Imprime & "..."
                    Else

                        Dim _Reg As Boolean

                        If _SoloEnviarDocDeSucursalDelDiablito Then
                            _Reg = _Sql.Fx_Cuenta_Registros("MAEEDO", "IDMAEEDO = " & _Idmaeedo & " And SUDO = '" & _Row_Eim.Item("ESUCURSAL") & "'")
                        Else
                            _Reg = _Sql.Fx_Cuenta_Registros("MAEEDO", "IDMAEEDO = " & _Idmaeedo)
                        End If

                        If _Reg Then
                            Consulta_sql += Fx_Inyectar_Formato(_Idmaeedo, _Tido, _Nudo, _Kofudo, _Row, True, False)
                        End If

                    End If

                Next

            Else

                If Not CBool(_Tbl_Conf_Impresion_Normal.Rows.Count) Then
                    _Error += " - No existe configuración para la salida de impresión normal para este tipo de documento en el diablito de impresiones automáticas" & vbCrLf
                End If

                For Each _Row As DataRow In _Tbl_Conf_Impresion_Normal.Rows

                    Dim _NombreEquipo_Imprime = _Row.Item("NombreEquipo_Imprime")

                    Dim _Query = "Select ESUCURSAL,EBODEGA 
                                    From CONFIEST
                                    Inner Join " & _Global_BaseBk & "Zw_EstacionesBkp On Modalidad_X_Defecto = MODALIDAD
                                    Where NombreEquipo = '" & _NombreEquipo_Imprime & "'"
                    Dim _Row_Eim As DataRow = _Sql.Fx_Get_DataRow(_Query)

                    If IsNothing(_Row_Eim) Then
                        Consulta_sql = "Falta la modalidad en la Estacion del Equipo: " & _NombreEquipo_Imprime & "..."
                    Else

                        Dim _Reg As Boolean

                        If _SoloEnviarDocDeSucursalDelDiablito Then
                            _Reg = _Sql.Fx_Cuenta_Registros("MAEEDO", "IDMAEEDO = " & _Idmaeedo & " And SUDO = '" & _Row_Eim.Item("ESUCURSAL") & "'")
                        Else
                            _Reg = _Sql.Fx_Cuenta_Registros("MAEEDO", "IDMAEEDO = " & _Idmaeedo)
                        End If

                        If _Reg Then
                            Consulta_sql += Fx_Inyectar_Formato(_Idmaeedo, _Tido, _Nudo, _Kofudo, _Row, False, False)
                        End If

                    End If

                Next

                If Not CBool(_Tbl_Conf_Impresion_Normal.Rows.Count) Then
                    _Error += " - No existe configuración para la salida de impresión de picking para este tipo de documento en el diablito de impresiones automáticas" & vbCrLf
                End If

                For Each _Row As DataRow In _Tbl_Conf_Impresion_Picking.Rows

                    Dim _Empresa = _Row.Item("Empresa")
                    Dim _Sucursal = _Row.Item("Sucursal_Picking")
                    Dim _Bodega = _Row.Item("Bodega_Picking")

                    Dim _Enviar As Boolean = CBool(_Sql.Fx_Cuenta_Registros("MAEDDO",
                                                                            "IDMAEEDO = " & _Idmaeedo &
                                                                            " And EMPRESA = '" & _Empresa & "'" &
                                                                            " And SULIDO = '" & _Sucursal & "'" &
                                                                            " And BOSULIDO = '" & _Bodega & "'"))
                    If _Enviar Then
                        Consulta_sql += Fx_Inyectar_Formato(_Idmaeedo, _Tido, _Nudo, _Kofudo, _Row, False, True)
                    End If

                Next

            End If

            If Not String.IsNullOrEmpty(_Error) Then
                _Error = "Tido: " & _Tido & vbCrLf & _Error
            End If

            If String.IsNullOrEmpty(Consulta_sql) Then
                Return False
            End If

            If Not _Sql.Fx_Ej_consulta_IDU(Consulta_sql, False) Then
                _Error += _Sql.Pro_Error
                Return False
            End If

            Return True

        Catch ex As Exception
            _Error = ex.Message
        End Try

    End Function

    Private Function Fx_Inyectar_Formato(_Idmaeedo As Integer,
                                         _Tido As String,
                                         _Nudo As String,
                                         _Kofudo As String,
                                         _Row_Conf As DataRow,
                                         _Nudonodefi As Boolean,
                                         _Picking As Boolean)

        Dim _NombreFormato As String
        Dim _NombreEquipo_Imprime As String
        Dim _Nro_Copias As Integer
        Dim _Impresora As String
        Dim _Imprimir_Voucher_TJV As Integer
        Dim _Imprimir_Voucher_TJV_Original As Integer
        Dim _Vale_Transitorio As Integer = Convert.ToInt32(_Nudonodefi)

        _NombreFormato = _Row_Conf.Item("NombreFormato")
        _NombreEquipo_Imprime = _Row_Conf.Item("NombreEquipo_Imprime")
        _Nro_Copias = _Row_Conf.Item("Nro_Copias")
        _Impresora = _Row_Conf.Item("Impresora")
        _Imprimir_Voucher_TJV = _Row_Conf.Item("Imprimir_Voucher_TJV")
        _Imprimir_Voucher_TJV_Original = _Row_Conf.Item("Imprimir_Voucher_TJV_Original")

        Dim _Consulta_sql = String.Empty

        _Consulta_sql += "-- INSERTANDO " & _Tido &
                                         vbCrLf &
                                         vbCrLf &
                                         "Insert Into " & _Global_BaseBk & "Zw_Demonio_Doc_Emitidos_Cola_Impresion" & Space(1) &
                                         "(NombreEquipo,Idmaeedo,Tido,Nudo,Funcionario,Fecha,NombreFormato,Nudonodefi,Picking,Nro_Copias_Impresion," &
                                         "Impresora,Imprimir_Voucher_TJV,Imprimir_Voucher_TJV_Original,Vale_Transitorio)" & vbCrLf &
                                         "Select '" & _NombreEquipo_Imprime & "'," & _Idmaeedo & ",'" & _Tido & "','" & _Nudo & "','" & _Kofudo & "',GetDate()," &
                                         "'" & _NombreFormato & "'," & Convert.ToInt32(_Nudonodefi) & "," & Convert.ToInt32(_Picking) &
                                         "," & _Nro_Copias & ",'" & _Impresora & "'," & Convert.ToInt32(_Imprimir_Voucher_TJV) &
                                         "," & Convert.ToInt32(_Imprimir_Voucher_TJV_Original) & "," & _Vale_Transitorio & vbCrLf & vbCrLf
        Return _Consulta_sql

    End Function

End Class
