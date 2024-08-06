Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Windows.Forms

Public Class Cl_Conteo

    Dim _Sql As New Class_SQL
    Dim Consulta_sql As String

    Public Property Zw_Inv_Hoja As New Zw_Inv_Hoja
    Public Property Ls_Zw_Inv_Hoja_Detalle As New List(Of Zw_Inv_Hoja_Detalle)

    Public Sub New()

    End Sub

    Function Fx_Llenar_Zw_Inv_Hoja(_Id As Integer) As LsValiciones.Mensajes

        Dim _Mensaje_Stem As New LsValiciones.Mensajes

        _Mensaje_Stem.EsCorrecto = False
        _Mensaje_Stem.Detalle = "Cargar Hoja"
        _Mensaje_Stem.Mensaje = String.Empty

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Inv_Hoja Where Id = " & _Id
        Dim _Row As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        If IsNothing(_Row) Then
            _Mensaje_Stem.Mensaje = "No se encontro el registro en la tabla Zw_Inv_Hoja con el Id " & _Id
            Return _Mensaje_Stem
        End If

        With Zw_Inv_Hoja

            .Id = _Row.Item("Id")
            .IdInventario = _Row.Item("IdInventario")
            .Nro_Hoja = _Row.Item("Nro_Hoja")
            .NombreEquipo = _Row.Item("NombreEquipo")
            .FechaCreacion = _Row.Item("FechaCreacion")
            .CodResponsable = _Row.Item("CodResponsable")
            .IdContador1 = _Row.Item("IdContador1")
            .IdContador2 = _Row.Item("IdContador2")
            .FechaLevantamiento = _Row.Item("FechaLevantamiento")
            .Reconteo = _Row.Item("Reconteo")

        End With

        Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Inv_Hoja_Detalle Where IdHoja = " & _Id
        Dim _Tbl As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

        For Each _Fila As DataRow In _Tbl.Rows

            Dim _Zw_Inv_Hoja_Detalle As New Zw_Inv_Hoja_Detalle

            With _Zw_Inv_Hoja_Detalle

                .Id = _Row.Item("Id")
                .IdHoja = _Row.Item("IdHoja")
                .Nro_Hoja = _Row.Item("Nro_Hoja")
                .IdInventario = _Row.Item("IdInventario")
                .Empresa = _Row.Item("Empresa")
                .Sucursal = _Row.Item("Sucursal")
                .Bodega = _Row.Item("Bodega")
                .Responsable = _Row.Item("Responsable")
                .IdContador1 = _Row.Item("IdContador1")
                .IdContador2 = _Row.Item("IdContador2")
                .Item_Hoja = _Row.Item("Item_Hoja")
                .IdSector = _Row.Item("IdSector")
                .Sector = _Row.Item("Sector")
                .Ubicacion = _Row.Item("CodUbicacion")
                .TipoConteo = _Row.Item("TipoConteo")
                .Codigo = _Row.Item("Codproducto")
                .EsSeriado = _Row.Item("EsSeriado")
                .NroSerie = _Row.Item("NroSerie")
                .FechaHoraToma = _Row.Item("FechaHoraToma")
                .Rtu = _Row.Item("Rtu")
                .RtuVariable = _Row.Item("RtuVariable")
                .Udtrpr = _Row.Item("Udtrpr")
                .Ud1 = _Row.Item("Ud1")
                .CantidadUd1 = _Row.Item("CantidadUd1")
                .Ud2 = _Row.Item("Ud2")
                .CantidadUd2 = _Row.Item("CantidadUd2")
                .Observaciones = _Row.Item("Observaciones")
                .Recontado = _Row.Item("Recontado")
                .Actualizado_por = _Row.Item("Actualizado_por")
                .Obs_Actualizacion = _Row.Item("Obs_Actualizacion")

                Ls_Zw_Inv_Hoja_Detalle.Add(_Zw_Inv_Hoja_Detalle)

            End With

        Next

        _Mensaje_Stem.EsCorrecto = True
        _Mensaje_Stem.Mensaje = "Registros cargados correctamente"

    End Function

    Function Fx_Nueva_Hoja(Zw_Inv_Inventario As Zw_Inv_Inventario,
                           _NombreEquipo As String,
                           _CodResponsable As String) As LsValiciones.Mensajes

        Dim _Mensaje_Stem As New LsValiciones.Mensajes

        _Mensaje_Stem.EsCorrecto = False
        _Mensaje_Stem.Detalle = "Nueva Hoja"
        _Mensaje_Stem.Mensaje = String.Empty

        With Zw_Inv_Hoja

            .Id = 0
            .IdInventario = Zw_Inv_Inventario.Id
            .Nro_Hoja = String.Empty
            .NombreEquipo = _NombreEquipo
            .CodResponsable = _CodResponsable
            .IdContador1 = 0
            .IdContador2 = 0

        End With

        _Mensaje_Stem.EsCorrecto = True
        _Mensaje_Stem.Mensaje = "Registros cargados correctamente"

    End Function

    Function Fx_Crear_Hoja(_Zw_Inv_Hoja As Zw_Inv_Hoja, Ls_Zw_Inv_Hoja_Detalle As List(Of Zw_Inv_Hoja_Detalle)) As LsValiciones.Mensajes

        Dim _Mensaje As New LsValiciones.Mensajes

        _Mensaje.Detalle = "Crear Hoja de Inventario"
        _Mensaje.EsCorrecto = False
        _Mensaje.Mensaje = String.Empty
        _Mensaje.Icono = MessageBoxIcon.Stop

        Consulta_sql = String.Empty

        Dim myTrans As SqlClient.SqlTransaction
        Dim Comando As SqlClient.SqlCommand

        Dim Cn2 As New SqlConnection
        Dim SQL_ServerClass As New Class_SQL '(Cadena_ConexionSQL_Server)

        SQL_ServerClass.Sb_Abrir_Conexion(Cn2)

        Try

            myTrans = Cn2.BeginTransaction()

            With _Zw_Inv_Hoja

                Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Inv_Hoja (IdInventario,Nro_Hoja,NombreEquipo,FechaCreacion,CodResponsable,IdContador1,IdContador2,FechaLevantamiento,Reconteo) Values " &
                               "(" & .IdInventario & ",'" & .Nro_Hoja & "','" & .NombreEquipo & "','" & .FechaCreacion & "','" & .CodResponsable & "'," & .IdContador1 & "," & .IdContador2 & ",'" & .FechaLevantamiento & "'," & Convert.ToInt32(.Reconteo) & ")"

                Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
                Comando.Transaction = myTrans
                Comando.ExecuteNonQuery()

                Dim dfd1 As System.Data.SqlClient.SqlDataReader = Comando.ExecuteReader()
                While dfd1.Read()
                    .Id = dfd1("Identity")
                End While
                dfd1.Close()

            End With

            For Each _Zw_Inv_Hoja_Detalle As Zw_Inv_Hoja_Detalle In Ls_Zw_Inv_Hoja_Detalle

                With _Zw_Inv_Hoja_Detalle

                    Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Inv_Hoja_Detalle (IdHoja,Nro_Hoja,IdInventario,Empresa,Sucursal,Bodega,Responsable,IdContador1,IdContador2,Item_Hoja,IdSector,Sector,Ubicacion,TipoConteo,Codproducto,EsSeriado,NroSerie,FechaHoraToma,Rtu,RtuVariable,Udtrpr,Ud1,CantidadUd1,Ud2,CantidadUd2,Observaciones,Recontado,Actualizado_por,Obs_Actualizacion) Values " &
                                   "(" & .IdHoja & ",'" & .Nro_Hoja & "'," & .IdInventario & ",'" & .Empresa & "','" & .Sucursal & "','" & .Bodega & "','" & .Responsable & "'," & .IdContador1 & "," & .IdContador2 & ",'" & .Item_Hoja & "'," & .IdSector & ",'" & .Ubicacion & "','" & .TipoConteo & "','" & .Codigo & "'," & Convert.ToInt32(.EsSeriado) & ",'" & .NroSerie & "','" & .FechaHoraToma & "','" & .Rtu & "','" & .RtuVariable & "','" & .Udtrpr & "','" & .Ud1 & "'," & .CantidadUd1 & ",'" & .Ud2 & "'," & .CantidadUd2 & ",'" & .Observaciones & "'," & Convert.ToInt32(.Recontado) & ",'" & .Actualizado_por & "','" & .Obs_Actualizacion & "')"

                    Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
                    Comando.Transaction = myTrans
                    Comando.ExecuteNonQuery()

                End With

            Next

            myTrans.Commit()
            SQL_ServerClass.Sb_Cerrar_Conexion(Cn2)

            Consulta_sql = "Select * From " & _Global_BaseBk & "Zw_Inv_Hoja Where Id = " & _Zw_Inv_Hoja.Id
            Dim _Row As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            _Mensaje.EsCorrecto = True
            _Mensaje.Mensaje = "Hoja creada correctamente"
            _Mensaje.Icono = MessageBoxIcon.Information
            _Mensaje.Detalle = "Hoja creada correctamente con el Id " & _Row.Item("Id")
            _Mensaje.Tag = _Row
            _Mensaje.Id = _Row.Item("Id")

        Catch ex As Exception

            _Mensaje.Mensaje = ex.Message
            If Not IsNothing(myTrans) Then myTrans.Rollback()

            SQL_ServerClass.Sb_Cerrar_Conexion(Cn2)

        End Try

        Return _Mensaje

    End Function

    Function Fx_Grabar_Nueva_Hoja() As LsValiciones.Mensajes

        Dim _Mensaje As New LsValiciones.Mensajes

        Consulta_sql = String.Empty

        Dim _Reg = _Sql.Fx_Cuenta_Registros(_Global_BaseBk & "Zw_Inv_Hoja",
                                            "IdInventario = " & Zw_Inv_Hoja.IdInventario & " And Nro_Hoja = '" & Zw_Inv_Hoja.Nro_Hoja & "'")

        If CBool(_Reg) Then
            _Mensaje.EsCorrecto = False
            _Mensaje.Detalle = "Crear Hoja"
            _Mensaje.Mensaje = "El Nro de Hoja " & Zw_Inv_Hoja.Nro_Hoja & " ya existe para este inventario"
            _Mensaje.Icono = MessageBoxIcon.Stop
            Return _Mensaje
        End If

        Dim myTrans As SqlClient.SqlTransaction
        Dim Comando As SqlClient.SqlCommand

        Dim Cn2 As New SqlConnection
        Dim SQL_ServerClass As New Class_SQL '(Cadena_ConexionSQL_Server)

        SQL_ServerClass.Sb_Abrir_Conexion(Cn2)

        Try

            myTrans = Cn2.BeginTransaction()

            With Zw_Inv_Hoja

                .Nro_Hoja = Fx_NvoNroHoja(.IdInventario)

                Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Inv_Hoja (IdInventario,Nro_Hoja,NombreEquipo,FechaCreacion,CodResponsable," &
                               "IdContador1,IdContador2,Reconteo) Values " &
                               "(" & .IdInventario & ",'" & .Nro_Hoja & "','" & .NombreEquipo & "','" & .FechaCreacion & "','" & .CodResponsable & "'" &
                               "," & .IdContador1 & "," & .IdContador2 & "," & Convert.ToInt32(.Reconteo) & ")"

                Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
                Comando.Transaction = myTrans
                Comando.ExecuteNonQuery()

                Comando = New System.Data.SqlClient.SqlCommand("SELECT @@IDENTITY AS 'Identity'", Cn2)
                Comando.Transaction = myTrans
                Dim dfd1 As System.Data.SqlClient.SqlDataReader = Comando.ExecuteReader()
                While dfd1.Read()
                    .Id = dfd1("Identity")
                End While
                dfd1.Close()

            End With

            For Each _Fila As Zw_Inv_Hoja_Detalle In Ls_Zw_Inv_Hoja_Detalle

                With _Fila

                    .IdHoja = Zw_Inv_Hoja.Id
                    .Nro_Hoja = Zw_Inv_Hoja.Nro_Hoja

                    If Not String.IsNullOrEmpty(.Codigo) Then

                        Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Inv_Hoja_Detalle (IdHoja,Nro_Hoja,IdInventario,Empresa,Sucursal,Bodega," &
                                       "Responsable,IdContador1,IdContador2,Item_Hoja,IdSector,Sector,Ubicacion,TipoConteo,Codigo,EsSeriado," &
                                       "NroSerie,FechaHoraToma,Rtu,RtuVariable,Udtrpr,Cantidad,Ud1,CantidadUd1,Ud2,CantidadUd2,Observaciones," &
                                       "Recontado,Actualizado_por,Obs_Actualizacion) Values " &
                                       "(" & .IdHoja & ",'" & .Nro_Hoja & "'," & .IdInventario & ",'" & .Empresa & "','" & .Sucursal & "','" & .Bodega & "'" &
                                       ",'" & .Responsable & "'," & .IdContador1 & "," & .IdContador2 & ",'" & .Item_Hoja & "'," & .IdSector &
                                       ",'" & .Sector & "','" & .Ubicacion & "','" & .TipoConteo & "','" & .Codigo & "'," & Convert.ToInt32(.EsSeriado) & ",'" & .NroSerie & "'" &
                                       ",Getdate(),'" & .Rtu & "','" & .RtuVariable & "','" & .Udtrpr & "'" &
                                       "," & De_Num_a_Tx_01(.Cantidad, False, 5) & ",'" & .Ud1 &
                                       "'," & De_Num_a_Tx_01(.CantidadUd1, False, 5) & ",'" & .Ud2 & "'," & De_Num_a_Tx_01(.CantidadUd2, False, 5) &
                                       ",'" & .Observaciones & "'," & Convert.ToInt32(.Recontado) & ",'" & .Actualizado_por & "','" & .Obs_Actualizacion & "')"

                        Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
                        Comando.Transaction = myTrans
                        Comando.ExecuteNonQuery()

                        Consulta_sql = "Update " & _Global_BaseBk & "Zw_Inv_FotoInventario Set Cant_Inventariada +=" & .Cantidad & vbCrLf &
                                       "Where IdInventario = " & .IdInventario & " And Codigo = '" & .Codigo & "'"

                        Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
                        Comando.Transaction = myTrans
                        Comando.ExecuteNonQuery()

                        Consulta_sql = "Update " & _Global_BaseBk & "Zw_Inv_FotoInventario Set Total_Costo_Inv = Cant_Inventariada*Costo" & vbCrLf &
                                       "Where IdInventario = " & .IdInventario & " And Codigo = '" & .Codigo & "'"

                        If Zw_Inv_Hoja.Reconteo Then

                            Consulta_sql = "Update " & _Global_BaseBk & "Zw_Inv_FotoInventario Set IdHojaUltReconteo = " & .IdHoja & vbCrLf &
                                           "Where IdInventario = " & .IdInventario & " And Codigo = '" & .Codigo & "'"

                            Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                        End If

                    End If

                End With

            Next

            myTrans.Commit()
            SQL_ServerClass.Sb_Cerrar_Conexion(Cn2)

            _Mensaje.EsCorrecto = True
            _Mensaje.Id = Zw_Inv_Hoja.Id
            _Mensaje.Detalle = "Crear Hoja"
            _Mensaje.Mensaje = "Documento grabado correctamente" & vbCrLf &
                               "Número de Hoja: " & Zw_Inv_Hoja.Nro_Hoja
            _Mensaje.Tag = Zw_Inv_Hoja

            _Mensaje.Icono = MessageBoxIcon.Information

        Catch ex As Exception

            _Mensaje.EsCorrecto = False
            _Mensaje.Detalle = "Error al grabar"
            _Mensaje.Mensaje = ex.Message
            _Mensaje.Icono = MessageBoxIcon.Stop

            If Not IsNothing(myTrans) Then myTrans.Rollback()

            SQL_ServerClass.Sb_Cerrar_Conexion(Cn2)

        End Try

        Return _Mensaje

    End Function

    Function Fx_Eliminar_Hoja(_IdHoja As Integer, _NombreEquipoEli As String, _CodFuncionarioEli As String) As LsValiciones.Mensajes

        Dim _Mensaje As New LsValiciones.Mensajes

        Consulta_sql = String.Empty

        'Dim _Reg = _Sql.Fx_Cuenta_Registros(_Global_BaseBk & "Zw_Inv_Hoja",
        '                                    "IdInventario = " & Zw_Inv_Hoja.IdInventario & " And Nro_Hoja = '" & Zw_Inv_Hoja.Nro_Hoja & "'")

        'If CBool(_Reg) Then
        '    _Mensaje_Stem.EsCorrecto = False
        '    _Mensaje_Stem.Detalle = "Crear Hoja"
        '    _Mensaje_Stem.Mensaje = "El Nro de Hoja " & Zw_Inv_Hoja.Nro_Hoja & " ya existe para este inventario"
        '    _Mensaje_Stem.Icono = MessageBoxIcon.Stop
        '    Return _Mensaje_Stem
        'End If

        Dim myTrans As SqlClient.SqlTransaction
        Dim Comando As SqlClient.SqlCommand

        Dim Cn2 As New SqlConnection
        Dim SQL_ServerClass As New Class_SQL '(Cadena_ConexionSQL_Server)

        SQL_ServerClass.Sb_Abrir_Conexion(Cn2)

        Try

            myTrans = Cn2.BeginTransaction()

            Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Inv_HojaEli (Id,IdInventario,Nro_Hoja,NombreEquipo,FechaCreacion,CodResponsable,IdContador1,IdContador2,FechaLevantamiento,Reconteo,CodFuncionarioEli,FechaEli,NombreEquipoEli)" & vbCrLf &
                           "Select Id,IdInventario,Nro_Hoja,NombreEquipo,FechaCreacion,CodResponsable,IdContador1,IdContador2,FechaLevantamiento,Reconteo,'" & _CodFuncionarioEli & "',GETDATE(),'" & _NombreEquipoEli & "'" & vbCrLf &
                           "From " & _Global_BaseBk & "Zw_Inv_Hoja" & vbCrLf &
                           "Where Id = " & _IdHoja
            Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Inv_HojaEli_Detalle(IdHoja,Nro_Hoja,IdInventario,Empresa,Sucursal,Bodega,Responsable,IdContador1,IdContador2,Item_Hoja,IdSector,Sector,Ubicacion,TipoConteo,Codigo,EsSeriado,NroSerie,FechaHoraToma,Rtu," & vbCrLf &
                           "RtuVariable,Udtrpr,Cantidad,Ud1,CantidadUd1,Ud2,CantidadUd2,Observaciones,Recontado,Actualizado_por,Obs_Actualizacion)" & vbCrLf &
                           "Select IdHoja,Nro_Hoja,IdInventario,Empresa,Sucursal,Bodega,Responsable,IdContador1,IdContador2,Item_Hoja,IdSector,Sector,Ubicacion,TipoConteo,Codigo,EsSeriado,NroSerie,FechaHoraToma,Rtu," & vbCrLf &
                           "RtuVariable,Udtrpr,Cantidad,Ud1,CantidadUd1,Ud2,CantidadUd2,Observaciones,Recontado,Actualizado_por,Obs_Actualizacion" & vbCrLf &
                           "From " & _Global_BaseBk & "Zw_Inv_Hoja_Detalle" & vbCrLf &
                           "Where IdHoja = " & _IdHoja
            Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            Consulta_sql = "Delete " & _Global_BaseBk & "Zw_Inv_Hoja Where Id = " & _IdHoja
            Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            Consulta_sql = "Delete " & _Global_BaseBk & "Zw_Inv_Hoja_Detalle Where IdHoja = " & _IdHoja
            Comando = New SqlClient.SqlCommand(Consulta_sql, Cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            myTrans.Commit()
            SQL_ServerClass.Sb_Cerrar_Conexion(Cn2)

            _Mensaje.EsCorrecto = True
            _Mensaje.Detalle = "Eliminar Hoja"
            _Mensaje.Mensaje = "Documento eliminado correctamente"
            _Mensaje.Icono = MessageBoxIcon.Information

        Catch ex As Exception

            _Mensaje.EsCorrecto = False
            _Mensaje.Detalle = "Error al eliminar"
            _Mensaje.Mensaje = ex.Message
            _Mensaje.Icono = MessageBoxIcon.Stop

            If Not IsNothing(myTrans) Then myTrans.Rollback()

            SQL_ServerClass.Sb_Cerrar_Conexion(Cn2)

        End Try

        Return _Mensaje

    End Function

    Function Fx_NvoNroHoja(_IdInventario As Integer) As String

        Dim _Nro_Hoja As String

        Dim _Sql As New Class_SQL '(Cadena_ConexionSQL_Server)
        Consulta_sql = "Select Max(Nro_Hoja) As Nro_Hoja From " & _Global_BaseBk & "Zw_Inv_Hoja Where IdInventario = " & _IdInventario

        Dim _Row As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)
        Dim _Ult_Nro_OT As String = NuloPorNro(_Row.Item("Nro_Hoja"), "")

        If String.IsNullOrEmpty(Trim(_Ult_Nro_OT)) Then
            _Ult_Nro_OT = "00000"
        End If

        _Nro_Hoja = Fx_Proximo_NroDocumento(_Ult_Nro_OT, 10)

        Return _Nro_Hoja

    End Function

End Class
