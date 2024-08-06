Imports System.Data
Imports System.Data.SqlClient

Public Class Clase_EliminarAnular_Documento

    Dim _Sql As New Class_SQL '(_Global_Cadena_De_Conexion_SQL)

    Enum _Accion_EA
        Anular
        Eliminar
        Modificar
    End Enum

    Function Fx_EliminarAnular_Doc(_Idmaeedo As Integer,
                                   _Cod_Func_Eliminador As String,
                                   _Accion As _Accion_EA,
                                   _Mostrar_Mensaje As Boolean) As Boolean

        Dim _FechaEliminacion = FechaDelServidor()


        If Not Revisar_Si_Se_Puede_Eliminar_El_Documento(_Idmaeedo, _Accion, _Mostrar_Mensaje) Then
            Return False
        End If

        Try

            Dim Fecha_Elimi As String = Format(_FechaEliminacion, "yyyyMMdd")

            Consulta_sql = "Select EMPRESA,SUDO,TIDO,NUDO,ENDO,SUENDO,FEEMDO,KOFUDO,VANEDO,VABRDO" & vbCrLf &
                           "FROM MAEEDO" & vbCrLf &
                           "WHERE IDMAEEDO = " & _Idmaeedo
            Dim Tabla_Doc As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

            With Tabla_Doc.Rows(0)

                Dim _Endo As String = Trim(.Item("ENDO"))
                Dim _Suendo As String = Trim(.Item("SUENDO"))
                Dim _Tido As String = Trim(.Item("TIDO"))
                Dim _Nudo As String = .Item("NUDO")
                Dim _Fecha_Doc_Origen As Date = .Item("FEEMDO")
                Dim _Fecha_Eliminacion As String = Fecha_Elimi
                Dim _Funcionario_Doc_Origen As String = .Item("KOFUDO")
                Dim _Funcionario_Eliminador As String = _Cod_Func_Eliminador
                Dim _Empresa As String = .Item("EMPRESA")
                Dim _Sucursal As String = .Item("SUDO")
                Dim _Neto_Doc_Origen As String = .Item("VANEDO")
                Dim _Bruto_Doc_Origen As String = .Item("VABRDO")

                Dim _Fecha_Ori As String = Format(_Fecha_Doc_Origen, "yyyyMMdd")



                If _Accion = _Accion_EA.Anular Then

                    Consulta_sql = "INSERT INTO ELIDDO SELECT * FROM MAEDDO WHERE MAEDDO.IDMAEEDO = " & _Idmaeedo & vbCrLf &
                                   "INSERT INTO ELIEDO SELECT * FROM MAEEDO WHERE MAEEDO.IDMAEEDO = " & _Idmaeedo & vbCrLf &
                                   "Update MAEEDO Set ESDO = 'N',LIBRO = '',KOFUAUDO = '" & _Cod_Func_Eliminador & "' Where IDMAEEDO =" & _Idmaeedo & vbCrLf & vbCrLf

                ElseIf _Accion = _Accion_EA.Eliminar Then

                    Consulta_sql = "INSERT INTO MAEELIMI (EMPRESA,TIDO,NUDO,ENDO,SUENDO,FEEMDO,FEELIDO,KOFUDO,VANEDO,VABRDO)" & vbCrLf &
                                   "SELECT EMPRESA,TIDO,NUDO,ENDO,SUENDO,FEEMDO,'" & _Fecha_Eliminacion &
                                   "',KOFUDO,VANEDO,VABRDO FROM MAEEDO" & vbCrLf &
                                   "Where IDMAEEDO =" & _Idmaeedo & vbCrLf &
                                   "INSERT INTO ELIDDO SELECT * FROM MAEDDO WHERE MAEDDO.IDMAEEDO = " & _Idmaeedo & vbCrLf &
                                   "INSERT INTO ELIEDO SELECT * FROM MAEEDO WHERE MAEEDO.IDMAEEDO = " & _Idmaeedo & vbCrLf &
                                   "DELETE MAEEDO WHERE IDMAEEDO =" & _Idmaeedo & vbCrLf

                ElseIf _Accion = _Accion_EA.Modificar Then

                    Consulta_sql = "DELETE MAEEDO WHERE IDMAEEDO =" & _Idmaeedo & vbCrLf

                End If

                Consulta_sql +=
                               "DELETE FROM MAEPOSLI" & vbCrLf &
                               "WHERE MAEPOSLI.IDMAEDDO IN (SELECT IDMAEDDO FROM MAEDDO WHERE IDMAEEDO=" & _Idmaeedo & ")" & vbCrLf &
                               "DELETE FROM MEVENTO WHERE ARCHIRVE='MAEEDO' AND IDRVE=" & _Idmaeedo & vbCrLf &
                               "DELETE FROM MAEIMLI WHERE IDMAEEDO =" & _Idmaeedo & vbCrLf &
                               "DELETE FROM MAEDTLI WHERE IDMAEEDO=" & _Idmaeedo & vbCrLf &
                               "DELETE FROM MEVENTO " &
                               "WHERE ARCHIRVE='MAEDDO' AND IDRVE IN (SELECT IDMAEDDO FROM MAEDDO WHERE IDMAEEDO=" & _Idmaeedo & ")" & vbCrLf &
                               "DELETE FROM MAEDDO WHERE IDMAEEDO=" & _Idmaeedo & vbCrLf &
                               "DELETE FROM MAEVEN WHERE IDMAEEDO=" & _Idmaeedo & vbCrLf &
                               "DELETE FROM MAEEDOOB WHERE IDMAEEDO=" & _Idmaeedo & vbCrLf &
                               "DELETE FROM TABPERMISO WHERE IDRST=" & _Idmaeedo & " AND ARCHIRST='MAEEDO'" & vbCrLf '& _
                ' "SELECT TOP 1 * FROM MAEDCR WITH (NOLOCK) WHERE IDMAEEDO=" & _Idmaeedo & vbCrLf & _
                ' "DELETE FROM MAEDCR WHERE IDMAEEDO=" & _Idmaeedo & vbCrLf


                Consulta_sql = Replace(Consulta_sql, "#Idmaeedo#", _Idmaeedo)

                Return _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(Consulta_sql)


            End With

        Catch ex As Exception

            'MessageBoxEx.Show(_Formulario, "Transaccion desecha", "Problema", _
            '                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
        End Try

    End Function

    Function Revisar_Si_Se_Puede_Eliminar_El_Documento(_Idmaeedo As Integer,
                                                       _Accion As _Accion_EA,
                                                       Optional _Mostrar_Mensaje As Boolean = False) As Boolean

        Consulta_sql = "Select Top 1 * From MAEEDO Where IDMAEEDO = " & _Idmaeedo
        Dim _RowMaeedo As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)
        Dim _Tido As String

        If (_RowMaeedo Is Nothing) Then
            Return False
        End If

        _Tido = _RowMaeedo.Item("TIDO")

        Dim Dst_Paso As New DataSet

        Consulta_sql = My.Resources.Rec_Documentos.Revisar_sutentatorio
        Consulta_sql = Replace(Consulta_sql, "#Idmaeedo#", _Idmaeedo)
        Consulta_sql = Replace(Consulta_sql, "#Tido#", _Tido)

        Dst_Paso = _Sql.Fx_Get_DataSet(Consulta_sql)

        For Each Tabla As DataTable In Dst_Paso.Tables

            If CBool(Tabla.Rows.Count) Then
                If _Mostrar_Mensaje Then
                    MsgBox("El documento es sustentatorio de otro documento" & vbCrLf &
                           "No es posible " & _Accion.ToString & " documento", MsgBoxStyle.Critical,
                           UCase(_Accion) & " DOCUMENTO")

                End If
                Return False
            End If
        Next
        Return True
    End Function


End Class
