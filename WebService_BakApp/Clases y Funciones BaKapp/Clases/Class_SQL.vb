Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Services

Public Class Class_SQL

    Dim _SQL_String_conexion As String = System.Configuration.ConfigurationManager.ConnectionStrings("db_bakapp").ToString()
    Dim _Error As String
    Dim _Cn As New SqlConnection

    Public ReadOnly Property Pro_Error() As String
        Get
            Return _Error
        End Get
    End Property

    Public Sub New()
        '_SQL_String_conexion = SQL_String_conexion
    End Sub

    Function Fx_Ej_consulta_IDU(ByVal ConsultaSql As String, _
                                Optional ByVal MostrarError As Boolean = True) As Boolean
        Try
            'Abrimos la conexión con la base de datos


            Sb_Abrir_Conexion(_Cn)
            'System.Windows.Forms.Application.DoEvents()
            Dim cmd As System.Data.SqlClient.SqlCommand
            cmd = New System.Data.SqlClient.SqlCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = ConsultaSql
            cmd.CommandTimeout = 0
            cmd.Connection = _Cn

            cmd.ExecuteNonQuery()
            'Cerramos la conexión con la base de datos
            Sb_Cerrar_Conexion(_Cn)

            'System.Windows.Forms.Application.DoEvents()
            Return True
        Catch ex As Exception
            'If MostrarError = True Then
            'MsgBox("No se realizo la operación: Insert, Update o Delete..." _
            '       , MsgBoxStyle.Critical, "Modificar tabla")
            'MsgBox(ex.Message)
            'End If
            _Error = ex.Message
            Return False
        End Try

    End Function

    Function Fx_Get_Tablas(ByVal _Consulta_sql As String) As DataTable

        Dim _Tbl As New DataTable
        _Error = String.Empty

        Try
            Sb_Abrir_Conexion(_Cn)

            Dim _SqlDa As New SqlDataAdapter

            _SqlDa = New SqlDataAdapter(_Consulta_sql, _Cn)
            _SqlDa.SelectCommand.CommandTimeout = 8000

            _SqlDa.Fill(_Tbl)
            Sb_Cerrar_Conexion(_Cn)

            ' retornar el dataTable
            Return _Tbl

            ' errores
        Catch ex As Exception
            _Error = ex.Message
        End Try

    End Function

    Function Fx_Get_DataRow(ByVal _Consulta_sql As String) As DataRow

        Try
            _Error = String.Empty
            Dim _Tbl As DataTable = Fx_Get_Tablas(_Consulta_sql)

            If CBool(_Tbl.Rows.Count) Then
                Return _Tbl.Rows(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            _Error = ex.Message
        End Try

    End Function

    Function Fx_Get_DataSet(ByVal _Consulta_sql As String, _
                            ByVal _Ds As DataSet, _
                            ByVal _Nombre_Tabla As String) As DataSet

        Dim _Tbl As New DataTable

        Try
            _Error = String.Empty
            Sb_Abrir_Conexion(_Cn)

            Dim daAuthors As New SqlDataAdapter(_Consulta_sql, _Cn)
            daAuthors.SelectCommand.CommandTimeout = 8000
            daAuthors.MissingSchemaAction = MissingSchemaAction.AddWithKey
            daAuthors.Fill(_Ds, _Nombre_Tabla)
            Sb_Cerrar_Conexion(_Cn)

            ' retornar el dataTable
            Return _Ds

            ' errores
        Catch ex As Exception
            _Error = ex.Message
        End Try


    End Function

    Function Fx_Get_DataSet(ByVal Consulta_sql As String) As DataSet

        Try
            _Error = String.Empty
            Sb_Abrir_Conexion(_Cn)

            Dim dt As DataTable = New DataTable()

            Dim _SqlDa As New SqlDataAdapter
            Dim _DataSt As New DataSet

            _SqlDa = New SqlDataAdapter(Consulta_sql, _Cn)
            _SqlDa.SelectCommand.CommandTimeout = 8000
            _SqlDa.Fill(_DataSt)

            Sb_Cerrar_Conexion(_Cn)

            Return _DataSt
            ' errores
        Catch ex As Exception
            _Error = ex.Message
        End Try


    End Function

    Function Fx_Extrae_Archivo_desde_BD(ByVal _Tabla As String, _
                                        ByVal _Campo As String, _
                                        ByVal _Condicion As String, _
                                        ByVal _Nom_Archivo As String, _
                                        ByVal _Dir_Temp As String) As Boolean

        Dim data As Byte() = Nothing

        Try
            ' Construimos los correspondientes objetos para
            ' conectarnos a la base de datos de SQL Server,
            ' utilizando la seguridad integrada de Windows NT.
            '
            Using cn As New SqlConnection
                Dim sCnn = _SQL_String_conexion
                cn.ConnectionString = sCnn
                Dim cmd As SqlCommand = cn.CreateCommand 'cnn.CreateCommand()
                ' Seleccionamos únicamente el campo que contiene
                ' los documentos, filtrándolo mediante su
                ' correspondiente campo identificador.
                '
                cmd.CommandText = "SELECT " & _Campo & " From " & _Tabla & " WHERE " & _Condicion
                ' Abrimos la conexión.
                cn.Open()
                ' Creamos un DataReader.
                Dim dr As SqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                cmd.CommandTimeout = 8000
                ' Leemos el registro.
                dr.Read()
                ' El tamaño del búfer debe ser el adecuado para poder
                ' escribir en el archivo todos los datos leídos.
                '
                ' Si el parámetro 'buffer' lo pasamos como Nothing, obtendremos
                ' la longitud del campo en bytes.
                '
                Dim bufferSize As Integer = Convert.ToInt32(dr.GetBytes(0, 0, Nothing, 0, 0))

                ' Creamos el array de bytes. Como el índice está
                ' basado en cero, la longitud del array será la
                ' longitud del campo menos una unidad.
                '
                data = New Byte(bufferSize - 1) {}

                ' Leemos el campo.
                '
                dr.GetBytes(0, 0, data, 0, bufferSize)

                ' Cerramos el objeto DataReader, e implícitamente la conexión.
                '
                dr.Close()

            End Using

            ' Procedemos a crear el archivo, en el ejemplo
            ' un documento de Microsoft Word.
            '
            If File.Exists(_Dir_Temp & "\" & _Nom_Archivo) Then File.Delete(_Dir_Temp & "\" & _Nom_Archivo)

            Using fs As New IO.FileStream(_Dir_Temp & "\" & _Nom_Archivo, IO.FileMode.CreateNew, IO.FileAccess.Write)

                ' Crea el escritor para la secuencia.
                Dim bw As New IO.BinaryWriter(fs)

                ' Escribir los datos en la secuencia.
                bw.Write(data)

            End Using


            'Sb_WriteBinaryFile(Me, _Dir_Temp & "\" & _Nom_Archivo, data)
            Return True
        Catch ex As Exception

        End Try

    End Function

    'System.Windows.Forms.Application.DoEvents()
    Sub Sb_Abrir_Conexion(ByVal _Cn As SqlConnection)

        Try
            _Error = String.Empty
            If _Cn.State = ConnectionState.Open Then
                ' Cerrar conexion
                _Cn.Close()
            End If

            _Cn.ConnectionString = _SQL_String_conexion
            _Cn.Open()
            'MsgBox(sCnn)

        Catch ex As SqlClient.SqlException 'Exception
            _Error = ex.Message
        End Try

    End Sub

    Sub Sb_Cerrar_Conexion(ByVal _Cn As SqlConnection)
        '_Error = String.Empty
        Try
            If _Cn.State = ConnectionState.Open Then
                '' Cerrar conexion
                _Cn.Close()
            End If
        Catch ex As Exception
            ' _Error = ex.Message
        End Try
    End Sub

    Function Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(ByVal Consulta_sql As String) As Boolean

        Dim myTrans As SqlClient.SqlTransaction
        Dim Comando As SqlClient.SqlCommand

        Dim _Fijar As Boolean
        Dim _Cn As New SqlConnection

        Try

            Sb_Abrir_Conexion(_Cn)

            If String.IsNullOrEmpty(_Error) Then

                myTrans = _Cn.BeginTransaction()

                Comando = New SqlClient.SqlCommand(Consulta_sql, _Cn)
                Comando.Transaction = myTrans
                Comando.ExecuteNonQuery()

                '**********************************'**********************************

                myTrans.Commit()
                Sb_Cerrar_Conexion(_Cn)

                _Fijar = True
                Return _Fijar
            Else
                Throw New Exception(_Error)
            End If

        Catch ex As Exception
            _Error = ex.Message
        Finally

            If Not _Fijar Then
                myTrans.Rollback()
            End If

        End Try


    End Function

    Function Fx_Trae_Dato(ByVal _Tabla As String, _
                         ByVal _Campo As String, _
                         Optional ByVal _Condicion As String = "", _
                         Optional ByVal _DevNumero As Boolean = False, _
                         Optional ByVal _MostrarError As Boolean = True, _
                         Optional ByVal _Dato_Default As String = "") As String
        Try

            Dim _Valor
            Dim _Valor_Si_No_Encuentra As String

            If Not String.IsNullOrEmpty(_Condicion) Then
                _Condicion = vbCrLf & "And " & _Condicion
            End If

            If _DevNumero Then

            End If
            'Then Valor_Si_No_Encuentra = 0

            Dim _Sql As String = "SELECT TOP (1) " & _Campo & " AS CAMPO FROM " & _Tabla & vbCrLf & _
                                 "Where 1 > 0" & _Condicion


            Dim _Tbl As DataTable = Fx_Get_Tablas(_Sql)

            Dim cuenta As Long = _Tbl.Rows.Count

            If CBool(_Tbl.Rows.Count) Then

                _Valor = _Tbl.Rows(0).Item("CAMPO")

                If _DevNumero Then
                    _Valor = NuloPorNro(_Valor, 0)
                Else
                    _Valor = NuloPorNro(_Valor, "")
                End If

            Else
                If _DevNumero Then
                    _Valor = 0
                Else
                    _Valor = ""
                End If
            End If

            If String.IsNullOrEmpty(_Valor) Then
                _Valor = _Dato_Default
            End If

            Return _Valor

        Catch ex As Exception
            If _MostrarError Then
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!!")
            Else
                Return ex.Message
            End If
        End Try

    End Function

    Function Fx_Cuenta_Registros(ByVal _Tabla As String, _
                                 Optional ByVal _Condicion As String = "") As Double

        If Not String.IsNullOrEmpty(_Condicion) Then
            _Condicion = vbCrLf & "And " & _Condicion
        End If

        Dim _Sql As String = "Select Count(*) As Cuenta From " & _Tabla & " Where 1 > 0 " & _Condicion

        Dim _RowTabpre As DataRow = Fx_Get_DataRow(_Sql)

        Dim _Cuenta As Double

        If (_RowTabpre Is Nothing) Then
            _Cuenta = 0
        Else
            _Cuenta = _RowTabpre.Item("Cuenta")
        End If

        Return _Cuenta

    End Function

    Function Fx_SqlDataReader(ByVal Consulta_sql As String) As SqlDataReader

        Sb_Abrir_Conexion(_Cn)
        Dim _Comando As SqlClient.SqlCommand

        _Comando = New SqlCommand(Consulta_sql, _Cn)
        Dim _Sql_DReader As SqlDataReader = _Comando.ExecuteReader()

        Return _Sql_DReader

    End Function

    Sub Sb_Eliminar_Tabla_De_Paso(ByVal _Tabla_Paso As String)

        Dim _ConsultaSql As String = "BEGIN TRY" & vbCrLf & _
                                     "DROP TABLE " & _Tabla_Paso & vbCrLf & _
                                     "End Try" & vbCrLf & _
                                     "BEGIN CATCH" & vbCrLf & _
                                     "END CATCH"

        Fx_Ej_consulta_IDU(_ConsultaSql, False)

    End Sub

    Function Fx_Probar_Conexion() As String

        Sb_Abrir_Conexion(_Cn)

        If String.IsNullOrEmpty(_Error) Then
            _Error = "Conexión OK"
        End If

        Return _Error

    End Function

    Function Fx_Existe_Tabla(ByVal _Tabla As String) As Boolean

        Dim _ConsultaSql As String

        If _Tabla.Contains(_Global_BaseBk) Then

            _Tabla = Replace(_Tabla, _Global_BaseBk, "")
            _ConsultaSql = "USE " & Replace(_Global_BaseBk, ".dbo.", "") & "
                            SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & _Tabla & "'"

        Else

            _ConsultaSql = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & _Tabla & "'"

        End If

        Dim _Tbl As DataTable = Fx_Get_Tablas(_ConsultaSql)

        Return _Tbl.Rows.Count

    End Function

End Class
