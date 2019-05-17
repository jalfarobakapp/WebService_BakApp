'Imports DevComponents.DotNetBar
'Imports Funciones_BakApp


Public Class Class_Conectar_Base_BakApp

    Dim _Bk_BaseDeDatos As String
    Dim _Existe_Base As Boolean
    Dim _Sql As New Class_SQL '(_Global_Cadena_Conexion_SQL_Server)
    'Dim _Formulario As Form
    Dim _Row_Tabcarac As DataRow

    Public ReadOnly Property Pro_Existe_Base() As Boolean
        Get
            Return _Existe_Base
        End Get
    End Property

    Public ReadOnly Property Pro_Row_Tabcarac() As DataRow
        Get
            Return _Row_Tabcarac
        End Get
    End Property

    Public Sub New()

        '_Formulario = Formulario
        Consulta_sql = "Select Top 1 *,'['+NOKOCARAC+'].dbo.' As Global_BaseBk From TABCARAC Where KOTABLA = 'BAKAPP' And KOCARAC = 'BASE'"
        _Row_Tabcarac = _Sql.Fx_Get_DataRow(Consulta_sql)

        _Existe_Base = Fx_Existe_Base_En_SQLServer(_Row_Tabcarac)

    End Sub


    Private Function Fx_Existe_Base_En_SQLServer(ByVal _Row As DataRow) As Boolean

        Dim _Reg As Boolean

        If Not (_Row Is Nothing) Then

            Dim _Nombre_Base = _Row_Tabcarac.Item("NOKOCARAC")
            _Reg = CBool(_Sql.Fx_Cuenta_Registros("sys.databases", "name = '" & _Nombre_Base & "'"))

        End If
        _Existe_Base = _Reg
        Return _Reg

    End Function

    Function Fx_Grabar_Base_Bakapp_En_Tabcarac() As Boolean

        Dim _Nokocarac As String


        '_Nokocarac = InputBox_Bk(_Formulario, "Ingrese el nombre de la" & vbCrLf & "base de datos BakApp", _
        '                          "Base de datos BakApp", "", False, _Tipo_Mayus_Minus.Normal, 0, True, _Tipo_Imagen.Texto, False)

        If _Nokocarac <> "@@Accion_Cancelada##" Then

            Dim _Reg As Boolean

            _Reg = CBool(_Sql.Fx_Cuenta_Registros("sys.databases", "name = '" & _Nokocarac & "'"))

            If _Reg Then

                Consulta_sql = "Delete TABCARAC Where KOTABLA = 'BAKAPP' And KOCARAC = 'BASE'" & vbCrLf & _
                               "Insert Into TABCARAC (KOTABLA,NOKOTABLA,KOCARAC,NOKOCARAC) VALUES " & _
                               "('BAKAPP','','BASE','" & _Nokocarac & "')"

                If _Sql.Fx_Ej_consulta_IDU(Consulta_sql) Then

                    Consulta_sql = "Select Top 1 *,NOKOCARAC+'.dbo.' As Global_BaseBk From TABCARAC Where KOTABLA = 'BAKAPP' And KOCARAC = 'BASE'"
                    _Row_Tabcarac = _Sql.Fx_Get_DataRow(Consulta_sql)

                    Return True
                End If

            End If

        Else
            Return False
        End If


    End Function

End Class
