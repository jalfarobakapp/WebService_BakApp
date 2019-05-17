Imports Microsoft.VisualBasic
Imports System
'Imports System.Windows.Forms
Imports System.Drawing
Imports System.Math
Imports System.Data
Imports System.Data.SqlClient
Imports System.Security.Cryptography
'Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Net
Imports System.IO
Imports System.Drawing.Printing
'Imports DevComponents.DotNetBar
Imports System.Globalization
Imports System.Text.RegularExpressions


Public Module Funciones


    Public PickingActivo As Boolean
    Public _MTS_Lista_activo As Boolean


    Public Function Hora_Grab_fx(ByVal _EsAjuste As Boolean, ByVal _Fecha As Date) As String

        Dim _HH_sistem As Date

        _HH_sistem = _Fecha

        Dim _HH, _MM, _SS As Double

        _HH = _HH_sistem.Hour
        _MM = _HH_sistem.Minute
        _SS = _HH_sistem.Second

        If _EsAjuste Then
            _HH = 23 : _MM = 59 : _SS = 59
        End If

        Dim _HoraGrab As String = Math.Round((_HH * 3600) + (_MM * 60) + _SS, 0)

        Return _HoraGrab

    End Function

    Function Fx_Genera_Licencia_BakApp(ByVal _RutEmpresa As String, _
                                        ByVal _FechaCaduca As Date, _
                                        ByVal _CantLicencias As Integer, _
                                        ByVal _Palabra_X As String) As String()

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

    Function numero_(ByVal Num As String, ByVal d As Integer) As String
        Dim i As Integer
        Dim nro As String
        nro = Len(RTrim$(Num))

        For i = nro To d - 1
            Num = "0" & Num
        Next

        Return RTrim$(Num)
    End Function

    Function QuitaEspacios_ParaCodigos(ByVal s As String, _
                           ByVal lon As Integer) As String

        Dim arr(lon - 1) As Char '= s.ToCharArray
        arr = s.ToCharArray
        Dim Contador = arr.Length - 1
        Dim _palabra As String

        ' arr = s.ToCharArray

        Do While (Contador >= 0)

            Dim _Asc As Integer
            Dim _Letra As String = arr(Contador)
            _Asc = Asc(_Letra)

            If _Asc <> 160 Then
                If Contador = arr.Length - 1 Then
                    _palabra = s
                Else
                    _palabra = Mid(s, 1, Contador)
                End If

                Exit Do
            End If

            If Contador = 0 Then

            End If

            Contador -= 1
        Loop

        Return _palabra
        ' Return corre
    End Function

    Function Ruta_conexion(ByVal Ruta As String) As String
        Try

            Dim texto As String
            Dim sr As New System.IO.StreamReader(Ruta)
            texto = sr.ReadToEnd()
            sr.Close()
            Return texto
        Catch ex As Exception

        End Try
    End Function

    Function LeeArchivo(ByVal Ruta As String) As String
        Dim texto As String
        Dim sr As New System.IO.StreamReader(Ruta)
        texto = sr.ReadToEnd()
        sr.Close()
        Return texto
    End Function

    Function Encripta_md5(ByVal TextoAEncriptar As String) As String
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

    Function Fx_Proximo_NroDocumento(ByVal _NrNumeroDoco As String) As String

        Dim _Pos As Integer = 0
        Dim _Prefijo As String
        Dim _Cadena As String = String.Empty
        Dim _NvoNumero1 As Integer
        Dim _NvoNumero2 As String

        Do While _Pos < 10
            _Pos += 1
            Dim _Caracter As String = Right(_NrNumeroDoco, 11 - _Pos)

            If IsNumeric(_Caracter) Then
                _Prefijo = Left(_NrNumeroDoco, _Pos - 1)
                _Cadena = Right(_NrNumeroDoco, 11 - _Pos)
                Exit Do
            End If

        Loop

        _NvoNumero1 = CInt(_Cadena) + 1
        _NvoNumero2 = _Prefijo & numero_(_NvoNumero1, Len(_Cadena))

        Return _NvoNumero2

    End Function

    Function RutDigito(ByVal numero As String) As String

        Dim cuenta, Suma, resto, Digito As Integer
        Dim dig As Decimal
        Suma = 0
        cuenta = 2

        Do
            dig = numero Mod 10
            numero = Int(numero / 10)
            Suma = Suma + (dig * cuenta)
            cuenta = cuenta + 1
            If cuenta = 8 Then cuenta = 2
        Loop Until numero = 0

        resto = Suma Mod 11
        Digito = 11 - resto

        Select Case Digito
            Case 10 : Return "K"
            Case 11 : Return "0"
            Case Else : Return Trim(Str(Digito))
        End Select

    End Function

    Function VerificaDigito(ByVal RUT As String) As Boolean
        Try

            Dim Rt(1) As String
            Rt = Split(RUT, "-")

            Dim DigitoVerdadero, Digi As String
            DigitoVerdadero = UCase(RutDigito(Rt(0)))
            Digi = UCase(Rt(1))


            If Trim(Digi) = Trim(DigitoVerdadero) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try

    End Function

    Function NuloPorNro(Of T)(ByVal value As T, ByVal defaultValue As T) As T

        Dim obj1 As Object = value
        Dim obj2 As Object = defaultValue

        Try
            If ((obj1 Is DBNull.Value) OrElse (obj1 Is Nothing)) Then
                ' Es NULL; devolvemos el valor por defecto siempre
                ' y cuando éste tampoco sea NULL.
                '
                If (Not obj2 Is DBNull.Value) Then
                    Return defaultValue

                Else
                    Return Nothing

                End If

            Else
                ' No es NULL ni Nothing; devolvemos el valor pasado.
                '

                Return value

            End If

        Catch ex As Exception
            Return Nothing

        End Try

    End Function

    Public Function Primerdiadelmes(ByVal fecha As Date) As Date
        Dim rtn As New Date
        rtn = fecha 'Date.Now
        rtn = rtn.AddDays(-rtn.Day + 1)
        Return rtn
    End Function

    Public Function ultimodiadelmes(ByVal fecha As Date) As Date
        Dim rtn As New Date
        rtn = fecha.Date ' fecha 'Date.Now
        rtn = rtn.AddDays(-rtn.Day + 1).AddMonths(1).AddDays(-1)
        Return rtn
    End Function

    Function es_numero(ByVal numero As String) As Boolean

        Dim w As Integer
        Dim lineax As String

        w = 0

        Select Case RTrim$(Mid(numero, 1, 1)) & RTrim$(Mid(numero, 2, 1))
            Case "00" : w = 1
            Case "01" : w = 1
            Case "02" : w = 1
            Case "03" : w = 1
            Case "04" : w = 1
            Case "05" : w = 1
            Case "06" : w = 1
            Case "07" : w = 1
            Case "08" : w = 1
            Case "09" : w = 1
        End Select

        If w = 1 Then
            es_numero = False
            Exit Function
        End If

        For w = 1 To Len(numero)
            lineax = RTrim$(Mid(numero, w, 1))

            If lineax = "" Then
                es_numero = False
                Exit Function
            End If

            If InStr("-.,1234567890", RTrim$(lineax)) = 0 Then
                es_numero = False
                Exit Function
            Else
                es_numero = True
            End If
        Next


    End Function

    Function SoloNumeros(ByVal Keyascii As Short, _
                        Optional ByVal _Solo_Nros As Boolean = True) As Short


        Dim _Sd = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator

        Dim T As String = Chr(Keyascii)
        ' Dim dd '= InStr("1234567890,.-", T)

        If _Solo_Nros Then
            'dd = InStr("1234567890", T)
            If InStr("1234567890", Chr(Keyascii)) = 0 Then
                SoloNumeros = 0
            Else
                SoloNumeros = Keyascii
            End If
        Else
            ' dd = InStr("1234567890,.-", T)
            If InStr("1234567890,.-", Chr(Keyascii)) = 0 Then
                SoloNumeros = 0
            Else
                SoloNumeros = Keyascii
            End If
        End If



        Select Case Keyascii
            Case 8
                SoloNumeros = Keyascii
            Case 13
                SoloNumeros = Keyascii
        End Select
    End Function

    Function SoloNumerosSinPuntosNiComas(ByVal Keyascii As Short) As Short
        If InStr("1234567890", Chr(Keyascii)) = 0 Then
            SoloNumerosSinPuntosNiComas = 0
        Else
            SoloNumerosSinPuntosNiComas = Keyascii
        End If
        Select Case Keyascii
            Case 8
                SoloNumerosSinPuntosNiComas = Keyascii
            Case 13
                SoloNumerosSinPuntosNiComas = Keyascii
        End Select
    End Function

    Function SoloLetrasNumeros(ByVal Keyascii As Short) As Short
        If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890,.-", Chr(Keyascii)) = 0 Then
            SoloLetrasNumeros = 0
        Else
            SoloLetrasNumeros = Keyascii
        End If
    End Function

    Function CrearArchivoTxt(ByVal Ruta As String, _
                             ByVal NombreArchivo As String, _
                             ByVal Cuerpo As String)
        Try

            Dim RutaArchivo As String = Ruta & NombreArchivo

            Dim oSW As New System.IO.StreamWriter(RutaArchivo)

            oSW.WriteLine(Cuerpo)
            oSW.Close()

            'aqui creo el archivo oculto,,, si no se pone este instrucion no pasa nada .. solo es para 
            'asignarles caracteristicas a los archivos 
            'quitalo como comentario y calalo
            'SetAttr(RutaArchivo, vbHidden)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Public Function _Global_Fx_Cambio_en_la_Grilla(ByVal _Tbl_Grilla As DataTable, _
                                                   Optional ByVal _Rev_Insertas As Boolean = True, _
                                                   Optional ByVal _Rev_Eliminadas As Boolean = True, _
                                                   Optional ByVal _Rev_Modificada As Boolean = True) As Boolean

        Dim _Modificado As Boolean

        For Each _Fila As DataRow In _Tbl_Grilla.Rows
            Select Case _Fila.RowState
                Case DataRowState.Added
                    If _Rev_Insertas Then _Modificado = True
                Case DataRowState.Deleted
                    If _Rev_Eliminadas Then _Modificado = True
                Case DataRowState.Detached
                    _Modificado = True
                Case DataRowState.Modified
                    If _Rev_Modificada Then _Modificado = True
            End Select

            If _Modificado Then Exit For
        Next

        Return _Modificado

    End Function

    Public Sub Sb_AddToLog(ByVal Accion As String, _
                           ByVal Descripcion As String, _
                           ByVal TxtLog As Object, _
                           Optional ByVal _Incluir_FechaHora As Boolean = True)
        If _Incluir_FechaHora Then
            TxtLog.Text += DateTime.Now.ToString() & " - " & Accion & " (" & Descripcion & ")" & vbCrLf
        Else
            TxtLog.Text += Accion & " (" & Descripcion & ")" & vbCrLf
        End If

        TxtLog.Select(TxtLog.Text.Length - 1, 0)
        TxtLog.ScrollToCaret()

    End Sub

    Function Generar_Filtro_IN_Arreglo(ByVal Arreglo() As String, _
                                       ByVal EsNumero As Boolean)

        Dim Cadena As String = String.Empty
        Dim Separador As String = ""

        If EsNumero Then
            Separador = "#"
        Else
            Separador = "@"
        End If

        'If (Tabla Is Nothing) Then Return "()"

        Dim i = 0
        For Each Valor As String In Arreglo
            If Not String.IsNullOrEmpty(Valor) Then
                Cadena = Cadena & Separador & Trim(Valor) & Separador '& Coma
                i += 1
            End If
        Next

        If EsNumero Then
            Cadena = Replace(Cadena, "##", ",")
            Cadena = Replace(Cadena, "#", "")
        Else
            Cadena = Replace(Cadena, "@@", "@,@")
            Cadena = Replace(Cadena, "@", "'")
        End If

        Cadena = "(" & Cadena & ")"

        Return Cadena

    End Function

    Function Rellenar(ByVal Cadena As String, _
                      ByVal CantCaracteres As Integer, _
                      ByVal Relleno As String, Optional ByVal Derecha As Boolean = True) As String
        Dim i As Integer
        Dim nro As String
        nro = Len(Cadena)

        Dim Cantidad As Integer = CantCaracteres - nro

        If Cantidad > 0 Then
            For i = 0 To Cantidad - 1
                If Derecha = True Then
                    Cadena = Cadena & Relleno
                Else
                    Cadena = Relleno & Cadena
                End If
            Next
        End If

        Return Cadena
    End Function

    Public Function De_Num_a_Tx_01(ByVal lNumero As Double, _
                               Optional ByVal bEntero As Boolean = False, _
                               Optional ByVal nDecimales As Integer = 2) As String
        '-------------------------------------------------§§§----'
        ' FUNCION PARA CONVERTIR UN NUMERO EN TEXTO
        '-------------------------------------------------§§§----'
        '
        On Error GoTo fin
        '
        Dim sNumero As String
        Dim nLong1 As Integer
        Dim nCont1 As Integer
        '
        If bEntero = True Then
            sNumero = CStr(Format(lNumero, "########0"))
            ''
        Else
            Select Case nDecimales
                Case -1 : sNumero = CStr(Format(lNumero, "########0.#########"))
                Case 1 : sNumero = CStr(Format(lNumero, "########0.#"))
                Case 2 : sNumero = CStr(Format(lNumero, "########0.0#"))
                Case 3 : sNumero = CStr(Format(lNumero, "########0.00#"))
                Case 4 : sNumero = CStr(Format(lNumero, "########0.000#"))
                Case 5 : sNumero = CStr(Format(lNumero, "########0.0000#"))
                Case 6 : sNumero = CStr(Format(lNumero, "########0.00000#"))
                Case 7 : sNumero = CStr(Format(lNumero, "########0.000000#"))
                Case 8 : sNumero = CStr(Format(lNumero, "########0.0000000#"))
                Case 9 : sNumero = CStr(Format(lNumero, "########0.00000000#"))
                Case 9 : sNumero = CStr(Format(lNumero, "########0.00000000#"))
                Case 10 : sNumero = CStr(Format(lNumero, "########0.000000000#"))
                Case 11 : sNumero = CStr(Format(lNumero, "########0.0000000000#"))
                Case 12 : sNumero = CStr(Format(lNumero, "########0.00000000000#"))
                Case Else : sNumero = CStr(Format(lNumero, "########0.00#"))
            End Select
            ''
        End If
        '
        nLong1 = Len(sNumero)
        '
        For nCont1 = 1 To nLong1
            If Mid$(sNumero, nCont1, 1) = "," Then Mid(sNumero, nCont1, 1) = "."
            ''
        Next nCont1
        '
        If bEntero = True Then
            De_Num_a_Tx_01 = sNumero
            ''
        ElseIf InStr(sNumero, ".") > 0 Then
            If (Len(sNumero) = InStr(sNumero, ".")) And (nDecimales = -1) Then
                De_Num_a_Tx_01 = Mid$(sNumero, 1, InStr(sNumero, ".") - 1)
                ''
            Else
                De_Num_a_Tx_01 = sNumero
                ''
            End If
            ''
        Else
            De_Num_a_Tx_01 = sNumero & ".0"
            ''
        End If
        '
        Exit Function
        '
fin:
        De_Num_a_Tx_01 = "###.###"
        ''
    End Function

    '‘———————————————— -§§§— - ’
    '‘ FUNCION PARA CONVERTIR UN TEXTO EN NUMERO DECIMAL
    '‘———————————————— -§§§— - ’

    Public Function De_Txt_a_Num_01(ByVal sTexto As String, _
                                       Optional ByVal nDecimales As Integer = 3, _
                                       Optional ByVal sP_Formato_Decimal As String = "") As Double
        '-------------------------------------------------§§§----'
        ' FUNCION PARA CONVERTIR UN TEXTO EN NUMERO DECIMAL
        '-------------------------------------------------§§§----'
        '
        Dim bCte2 As Boolean
        '
        Dim nContador1 As Integer
        Dim nContador2 As Integer
        Dim nLong_Total As Integer
        Dim nPos_Punto As Integer
        Dim nCte1 As Integer
        Dim nDecimal As Integer
        '
        Dim lNumeruco As Double
        '
        Dim sNumero As String
        Dim sL_Aux_01 As String
        '
        Dim sL_Array_Pto_01() As String
        Dim sL_Array_Coma_01() As String
        '
        On Error GoTo Error_Numero
        '
        '-------------------------------------------------§§§----'
        Select Case sP_Formato_Decimal
            Case "."    ' USAMOS "." COMO SEPARADOR DE DECIMALES
                ' Y LA "," LA ELIMINAMOS
                sL_Array_Pto_01 = Split(sTexto, ".")
                sL_Array_Coma_01 = Split(sTexto, ",")
                '
                sL_Aux_01 = ""
                For nContador1 = LBound(sL_Array_Coma_01) To UBound(sL_Array_Coma_01)
                    sL_Aux_01 = sL_Aux_01 & sL_Array_Coma_01(nContador1)
                    ''
                Next nContador1
                '
                sTexto = sL_Aux_01
                ''
            Case ","    ' USAMOS "," COMO SEPARADOR DE DECIMALES
                ' Y EL "." LE ELIMINAMOS
                sL_Array_Pto_01 = Split(sTexto, ".")
                sL_Array_Coma_01 = Split(sTexto, ",")
                '
                sL_Aux_01 = ""
                For nContador1 = LBound(sL_Array_Pto_01) To UBound(sL_Array_Pto_01)
                    sL_Aux_01 = sL_Aux_01 & sL_Array_Pto_01(nContador1)
                    ''
                Next nContador1
                '
                sTexto = sL_Aux_01
                ''
        End Select
        '-------------------------------------------------§§§----'
        '
        lNumeruco = 0
        '
        If nDecimales >= 0 Then
            nDecimal = nDecimales
            ''
        Else
            nDecimal = 3
            ''
        End If
        '
        sTexto = Trim(sTexto)
        '
        If InStr(1, sTexto, "-") > 0 Then
            'Es un numero negativo
            bCte2 = True
            sTexto = Mid$(sTexto, 2)
            ''
        ElseIf InStr(1, sTexto, "+") > 0 Then
            'Es un numero positivo (con signo)
            bCte2 = False
            sTexto = Mid$(sTexto, 2)
            ''
        Else
            'Es un numero positivo
            bCte2 = False
            ''
        End If
        '
        nLong_Total = Len(sTexto)
        '
        For nContador1 = 1 To nLong_Total
            If Mid(sTexto, nContador1, 1) = "," Then Mid(sTexto, nContador1, 1) = "."
            ''
        Next nContador1
        '
        If InStr(1, sTexto, ".") <= 0 Then sTexto = sTexto & ".0"
        '
        nPos_Punto = InStr(1, sTexto, ".")
        '
        nContador2 = 0
        For nContador1 = 1 To nLong_Total
            If Mid$(sTexto, nContador1, 1) <> "." Then
                'No estamos en el caracte "."
                If nContador1 < nPos_Punto And nPos_Punto <> 0 Then
                    nCte1 = 1
                    ''
                Else
                    nContador2 = nContador2 + 1
                    nCte1 = 0
                    ''
                End If
                '
                sNumero = Mid$(sTexto, nContador1, 1)
                '
                If nContador2 > nDecimal Then
                    If sNumero > 5 Then lNumeruco = lNumeruco + (CSng(1) * (10 ^ (nPos_Punto - nContador1 - nCte1 + 1)))
                    nContador1 = nLong_Total
                    ''
                Else
                    lNumeruco = lNumeruco + (CSng(sNumero) * (10 ^ (nPos_Punto - nContador1 - nCte1)))
                    ''
                End If
                ''
            End If
            ''
        Next nContador1
        '
        If bCte2 = True Then
            De_Txt_a_Num_01 = (-1) * lNumeruco
            ''
        Else
            De_Txt_a_Num_01 = (1) * lNumeruco
            ''
        End If
        '
        If (nDecimales >= 0) Then De_Txt_a_Num_01 = Round(De_Txt_a_Num_01, nDecimales)
        '
        Exit Function
        '
Error_Numero:
        '
        '-------------------------------------------------§§§----'
        ' ERROR DE NUMERO
        '-------------------------------------------------§§§----'
        '
        De_Txt_a_Num_01 = -1.75E+308
        ''
    End Function

    Function llena_tabla_sola(ByVal Arreglo(,) As String)

        Dim dt As New DataTable
        dt.Columns.Add("Padre")
        dt.Columns.Add("Hijo")

        Dim dr As DataRow = dt.NewRow

        ' LBound(Arreglo)

        For i = 0 To Arreglo.GetUpperBound(0)
            dr = dt.NewRow
            dr("Padre") = Arreglo(i, 0)
            dr("Hijo") = Arreglo(i, 1)
            dt.Rows.Add(dr)
            dt.AcceptChanges()
        Next

        Return dt
    End Function

    Function Cuentadias(ByVal FechaInicio As Date, _
                    ByVal FechaFin As Date, _
                    ByVal Diadelasemana As FirstDayOfWeek) As Integer

        Dim n As Integer
        Dim Fechaini As Date = FechaInicio

        Do Until FechaFin < Fechaini

            If Weekday(Fechaini) = Diadelasemana Then
                n = n + 1
            End If
            Fechaini = Fechaini.AddDays(1)
        Loop
        Return n

    End Function

    Function Fx_Crea_Tabla_Con_Filtro(ByVal dt As DataTable, ByVal filter As String, ByVal sort As String) As DataTable

        Dim rows As DataRow()

        Dim dtNew As DataTable

        ' copy table structure
        dtNew = dt.Clone()

        ' sort and filter data
        rows = dt.Select(filter, sort)

        ' fill dtNew with selected rows
        For Each dr As DataRow In rows
            dtNew.ImportRow(dr)
        Next

        ' return filtered dt
        Return dtNew

    End Function

    Private Function BuscarTextoGrilla(ByVal Texto As String, ByVal Busqueda As String) As Boolean
        Dim i As Integer
        i = InStr(1, Texto, Busqueda)
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Fx_Rellena_ceros(ByVal _NroDoc As String, _
                                    ByVal _NroCaracateres As Integer, _
                                    Optional ByVal _Suma_uno As Boolean = False) As String

        Dim _Contador = 1
        Dim _Tot_carac = Len(_NroDoc)


        Do While _Contador < _NroCaracateres
            Dim _Pl = Microsoft.VisualBasic.Strings.Right(_NroDoc, _Contador)
            If Not IsNumeric(_Pl) Then
                Exit Do
            End If

            _Contador += 1
        Loop

        Dim _Cadena As String
        Dim _Cadena2 = Microsoft.VisualBasic.Strings.Right(_NroDoc, _Contador - 1)

        If _Cadena2 = _NroDoc Then
            _Cadena = numero_(_Cadena2, _NroCaracateres)
            Return _Cadena
        End If


        Dim _Cadena1 = Mid(_NroDoc, 1, _Tot_carac - (_Contador - 1))

        If _Suma_uno Then _Cadena2 += 1

        If String.IsNullOrEmpty(_Cadena2) Then
            Return _Cadena1
        End If

        Dim _nr = Len(_Cadena1)

        _Cadena = _Cadena1 & numero_(_Cadena2, _NroCaracateres - _nr)

        Return _Cadena

    End Function

    Function _Dev_HoraGrab(ByVal Hora As String)

        Dim _HH, _MM, _SS As Double
        Dim _Horagrab As Integer

        _HH = Mid(Hora, 1, 2)
        _MM = Mid(Hora, 4, 2)
        _SS = Mid(Hora, 7, 2)

        _Horagrab = Math.Round((_HH * 3600) + (_MM * 60) + _SS, 0)

        Return _Horagrab

    End Function

    Function Fx_Dias_Habiles(ByVal _Fecha_inicial As Date, ByVal _Fecha_final As Date) As Integer

        Dim dias As Integer
        _Fecha_inicial = DateAdd(DateInterval.Day, 1, _Fecha_inicial) 'agrego un dia adicional para la cuenta ya veraz porque 

        Dim dha As Integer = DateDiff(DateInterval.Day, _Fecha_inicial, _Fecha_final)

        Dim _Dia As Integer
        For _x = 0 To dha '- 1
            _Dia = Weekday(_Fecha_inicial)
            If _Dia <> "1" And _Dia <> "7" Then
                dias += 1
            End If
            _Fecha_inicial = DateAdd(DateInterval.Day, 1, _Fecha_inicial)
        Next

        Return dias

    End Function

    Enum Opcion_Dias
        Habiles
        Lunes
        Marte
        Miercoles
        Jueves
        Viernes
        Sabado
        Domingo
        Todos
    End Enum

    Function Fx_Cuenta_Dias(ByVal _Fecha_inicial As Date, _
                            ByVal _Fecha_final As Date, _
                            ByVal _Dias_a_contar As Opcion_Dias) As Integer

        Dim dias As Integer
        ' _Fecha_inicial = DateAdd(DateInterval.Day, 1, _Fecha_inicial) 'agrego un dia adicional para la cuenta ya veraz porque 

        Dim dha As Integer = DateDiff(DateInterval.Day, _Fecha_inicial, _Fecha_final)

        Dim _Dia As Integer
        For _x = 0 To dha '- 1
            _Dia = Weekday(_Fecha_inicial)

            Select Case _Dias_a_contar
                Case Opcion_Dias.Habiles
                    If _Dia <> "1" And _Dia <> "7" Then
                        dias += 1
                    End If
                Case Opcion_Dias.Todos
                    dias = dha 'dias += 1
                    Exit For
                Case Else
                    If _Dia = _Dias_a_contar Then
                        dias += 1
                    End If
            End Select

            _Fecha_inicial = DateAdd(DateInterval.Day, 1, _Fecha_inicial)

        Next

        Return dias

    End Function

    Function Fx_Validar_Email(ByVal email As String) As Boolean

        If email = String.Empty Then Return False
        ' Compruebo si el formato de la dirección es correcto.
        Dim re As Regex = New Regex("^[\w._%-]+@[\w.-]+\.[a-zA-Z]{2,4}$")
        Dim m As Match = re.Match(email)
        Return (m.Captures.Count <> 0)

    End Function

    Function Fx_Validar_Sitio_Web(ByVal _Sitio As String) As String 'As Boolean

        Dim Peticion As System.Net.WebRequest
        Dim Respuesta As System.Net.HttpWebResponse

        Dim _Respuestas As String

        Try
            Peticion = System.Net.WebRequest.Create(_Sitio) 'La direccion debe tener el formato ('http://www.direccion.com, es, net, org, vns, etc...))
            Respuesta = Peticion.GetResponse()
            _Respuestas = Respuesta.StatusDescription
            ' Return True
        Catch ex As System.Net.WebException
            _Respuestas = ex.Message
            If ex.Status = Net.WebExceptionStatus.NameResolutionFailure Then

                'Return False
            End If
        End Try

        Return _Respuestas

    End Function

    Function Fx_Validar_Impresora(ByVal _Impresora As String) As Boolean

        Dim pd As New PrintDocument

        For i = 1 To PrinterSettings.InstalledPrinters.Count '– 1

            Dim _Impresora_De_Lista = PrinterSettings.InstalledPrinters.Item(i - 1).ToString '_Lista_Impresoras.Items.Item(i - 1).ToString

            If Trim(_Impresora) = Trim(_Impresora_De_Lista) Then
                Return True
            End If

        Next

    End Function

    Function Traer_Numero_Documento(ByVal _TipoDoc As String, _
                                    ByVal _NumeroDoc As String, _
                                    ByVal _Modalidad_Seleccionada As String, _
                                    ByVal _Empresa As String)

        Dim _NrNumeroDoco As String

        Try

            Dim _Existe_Doc As Integer
            Dim _Arr_Nudo(1) As String


            Dim _Sql As New Class_SQL '(_Global_Cadena_De_Conexion_SQL)


            _Modalidad_Seleccionada = Trim(_Modalidad_Seleccionada)

            If String.IsNullOrEmpty(Trim(_NumeroDoc)) Then
                _NrNumeroDoco = _Sql.Fx_Trae_Dato("CONFIEST", _TipoDoc, "EMPRESA = '" & _Empresa &
                                             "' AND MODALIDAD = '" & _Modalidad_Seleccionada & "'") 'FUNCIONARIO & numero_(Trim(Str(CantOCCFuncionario)), 7)
            Else
                _NrNumeroDoco = _NumeroDoc
            End If

            _Existe_Doc = _Sql.Fx_Cuenta_Registros("MAEEDO", "TIDO = '" & _TipoDoc & "' And NUDO = '" & _NrNumeroDoco & "'")

            Dim Continua As Boolean = True
            Dim Contador = 0

            If String.IsNullOrEmpty(Trim(_NrNumeroDoco)) Then

                _NrNumeroDoco = _Sql.Fx_Trae_Dato("CONFIEST", _TipoDoc, "EMPRESA = '" & _Empresa & "' AND MODALIDAD = '  '")
                ' ** REVISA LA CONEXION
                'If _Sql.Pro_Error_Conexion Then Throw New Exception("Error de conexión")
                _Existe_Doc = 0

            ElseIf _NrNumeroDoco = "0000000000" Then

                _NrNumeroDoco = _Sql.Fx_Trae_Dato("MAEEDO", "COALESCE(MAX(NUDO),'0000000000')", "TIDO = '" & _TipoDoc & "'")
                ' ** REVISA LA CONEXION
                'If _Sql.Pro_Error_Conexion Then Throw New Exception("Error de conexión")

                _NrNumeroDoco = Fx_Rellena_ceros(_NrNumeroDoco, 10, True)

                _Existe_Doc = 0

            Else
                Do While CBool(_Existe_Doc)

                    Dim _Proximo_Nro As String = Fx_Proximo_NroDocumento(_NrNumeroDoco)
                    Consulta_sql = "UPDATE CONFIEST SET " & _TipoDoc & " = '" & _Proximo_Nro & "' WHERE EMPRESA = '" & _Empresa &
                                   "' AND  MODALIDAD = '" & _Modalidad_Seleccionada & "'"

                    _Sql.Fx_Ej_consulta_IDU(Consulta_sql)
                    ' ** REVISA LA CONEXION
                    'If _Sql.Pro_Error_Conexion Then Throw New Exception("Error de conexión")

                    _NrNumeroDoco = _Sql.Fx_Trae_Dato("CONFIEST", _TipoDoc, "EMPRESA = '" & _Empresa &
                                             "' AND MODALIDAD = '" & _Modalidad_Seleccionada & "'")
                    ' ** REVISA LA CONEXION
                    'If _Sql.Pro_Error_Conexion Then Throw New Exception("Error de conexión")

                    _Existe_Doc = _Sql.Fx_Cuenta_Registros("MAEEDO", "TIDO = '" & _TipoDoc & "' And NUDO = '" & _NrNumeroDoco & "'")
                    ' ** REVISA LA CONEXION
                    'If _Sql.Pro_Error_Conexion Then Throw New Exception("Error de conexión")
                Loop

            End If


            If CBool(_Existe_Doc) Then

                'If _Mostrar_Mensaje Then
                ' ** REVISA LA CONEXION
                'Throw New Exception("Problema, númeración existente con otra modalidad")
                'MsgBox("", MsgBoxStyle.Critical, "Validación")
                'End If

                _NrNumeroDoco = String.Empty

            End If

        Catch ex As Exception
            _NrNumeroDoco = String.Empty
        End Try

        Return _NrNumeroDoco

    End Function

    Function FechaDelServidor() As Date

        Dim _Sql As New Class_SQL '(_Global_Cadena_De_Conexion_SQL)
        Consulta_sql = "select getdate() As Fecha_Servidor"
        'Dim _RowFecha As DataRow
        Dim Fecha_Servidor As Date = _Sql.Fx_Get_DataRow(Consulta_sql).Item("Fecha_Servidor")

        Return Fecha_Servidor

    End Function

    Function Generar_Filtro_IN(ByVal Tabla As DataTable, _
                               ByVal _CodChk As String, _
                               ByVal _CodCampo As String, _
                               ByVal _EsNumero As Boolean, _
                               ByVal _TieneChk As Boolean, _
                               Optional ByVal _Separador As String = "''")

        Dim Cadena As String = String.Empty
        Dim Separador As String = ""

        If _EsNumero Then
            Separador = "#"
        Else
            Separador = "@"
        End If

        If (Tabla Is Nothing) Then Return "()"

        Dim i = 0
        For Each Rd As DataRow In Tabla.Rows

            Dim Estado As DataRowState = Rd.RowState

            If Estado <> DataRowState.Deleted Then
                Dim _Cadena As String = Rd.Item(_CodCampo).ToString()
                Dim _Encadenar As Boolean = False

                If _TieneChk Then
                    If Rd.Item(_CodChk) Then
                        _Encadenar = True
                    End If
                Else
                    If Not String.IsNullOrEmpty(Trim(_Cadena)) Then _Encadenar = True
                End If

                If _Encadenar Then
                    Cadena = Cadena & Separador & Trim(Rd.Item(_CodCampo).ToString) & Separador '& Coma
                End If
            End If
            i += 1
        Next

        If _EsNumero Then
            Cadena = Replace(Cadena, "##", ",")
            Cadena = Replace(Cadena, "#", "")
        Else
            Cadena = Replace(Cadena, "@@", "@,@")
            Cadena = Replace(Cadena, "@", _Separador)
        End If

        Cadena = "(" & Cadena & ")"

        Return Cadena

    End Function

    
    Function Fx_TraeClaveRD(ByVal Texto As String) As String

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

End Module





