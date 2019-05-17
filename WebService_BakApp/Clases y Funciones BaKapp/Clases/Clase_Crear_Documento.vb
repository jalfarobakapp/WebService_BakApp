Imports System.Data
Imports System.Data.SqlClient
'Imports System.Windows.Forms


Public Class Clase_Crear_Documento

    Dim _Sql As New Class_SQL '(_Global_Cadena_De_Conexion_SQL)

#Region "VARIABLES ENCABEZADO"

    Dim _Global_BaseBk As String
    Dim _Funcionario As String

    Dim _Modalidad As String
    Public _Nudo As String 'NUDO
    Dim _Idmaeedo As Integer

    Dim _Empresa As String
    Dim _Sudo As String         ' Sucursal documento -SUDO

    Dim _Tido As String         ' Tipo de documento TIDO
    Dim _Endo As String         ' Codigo Entidad -ENDO
    Dim _Suendo As String       ' Sucursal Entidad -SUENDO
    Dim _Endofi As String
    Dim _Tigedo As String       ' Tipo de generacion del documento E o I, desde TABTIDO

    Dim _Feemdo As String       ' Fecha emisión - FEEMDO
    Dim _FechaEmision As Date

    Dim _Kofudo As String       ' Responzable del documento
    Dim _Luvtven As String      ' Centro de Costo
    Dim _Caprco As String       ' Cantidad total productos del documento
    Dim _Caprad As String       ' Cantidad despachada Encabezado 
    'Dim _Caprad_Enc As String  ' Cantidad despachada Encabezado
    Dim _Meardo As String       ' Tipo Moneda del documento NETO o BRUTO
    Dim _Modo As String         ' Código moneda del documento
    Dim _Timodo As String       ' Tipo Moneda del documento: Nacional/Extranjera
    Dim _Tamodo As String       ' Valor de la moneda del documento "Dolar"
    Dim _Vaivdo As String       ' Iva 
    Dim _Vaimdo As String
    Dim _Vanedo As String       ' Neto
    Dim _Vabrdo As String       ' Bruto

    Dim _Fe01vedo As String     ' Fecha 1er Vencimiento Fecha_1er_Vencimiento 
    Dim _Feulvedo As String     ' Fecha Ultimo Vencimiento FEULVEDO FechaUltVencimiento

    Dim _Nuvedo As String       ' Numero de vencimientos
    Dim _Feer As String         ' Fecha esperada de recepcion --
    Dim _Subtido As String      ' AJU si es ajuste
    Dim _Marca As String        ' 1 si es ajuste
    Dim _Lisactiva As String    ' Lista de precios del documento

    'TipoDoc
    'NroDocumento
    'Nombre_Entidad
    'Dias_1er_Vencimiento
    'Dias_Vencimiento
    'ListaPrecios
    'CodSucEntidadFisica
    'Nombre_Entidad_Fisica
    'Contacto_Ent
    'CodFuncionario
    'NomFuncionario
    'Centro_Costo
    'Moneda_Doc
    'Valor_Dolar
    'TotalNetoDoc
    'TotalIvaDoc
    'TotalIlaDoc
    'TotalBrutoDoc
    'CantTotal
    'CantDesp
    'DocEn_Neto_Bruto
    'TipoMoneda
#End Region

    Public Sub New(ByVal Global_BaseBk As String, ByVal Funcionario As String)
        _Global_BaseBk = Global_BaseBk
        _Funcionario = Funcionario
    End Sub

#Region "VARIABLES DETALLE DEL DOCUMENTO"

    Dim Id_Linea As Integer

    Dim _Archirst As String
    Dim _Idrst As String

    Dim _Nulido As String
    Dim _Bosulido As String
    Dim _Sulido As String
    Dim _Koprct As String
    Dim _Nokopr As String
    Dim _Rludpr As String
    Dim _Kofulido As String
    Dim _Udtrpr As Integer
    Dim _Ud01pr As String
    Dim _Ud02pr As String

    Dim _Caprco1 As String
    Dim _Caprco2 As String
    Dim _Caprad1 As String
    Dim _Caprad2 As String
    Dim _Caprex1 As String
    Dim _Caprex2 As String
    Dim _Caprnc1 As String
    Dim _Caprnc2 As String
    Dim _Cafaco As String

    Dim _Alternat As String

    Dim _Mopppr As String
    Dim _Timopppr As String
    Dim _Tamopppr As String

    Dim _Koltpr As String
    Dim _Tipr As String
    Dim _Prct As Boolean
    Dim _Tict As String
    Dim _Nudtli As Integer
    Dim _Nuimli As Integer


    Dim _Ppprnelt As String
    Dim _Ppprne As String
    Dim _Ppprbrlt As String
    Dim _Ppprbr As String
    Dim _Ppprnere1 As String
    Dim _Ppprnere2 As String
    Dim _Poimglli As String
    Dim _Vaimli As String
    Dim _Podtglli As String
    Dim _Vadtneli As String
    Dim _Vadtbrli As String
    Dim _Poivli As String
    Dim _Vaivli As String
    Dim _Vaneli As String
    Dim _Vabrli As String

    Dim _Ppprpm As String
    Dim _Ppprmsuc As String
    Dim _Ppprpmifrs As String

    Dim _Feemli As String
    Dim _Feerli As String

    Dim _Operacion As String
    Dim _Potencia As String

    Dim _Eslido As String
    Dim _Lincondesp As Boolean
    Dim _Kofuaulido As String
    Dim _Observa As String

    Dim _Emprepa As String
    Dim _Tidopa As String
    Dim _Nudopa As String
    Dim _Endopa As String
    Dim _Nulidopa As String


#End Region

#Region "VARIABLES PIE DEL DOCUMENTO,OBSERVACIONES"


    Dim _Obdo As String         ' Observacion al documento --
    Dim _Cpdo As String         ' Condiciones de pago del documento
    Dim _Diendesp As String     ' Dirección entidad de despacho 
    Dim _Ocdo As String         ' Orden de compra del documento
    Dim Obs(34) As String       ' Textos del 1 al 35


#End Region

#Region "FUNCION CREAR DOCUMENTO RANDOM DEFINITIVO"

    Function Fx_Crear_Documento(ByVal Tipo_de_documento As String,
                               ByVal Numero_de_documento As String,
                               ByVal _Es_ValeTransitorio As Boolean,
                               ByVal _Es_Documento_Electronico As Boolean,
                               ByVal Bd_Documento As DataSet,
                               Optional ByVal EsAjuste As Boolean = False,
                               Optional ByVal _Cambiar_Numeracion_Confiest As Boolean = True) As String

        Dim cn2 As New SqlConnection

        Dim myTrans As SqlClient.SqlTransaction
        Dim Comando As SqlClient.SqlCommand

        Dim _Row_Encabezado As DataRow = Bd_Documento.Tables("Encabezado_Doc").Rows(0)
        Dim Tbl_Detalle As DataTable = Bd_Documento.Tables("Detalle_Doc")

        Dim _Empresa = _Row_Encabezado.Item("EMPRESA")

        Try

            _Sql.Sb_Abrir_Conexion(cn2)

            With _Row_Encabezado

                Dim _Modalidad As String = .Item("Modalidad")
                _Tido = .Item("TipoDoc")
                _Subtido = .Item("Subtido")

                .Item("NroDocumento") = Numero_de_documento
                _Nudo = Numero_de_documento

                If String.IsNullOrEmpty(Trim(_Nudo)) Then
                    Return 0
                End If

                _Empresa = .Item("Empresa")
                _Sudo = .Item("Sucursal")
                _Kofudo = .Item("CodFuncionario")


                _Endo = .Item("CodEntidad")
                _Suendo = .Item("CodSucEntidad")

                _Feemdo = Format(.Item("FechaEmision"), "yyyyMMdd")
                _Lisactiva = .Item("ListaPrecios")
                _Caprco = De_Num_a_Tx_01(.Item("CantTotal"), 5)
                _Caprad = De_Num_a_Tx_01(.Item("CantDesp"), 5)

                _Luvtven = .Item("Centro_Costo")
                _Modo = .Item("Moneda_Doc")
                _Meardo = .Item("DocEn_Neto_Bruto")
                _Tamodo = De_Num_a_Tx_01(.Item("Valor_Dolar"), False, 5)
                _Timodo = .Item("TipoMoneda")

                Dim _Vanedo_2 = .Item("TotalNetoDoc")
                Dim _Vaivdo_2 = .Item("TotalIvaDoc")
                Dim _Vaimdo_2 = .Item("TotalIlaDoc")
                Dim _Vabrdo_2 = .Item("TotalBrutoDoc")


                _Vanedo = De_Num_a_Tx_01(.Item("TotalNetoDoc"), False, 5)
                _Vaivdo = De_Num_a_Tx_01(.Item("TotalIvaDoc"), False, 5)
                _Vaimdo = De_Num_a_Tx_01(.Item("TotalIlaDoc"), False, 5)
                _Vabrdo = De_Num_a_Tx_01(.Item("TotalBrutoDoc"), False, 0)

                _Fe01vedo = Format(.Item("Fecha_1er_Vencimiento"), "yyyyMMdd")
                _Feulvedo = Format(.Item("FechaUltVencimiento"), "yyyyMMdd")

                _Feer = Format(.Item("FechaRecepcion"), "yyyyMMdd")
                _Feerli = Format(.Item("FechaRecepcion"), "yyyyMMdd")

                '------------------------------------------------------------------------------------------------------------


            End With

            Consulta_sql = "Select Top 1 * From TABTIDO Where TIDO = '" & _Tido & "'"
            Dim _Row_Tabtido As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

            Dim _Signo = String.Empty
            Dim _Fiad As Integer = _Row_Tabtido.Item("FIAD")
            Dim _Fico As Integer = _Row_Tabtido.Item("FICO")
            _Tigedo = _Row_Tabtido.Item("TIGEDO")

            If CBool(_Fico) Then
                If _Fico = 1 Then
                    _Signo = "+"
                ElseIf _Fico = -1 Then
                    _Signo = "-"
                End If
                _Lincondesp = True
            Else
                If _Fiad = 1 Then
                    _Signo = "+"
                ElseIf _Fiad = -1 Then
                    _Signo = "-"
                End If
            End If


            myTrans = cn2.BeginTransaction()

            Consulta_sql = "INSERT INTO MAEEDO ( EMPRESA,TIDO,NUDO,ENDO,SUENDO )" & vbCrLf &
                           "VALUES ( '" & _Empresa & "','" & _Tido & "','" & _Nudo &
                           "','" & _Endo & "','" & _Suendo & "')"

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            Comando = New SqlCommand("SELECT @@IDENTITY AS 'Identity'", cn2)
            Comando.Transaction = myTrans
            Dim dfd1 As SqlDataReader = Comando.ExecuteReader()
            While dfd1.Read()
                _Idmaeedo = dfd1("Identity")
            End While
            dfd1.Close()

            Bd_Documento.Tables("Detalle_Doc").Dispose()


            Dim Contador As Integer = 1

            For Each FDetalle As DataRow In Tbl_Detalle.Rows

                Dim Estado As DataRowState = FDetalle.RowState

                If Not Estado = DataRowState.Deleted Then

                    With FDetalle

                        Id_Linea = .Item("Id")

                        _Nulido = numero_(Contador, 5)

                        _Bosulido = .Item("Bodega")
                        _Koprct = .Item("Codigo")
                        _Nokopr = .Item("Descripcion")
                        _Rludpr = De_Num_a_Tx_01(.Item("Rtu"), False, 5)
                        _Sulido = .Item("Sucursal")
                        _Kofulido = .Item("CodFuncionario") 'FUNCIONARIO ' Codigo de funcionario
                        _Tict = .Item("Tict")
                        _Prct = .Item("Prct")
                        _Caprco1 = De_Num_a_Tx_01(.Item("CantUd1"), False, 5) ' Cantidad de la primera unidad
                        _Caprco2 = De_Num_a_Tx_01(.Item("CantUd2"), False, 5) ' Cantidad de la segunda unidad
                        _Tipr = .Item("Tipr")
                        _Lincondesp = .Item("Lincondest")

                        If _Lincondesp Then
                            _Caprad1 = _Caprco1 ' Cantidad que mueve Stock Fisico, según el producto.
                            _Caprad2 = _Caprco2 ' Cantidad que mueve Stock Fisico, según el producto.
                        Else
                            _Caprad1 = De_Num_a_Tx_01(NuloPorNro(.Item("CDespUd1"), 0), False, 5) ' Cantidad que mueve Stock Fisico, según el producto.
                            _Caprad2 = De_Num_a_Tx_01(NuloPorNro(.Item("CDespUd2"), 0), False, 5) ' Cantidad que mueve Stock Fisico, según el producto.
                        End If

                        _Eslido = NuloPorNro(.Item("Estado"), "")

                        _Caprex1 = 0 ' Cantidad  
                        _Caprex2 = 0
                        _Caprnc1 = 0 ' Cantidad de Nota de credito
                        _Caprnc2 = 0

                        _Udtrpr = .Item("UnTrans")  ' Unidad de la transaccion
                        _Ud01pr = .Item("Ud01PR")
                        _Ud02pr = .Item("Ud02PR")
                        _Koltpr = .Item("CodLista") 'LISTADEPRECIO
                        _Mopppr = .Item("Moneda") 'trae_dato(tb, cn1, "KOMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        _Timopppr = .Item("Tipo_Moneda") 'trae_dato(tb, cn1, "TIMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        _Tamopppr = De_Num_a_Tx_01(.Item("Tipo_Cambio"), False, 5) 'De_Num_a_Tx_01(trae_dato(tb, cn1, "VAMO", "TABMO", "NOKOMO LIKE '%PESO%'"), False, 5)
                        _Ppprnelt = De_Num_a_Tx_01(.Item("PrecioNetoUd"), False, 5)
                        _Podtglli = De_Num_a_Tx_01(.Item("DescuentoPorc"), False, 5)
                        _Poimglli = De_Num_a_Tx_01(.Item("PorIla"), False, 5)

                        _Operacion = .Item("Operacion")
                        _Potencia = De_Num_a_Tx_01(.Item("Potencia"), False, 5)

                        Dim Campo As String = "Precio"

                        _Ppprne = De_Num_a_Tx_01(.Item("PrecioNetoUd"), False, 5)
                        _Ppprbr = De_Num_a_Tx_01(.Item("PrecioBrutoUd"), False, 5)
                        _Ppprnelt = De_Num_a_Tx_01(NuloPorNro(Of Double)(.Item("PrecioNetoUdLista"), 0), False, 5)
                        _Ppprbrlt = De_Num_a_Tx_01(.Item("PrecioBrutoUdLista"), False, 0) ' PRECIO BRUTO DE LA LISTA

                        _Poivli = De_Num_a_Tx_01(.Item("PorIva"), True)
                        _Nudtli = De_Num_a_Tx_01(.Item("NroDscto"), True)

                        _Nuimli = De_Num_a_Tx_01(.Item("NroImpuestos"), True)

                        _Vadtneli = De_Num_a_Tx_01(.Item("DsctoNeto"), False, 5)
                        _Vadtbrli = De_Num_a_Tx_01(.Item("DsctoBruto"), False, 5) 'ValDscto
                        _Vaneli = De_Num_a_Tx_01(.Item("ValNetoLinea"), False, 5)
                        _Vaimli = De_Num_a_Tx_01(.Item("ValIlaLinea"), False, 5)
                        _Vaivli = De_Num_a_Tx_01(.Item("ValIvaLinea"), False, 5)
                        _Vabrli = De_Num_a_Tx_01(Math.Round(.Item("ValBrutoLinea"), 0), False, 5)
                        _Feemli = _Feemdo 'Format(Now.Date, "yyyyMMdd") '""20121127"
                        _Feerli = _Feemdo 'Format(Now.Date, "yyyyMMdd")

                        _Kofuaulido = NuloPorNro(.Item("CodFunAutoriza"), "")
                        _Observa = NuloPorNro(.Item("Observa"), "")

                        _Ppprnere1 = De_Num_a_Tx_01(.Item("PrecioNetoRealUd1"), False, 5)
                        _Ppprnere2 = De_Num_a_Tx_01(.Item("PrecioNetoRealUd2"), False, 5)

                        ' Costos del producto, solo deberia ser efectivo en las ventas
                        _Ppprpm = De_Num_a_Tx_01(NuloPorNro(.Item("PmLinea"), 0), False, 5)
                        _Ppprmsuc = De_Num_a_Tx_01(NuloPorNro(.Item("PmSucLinea"), 0), False, 5)
                        _Ppprpmifrs = De_Num_a_Tx_01(NuloPorNro(.Item("PmIFRS"), 0), False, 5)

                        _Cafaco = 0

                        _Alternat = NuloPorNro(.Item("CodigoProv"), "")

                        Dim _TipoValor As String = NuloPorNro(.Item("TipoValor"), "")



                        If _Prct Then ' ES CONCEPTO

                            If Not String.IsNullOrEmpty(Trim(_Tict)) Then

                                Dim TipoValor = _TipoValor

                                _Caprco2 = 0
                                _Caprad2 = 0
                                _Cafaco = 0
                                _Ppprnelt = 0
                                _Ppprne = 0
                                _Ppprbrlt = 0
                                _Ppprbr = 0
                                _Prct = 1
                                _Ppprpm = 0
                                _Ppprmsuc = 0
                                _Ppprpmifrs = 0
                                _Lincondesp = 1
                                _Nudtli = 0
                                _Eslido = "C"
                                _Lincondesp = True

                            End If

                            _Idrst = 0

                        Else

                            If _Tido <> "COV" Then

                                If _Tido = "OCC" Then

                                    Consulta_sql = "UPDATE MAEST SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," &
                                                                    "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf &
                                                                    "WHERE EMPRESA='" & _Empresa &
                                                                    "' AND KOSU='" & _Sulido &
                                                                    "' AND KOBO='" & _Bosulido &
                                                                    "' AND KOPR='" & _Koprct & "'" & vbCrLf &
                                                   "UPDATE MAEPREM SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," &
                                                                    "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf &
                                                                    "WHERE EMPRESA='" & _Empresa &
                                                                    "' AND KOPR='" & _Koprct & "'" & vbCrLf &
                                                   "UPDATE MAEPR SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," &
                                                                    "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf &
                                                                    "WHERE KOPR='" & _Koprct & "'"

                                    Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                    Comando.Transaction = myTrans
                                    Comando.ExecuteNonQuery()

                                ElseIf _Tido = "NVV" Then

                                    Consulta_sql = "UPDATE MAEST SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," &
                                                                   "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf &
                                                                   "WHERE EMPRESA='" & _Empresa &
                                                                   "' AND KOSU='" & _Sulido &
                                                                   "' AND KOBO='" & _Bosulido &
                                                                   "' AND KOPR='" & _Koprct & "'" & vbCrLf &
                                                  "UPDATE MAEPREM SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," &
                                                                   "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf &
                                                                   "WHERE EMPRESA='" & _Empresa &
                                                                   "' AND KOPR='" & _Koprct & "'" & vbCrLf &
                                                  "UPDATE MAEPR SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," &
                                                                   "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf &
                                                                   "WHERE KOPR='" & _Koprct & "'"

                                    Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                    Comando.Transaction = myTrans
                                    Comando.ExecuteNonQuery()

                                Else

                                    If _Lincondesp Then

                                        Consulta_sql = "UPDATE MAEPREM SET" & vbCrLf &
                                                       "STFI1 = STFI1 " & _Signo & " " & _Caprco1 & ",STFI2 = STFI2 " & _Signo & " " & _Caprco2 & vbCrLf &
                                                       "WHERE EMPRESA = '" & _Empresa & "' AND KOPR = '" & _Koprct & "'" &
                                                       vbCrLf &
                                                       vbCrLf &
                                                       "UPDATE MAEPR SET  STFI1 = STFI1 " & _Signo & " " & _Caprco1 & ",STFI2 = STFI2 " & _Signo & " " & _Caprco2 & vbCrLf &
                                                       "WHERE KOPR = '" & _Koprct & "'" &
                                                       vbCrLf &
                                                       vbCrLf &
                                                       "UPDATE MAEST SET STFI1 = STFI1 " & _Signo & " " & _Caprco1 & ",STFI2 = STFI2 " & _Signo & " " & _Caprco2 & vbCrLf &
                                                       "WHERE EMPRESA='" & _Empresa & "' AND KOSU='" & _Sulido &
                                                       "' AND KOBO='" & _Bosulido & "' AND KOPR='" & _Koprct & "'" &
                                                       vbCrLf &
                                                       vbCrLf &
                                                       "UPDATE MAEPMSUC SET STFI1 = STFI1 " & _Signo & " 1,STFI2 = STFI2 " & _Signo & " 1" & vbCrLf &
                                                       "WHERE EMPRESA = '" & _Empresa & "' AND KOSU = '" & _Sulido & "' AND KOPR = '" & _Koprct & "'"

                                    Else

                                        Consulta_sql = "UPDATE MAEPREM SET" & vbCrLf &
                                                       "STDV1 = STDV1 + " & _Caprco1 & ",STDV2 = STDV2 + " & _Caprco2 & vbCrLf &
                                                       "WHERE EMPRESA = '" & _Empresa & "' AND KOPR = '" & _Koprct & "'" & vbCrLf &
                                                       "UPDATE MAEPR SET  STDV1 = STDV1 + " & _Caprco1 & ",STDV2 = STDV2 + " & _Caprco2 &
                                                       vbCrLf &
                                                       "WHERE KOPR = '" & _Koprct & "'" & vbCrLf &
                                                       "UPDATE MAEST SET STDV1 = STDV1 + " & _Caprco1 & ",STDV2 = STDV2 + " & _Caprco2 &
                                                       vbCrLf &
                                                       "WHERE EMPRESA='" & _Empresa & "' AND KOSU='" & _Sudo &
                                                       "' AND KOBO='" & _Bosulido & "' AND KOPR='" & _Koprct & "'"

                                        _Caprad1 = 0
                                        _Caprad2 = 0


                                    End If

                                    Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                    Comando.Transaction = myTrans
                                    Comando.ExecuteNonQuery()

                                End If

                                _Idrst = Val(NuloPorNro(.Item("Idmaeddo_Dori"), ""))

                                'EMPREPA,TIDOPA,NUDOPA,ENDOPA,NULIDOPA

                                If CBool(_Idrst) Then

                                    Consulta_sql = "Select Top 1 * From MAEDDO Where IDMAEDDO = " & _Idrst

                                    Dim _Tbl_Doc_Origen As DataTable = _Sql.Fx_Get_Tablas(Consulta_sql)

                                    _Emprepa = _Tbl_Doc_Origen.Rows(0).Item("EMPRESA")
                                    _Tidopa = _Tbl_Doc_Origen.Rows(0).Item("TIDO")
                                    _Nudopa = _Tbl_Doc_Origen.Rows(0).Item("NUDO")
                                    _Endopa = _Tbl_Doc_Origen.Rows(0).Item("ENDO")
                                    _Nulidopa = _Tbl_Doc_Origen.Rows(0).Item("NULIDO")

                                    Dim _Caprnc1_Ori As Double = _Tbl_Doc_Origen.Rows(0).Item("CAPRNC1")
                                    Dim _Caprnc2_Ori As Double = _Tbl_Doc_Origen.Rows(0).Item("CAPRNC2")
                                    Dim _Caprex1_Ori As Double = _Tbl_Doc_Origen.Rows(0).Item("CAPREX1")
                                    Dim _Caprex2_Ori As Double = _Tbl_Doc_Origen.Rows(0).Item("CAPREX2")

                                    Dim _CantUd1_Dori As Double = .Item("CantUd1_Dori")
                                    Dim _CantUd2_Dori As Double = .Item("CantUd2_Dori")

                                    Dim _Cant_MovUd1_Ext As String
                                    Dim _Cant_MovUd2_Ext As String

                                    Dim _CantUd1 As Double = .Item("CantUd1")
                                    Dim _CantUd2 As Double = .Item("CantUd2")

                                    If _CantUd1_Dori < _CantUd1 Then
                                        _Cant_MovUd1_Ext = De_Num_a_Tx_01(_CantUd1_Dori, False, 5)
                                        _Cant_MovUd2_Ext = De_Num_a_Tx_01(_CantUd2_Dori, False, 5)
                                    Else
                                        _Cant_MovUd1_Ext = _CantUd1
                                        _Cant_MovUd2_Ext = _CantUd2
                                    End If

                                    _Archirst = "MAEDDO"

                                    If _Tido = "NCC" Or _Tido = "NCV" Then

                                        Consulta_sql = "UPDATE MAEDDO SET CAPRNC1=CAPRNC1+" & _Cant_MovUd1_Ext &
                                                                        ",CAPRNC2=CAPRNC2+" & _Cant_MovUd2_Ext & "," &
                                                       "'ESLIDO = " & vbCrLf &
                                                       "CASE" & vbCrLf &
                                                       "'WHEN UDTRPR='1' AND CAPRCO1-CAPRAD1-CAPREX1=0 THEN 'C'" & vbCrLf &
                                                       "'WHEN UDTRPR='2' AND CAPRCO2-CAPRAD2-CAPREX2=0 THEN 'C'" & vbCrLf &
                                                       "'ELSE ''" & vbCrLf &
                                                       "END" & vbCrLf &
                                                       "WHERE IDMAEDDO = " & _Idrst

                                    Else

                                        Consulta_sql = "UPDATE MAEDDO SET CAPREX1=CAPREX1+" & _Cant_MovUd1_Ext &
                                                                        ",CAPREX2=CAPREX2+" & _Cant_MovUd2_Ext & "," &
                                                       "ESLIDO = " &
                                                       "CASE" & vbCrLf &
                                                       "WHEN UDTRPR='1' AND " &
                                                       "ROUND(CAPRCO1-CAPRAD1-(CAPREX1+" & _Cant_MovUd1_Ext & "),5)=0 THEN 'C'" & vbCrLf &
                                                       "WHEN UDTRPR='2' AND " &
                                                       "ROUND(CAPRCO2-CAPRAD2-(CAPREX2+" & _Cant_MovUd2_Ext & "),5)=0 THEN 'C'" & vbCrLf &
                                                       "Else ''" & vbCrLf &
                                                       "End" & vbCrLf &
                                                       "WHERE IDMAEDDO = " & _Idrst  '1398920

                                    End If

                                    Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                    Comando.Transaction = myTrans
                                    Comando.ExecuteNonQuery()

                                End If

                                'Else
                                '_Idrst = 0
                            End If

                        End If

                        Consulta_sql =
                              "INSERT INTO MAEDDO (IDMAEEDO,ARCHIRST,IDRST,EMPRESA,TIDO,NUDO,ENDO,SUENDO,LILG,NULIDO," & vbCrLf &
                              "SULIDO,BOSULIDO,LUVTLIDO,KOFULIDO,TIPR,TICT,PRCT,KOPRCT,UDTRPR,RLUDPR,CAPRCO1," & vbCrLf &
                              "UD01PR,CAPRCO2,UD02PR,CAPRAD1,CAPRAD2,KOLTPR,MOPPPR,TIMOPPPR,TAMOPPPR,NUIMLI,NUDTLI," & vbCrLf &
                              "PODTGLLI,POIMGLLI,VAIMLI,VADTNELI,VADTBRLI,POIVLI,VAIVLI,VANELI,VABRLI,TIGELI," & vbCrLf &
                              "EMPREPA,TIDOPA,NUDOPA,ENDOPA,NULIDOPA," & vbCrLf &
                              "FEEMLI,FEERLI,PPPRNELT,PPPRNE,PPPRBRLT,PPPRBR,PPPRPM,PPPRNERE1,PPPRNERE2,CAFACO," & vbCrLf &
                              "FVENLOTE,FCRELOTE,NOKOPR,ALTERNAT,TASADORIG,CUOGASDIF,LINCONDESP,OPERACION,POTENCIA,ESLIDO,OBSERVA,KOFUAULIDO)" & vbCrLf &
                       "VALUES (" & _Idmaeedo & ",'" & _Archirst & "'," & _Idrst & ",'" & _Empresa & "','" & _Tido & "','" & _Nudo & "','" & _Endo &
                              "','" & _Suendo & "','SI','" & _Nulido & "','" & _Sulido & "','" & _Bosulido &
                              "','" & _Luvtven & "','" & _Kofulido & "','" & _Tipr & "','" & _Tict & "'," & CInt(_Prct) & ",'" & _Koprct &
                              "'," & _Udtrpr & "," & _Rludpr & "," & _Caprco1 & ",'" & _Ud01pr & "'," & _Caprco2 &
                              ",'" & _Ud02pr & "'," & _Caprad1 & "," & _Caprad2 & ",'TABPP" & _Koltpr & "'" &
                              ",'" & _Mopppr & "','" & _Timopppr & "'," & _Tamopppr &
                              "," & _Nuimli & "," & _Nudtli & "," & _Podtglli & "," & _Poimglli & "," & _Vaimli &
                              "," & _Vadtneli & "," & _Vadtbrli & "," & _Poivli & "," & _Vaivli & "," & _Vaneli &
                              "," & _Vabrli & ",'I'," &
                              "'" & _Emprepa & "','" & _Tidopa & "','" & _Nudopa & "','" & _Endopa & "','" & _Nulidopa & "'," &
                              "'" & _Feemli & "','" & _Feerli & "'," & _Ppprnelt & "," & _Ppprne &
                              "," & _Ppprbrlt & "," & _Ppprbr & "," & _Ppprpm & "," & _Ppprnere1 & "," & _Ppprnere2 &
                              "," & _Cafaco & ",NULL,NULL,'" & _Nokopr & "','" & _Alternat & "',1.00000,0," & CInt(_Lincondesp) * -1 &
                              ",'" & _Operacion & "'," & _Potencia & ",'" & _Eslido & "','" & _Observa & "','" & _Kofuaulido & "')"

                        Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                        Comando.Transaction = myTrans
                        Comando.ExecuteNonQuery()


                        Comando = New SqlCommand("SELECT @@IDENTITY AS 'Identity'", cn2)
                        Comando.Transaction = myTrans
                        dfd1 = Comando.ExecuteReader()
                        Dim _Idmaeddo As Integer
                        While dfd1.Read()
                            _Idmaeddo = dfd1("Identity")
                        End While
                        dfd1.Close()

                        ' *** PM POR SUCURSAL SI ES QUE EXISTE EL CAMPO
                        Dim _Reg As Integer = _Sql.Fx_Cuenta_Registros("INFORMATION_SCHEMA.COLUMNS",
                                                              "COLUMN_NAME LIKE 'PPPRPMSUC' AND TABLE_NAME = 'MAEDDO'")

                        If CBool(_Reg) Then

                            Consulta_sql = "UPDATE MAEDDO SET PPPRPMSUC = " & _Ppprmsuc & vbCrLf &
                                           "WHERE IDMAEDDO = " & _Idmaeddo

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()
                        End If
                        '*************************************************************************************************

                        ' *** PMIFRS SI ES QUE EXISTE EL CAMPO
                        _Reg = _Sql.Fx_Cuenta_Registros("INFORMATION_SCHEMA.COLUMNS",
                                                         "COLUMN_NAME LIKE 'PMIFRS' AND TABLE_NAME = 'MAEPREM'")

                        If CBool(_Reg) Then

                            Consulta_sql = "UPDATE MAEDDO SET PPPRPMIFRS = " & _Ppprpmifrs & vbCrLf &
                                           "WHERE IDMAEDDO=" & _Idmaeddo

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()
                        End If
                        '*************************************************************************************************

                    End With


                    ' TABLA DE IMPUESTOS

                    Dim Tbl_Impuestos As DataTable = Bd_Documento.Tables("Impuestos_Doc")

                    If Val(_Nuimli) > 0 And String.IsNullOrEmpty(Trim(_Tict)) Then

                        For Each FImpto As DataRow In Tbl_Impuestos.Select("Id = " & Id_Linea)

                            Dim _Poimli As String = De_Num_a_Tx_01(FImpto.Item("Poimli").ToString, False, 5)
                            Dim _Koimli As String = FImpto.Item("Koimli").ToString
                            Dim _Vaimli = De_Num_a_Tx_01(FImpto.Item("Vaimli").ToString, False, 5)

                            Consulta_sql = "INSERT INTO MAEIMLI(IDMAEEDO,NULIDO,KOIMLI,POIMLI,VAIMLI,LILG) VALUES " & vbCrLf &
                                           "(" & _Idmaeedo & ",'" & _Nulido & "','" & _Koimli & "'," & _Poimli & "," & _Vaimli & ",'')"

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                            '-- 3RA TRANSACCION--------------------------------------------------------------------
                        Next
                    End If



                    ' TABLA DE DESCUENTOS
                    Dim Tbl_Descuentos As DataTable = Bd_Documento.Tables("Descuentos_Doc")
                    _Nudtli = Tbl_Descuentos.Rows.Count
                    If Val(_Nudtli) > 0 And String.IsNullOrEmpty(Trim(_Tict)) Then

                        For Each FDscto As DataRow In Tbl_Descuentos.Select("Id = " & Id_Linea)

                            Dim _Podt = De_Num_a_Tx_01(FDscto.Item("Podt").ToString, False, 5)
                            Dim _Vadt = De_Num_a_Tx_01(FDscto.Item("Vadt").ToString, False, 5)

                            Consulta_sql = "INSERT INTO MAEDTLI (IDMAEEDO,NULIDO,KODT,PODT,VADT)" & vbCrLf &
                                   "values (" & _Idmaeedo & ",'" & _Nulido & "','D_SIN_TIPO'," & _Podt & "," & _Vadt & ")"


                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                            '-- 4TA TRANSACCION--------------------------------------------------------------------
                        Next
                    End If

                    Contador += 1
                End If
            Next


            'TABLA DE VENCIMIENTOS

            Consulta_sql = Fx_Vencimientos(_Row_Encabezado)

            If Not String.IsNullOrEmpty(Consulta_sql) Then

                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                Comando.Transaction = myTrans
                Comando.ExecuteNonQuery()

            End If

            If _Nuvedo = 0 Then _Nuvedo = 1

            Dim _HoraGrab As String

            'Dim _HH_sistem As Date

            '_HH_sistem = FechaDelServidor()
            '_HoraGrab = _HH_sistem.Hour

            'Dim _HH, _MM, _SS As Double

            '_HH = _HH_sistem.Hour
            '_MM = _HH_sistem.Minute
            '_SS = _HH_sistem.Second

            If EsAjuste Then
                _Marca = 1 ' Generalmente se marcan las GRI o GDI que son por ajuste
                _Subtido = "AJU" ' Generalmente se Marcan las Guias de despacho o recibo
                '_HH = 23 : _MM = 59 : _SS = 59
            Else
                _Marca = String.Empty
                _Subtido = String.Empty
            End If



            _HoraGrab = Hora_Grab_fx(EsAjuste, FechaDelServidor) 'Math.Round((_HH * 3600) + (_MM * 60) + _SS, 0)


            'Consulta_sql = "Declare @HoraGrab Int" & vbCrLf & _
            '               "set @HoraGrab = convert(money,substring(convert(varchar(20),getdate(),114),1,2)) * 3600 +" & vbCrLf & _
            '               "convert(money,substring(convert(varchar(20),getdate(),114),4,2)) * 60 + " & vbCrLf & _
            '               "Convert(money, substring(Convert(varchar(20), getdate(), 114), 7, 2))" & vbCrLf & vbCrLf & _



            Dim _Espgdo As String

            Select Case _Tido
                Case "COV", "GAR", "GDD", "GDI", "GDP", "GDV", "GRC", "GRD", "GRI", "GRP", "GTI", "NVV", "OCC"
                    _Espgdo = "S"
                Case Else
                    _Espgdo = "P"
            End Select

            ' HAY QUE PONER EL CAMPO TIPO DE MONEDA  "TIMODO"
            Consulta_sql = "UPDATE MAEEDO SET SUENDO='" & _Suendo & "',TIGEDO='I',SUDO='" & _Sudo &
                          "',FEEMDO='" & _Feemdo & "',KOFUDO='" & _Kofudo & "',ESPGDO='" & _Espgdo & "',CAPRCO=" & _Caprco &
                          ",CAPRAD=" & _Caprad & ",MEARDO = '" & _Meardo & "',MODO = '" & _Modo &
                          "',TIMODO = '" & _Timodo & "',TAMODO = " & _Tamodo & ",VAIVDO = " & _Vaivdo & ",VAIMDO = " & _Vaimdo & vbCrLf &
                          ",VANEDO = " & _Vanedo & ",VABRDO = " & _Vabrdo & ",FE01VEDO = '" & _Fe01vedo &
                          "',FEULVEDO = '" & _Feulvedo & "',NUVEDO = " & _Nuvedo & ",FEER = '" & _Feer &
                          "',KOTU = '1',LCLV = NULL,LAHORA = GETDATE(), DESPACHO = 1,HORAGRAB = " & _HoraGrab &
                          ",FECHATRIB = NULL,SUBTIDO = '" & _Subtido &
                          "',MARCA = '" & _Marca & "',ESDO = '',NUDONODEFI = " & CInt(_Es_ValeTransitorio) &
                          ",TIDOELEC = " & CInt(_Es_Documento_Electronico) & ",LUVTDO = '" & _Luvtven & "'" & vbCrLf &
                          "WHERE IDMAEEDO=" & _Idmaeedo

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()


            Dim Reg As Integer = _Sql.Fx_Cuenta_Registros("INFORMATION_SCHEMA.COLUMNS",
                                                         "COLUMN_NAME LIKE 'LISACTIVA' AND TABLE_NAME = 'MAEEDO'")

            If CBool(Reg) Then

                Consulta_sql = "UPDATE MAEEDO SET LISACTIVA = 'TABPP" & _Lisactiva & "'" & vbCrLf &
                               "WHERE IDMAEEDO=" & _Idmaeedo

                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                Comando.Transaction = myTrans
                Comando.ExecuteNonQuery()

            End If


            '========================================== CERRAR DOCUMENTOS ASOCIADOS ============================================
            If _Tido <> "COV" Then

                Dim Fl As String = Generar_Filtro_IN(Tbl_Detalle, "", "Idmaeedo_Dori", False, False, "")

                If Fl = "()" Then Fl = "(0)"

                Consulta_sql = "SELECT DISTINCT IDMAEEDO FROM MAEDDO WHERE IDMAEEDO IN " & Fl
                Dim _TblOrigen As DataTable = _Sql.Fx_Get_Tablas(Consulta_sql)

                'Idmaeedo_Dori

                If CBool(_TblOrigen.Rows.Count) Then

                    Dim _Sum_Caprco As Double

                    For Each _Fila_Idmaeedo As DataRow In _TblOrigen.Rows

                        _Sum_Caprco = 0

                        Dim _Idmaeedo_Origen = _Fila_Idmaeedo.Item("IDMAEEDO")

                        For Each _Fila As DataRow In Tbl_Detalle.Rows

                            Dim Idmaeedo_Dori = _Fila.Item("Idmaeedo_Dori")

                            If _Idmaeedo_Origen = Idmaeedo_Dori Then

                                Dim _Idrst = Val(_Fila.Item("Idmaeddo_Dori"))

                                If CBool(_Idrst) Then

                                    Dim _CantUd1_Dori As Double = _Fila.Item("CantUd1_Dori")
                                    Dim _CantUd2_Dori As Double = _Fila.Item("CantUd2_Dori")

                                    Dim _Cant_MovUd1_Ext As String
                                    Dim _Cant_MovUd2_Ext As String

                                    Dim _CantUd1 As Double = _Fila.Item("CantUd1")
                                    Dim _CantUd2 As Double = _Fila.Item("CantUd2")

                                    If _CantUd1_Dori < _CantUd1 Then
                                        _Cant_MovUd1_Ext = _CantUd1_Dori
                                        _Cant_MovUd2_Ext = _CantUd2_Dori
                                    Else
                                        _Cant_MovUd1_Ext = _CantUd1
                                        _Cant_MovUd2_Ext = _CantUd2
                                    End If

                                    _Sum_Caprco += _Cant_MovUd1_Ext '_Fila.Item("CantUd1") 'De_Num_a_Tx_01(_Fila.Item("CantUd1"), False, 5) ' Cantidad de la primera unidad

                                End If

                            End If

                        Next

                        If CBool(_Sum_Caprco) Then

                            Dim _Sum_Caprco_str As String = De_Num_a_Tx_01(_Sum_Caprco, False, 5)

                            If _Tido = "NCV" Or _Tido = "NCC" Then

                                Consulta_sql = "UPDATE MAEEDO SET CAPREX=CAPREX+0,CAPRNC=CAPRNC+" & _Sum_Caprco_str & ",CAPRAD=CAPRAD+0," & vbCrLf &
                                               "ESDO=CASE" & vbCrLf &
                                               "WHEN ROUND(CAPRCO-CAPRAD-CAPREX-(0)-(0),5)=0 THEN 'C'" & vbCrLf & "ELSE ''" & vbCrLf & "END," & vbCrLf &
                                               "ESFADO=CASE" & vbCrLf &
                                               "WHEN CAPRCO-CAPRAD-CAPREX-(0)-(0)=0 THEN 'F'" & vbCrLf & "ELSE ESFADO" & vbCrLf & "END" & vbCrLf &
                                               "WHERE IDMAEEDO= " & _Idmaeedo_Origen
                            Else
                                Consulta_sql = "UPDATE MAEEDO SET CAPREX=CAPREX+" & _Sum_Caprco_str & ",CAPRNC=CAPRNC+0,CAPRAD=CAPRAD+0," & vbCrLf &
                                               "ESDO=CASE " & vbCrLf &
                                               "WHEN ROUND(CAPRCO-CAPRAD-CAPREX-(0)-(" & _Sum_Caprco_str & "),5)=0 THEN 'C'" & vbCrLf & "ELSE ''" & vbCrLf & "END," & vbCrLf &
                                               "ESFADO=" & vbCrLf &
                                               "CASE" & vbCrLf &
                                               "WHEN CAPRCO-CAPRAD-CAPREX-(0)-(" & _Sum_Caprco_str & ")=0 THEN 'F'  " & vbCrLf & "ELSE ESFADO" & vbCrLf & "End" & vbCrLf &
                                               "WHERE IDMAEEDO = " & _Idmaeedo_Origen
                            End If

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                        End If

                    Next

                End If

            End If


            '=========================================== OBSERVACIONES ==================================================================

            Dim Tbl_Observaciones As DataTable = Bd_Documento.Tables("Observaciones_Doc")

            With Tbl_Observaciones

                _Obdo = .Rows(0).Item("Observaciones")
                _Cpdo = .Rows(0).Item("Forma_pago")
                _Ocdo = .Rows(0).Item("Orden_compra")

                For i = 0 To 34

                    Dim Campo As String = "Obs" & i + 1
                    Obs(i) = .Rows(0).Item(Campo)

                Next

            End With

            Consulta_sql = "INSERT INTO MAEEDOOB (IDMAEEDO,OBDO,CPDO,OCDO,DIENDESP,TEXTO1,TEXTO2,TEXTO3,TEXTO4,TEXTO5,TEXTO6," & vbCrLf &
                           "TEXTO7,TEXTO8,TEXTO9,TEXTO10,TEXTO11,TEXTO12,TEXTO13,TEXTO14,TEXTO15,CARRIER,BOOKING,LADING,AGENTE," & vbCrLf &
                           "MEDIOPAGO,TIPOTRANS,KOPAE,KOCIE,KOCME,FECHAE,HORAE,KOPAD,KOCID,KOCMD,FECHAD,HORAD,OBDOEXPO,MOTIVO," & vbCrLf &
                           "TEXTO16,TEXTO17,TEXTO18,TEXTO19,TEXTO20,TEXTO21,TEXTO22,TEXTO23,TEXTO24,TEXTO25,TEXTO26,TEXTO27," & vbCrLf &
                           "TEXTO28,TEXTO29,TEXTO30,TEXTO31,TEXTO32,TEXTO33,TEXTO34,TEXTO35) VALUES " & vbCrLf &
                           "(" & _Idmaeedo & ",'" & _Obdo & "','" & _Cpdo & "','" & _Ocdo & "','','" & Obs(0) & "','" & Obs(1) &
                           "','" & Obs(2) & "','" & Obs(3) & "','" & Obs(4) & "','" & Obs(5) & "','" & Obs(6) & "','" & Obs(7) &
                           "','" & Obs(8) & "','" & Obs(9) & "','" & Obs(10) & "','" & Obs(11) & "','" & Obs(12) & "','" & Obs(13) &
                           "','" & Obs(14) & "','','','','','','','','','',GETDATE(),'','','','',GETDATE(),'','','','" & Obs(15) &
                           "','" & Obs(16) & "','" & Obs(17) & "','" & Obs(18) & "','" & Obs(19) &
                           "','" & Obs(20) & "','" & Obs(21) & "','" & Obs(22) & "','" & Obs(23) & "','" & Obs(24) &
                           "','" & Obs(25) & "','" & Obs(26) & "','" & Obs(27) & "','" & Obs(28) & "','" & Obs(29) &
                           "','" & Obs(30) & "','" & Obs(31) & "','" & Obs(32) & "','" & Obs(33) & "','" & Obs(34) & "')"

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            ' ====================================================================================================================================

            If _Cambiar_Numeracion_Confiest Then
                Consulta_sql = Fx_Cambiar_Numeracion_Modalidad(_Tido, _Nudo, _Modalidad, _Empresa)

                If Not String.IsNullOrEmpty(Consulta_sql) Then

                    Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                    Comando.Transaction = myTrans
                    Comando.ExecuteNonQuery()

                End If
            End If

            myTrans.Commit()

            Return _Idmaeedo

        Catch ex As Exception

            Dim _Erro_VB As String = ex.Message & vbCrLf & ex.StackTrace & vbCrLf &
                                     "Código: " & _Koprct
            'Clipboard.SetText(_Erro_VB)

            My.Computer.FileSystem.WriteAllText("ArchivoSalida", ex.Message & vbCrLf & ex.StackTrace, False)
            'MessageBoxEx.Show(ex.Message, "Error", _
            '                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
            myTrans.Rollback()
            MsgBox("Transacción desecha", MsgBoxStyle.Critical, "BakApp")
            'MessageBoxEx.Show("Transaccion desecha", "Problema", _
            '                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
            'SQL_ServerClass.CerrarConexion(cn2)
            Return _Erro_VB '0
        Finally
            _Sql.Sb_Cerrar_Conexion(cn2)
        End Try

    End Function

    Function Fx_Crear_Documento_Old(ByVal Tipo_de_documento As String,
                               ByVal Numero_de_documento As String,
                               ByVal _Es_ValeTransitorio As Boolean,
                               ByVal _Es_Documento_Electronico As Boolean,
                               ByVal Bd_Documento As DataSet,
                               Optional ByVal EsAjuste As Boolean = False,
                               Optional ByVal _Cambiar_Numeracion_Confiest As Boolean = True) As String

        Dim cn2 As New SqlConnection

        Dim myTrans As SqlClient.SqlTransaction
        Dim Comando As SqlClient.SqlCommand

        Dim _Row_Encabezado As DataRow = Bd_Documento.Tables("Encabezado_Doc").Rows(0)


        Try

            _Sql.Sb_Abrir_Conexion(cn2)

            With _Row_Encabezado

                Dim _Modalidad As String = .Item("Modalidad")
                _Tido = .Item("TipoDoc")

                .Item("NroDocumento") = Numero_de_documento
                _Nudo = .Item("NroDocumento")

                If String.IsNullOrEmpty(Trim(_Nudo)) Then
                    Return 0
                End If

                _Empresa = .Item("Empresa").ToString
                _Sudo = .Item("Sucursal")
                _Kofudo = .Item("CodFuncionario")


                _Endo = .Item("CodEntidad")
                _Suendo = .Item("CodSucEntidad")

                _Feemdo = Format(.Item("FechaEmision"), "yyyyMMdd")
                _Lisactiva = .Item("ListaPrecios")
                _Caprco = De_Num_a_Tx_01(.Item("CantTotal"), 5)
                _Caprad = De_Num_a_Tx_01(.Item("CantDesp"), 5)

                _Luvtven = .Item("Centro_Costo")
                _Modo = .Item("Moneda_Doc")
                _Meardo = .Item("DocEn_Neto_Bruto")
                _Tamodo = De_Num_a_Tx_01(.Item("Valor_Dolar"), False, 5)
                _Timodo = .Item("TipoMoneda")

                Dim _Vanedo_2 = .Item("TotalNetoDoc")
                Dim _Vaivdo_2 = .Item("TotalIvaDoc")
                Dim _Vaimdo_2 = .Item("TotalIlaDoc")
                Dim _Vabrdo_2 = .Item("TotalBrutoDoc")


                _Vanedo = De_Num_a_Tx_01(.Item("TotalNetoDoc"), False, 5)
                _Vaivdo = De_Num_a_Tx_01(.Item("TotalIvaDoc"), False, 5)
                _Vaimdo = De_Num_a_Tx_01(.Item("TotalIlaDoc"), False, 5)
                _Vabrdo = De_Num_a_Tx_01(.Item("TotalBrutoDoc"), False, 0)

                _Fe01vedo = Format(.Item("Fecha_1er_Vencimiento"), "yyyyMMdd")
                _Feulvedo = Format(.Item("FechaUltVencimiento"), "yyyyMMdd")

                _Feer = Format(.Item("FechaRecepcion"), "yyyyMMdd")
                _Feerli = Format(.Item("FechaRecepcion"), "yyyyMMdd")

                '------------------------------------------------------------------------------------------------------------


            End With


            myTrans = cn2.BeginTransaction()

            Consulta_sql = "INSERT INTO MAEEDO ( EMPRESA,TIDO,NUDO,ENDO,SUENDO )" & vbCrLf &
                           "VALUES ( '" & _Empresa & "','" & _Tido & "','" & _Nudo &
                           "','" & _Endo & "','" & _Suendo & "')"

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            Comando = New SqlCommand("SELECT @@IDENTITY AS 'Identity'", cn2)
            Comando.Transaction = myTrans
            Dim dfd1 As SqlDataReader = Comando.ExecuteReader()
            While dfd1.Read()
                _Idmaeedo = dfd1("Identity")
            End While
            dfd1.Close()

            Bd_Documento.Tables("Detalle_Doc").Dispose()
            Dim Tbl_Detalle As DataTable = Bd_Documento.Tables("Detalle_Doc")

            Dim Contador As Integer = 1

            For Each FDetalle As DataRow In Tbl_Detalle.Rows

                Dim Estado As DataRowState = FDetalle.RowState

                If Not Estado = DataRowState.Deleted Then

                    With FDetalle



                        Id_Linea = .Item("Id")


                        _Nulido = numero_(Contador, 5)

                        _Bosulido = .Item("Bodega")
                        _Koprct = .Item("Codigo")
                        _Nokopr = .Item("Descripcion")
                        _Rludpr = De_Num_a_Tx_01(.Item("Rtu"), False, 5)
                        _Sulido = .Item("Sucursal")
                        _Kofulido = .Item("CodFuncionario") 'FUNCIONARIO ' Codigo de funcionario
                        _Tict = .Item("Tict")
                        _Prct = .Item("Prct")
                        _Caprco1 = De_Num_a_Tx_01(.Item("CantUd1"), False, 5) ' Cantidad de la primera unidad
                        _Caprco2 = De_Num_a_Tx_01(.Item("CantUd2"), False, 5) ' Cantidad de la segunda unidad
                        _Tipr = .Item("Tipr")
                        _Lincondesp = .Item("Lincondest")

                        If _Lincondesp Then
                            _Caprad1 = _Caprco1 ' Cantidad que mueve Stock Fisico, según el producto.
                            _Caprad2 = _Caprco2 ' Cantidad que mueve Stock Fisico, según el producto.
                        Else
                            _Caprad1 = De_Num_a_Tx_01(NuloPorNro(.Item("CDespUd1"), 0), False, 5) ' Cantidad que mueve Stock Fisico, según el producto.
                            _Caprad2 = De_Num_a_Tx_01(NuloPorNro(.Item("CDespUd2"), 0), False, 5) ' Cantidad que mueve Stock Fisico, según el producto.
                        End If

                        _Eslido = NuloPorNro(.Item("Estado"), "")

                        _Caprex1 = 0 ' Cantidad  
                        _Caprex2 = 0
                        _Caprnc1 = 0 ' Cantidad de Nota de credito
                        _Caprnc2 = 0

                        _Udtrpr = .Item("UnTrans")  ' Unidad de la transaccion
                        _Ud01pr = .Item("Ud01PR")
                        _Ud02pr = .Item("Ud02PR")
                        _Koltpr = .Item("CodLista") 'LISTADEPRECIO
                        _Mopppr = .Item("Moneda") 'trae_dato(tb, cn1, "KOMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        _Timopppr = .Item("Tipo_Moneda") 'trae_dato(tb, cn1, "TIMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        _Tamopppr = De_Num_a_Tx_01(.Item("Tipo_Cambio"), False, 5) 'De_Num_a_Tx_01(trae_dato(tb, cn1, "VAMO", "TABMO", "NOKOMO LIKE '%PESO%'"), False, 5)
                        _Ppprnelt = De_Num_a_Tx_01(.Item("PrecioNetoUd"), False, 5)
                        _Podtglli = De_Num_a_Tx_01(.Item("DescuentoPorc"), False, 5)
                        _Poimglli = De_Num_a_Tx_01(.Item("PorIla"), False, 5)

                        _Operacion = .Item("Operacion")
                        _Potencia = De_Num_a_Tx_01(.Item("Potencia"), False, 5)

                        Dim Campo As String = "Precio"

                        _Ppprne = De_Num_a_Tx_01(.Item("PrecioNetoUd"), False, 5)
                        _Ppprbr = De_Num_a_Tx_01(.Item("PrecioBrutoUd"), False, 5)
                        _Ppprnelt = De_Num_a_Tx_01(NuloPorNro(Of Double)(.Item("PrecioNetoUdLista"), 0), False, 5)
                        _Ppprbrlt = De_Num_a_Tx_01(.Item("PrecioBrutoUdLista"), False, 0) ' PRECIO BRUTO DE LA LISTA

                        _Poivli = De_Num_a_Tx_01(.Item("PorIva"), True)
                        _Nudtli = De_Num_a_Tx_01(.Item("NroDscto"), True)

                        _Nuimli = De_Num_a_Tx_01(.Item("NroImpuestos"), True)

                        _Vadtneli = De_Num_a_Tx_01(.Item("DsctoNeto"), False, 5)
                        _Vadtbrli = De_Num_a_Tx_01(.Item("DsctoBruto"), False, 5) 'ValDscto
                        _Vaneli = De_Num_a_Tx_01(.Item("ValNetoLinea"), False, 5)
                        _Vaimli = De_Num_a_Tx_01(.Item("ValIlaLinea"), False, 5)
                        _Vaivli = De_Num_a_Tx_01(.Item("ValIvaLinea"), False, 5)
                        _Vabrli = De_Num_a_Tx_01(Math.Round(.Item("ValBrutoLinea"), 0), False, 5)
                        _Feemli = _Feemdo 'Format(Now.Date, "yyyyMMdd") '""20121127"
                        _Feerli = _Feemdo 'Format(Now.Date, "yyyyMMdd")

                        _Observa = NuloPorNro(.Item("Observa"), "")

                        _Ppprnere1 = De_Num_a_Tx_01(.Item("PrecioNetoRealUd1"), False, 5)
                        _Ppprnere2 = De_Num_a_Tx_01(.Item("PrecioNetoRealUd2"), False, 5)
                        _Ppprpm = De_Num_a_Tx_01(NuloPorNro(.Item("PmLinea"), 0), False, 5)
                        _Ppprmsuc = De_Num_a_Tx_01(NuloPorNro(.Item("PmSucLinea"), 0), False, 5)
                        _Ppprpmifrs = De_Num_a_Tx_01(NuloPorNro(.Item("PmIFRS"), 0), False, 5)


                        _Alternat = NuloPorNro(.Item("CodigoProv"), "")

                        Dim _TipoValor As String = NuloPorNro(.Item("TipoValor"), "")



                        If Not String.IsNullOrEmpty(Trim(_Tict)) Then
                            Dim TipoValor = _TipoValor 'trae_dato(tb, cn1, "TipoValor", "ZW_Bkp_Configuracion")

                            If TipoValor = "N" Then
                                _Caprco1 = _Vadtneli
                                _Vadtbrli = De_Txt_a_Num_01((_Vabrli), 0) * -1
                            Else
                                _Vadtneli = De_Num_a_Tx_01(De_Txt_a_Num_01((_Vaneli), 5) * -1, False, 5)
                                _Caprco1 = _Vadtbrli
                            End If

                            _Caprco2 = 0
                            _Caprad2 = 0
                            _Cafaco = 0
                            _Ppprnelt = 0
                            _Ppprne = 0
                            _Ppprbrlt = 0
                            _Ppprbr = 0
                            _Prct = 1
                            _Ppprpm = 0
                            _Ppprmsuc = 0
                            _Ppprpmifrs = 0
                            _Lincondesp = 1
                            _Nudtli = 0
                            _Eslido = "C"
                        Else
                            _Cafaco = _Caprco1
                        End If

                        If _Tido <> "COV" Then

                            If _Tido = "OCC" Then

                                Consulta_sql = "UPDATE MAEST SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," &
                                                                "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf &
                                                                "WHERE EMPRESA='" & _Empresa &
                                                                "' AND KOSU='" & _Sulido &
                                                                "' AND KOBO='" & _Bosulido &
                                                                "' AND KOPR='" & _Koprct & "'" & vbCrLf &
                                               "UPDATE MAEPREM SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," &
                                                                "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf &
                                                                "WHERE EMPRESA='" & _Empresa &
                                                                "' AND KOPR='" & _Koprct & "'" & vbCrLf &
                                               "UPDATE MAEPR SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," &
                                                                "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf &
                                                                "WHERE KOPR='" & _Koprct & "'"

                                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                Comando.Transaction = myTrans
                                Comando.ExecuteNonQuery()

                            ElseIf _Tido = "NVV" Then

                                Consulta_sql = "UPDATE MAEST SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," &
                                                               "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf &
                                                               "WHERE EMPRESA='" & _Empresa &
                                                               "' AND KOSU='" & _Sulido &
                                                               "' AND KOBO='" & _Bosulido &
                                                               "' AND KOPR='" & _Koprct & "'" & vbCrLf &
                                              "UPDATE MAEPREM SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," &
                                                               "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf &
                                                               "WHERE EMPRESA='" & _Empresa &
                                                               "' AND KOPR='" & _Koprct & "'" & vbCrLf &
                                              "UPDATE MAEPR SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," &
                                                               "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf &
                                                               "WHERE KOPR='" & _Koprct & "'"

                                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                Comando.Transaction = myTrans
                                Comando.ExecuteNonQuery()

                            Else

                                If _Lincondesp Then

                                    Consulta_sql = "UPDATE MAEPREM SET" & vbCrLf &
                                                  "STFI1 = STFI1 - " & _Caprco1 & ",STFI2 =  - " & _Caprco2 & vbCrLf &
                                                  "WHERE EMPRESA = '" & _Empresa & "' AND KOPR = '" & _Koprct & "'" & vbCrLf &
                                                  "UPDATE MAEPR SET  STFI1 = STFI1 - " & _Caprco1 & ",STFI2 = - " & _Caprco2 & vbCrLf &
                                                  "WHERE KOPR = '" & _Koprct & "'" & vbCrLf &
                                                  "UPDATE MAEST SET STFI1 = STFI1 - " & _Caprco1 & ",STFI2 =  - " & _Caprco2 & vbCrLf &
                                                  "WHERE EMPRESA='" & _Empresa & "' AND KOSU='" & _Sudo &
                                                  "' AND KOBO='" & _Bosulido & "' AND KOPR='" & _Koprct & "'"

                                    _Caprad1 = _Caprco1
                                    _Caprad2 = _Caprco2


                                Else

                                    Consulta_sql = "UPDATE MAEPREM SET" & vbCrLf &
                                                   "STDV1 = STDV1 + " & _Caprco1 & ",STDV2 =  + " & _Caprco2 & vbCrLf &
                                                   "WHERE EMPRESA = '" & _Empresa & "' AND KOPR = '" & _Koprct & "'" & vbCrLf &
                                                   "UPDATE MAEPR SET  STDV1 = STDV1 + " & _Caprco1 & ",STDV2 = + " & _Caprco2 & vbCrLf &
                                                   "WHERE KOPR = '" & _Koprct & "'" & vbCrLf &
                                                   "UPDATE MAEST SET STDV1 = STDV1 + " & _Caprco1 & ",STDV2 =  + " & _Caprco2 & vbCrLf &
                                                   "WHERE EMPRESA='" & _Empresa & "' AND KOSU='" & _Sudo &
                                                   "' AND KOBO='" & _Bosulido & "' AND KOPR='" & _Koprct & "'"

                                    _Caprad1 = 0
                                    _Caprad2 = 0


                                End If

                                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                Comando.Transaction = myTrans
                                Comando.ExecuteNonQuery()

                            End If
                        End If

                        '_Mopppr = .Item("Moneda") 'trae_dato(tb, cn1, "KOMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        '_Timopppr = .Item("Tipo_Moneda") 'trae_dato(tb, cn1, "TIMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        '_Tamopppr = .Item("Tipo_Cambio") 'De_Num_a_Tx_01(trae_dato(tb, cn1, "VAMO", "TABMO", "NOKOMO LIKE '%PESO%'"), False, 5)

                        If _Tido <> "COV" Then

                            _Idrst = Val(NuloPorNro(.Item("Idmaeddo_Dori"), ""))

                            'EMPREPA,TIDOPA,NUDOPA,ENDOPA,NULIDOPA

                            If CBool(_Idrst) Then

                                Consulta_sql = "Select Top 1 * From MAEDDO Where IDMAEDDO = " & _Idrst

                                Dim _Tbl_Doc_Origen As DataTable = _Sql.Fx_Get_Tablas(Consulta_sql)

                                _Emprepa = _Tbl_Doc_Origen.Rows(0).Item("EMPRESA")
                                _Tidopa = _Tbl_Doc_Origen.Rows(0).Item("TIDO")
                                _Nudopa = _Tbl_Doc_Origen.Rows(0).Item("NUDO")
                                _Endopa = _Tbl_Doc_Origen.Rows(0).Item("ENDO")
                                _Nulidopa = _Tbl_Doc_Origen.Rows(0).Item("NULIDO")

                                Dim _Caprnc1_Ori As Double = _Tbl_Doc_Origen.Rows(0).Item("CAPRNC1")
                                Dim _Caprnc2_Ori As Double = _Tbl_Doc_Origen.Rows(0).Item("CAPRNC2")
                                Dim _Caprex1_Ori As Double = _Tbl_Doc_Origen.Rows(0).Item("CAPREX1")
                                Dim _Caprex2_Ori As Double = _Tbl_Doc_Origen.Rows(0).Item("CAPREX2")

                                Dim _CantUd1_Dori As Double = .Item("CantUd1_Dori")
                                Dim _CantUd2_Dori As Double = .Item("CantUd2_Dori")

                                Dim _Cant_MovUd1_Ext As String
                                Dim _Cant_MovUd2_Ext As String

                                Dim _CantUd1 As Double = .Item("CantUd1")
                                Dim _CantUd2 As Double = .Item("CantUd2")

                                If _CantUd1_Dori < _CantUd1 Then
                                    _Cant_MovUd1_Ext = De_Num_a_Tx_01(_CantUd1_Dori, False, 5)
                                    _Cant_MovUd2_Ext = De_Num_a_Tx_01(_CantUd2_Dori, False, 5)
                                Else
                                    _Cant_MovUd1_Ext = De_Num_a_Tx_01(_CantUd1, False, 5)
                                    _Cant_MovUd2_Ext = De_Num_a_Tx_01(_CantUd2, False, 5)
                                End If

                                _Archirst = "MAEDDO"

                                If _Tido = "NCC" Or _Tido = "NCV" Then

                                    Consulta_sql = "UPDATE MAEDDO SET CAPRNC1=CAPRNC1+" & _Cant_MovUd1_Ext &
                                                                    ",CAPRNC2=CAPRNC2+" & _Cant_MovUd2_Ext & "," &
                                                   "'ESLIDO = " & vbCrLf &
                                                   "CASE" & vbCrLf &
                                                   "'WHEN UDTRPR='1' AND CAPRCO1-CAPRAD1-CAPREX1=0 THEN 'C'" & vbCrLf &
                                                   "'WHEN UDTRPR='2' AND CAPRCO2-CAPRAD2-CAPREX2=0 THEN 'C'" & vbCrLf &
                                                   "'ELSE ''" & vbCrLf &
                                                   "END" & vbCrLf &
                                                   "WHERE IDMAEDDO = " & _Idrst

                                Else

                                    Consulta_sql = "UPDATE MAEDDO SET CAPREX1=CAPREX1+" & _Cant_MovUd1_Ext &
                                                                    ",CAPREX2=CAPREX2+" & _Cant_MovUd2_Ext & "," &
                                                   "ESLIDO = " &
                                                   "CASE" & vbCrLf &
                                                   "WHEN UDTRPR='1' AND " &
                                                   "ROUND(CAPRCO1-CAPRAD1-(CAPREX1+" & _Cant_MovUd1_Ext & "),5)=0 THEN 'C'" & vbCrLf &
                                                   "WHEN UDTRPR='2' AND " &
                                                   "ROUND(CAPRCO2-CAPRAD2-(CAPREX2+" & _Cant_MovUd2_Ext & "),5)=0 THEN 'C'" & vbCrLf &
                                                   "Else ''" & vbCrLf &
                                                   "End" & vbCrLf &
                                                   "WHERE IDMAEDDO = " & _Idrst  '1398920

                                End If

                                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                Comando.Transaction = myTrans
                                Comando.ExecuteNonQuery()

                            End If
                        Else
                            _Idrst = 0
                        End If

                        Consulta_sql =
                              "INSERT INTO MAEDDO (IDMAEEDO,ARCHIRST,IDRST,EMPRESA,TIDO,NUDO,ENDO,SUENDO,LILG,NULIDO," & vbCrLf &
                              "SULIDO,BOSULIDO,LUVTLIDO,KOFULIDO,TIPR,TICT,PRCT,KOPRCT,UDTRPR,RLUDPR,CAPRCO1," & vbCrLf &
                              "UD01PR,CAPRCO2,UD02PR,CAPRAD1,CAPRAD2,KOLTPR,MOPPPR,TIMOPPPR,TAMOPPPR,NUIMLI,NUDTLI," & vbCrLf &
                              "PODTGLLI,POIMGLLI,VAIMLI,VADTNELI,VADTBRLI,POIVLI,VAIVLI,VANELI,VABRLI,TIGELI," & vbCrLf &
                              "EMPREPA,TIDOPA,NUDOPA,ENDOPA,NULIDOPA," & vbCrLf &
                              "FEEMLI,FEERLI,PPPRNELT,PPPRNE,PPPRBRLT,PPPRBR,PPPRPM,PPPRNERE1,PPPRNERE2,CAFACO," & vbCrLf &
                              "FVENLOTE,FCRELOTE,NOKOPR,ALTERNAT,TASADORIG,CUOGASDIF,OPERACION,POTENCIA,ESLIDO,OBSERVA)" & vbCrLf &
                       "VALUES (" & _Idmaeedo & ",'" & _Archirst & "'," & _Idrst & ",'" & _Empresa & "','" & _Tido & "','" & _Nudo & "','" & _Endo &
                              "','" & _Suendo & "','SI','" & _Nulido & "','" & _Sudo & "','" & _Bosulido &
                              "','" & _Luvtven & "','" & _Kofudo & "','" & _Tipr & "','" & _Tict & "'," & CInt(_Prct) & ",'" & _Koprct &
                              "'," & _Udtrpr & "," & _Rludpr & "," & _Caprco1 & ",'" & _Ud01pr & "'," & _Caprco2 &
                              ",'" & _Ud02pr & "'," & _Caprad1 & "," & _Caprad2 & ",'TABPP" & _Koltpr & "'" &
                              ",'" & _Mopppr & "','" & _Timopppr & "'," & _Tamopppr &
                              "," & _Nuimli & "," & _Nudtli & "," & _Podtglli & "," & _Poimglli & "," & _Vaimli &
                              "," & _Vadtneli & "," & _Vadtbrli & "," & _Poivli & "," & _Vaivli & "," & _Vaneli &
                              "," & _Vabrli & ",'I'," &
                              "'" & _Emprepa & "','" & _Tidopa & "','" & _Nudopa & "','" & _Endopa & "','" & _Nulidopa & "'," &
                              "'" & _Feemli & "','" & _Feerli & "'," & _Ppprnelt & "," & _Ppprne &
                              "," & _Ppprbrlt & "," & _Ppprbr & "," & _Ppprpm & "," & _Ppprnere1 & "," & _Ppprnere2 &
                              "," & _Cafaco & ",NULL,NULL,'" & _Nokopr & "','" & _Alternat & "',1.00000,0" &
                              ",'" & _Operacion & "'," & _Potencia & ",'" & _Eslido & "',' " & _Observa & "')"

                        Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                        Comando.Transaction = myTrans
                        Comando.ExecuteNonQuery()


                        Comando = New SqlCommand("SELECT @@IDENTITY AS 'Identity'", cn2)
                        Comando.Transaction = myTrans
                        dfd1 = Comando.ExecuteReader()
                        Dim _Idmaeddo As Integer
                        While dfd1.Read()
                            _Idmaeddo = dfd1("Identity")
                        End While
                        dfd1.Close()

                        ' *** PM POR SUCURSAL SI ES QUE EXISTE EL CAMPO
                        Dim _Reg As Integer = _Sql.Fx_Cuenta_Registros("INFORMATION_SCHEMA.COLUMNS",
                                                              "COLUMN_NAME LIKE 'PPPRPMSUC' AND TABLE_NAME = 'MAEDDO'")

                        If CBool(_Reg) Then

                            Consulta_sql = "UPDATE MAEDDO SET PPPRPMSUC = " & _Ppprmsuc & vbCrLf &
                                           "WHERE IDMAEDDO = " & _Idmaeddo

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()
                        End If
                        '*************************************************************************************************

                        ' *** PMIFRS SI ES QUE EXISTE EL CAMPO
                        _Reg = _Sql.Fx_Cuenta_Registros("INFORMATION_SCHEMA.COLUMNS",
                                                         "COLUMN_NAME LIKE 'PMIFRS' AND TABLE_NAME = 'MAEPREM'")

                        If CBool(_Reg) Then

                            Consulta_sql = "UPDATE MAEDDO SET PPPRPMIFRS = " & _Ppprpmifrs & vbCrLf &
                                           "WHERE IDMAEDDO=" & _Idmaeddo

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()
                        End If
                        '*************************************************************************************************

                    End With


                    ' TABLA DE IMPUESTOS

                    Dim Tbl_Impuestos As DataTable = Bd_Documento.Tables("Impuestos_Doc")

                    If Val(_Nuimli) > 0 And String.IsNullOrEmpty(Trim(_Tict)) Then

                        For Each FImpto As DataRow In Tbl_Impuestos.Select("Id = " & Id_Linea)

                            Dim _Poimli As String = De_Num_a_Tx_01(FImpto.Item("Poimli").ToString, False, 5)
                            Dim _Koimli As String = FImpto.Item("Koimli").ToString
                            Dim _Vaimli = De_Num_a_Tx_01(FImpto.Item("Vaimli").ToString, False, 5)

                            Consulta_sql = "INSERT INTO MAEIMLI(IDMAEEDO,NULIDO,KOIMLI,POIMLI,VAIMLI,LILG) VALUES " & vbCrLf &
                                           "(" & _Idmaeedo & ",'" & _Nulido & "','" & _Koimli & "'," & _Poimli & "," & _Vaimli & ",'')"

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                            '-- 3RA TRANSACCION--------------------------------------------------------------------
                        Next
                    End If



                    ' TABLA DE DESCUENTOS
                    Dim Tbl_Descuentos As DataTable = Bd_Documento.Tables("Descuentos_Doc")
                    _Nudtli = Tbl_Descuentos.Rows.Count
                    If Val(_Nudtli) > 0 And String.IsNullOrEmpty(Trim(_Tict)) Then

                        For Each FDscto As DataRow In Tbl_Descuentos.Select("Id = " & Id_Linea)

                            Dim _Podt = De_Num_a_Tx_01(FDscto.Item("Podt").ToString, False, 5)
                            Dim _Vadt = De_Num_a_Tx_01(FDscto.Item("Vadt").ToString, False, 5)

                            Consulta_sql = "INSERT INTO MAEDTLI (IDMAEEDO,NULIDO,KODT,PODT,VADT)" & vbCrLf &
                                   "values (" & _Idmaeedo & ",'" & _Nulido & "','D_SIN_TIPO'," & _Podt & "," & _Vadt & ")"


                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                            '-- 4TA TRANSACCION--------------------------------------------------------------------
                        Next
                    End If

                    Contador += 1
                End If
            Next


            'TABLA DE VENCIMIENTOS

            Consulta_sql = Fx_Vencimientos(_Row_Encabezado)

            If Not String.IsNullOrEmpty(Consulta_sql) Then

                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                Comando.Transaction = myTrans
                Comando.ExecuteNonQuery()

            End If

            If _Nuvedo = 0 Then _Nuvedo = 1

            Dim _HoraGrab As String
            'Dim _HH_sistem As Date

            '_HH_sistem = FechaDelServidor()
            '_HoraGrab = _HH_sistem.Hour

            'Dim _HH, _MM, _SS As Double

            '_HH = _HH_sistem.Hour
            '_MM = _HH_sistem.Minute
            '_SS = _HH_sistem.Second

            If EsAjuste Then
                _Marca = 1 ' Generalmente se marcan las GRI o GDI que son por ajuste
                _Subtido = "AJU" ' Generalmente se Marcan las Guias de despacho o recibo
                '_HH = 23 : _MM = 59 : _SS = 59
            Else
                _Marca = String.Empty
                _Subtido = String.Empty
            End If



            _HoraGrab = Hora_Grab_fx(EsAjuste, FechaDelServidor) 'Math.Round((_HH * 3600) + (_MM * 60) + _SS, 0)


            'Consulta_sql = "Declare @HoraGrab Int" & vbCrLf & _
            '               "set @HoraGrab = convert(money,substring(convert(varchar(20),getdate(),114),1,2)) * 3600 +" & vbCrLf & _
            '               "convert(money,substring(convert(varchar(20),getdate(),114),4,2)) * 60 + " & vbCrLf & _
            '               "Convert(money, substring(Convert(varchar(20), getdate(), 114), 7, 2))" & vbCrLf & vbCrLf & _



            Dim _Espgdo As String = "P"
            If _Tido = "OCC" Then _Espgdo = "S"
            ' HAY QUE PONER EL CAMPO TIPO DE MONEDA  "TIMODO"
            Consulta_sql = "UPDATE MAEEDO SET SUENDO='" & _Suendo & "',TIGEDO='I',SUDO='" & _Sudo &
                          "',FEEMDO='" & _Feemdo & "',KOFUDO='" & _Kofudo & "',ESPGDO='" & _Espgdo & "',CAPRCO=" & _Caprco &
                          ",CAPRAD=" & _Caprad & ",MEARDO = '" & _Meardo & "',MODO = '" & _Modo &
                          "',TIMODO = '" & _Timodo & "',TAMODO = " & _Tamodo & ",VAIVDO = " & _Vaivdo & ",VAIMDO = " & _Vaimdo & vbCrLf &
                          ",VANEDO = " & _Vanedo & ",VABRDO = " & _Vabrdo & ",FE01VEDO = '" & _Fe01vedo &
                          "',FEULVEDO = '" & _Feulvedo & "',NUVEDO = " & _Nuvedo & ",FEER = '" & _Feer &
                          "',KOTU = '1',LCLV = NULL,LAHORA = GETDATE(), DESPACHO = 1,HORAGRAB = " & _HoraGrab &
                          ",FECHATRIB = NULL,SUBTIDO = '" & _Subtido &
                          "',MARCA = '" & _Marca & "',ESDO = '',NUDONODEFI = " & CInt(_Es_ValeTransitorio) &
                          ",TIDOELEC = " & CInt(_Es_Documento_Electronico) & ",LUVTDO = '" & _Luvtven & "'" & vbCrLf &
                          "WHERE IDMAEEDO=" & _Idmaeedo

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()


            Dim Reg As Integer = _Sql.Fx_Cuenta_Registros("INFORMATION_SCHEMA.COLUMNS",
                                                         "COLUMN_NAME LIKE 'LISACTIVA' AND TABLE_NAME = 'MAEEDO'")

            If CBool(Reg) Then

                Consulta_sql = "UPDATE MAEEDO SET LISACTIVA = 'TABPP" & _Lisactiva & "'" & vbCrLf &
                               "WHERE IDMAEEDO=" & _Idmaeedo

                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                Comando.Transaction = myTrans
                Comando.ExecuteNonQuery()

            End If


            '========================================== CERRAR DOCUMENTOS ASOCIADOS ============================================
            If _Tido <> "COV" Then

                Dim Fl As String = Generar_Filtro_IN(Tbl_Detalle, "", "Idmaeedo_Dori", False, False, "")

                If Fl = "()" Then Fl = "(0)"

                Consulta_sql = "SELECT DISTINCT IDMAEEDO FROM MAEDDO WHERE IDMAEEDO IN " & Fl
                Dim _TblOrigen As DataTable = _Sql.Fx_Get_Tablas(Consulta_sql)

                'Idmaeedo_Dori

                If CBool(_TblOrigen.Rows.Count) Then

                    Dim _Sum_Caprco As Double

                    For Each _Fila_Idmaeedo As DataRow In _TblOrigen.Rows

                        _Sum_Caprco = 0

                        Dim _Idmaeedo_Origen = _Fila_Idmaeedo.Item("IDMAEEDO")

                        For Each _Fila As DataRow In Tbl_Detalle.Rows

                            Dim Idmaeedo_Dori = _Fila.Item("Idmaeedo_Dori")

                            If _Idmaeedo_Origen = Idmaeedo_Dori Then

                                Dim _Idrst = Val(_Fila.Item("Idmaeddo_Dori"))

                                If CBool(_Idrst) Then

                                    Dim _CantUd1_Dori As Double = _Fila.Item("CantUd1_Dori")
                                    Dim _CantUd2_Dori As Double = _Fila.Item("CantUd2_Dori")

                                    Dim _Cant_MovUd1_Ext As String
                                    Dim _Cant_MovUd2_Ext As String

                                    Dim _CantUd1 As Double = _Fila.Item("CantUd1")
                                    Dim _CantUd2 As Double = _Fila.Item("CantUd2")

                                    If _CantUd1_Dori < _CantUd1 Then
                                        _Cant_MovUd1_Ext = _CantUd1_Dori
                                        _Cant_MovUd2_Ext = _CantUd2_Dori
                                    Else
                                        _Cant_MovUd1_Ext = _CantUd1
                                        _Cant_MovUd2_Ext = _CantUd2
                                    End If

                                    _Sum_Caprco += _Cant_MovUd1_Ext '_Fila.Item("CantUd1") 'De_Num_a_Tx_01(_Fila.Item("CantUd1"), False, 5) ' Cantidad de la primera unidad

                                End If

                            End If

                        Next

                        If CBool(_Sum_Caprco) Then

                            Dim _Sum_Caprco_str As String = De_Num_a_Tx_01(_Sum_Caprco, False, 5)

                            If _Tido = "NCV" Or _Tido = "NCC" Then

                                Consulta_sql = "UPDATE MAEEDO SET CAPREX=CAPREX+0,CAPRNC=CAPRNC+" & _Sum_Caprco_str & ",CAPRAD=CAPRAD+0," & vbCrLf &
                                               "ESDO=CASE" & vbCrLf &
                                               "WHEN ROUND(CAPRCO-CAPRAD-CAPREX-(0)-(0),5)=0 THEN 'C'" & vbCrLf & "ELSE ''" & vbCrLf & "END," & vbCrLf &
                                               "ESFADO=CASE" & vbCrLf &
                                               "WHEN CAPRCO-CAPRAD-CAPREX-(0)-(0)=0 THEN 'F'" & vbCrLf & "ELSE ESFADO" & vbCrLf & "END" & vbCrLf &
                                               "WHERE IDMAEEDO= " & _Idmaeedo_Origen
                            Else
                                Consulta_sql = "UPDATE MAEEDO SET CAPREX=CAPREX+" & _Sum_Caprco_str & ",CAPRNC=CAPRNC+0,CAPRAD=CAPRAD+0," & vbCrLf &
                                               "ESDO=CASE " & vbCrLf &
                                               "WHEN ROUND(CAPRCO-CAPRAD-CAPREX-(0)-(" & _Sum_Caprco_str & "),5)=0 THEN 'C'" & vbCrLf & "ELSE ''" & vbCrLf & "END," & vbCrLf &
                                               "ESFADO=" & vbCrLf &
                                               "CASE" & vbCrLf &
                                               "WHEN CAPRCO-CAPRAD-CAPREX-(0)-(" & _Sum_Caprco_str & ")=0 THEN 'F'  " & vbCrLf & "ELSE ESFADO" & vbCrLf & "End" & vbCrLf &
                                               "WHERE IDMAEEDO = " & _Idmaeedo_Origen
                            End If

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                        End If

                    Next

                End If

            End If


            '=========================================== OBSERVACIONES ==================================================================

            Dim Tbl_Observaciones As DataTable = Bd_Documento.Tables("Observaciones_Doc")

            With Tbl_Observaciones

                _Obdo = .Rows(0).Item("Observaciones")
                _Cpdo = .Rows(0).Item("Forma_pago")
                _Ocdo = .Rows(0).Item("Orden_compra")

                For i = 0 To 34

                    Dim Campo As String = "Obs" & i + 1
                    Obs(i) = .Rows(0).Item(Campo)

                Next

            End With

            Consulta_sql = "INSERT INTO MAEEDOOB (IDMAEEDO,OBDO,CPDO,OCDO,DIENDESP,TEXTO1,TEXTO2,TEXTO3,TEXTO4,TEXTO5,TEXTO6," & vbCrLf &
                           "TEXTO7,TEXTO8,TEXTO9,TEXTO10,TEXTO11,TEXTO12,TEXTO13,TEXTO14,TEXTO15,CARRIER,BOOKING,LADING,AGENTE," & vbCrLf &
                           "MEDIOPAGO,TIPOTRANS,KOPAE,KOCIE,KOCME,FECHAE,HORAE,KOPAD,KOCID,KOCMD,FECHAD,HORAD,OBDOEXPO,MOTIVO," & vbCrLf &
                           "TEXTO16,TEXTO17,TEXTO18,TEXTO19,TEXTO20,TEXTO21,TEXTO22,TEXTO23,TEXTO24,TEXTO25,TEXTO26,TEXTO27," & vbCrLf &
                           "TEXTO28,TEXTO29,TEXTO30,TEXTO31,TEXTO32,TEXTO33,TEXTO34,TEXTO35) VALUES " & vbCrLf &
                           "(" & _Idmaeedo & ",'" & _Obdo & "','" & _Cpdo & "','" & _Ocdo & "','','" & Obs(0) & "','" & Obs(1) &
                           "','" & Obs(2) & "','" & Obs(3) & "','" & Obs(4) & "','" & Obs(5) & "','" & Obs(6) & "','" & Obs(7) &
                           "','" & Obs(8) & "','" & Obs(9) & "','" & Obs(10) & "','" & Obs(11) & "','" & Obs(12) & "','" & Obs(13) &
                           "','" & Obs(14) & "','','','','','','','','','',GETDATE(),'','','','',GETDATE(),'','','','" & Obs(15) &
                           "','" & Obs(16) & "','" & Obs(17) & "','" & Obs(18) & "','" & Obs(19) &
                           "','" & Obs(20) & "','" & Obs(21) & "','" & Obs(22) & "','" & Obs(23) & "','" & Obs(24) &
                           "','" & Obs(25) & "','" & Obs(26) & "','" & Obs(27) & "','" & Obs(28) & "','" & Obs(29) &
                           "','" & Obs(30) & "','" & Obs(31) & "','" & Obs(32) & "','" & Obs(33) & "','" & Obs(34) & "')"

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            ' ====================================================================================================================================

            If _Cambiar_Numeracion_Confiest Then
                Consulta_sql = Fx_Cambiar_Numeracion_Modalidad(_Tido, _Nudo, _Modalidad, _Empresa)

                If Not String.IsNullOrEmpty(Consulta_sql) Then

                    Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                    Comando.Transaction = myTrans
                    Comando.ExecuteNonQuery()

                End If
            End If

            myTrans.Commit()

            'Tbl_Detalle()
            Return _Idmaeedo

        Catch ex As Exception

            Dim _Erro_VB As String = ex.Message & vbCrLf & ex.StackTrace & vbCrLf &
                                     "Código: " & _Koprct
            'Clipboard.SetText(_Erro_VB)

            My.Computer.FileSystem.WriteAllText("ArchivoSalida", ex.Message & vbCrLf & ex.StackTrace, False)
            'MessageBoxEx.Show(ex.Message, "Error", _
            '                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
            myTrans.Rollback()
            MsgBox("Transacción desecha", MsgBoxStyle.Critical, "BakApp")
            'MessageBoxEx.Show("Transaccion desecha", "Problema", _
            '                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
            'SQL_ServerClass.CerrarConexion(cn2)
            Return _Erro_VB '0
        Finally
            _Sql.Sb_Cerrar_Conexion(cn2)
        End Try

    End Function
    Private Function Fx_Vencimientos(ByVal _RowEncabezado As DataRow) As String

        Dim _SqlQuery As String

        Dim _Tido = _RowEncabezado.Item("TipoDoc")
        Dim _TotalBrutoDoc As Double = _RowEncabezado.Item("TotalBrutoDoc")

        Dim _FechaEmision As Date = _RowEncabezado.Item("FechaEmision")
        Dim _Fecha_1er_Vencimiento As Date = _RowEncabezado.Item("Fecha_1er_Vencimiento")
        Dim _Cuotas As Integer = _RowEncabezado.Item("Cuotas")
        Dim _Dias_Vencimiento As Integer = _RowEncabezado.Item("Dias_Vencimiento")

        If _Cuotas = 0 Then _Cuotas = 1
        Dim _Aplica_Vencimientos As Boolean

        Select Case Mid(_Tido, 1, 1)

            Case "B", "F"
                _Aplica_Vencimientos = True
            Case Else
                _Aplica_Vencimientos = False

        End Select

        _Nuvedo = _Cuotas

        If _Aplica_Vencimientos Then

            Dim Cuotas_(_Cuotas - 1) As Date
            Cuotas_(0) = _Fecha_1er_Vencimiento

            Dim _FechasVenci As Date = _FechaEmision
            Dim _dias As Integer
            'If Cuotas > 1 Then

            Dim _Resultado As Double = _TotalBrutoDoc / _Cuotas
            Dim _Vave As Double = Math.Round(_TotalBrutoDoc / _Cuotas, 0)

            'If (Resultado Mod 1) = 0 Then
            'Valor_Cuota = Resultado
            'End If

            _dias = _Dias_Vencimiento

            For i = 1 To _Cuotas

                _FechasVenci = DateAdd(DateInterval.Day, _dias, _FechasVenci)
                Cuotas_(i - 1) = _FechasVenci
                _dias = _Dias_Vencimiento


                If i = _Cuotas Then
                    Dim rS = _Vave * _Cuotas
                    rS = _TotalBrutoDoc - rS
                    rS = _Vave + rS
                    _Vave = rS
                End If

                If i = 1 Then
                    _FechasVenci = _Fecha_1er_Vencimiento
                Else
                    _FechasVenci = _FechasVenci
                End If

                Dim _Feve As String = Format(_FechasVenci, "yyyyMMdd")
                Dim _Espgve As String = String.Empty
                Dim _Vaabve As String = 0
                Dim _Archirst As String = String.Empty
                Dim _porestpag As String = 0
                Dim __Observa As String = String.Empty

                _SqlQuery += "INSERT INTO MAEVEN (IDMAEEDO,FEVE,ESPGVE,VAVE,VAABVE,ARCHIRST,PORESTPAG,OBSERVA)" & vbCrLf & _
                               "values (" & _Idmaeedo & ",'" & _Feve & "','" & _Espgve & "'," & De_Num_a_Tx_01(_Vave, False, 5) & "," & _Vaabve & _
                               ",'" & _Archirst & "'," & _porestpag & ",'" & __Observa & "')" & vbCrLf



            Next

        End If

        Return _SqlQuery

    End Function


#End Region

    Function Fx_Cambiar_Numeracion_Modalidad(ByVal _Tido As String, _
                                             ByVal _Nudo As String, _
                                             ByVal _Empresa As String, _
                                             ByVal _Modalidad As String) As String


        Dim _Consulta_sql = "Select " & _Tido & " From CONFIEST Where MODALIDAD = '" & _Modalidad & "'"
        Dim _Tbl As DataTable = _Sql.Fx_Get_Tablas(_Consulta_sql) 'get_Tablas(_Consulta_sql, cn1)

        Dim _Nudo_Modalidad As String

        _Consulta_sql = String.Empty

        If CBool(_Tbl.Rows.Count) Then

            _Nudo_Modalidad = _Tbl.Rows(0).Item(_Tido)

            Dim Continua As Boolean = True

            If Not String.IsNullOrEmpty(Trim(_Nudo_Modalidad)) Then

                Dim _ProxNumero As String = Fx_Proximo_NroDocumento(_Nudo)

                _Consulta_sql = "UPDATE CONFIEST SET " & _Tido & " = '" & _ProxNumero & "'" & vbCrLf & _
                                "WHERE EMPRESA = '" & _Empresa & "' AND  MODALIDAD = '" & _Modalidad & "'"

            End If

        End If


        Return _Consulta_sql
        'Dim _NrNumeroDoco As String = trae_dato(tb, cn1, _TipoDoc, "CONFIEST", "EMPRESA = '" & ModEmpresa & _
        '                              "' AND MODALIDAD = '" & Modalidad & "'") 'FUNCIONARIO & numero_(Trim(Str(CantOCCFuncionario)), 7)

    End Function

#Region "FUNCION CREAR DOCUMENTO RANDOM KASI"

    Function Fx_Crear_Documento_KASI(ByVal _Tipo_de_documento As String, _
                                     ByVal _Numero_de_documento As String, _
                                     ByVal _Es_Documento_Electronico As Boolean, _
                                     ByVal _Bd_Documento As DataSet, _
                                     Optional ByVal _EsAjuste As Boolean = False) As Integer



        Dim myTrans As SqlClient.SqlTransaction
        Dim Comando As SqlClient.SqlCommand
        Dim cn2 As New SqlConnection

        Dim Tbl_Encabezado As DataTable = _Bd_Documento.Tables("Encabezado_Doc")


        Try

            _Sql.Sb_Abrir_Conexion(cn2)

            With Tbl_Encabezado

                Dim _Modalidad As String = .Rows(0).Item("Modalidad")
                _Tido = .Rows(0).Item("TipoDoc")
                _Numero_de_documento = Traer_Numero_Documento(_Tido, .Rows(0).Item("NroDocumento"), _Modalidad, _Empresa)

                If String.IsNullOrEmpty(_Numero_de_documento) Then
                    Return 0
                End If

                .Rows(0).Item("NroDocumento") = _Numero_de_documento
                _Nudo = .Rows(0).Item("NroDocumento")

                If String.IsNullOrEmpty(Trim(_Nudo)) Then
                    Return 0
                End If

                _Empresa = .Rows(0).Item("Empresa").ToString
                _Sudo = .Rows(0).Item("Sucursal")
                _Kofudo = .Rows(0).Item("CodFuncionario")


                _Endo = .Rows(0).Item("CodEntidad")
                _Suendo = .Rows(0).Item("CodSucEntidad")

                _Feemdo = Format(.Rows(0).Item("FechaEmision"), "yyyyMMdd")
                _Lisactiva = .Rows(0).Item("ListaPrecios")
                _Caprco = De_Num_a_Tx_01(.Rows(0).Item("CantTotal"), 5)
                _Caprad = De_Num_a_Tx_01(.Rows(0).Item("CantDesp"), 5)

                _Luvtven = .Rows(0).Item("Centro_Costo")
                _Modo = .Rows(0).Item("Moneda_Doc")
                _Meardo = .Rows(0).Item("DocEn_Neto_Bruto")
                _Tamodo = De_Num_a_Tx_01(.Rows(0).Item("Valor_Dolar"), False, 5)
                _Timodo = .Rows(0).Item("TipoMoneda")

                Dim _Vanedo_2 = .Rows(0).Item("TotalNetoDoc")
                Dim _Vaivdo_2 = .Rows(0).Item("TotalIvaDoc")
                Dim _Vaimdo_2 = .Rows(0).Item("TotalIlaDoc")
                Dim _Vabrdo_2 = .Rows(0).Item("TotalBrutoDoc")


                _Vanedo = De_Num_a_Tx_01(.Rows(0).Item("TotalNetoDoc"), False, 5)
                _Vaivdo = De_Num_a_Tx_01(.Rows(0).Item("TotalIvaDoc"), False, 5)
                _Vaimdo = De_Num_a_Tx_01(.Rows(0).Item("TotalIlaDoc"), False, 5)
                _Vabrdo = De_Num_a_Tx_01(.Rows(0).Item("TotalBrutoDoc"), False, 5)

                _Fe01vedo = Format(.Rows(0).Item("Fecha_1er_Vencimiento"), "yyyyMMdd")
                _Feulvedo = Format(.Rows(0).Item("FechaUltVencimiento"), "yyyyMMdd")

                _Feer = Format(.Rows(0).Item("FechaRecepcion"), "yyyyMMdd")
                _Feerli = Format(.Rows(0).Item("FechaRecepcion"), "yyyyMMdd")

                '------------------------------------------------------------------------------------------------------------


            End With


            myTrans = cn2.BeginTransaction()


            Consulta_sql = "INSERT INTO KASIEDO ( EMPRESA,TIDO,NUDO,ENDO,SUENDO )" & vbCrLf & _
                           "VALUES ( '" & _Empresa & "','" & _Tido & "','" & _Nudo & _
                           "','" & _Endo & "','" & _Suendo & "')"

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            Comando = New SqlCommand("SELECT @@IDENTITY AS 'Identity'", cn2)
            Comando.Transaction = myTrans
            Dim dfd1 As SqlDataReader = Comando.ExecuteReader()
            While dfd1.Read()
                _Idmaeedo = dfd1("Identity")
            End While
            dfd1.Close()

            _Bd_Documento.Tables("Detalle_Doc").Dispose()
            Dim Tbl_Detalle As DataTable = _Bd_Documento.Tables("Detalle_Doc")

            Dim Contador As Integer = 1

            For Each FDetalle As DataRow In Tbl_Detalle.Rows
                Dim Estado As DataRowState = FDetalle.RowState
                If Not Estado = DataRowState.Deleted Then

                    With FDetalle



                        Id_Linea = .Item("Id")


                        _Nulido = numero_(Contador, 5)

                        _Bosulido = .Item("Bodega")
                        _Koprct = .Item("Codigo")
                        _Nokopr = .Item("Descripcion")
                        _Rludpr = De_Num_a_Tx_01(.Item("Rtu"), False, 5)
                        _Sulido = .Item("Sucursal")
                        _Kofulido = _Funcionario 'FUNCIONARIO ' Codigo de funcionario
                        _Tict = .Item("Tict")
                        _Prct = .Item("Prct")
                        _Caprco1 = De_Num_a_Tx_01(.Item("CantUd1"), False, 5) ' Cantidad de la primera unidad
                        _Caprco2 = De_Num_a_Tx_01(.Item("CantUd2"), False, 5) ' Cantidad de la segunda unidad
                        _Tipr = .Item("Tipr")
                        _Lincondesp = .Item("Lincondest")

                        'CantidadTotal = CantidadTotal + Val(CAPRCO1)
                        If _Lincondesp Then
                            _Caprad1 = _Caprco1 ' Cantidad que mueve Stock Fisico, según el producto.
                            _Caprad2 = _Caprco2 ' Cantidad que mueve Stock Fisico, según el producto.
                        Else
                            _Caprad1 = De_Num_a_Tx_01(NuloPorNro(.Item("CDespUd1"), 0), False, 5) ' Cantidad que mueve Stock Fisico, según el producto.
                            _Caprad2 = De_Num_a_Tx_01(NuloPorNro(.Item("CDespUd2"), 0), False, 5) ' Cantidad que mueve Stock Fisico, según el producto.
                        End If

                        _Eslido = NuloPorNro(.Item("Estado"), "")

                        _Caprex1 = 0 ' Cantidad  
                        _Caprex2 = 0
                        _Caprnc1 = 0 ' Cantidad de Nota de credito
                        _Caprnc2 = 0

                        _Udtrpr = .Item("UnTrans")  ' Unidad de la transaccion
                        _Ud01pr = .Item("Ud01PR")
                        _Ud02pr = .Item("Ud02PR")
                        _Koltpr = .Item("CodLista") 'LISTADEPRECIO
                        _Mopppr = .Item("Moneda") 'trae_dato(tb, cn1, "KOMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        _Timopppr = .Item("Tipo_Moneda") 'trae_dato(tb, cn1, "TIMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        _Tamopppr = De_Num_a_Tx_01(.Item("Tipo_Cambio"), False, 5) 'De_Num_a_Tx_01(trae_dato(tb, cn1, "VAMO", "TABMO", "NOKOMO LIKE '%PESO%'"), False, 5)
                        _Ppprnelt = De_Num_a_Tx_01(.Item("PrecioNetoUd"), False, 5)
                        _Podtglli = De_Num_a_Tx_01(.Item("DescuentoPorc"), False, 5)
                        _Poimglli = De_Num_a_Tx_01(.Item("PorIla"), False, 5)

                        _Operacion = .Item("Operacion")
                        _Potencia = De_Num_a_Tx_01(.Item("Potencia"), False, 5)

                        Dim Campo As String = "Precio"

                        _Ppprne = De_Num_a_Tx_01(.Item("PrecioNetoUd"), False, 5)
                        _Ppprbr = De_Num_a_Tx_01(.Item("PrecioBrutoUd"), False, 5)
                        _Ppprnelt = De_Num_a_Tx_01(NuloPorNro(Of Double)(.Item("PrecioNetoUdLista"), 0), False, 5)
                        _Ppprbrlt = De_Num_a_Tx_01(.Item("PrecioBrutoUdLista"), False, 0) ' PRECIO BRUTO DE LA LISTA

                        _Poivli = De_Num_a_Tx_01(.Item("PorIva"), True)
                        _Nudtli = De_Num_a_Tx_01(.Item("NroDscto"), True)

                        _Nuimli = De_Num_a_Tx_01(.Item("NroImpuestos"), True)

                        _Vadtneli = De_Num_a_Tx_01(.Item("DsctoNeto"), False, 5)
                        _Vadtbrli = De_Num_a_Tx_01(.Item("DsctoBruto"), False, 5) 'ValDscto
                        _Vaneli = De_Num_a_Tx_01(.Item("ValNetoLinea"), False, 5)
                        _Vaimli = De_Num_a_Tx_01(.Item("ValIlaLinea"), False, 5)
                        _Vaivli = De_Num_a_Tx_01(.Item("ValIvaLinea"), False, 5)
                        _Vabrli = De_Num_a_Tx_01(.Item("ValBrutoLinea"), False, 5)
                        _Feemli = _Feemdo 'Format(Now.Date, "yyyyMMdd") '""20121127"
                        _Feerli = _Feemdo 'Format(Now.Date, "yyyyMMdd")

                        _Observa = NuloPorNro(.Item("Observa"), "")

                        _Ppprnere1 = De_Num_a_Tx_01(.Item("PrecioNetoRealUd1"), False, 5)
                        _Ppprnere2 = De_Num_a_Tx_01(.Item("PrecioNetoRealUd2"), False, 5)
                        _Ppprpm = De_Num_a_Tx_01(NuloPorNro(.Item("PmLinea"), 0), False, 5)
                        _Ppprmsuc = De_Num_a_Tx_01(NuloPorNro(.Item("PmSucLinea"), 0), False, 5)

                        _Alternat = NuloPorNro(.Item("CodigoProv"), "")

                        Dim _TipoValor As String = NuloPorNro(.Item("TipoValor"), "")



                        If Not String.IsNullOrEmpty(Trim(_Tict)) Then
                            Dim TipoValor = _TipoValor 'trae_dato(tb, cn1, "TipoValor", "ZW_Bkp_Configuracion")

                            If TipoValor = "N" Then
                                _Caprco1 = _Vadtneli
                                _Vadtbrli = De_Txt_a_Num_01((_Vabrli), 0) * -1
                            Else
                                _Vadtneli = De_Num_a_Tx_01(De_Txt_a_Num_01((_Vaneli), 5) * -1, False, 5)
                                _Caprco1 = _Vadtbrli
                            End If

                            _Caprco2 = 0
                            _Caprad2 = 0
                            _Cafaco = 0
                            _Ppprnelt = 0
                            _Ppprne = 0
                            _Ppprbrlt = 0
                            _Ppprbr = 0
                            _Prct = 1
                            _Ppprpm = 0
                            _Ppprmsuc = 0
                            _Lincondesp = 1
                            _Nudtli = 0
                            _Eslido = "C"
                        Else
                            _Cafaco = _Caprco1
                        End If


                        '_Mopppr = .Item("Moneda") 'trae_dato(tb, cn1, "KOMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        '_Timopppr = .Item("Tipo_Moneda") 'trae_dato(tb, cn1, "TIMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        '_Tamopppr = .Item("Tipo_Cambio") 'De_Num_a_Tx_01(trae_dato(tb, cn1, "VAMO", "TABMO", "NOKOMO LIKE '%PESO%'"), False, 5)

                        Consulta_sql = _
                              "INSERT INTO KASIDDO (IDMAEDDO,IDMAEEDO,ARCHIRST,IDRST,EMPRESA,TIDO,NUDO,ENDO,SUENDO,LILG,NULIDO," & vbCrLf & _
                              "SULIDO,BOSULIDO,LUVTLIDO,KOFULIDO,TIPR,TICT,PRCT,KOPRCT,UDTRPR,RLUDPR,CAPRCO1," & vbCrLf & _
                              "UD01PR,CAPRCO2,UD02PR,CAPRAD1,CAPRAD2,KOLTPR,MOPPPR,TIMOPPPR,TAMOPPPR,NUIMLI,NUDTLI," & vbCrLf & _
                              "PODTGLLI,POIMGLLI,VAIMLI,VADTNELI,VADTBRLI,POIVLI,VAIVLI,VANELI,VABRLI,TIGELI," & vbCrLf & _
                              "FEEMLI,FEERLI,PPPRNELT,PPPRNE,PPPRBRLT,PPPRBR,PPPRPM,PPPRNERE1,PPPRNERE2,CAFACO," & vbCrLf & _
                              "FVENLOTE,FCRELOTE,NOKOPR,ALTERNAT,TASADORIG,CUOGASDIF,LINCONDESP,OPERACION,POTENCIA,ESLIDO,OBSERVA)" & vbCrLf & _
                       "VALUES (0," & _Idmaeedo & ",'',0,'" & _Empresa & "','','','" & _Endo & _
                              "','" & _Suendo & "','SI','" & _Nulido & "','" & _Sudo & "','" & _Bosulido & _
                              "','" & _Luvtven & "','" & _Kofudo & "','" & _Tipr & "','" & _Tict & "'," & CInt(_Prct) & ",'" & _Koprct & _
                              "'," & _Udtrpr & "," & _Rludpr & "," & _Caprco1 & ",'" & _Ud01pr & "'," & _Caprco2 & _
                              ",'" & _Ud02pr & "'," & _Caprad1 & "," & _Caprad2 & ",'TABPP" & _Koltpr & "'" & _
                              ",'" & _Mopppr & "','" & _Timopppr & "'," & _Tamopppr & _
                              "," & _Nuimli & "," & _Nudtli & "," & _Podtglli & "," & _Poimglli & "," & _Vaimli & _
                              "," & _Vadtneli & "," & _Vadtbrli & "," & _Poivli & "," & _Vaivli & "," & _Vaneli & _
                              "," & _Vabrli & ",'I','" & _Feemli & "','" & _Feerli & "'," & _Ppprnelt & "," & _Ppprne & _
                              "," & _Ppprbrlt & "," & _Ppprbr & "," & _Ppprpm & "," & _Ppprnere1 & "," & _Ppprnere2 & _
                              "," & _Cafaco & ",NULL,NULL,'" & _Nokopr & "','" & _Alternat & "',0,0," & CInt(_Lincondesp) & _
                              ",'" & _Operacion & "'," & _Potencia & ",'" & _Eslido & "',' " & _Observa & "')"

                        Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                        Comando.Transaction = myTrans
                        Comando.ExecuteNonQuery()

                        Dim _Reg As Integer = _Sql.Fx_Cuenta_Registros("INFORMATION_SCHEMA.COLUMNS", _
                                                              "COLUMN_NAME LIKE 'PPPRPMSUC' AND TABLE_NAME = 'KASIDDO'")

                        If CBool(_Reg) Then
                            Comando = New SqlCommand("SELECT @@IDENTITY AS 'Identity'", cn2)
                            Comando.Transaction = myTrans
                            dfd1 = Comando.ExecuteReader()
                            Dim _Idmaeddo As Integer
                            While dfd1.Read()
                                _Idmaeddo = dfd1("Identity")
                            End While
                            dfd1.Close()

                            Consulta_sql = "UPDATE KASIDDO SET PPPRPMSUC = " & _Ppprmsuc & vbCrLf & _
                                           "WHERE IDMAEDDO = " & _Idmaeddo

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()
                        End If

                    End With


                    ' TABLA DE IMPUESTOS

                    Dim Tbl_Impuestos As DataTable = _Bd_Documento.Tables("Impuestos_Doc")

                    If Val(_Nuimli) > 0 And String.IsNullOrEmpty(Trim(_Tict)) Then

                        For Each FImpto As DataRow In Tbl_Impuestos.Select("Id = " & Id_Linea)

                            Dim _Poimli As String = De_Num_a_Tx_01(FImpto.Item("Poimli").ToString, False, 5)
                            Dim _Koimli As String = FImpto.Item("Koimli").ToString
                            Dim _Vaimli = De_Num_a_Tx_01(FImpto.Item("Vaimli").ToString, False, 5)

                            Consulta_sql = "INSERT INTO KASIIMLI(IDMAEEDO,NULIDO,KOIMLI,POIMLI,VAIMLI,LILG) VALUES " & vbCrLf & _
                                           "(" & _Idmaeedo & ",'" & _Nulido & "','" & _Koimli & "'," & _Poimli & "," & _Vaimli & ",'')"

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                            '-- 3RA TRANSACCION--------------------------------------------------------------------
                        Next
                    End If



                    ' TABLA DE DESCUENTOS
                    Dim Tbl_Descuentos As DataTable = _Bd_Documento.Tables("Descuentos_Doc")
                    _Nudtli = Tbl_Descuentos.Rows.Count
                    If Val(_Nudtli) > 0 And String.IsNullOrEmpty(Trim(_Tict)) Then

                        For Each FDscto As DataRow In Tbl_Descuentos.Select("Id = " & Id_Linea)

                            Dim _Podt = De_Num_a_Tx_01(FDscto.Item("Podt").ToString, False, 5)
                            Dim _Vadt = De_Num_a_Tx_01(FDscto.Item("Vadt").ToString, False, 5)

                            Consulta_sql = "INSERT INTO KASIDTLI (IDMAEEDO,NULIDO,KODT,PODT,VADT)" & vbCrLf & _
                                   "values (" & _Idmaeedo & ",'" & _Nulido & "','D_SIN_TIPO'," & _Podt & "," & _Vadt & ")"


                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                            '-- 4TA TRANSACCION--------------------------------------------------------------------
                        Next
                    End If

                    Contador += 1
                End If
            Next



            'TABLA DE VENCIMIENTOS

            'Dim Tbl_Vencimientos As DataTable = _Bd_Documento.Tables("Vencimientos_Doc")

            '_Nuvedo = Tbl_Vencimientos.Rows.Count

            'For Each FVencimientos As DataRow In Tbl_Vencimientos.Rows

            'Dim _Feve As String = Format(FVencimientos.Item("Fecha_Vencimiento"), "yyyyMMdd")
            'Dim _Espgve As String = String.Empty 'FilaX.Item("Estado_Pago").ToString
            'Dim _Vave As String = De_Num_a_Tx_01(FVencimientos.Item("Valor_Vencimiento").ToString, False, 5)
            'Dim _Vaabve As String = De_Num_a_Tx_01(FVencimientos.Item("Valor_Abonado").ToString, False, 5)
            'Dim _Archirst As String = String.Empty 'FilaX.Item("Archirst").ToString
            'Dim _porestpag As String = 0 'De_Num_a_Tx_01(FilaX.Item("Porestpag").ToString, False, 5)
            'Dim __Observa As String = String.Empty 'FilaX.Item("Archirst").ToString

            'Consulta_sql = "INSERT INTO MAEVEN (IDMAEEDO,FEVE,ESPGVE,VAVE,VAABVE,ARCHIRST,PORESTPAG,OBSERVA)" & vbCrLf & _
            '               "values (" & _Idmaeedo & ",'" & _Feve & "','" & _Espgve & "'," & _Vave & "," & _Vaabve & _
            '               ",'" & _Archirst & "'," & _porestpag & ",'" & __Observa & "')"

            'Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            'Comando.Transaction = myTrans
            'Comando.ExecuteNonQuery()

            'Next




            If _Nuvedo = 0 Then _Nuvedo = 1

            Dim _HoraGrab As String

            If _EsAjuste Then
                _Marca = 1 ' Generalmente se marcan las GRI o GDI que son por ajuste
                _Subtido = "AJU" ' Generalmente se Marcan las Guias de despacho o recibo
                '_HH = 23 : _MM = 59 : _SS = 59
            Else
                _Marca = String.Empty
                _Subtido = String.Empty
            End If

            _HoraGrab = Hora_Grab_fx(_EsAjuste, _FechaEmision) 'Math.Round((_HH * 3600) + (_MM * 60) + _SS, 0)



            Dim _Espgdo As String = "P"
            If _Tido = "OCC" Then _Espgdo = "S"
            ' HAY QUE PONER EL CAMPO TIPO DE MONEDA  "TIMODO"
            Consulta_sql = "UPDATE KASIEDO SET SUENDO='" & _Suendo & "',TIGEDO='I',SUDO='" & _Sudo & _
                           "',FEEMDO='" & _Feemdo & "',KOFUDO='" & _Kofudo & "',ESPGDO='" & _Espgdo & "',CAPRCO=" & _Caprco & _
                           ",CAPRAD=" & _Caprad & ",MEARDO = '" & _Meardo & "',MODO = '" & _Modo & _
                           "',TIMODO = '" & _Timodo & "',TAMODO = " & _Tamodo & ",VAIVDO = " & _Vaivdo & ",VAIMDO = " & _Vaimdo & vbCrLf & _
                           ",VANEDO = " & _Vanedo & ",VABRDO = " & _Vabrdo & ",FE01VEDO = '" & _Fe01vedo & _
                           "',FEULVEDO = '" & _Feulvedo & "',NUVEDO = " & _Nuvedo & ",FEER = '" & _Feer & _
                           "',KOTU = '1',LCLV = NULL,LAHORA = GETDATE(), DESPACHO = 1,HORAGRAB = " & _HoraGrab & _
                           ",FECHATRIB = NULL,NUMOPERVEN = 0,FLIQUIFCV = '" & _Feemdo & "',SUBTIDO = '" & _Subtido & _
                           "',MARCA = '" & _Marca & "',ESDO = ''" & _
                           ",TIDOELEC = " & CInt(_Es_Documento_Electronico) & ",LUVTDO = '" & _Luvtven & "'" & vbCrLf & _
                           "WHERE IDMAEEDO=" & _Idmaeedo


            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()


            Dim Reg As Integer = _Sql.Fx_Cuenta_Registros("INFORMATION_SCHEMA.COLUMNS", _
                                                         "COLUMN_NAME LIKE 'LISACTIVA' AND TABLE_NAME = 'KASIEDO'")

            If CBool(Reg) Then

                Consulta_sql = "UPDATE KASIEDO SET LISACTIVA = 'TABPP" & _Lisactiva & "'" & vbCrLf & _
                               "WHERE IDMAEEDO=" & _Idmaeedo

                Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                Comando.Transaction = myTrans
                Comando.ExecuteNonQuery()

            End If

            '=========================================== OBSERVACIONES ==================================================================

            Dim Tbl_Observaciones As DataTable = _Bd_Documento.Tables("Observaciones_Doc")

            With Tbl_Observaciones
                _Obdo = .Rows(0).Item("Observaciones")
                _Cpdo = .Rows(0).Item("Forma_pago")
                _Ocdo = .Rows(0).Item("Orden_compra")


                For i = 0 To 34

                    Dim Campo As String = "Obs" & i + 1
                    Obs(i) = .Rows(0).Item(Campo)

                Next

            End With

            Consulta_sql = "INSERT INTO KASIEDOB (IDMAEEDO,OBDO,CPDO,OCDO,DIENDESP,TEXTO1,TEXTO2,TEXTO3,TEXTO4,TEXTO5,TEXTO6," & vbCrLf & _
                           "TEXTO7,TEXTO8,TEXTO9,TEXTO10,TEXTO11,TEXTO12,TEXTO13,TEXTO14,TEXTO15,CARRIER,BOOKING,LADING,AGENTE," & vbCrLf & _
                           "MEDIOPAGO,TIPOTRANS,KOPAE,KOCIE,KOCME,FECHAE,HORAE,KOPAD,KOCID,KOCMD,FECHAD,HORAD,OBDOEXPO,MOTIVO," & vbCrLf & _
                           "TEXTO16,TEXTO17,TEXTO18,TEXTO19,TEXTO20,TEXTO21,TEXTO22,TEXTO23,TEXTO24,TEXTO25,TEXTO26,TEXTO27," & vbCrLf & _
                           "TEXTO28,TEXTO29,TEXTO30,TEXTO31,TEXTO32,TEXTO33,TEXTO34,TEXTO35,PLACAPAT) VALUES " & vbCrLf & _
                           "(" & _Idmaeedo & ",'" & _Obdo & "','" & _Cpdo & "','" & _Ocdo & "','','" & Obs(0) & "','" & Obs(1) & _
                           "','" & Obs(2) & "','" & Obs(3) & "','" & Obs(4) & "','" & Obs(5) & "','" & Obs(6) & "','" & Obs(7) & _
                           "','" & Obs(8) & "','" & Obs(9) & "','" & Obs(10) & "','" & Obs(11) & "','" & Obs(12) & "','" & Obs(13) & _
                           "','" & Obs(14) & "','','','','','','','','','',GETDATE(),'','','','',GETDATE(),'','','','" & Obs(15) & _
                           "','" & Obs(16) & "','" & Obs(17) & "','" & Obs(18) & "','" & Obs(19) & _
                           "','" & Obs(20) & "','" & Obs(21) & "','" & Obs(22) & "','" & Obs(23) & "','" & Obs(24) & _
                           "','" & Obs(25) & "','" & Obs(26) & "','" & Obs(27) & "','" & Obs(28) & "','" & Obs(29) & _
                           "','" & Obs(30) & "','" & Obs(31) & "','" & Obs(32) & "','" & Obs(33) & "','" & Obs(34) & _
                           "','')"

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            ' ====================================================================================================================================

            myTrans.Commit()

            Return _Idmaeedo

        Catch ex As Exception

            'Dim _Erro_VB As String = ex.Message & vbCrLf & ex.StackTrace & vbCrLf & _
            '                        "Código: " & _Koprct
            'Clipboard.SetText(_Erro_VB)

            'My.Computer.FileSystem.WriteAllText("ArchivoSalida", ex.Message & vbCrLf & ex.StackTrace, False)
            'MessageBoxEx.Show(ex.Message, "Error", _
            '                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
            myTrans.Rollback()

            'MessageBoxEx.Show("Transaccion desecha", "Problema", _
            '                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
            'SQL_ServerClass.CerrarConexion(cn2)
            Return 0
        Finally
            _Sql.Sb_Cerrar_Conexion(cn2)
        End Try

    End Function

#End Region


#Region "FUNCION CREAR DOCUMENTO RANDOM CASI DEFINITIVO"

    Function Fx_Crear_Documento_En_BakApp_Casi(ByVal Bd_Documento As DataSet, _
                                               Optional ByVal EsAjuste As Boolean = False) As Integer

        Dim _Id_DocEnc As Integer

        Dim _Modalidad As String
        Dim _TipoDoc As String
        Dim _NroDocumento As String
        Dim _Es_ValeTransitorio As Boolean
        Dim _Es_Documento_Electronico As Boolean
        Dim _Sucursal As String
        Dim _CodFuncionario As String
        Dim _CodEntidad As String
        Dim _CodSucEntidad As String
        Dim _FechaEmision As String
        Dim _ListaPrecios As String
        Dim _CantTotal As String
        Dim _CantDesp As String
        Dim _Centro_Costo As String
        Dim _Moneda_Doc As String
        Dim _DocEn_Neto_Bruto As String
        Dim _Valor_Dolar As String
        Dim _Timodo As String
        Dim _TotalNetoDoc As String
        Dim _TotalIvaDoc As String
        Dim _TotalIlaDoc As String
        Dim _TotalBrutoDoc As String
        Dim _Fecha_1er_Vencimiento As String
        Dim _FechaUltVencimiento As String
        Dim _FechaRecepcion As String

        Dim _Nombre_Entidad As String
        Dim _Cuotas As String
        Dim _Dias_1er_Vencimiento As String
        Dim _Dias_Vencimiento As String
        Dim _CodEntidadFisica As String
        Dim _CodSucEntidadFisica As String
        Dim _Nombre_Entidad_Fisica As String
        Dim _Contacto_Ent As String
        Dim _NomFuncionario As String
        Dim _Es_Electronico As String
        Dim _TipoMoneda As String
        Dim _Vizado As String



        Dim _NroLinea As String
        Dim _Bodega As String
        Dim _Codigo As String
        Dim _Descripcion As String
        Dim _Rtu As String
        Dim _Sucursal_Li As String
        Dim _CodVendedor As String
        Dim _Tict As String
        Dim _Prct As String
        Dim _Tipr As String
        Dim _Cantidad As String
        Dim _Precio As String
        Dim _CantUd1 As String
        Dim _CantUd2 As String
        Dim _Lincondest As String
        Dim _CDespUd1 As String
        Dim _CDespUd2 As String
        Dim _Estado As String
        Dim _UnTrans As String
        Dim _UdTrans As String
        Dim _Ud01PR As String
        Dim _Ud02PR As String
        Dim _CodLista As String
        Dim _Moneda As String
        Dim _Tipo_Moneda As String
        Dim _Tipo_Cambio As String
        Dim _DescuentoPorc As String
        Dim _DescuentoValor As String
        Dim _PorIla As String
        Dim _Operacion As String
        Dim _Potencia As String

        'Dim Campo As String = "Precio"

        Dim _PrecioNetoUd As String
        Dim _PrecioBrutoUd As String
        Dim _PrecioNetoUdLista As String
        Dim _PrecioBrutoUdLista As String
        Dim _PorIva As String
        Dim _NroDscto As String
        Dim _NroImpuestos As String
        Dim _DsctoRealPorc As String
        Dim _DsctoRealValor As String
        Dim _DsctoNeto As String
        Dim _DsctoBruto As String
        Dim _ValSubNetoLinea As String
        Dim _StockBodega As String
        Dim _UbicacionBod As String
        Dim _SubTotal As String
        Dim _ValNetoLinea As String
        Dim _ValIlaLinea As String
        Dim _ValIvaLinea As String
        Dim _ValBrutoLinea As String
        Dim _FechaEmision_Linea As String
        Dim _FechaRecepcion_Linea As String
        Dim _Observa As String
        Dim _PrecioNetoRealUd1 As String
        Dim _PrecioNetoRealUd2 As String
        Dim _PmLinea As String
        Dim _PmSucLinea As String
        Dim _CodigoProv As String
        Dim _TipoValor As String
        Dim _ValVtaDescMax As String
        Dim _ValVtaStockInf As String
        Dim _DescMaximo As String

        Dim _Idmaeedo_Dori As String
        Dim _Idmaeddo_Dori As String
        Dim _CantUd1_Dori As String
        Dim _CantUd2_Dori As String

        Dim _Tidopa As String
        Dim _NudoPa As String

        Dim _DsctoGlobalSuperado As String
        Dim _Tiene_Dscto As String
        Dim _CantidadCalculo As String
        Dim _PrecioCalculo As String
        Dim _OCCGenerada As String
        Dim _Bloqueapr As String
        Dim _CodFunAutoriza As String
        Dim _CodPermiso As String
        Dim _Nuevo_Producto As String
        Dim _Solicitado_bodega As String


        Dim myTrans As SqlClient.SqlTransaction
        Dim Comando As SqlClient.SqlCommand
        Dim cn2 As New SqlConnection


        Dim Tbl_Encabezado As DataTable = Bd_Documento.Tables("Encabezado_Doc")


        Try

            _Sql.Sb_Abrir_Conexion(cn2)

            With Tbl_Encabezado.Rows(0)

                _Modalidad = .Item("Modalidad")

                _TipoDoc = .Item("TipoDoc")
                _NroDocumento = Traer_Numero_Documento(_Tido, .Item("NroDocumento"), _Modalidad, _Empresa)

                If String.IsNullOrEmpty(_NroDocumento) Then
                    Return 0
                End If

                .Item("NroDocumento") = _NroDocumento
                '_Nudo = .Item("NroDocumento")

                If String.IsNullOrEmpty(Trim(_NroDocumento)) Then
                    Return 0
                End If

                _Empresa = .Item("Empresa").ToString

                _Sucursal = .Item("Sucursal")
                _CodFuncionario = .Item("CodFuncionario")


                _CodEntidad = .Item("CodEntidad")
                _CodSucEntidad = .Item("CodSucEntidad")

                _FechaEmision = Format(.Item("FechaEmision"), "yyyyMMdd")
                _ListaPrecios = .Item("ListaPrecios")
                _CantTotal = De_Num_a_Tx_01(.Item("CantTotal"), 5)
                _CantDesp = De_Num_a_Tx_01(.Item("CantDesp"), 5)

                _Centro_Costo = NuloPorNro(.Item("Centro_Costo"), "")
                _Moneda_Doc = .Item("Moneda_Doc")
                _DocEn_Neto_Bruto = .Item("DocEn_Neto_Bruto")
                _Valor_Dolar = De_Num_a_Tx_01(.Item("Valor_Dolar"), False, 5)
                _Tipo_Moneda = .Item("TipoMoneda")

                Dim _TotalNetoDoc_2 = .Item("TotalNetoDoc")
                Dim _TotalIvaDoc_2 = .Item("TotalIvaDoc")
                Dim _TotalIlaDoc_2 = .Item("TotalIlaDoc")
                Dim _TotalBrutoDoc_2 = .Item("TotalBrutoDoc")


                _TotalNetoDoc = De_Num_a_Tx_01(.Item("TotalNetoDoc"), False, 5)
                _TotalIvaDoc = De_Num_a_Tx_01(.Item("TotalIvaDoc"), False, 5)
                _TotalIlaDoc = De_Num_a_Tx_01(.Item("TotalIlaDoc"), False, 5)
                _TotalBrutoDoc = De_Num_a_Tx_01(.Item("TotalBrutoDoc"), False, 5)

                _Fecha_1er_Vencimiento = Format(.Item("Fecha_1er_Vencimiento"), "yyyyMMdd")
                _FechaUltVencimiento = Format(.Item("FechaUltVencimiento"), "yyyyMMdd")

                _FechaRecepcion = Format(.Item("FechaRecepcion"), "yyyyMMdd")

                _Nombre_Entidad = .Item("Nombre_Entidad").ToString
                _Cuotas = .Item("Cuotas").ToString
                _Dias_1er_Vencimiento = .Item("Dias_1er_Vencimiento").ToString
                _Dias_Vencimiento = .Item("Dias_Vencimiento").ToString
                _CodEntidadFisica = NuloPorNro(.Item("CodEntidadFisica").ToString, "")
                _CodSucEntidadFisica = NuloPorNro(.Item("CodSucEntidadFisica").ToString, "")
                _Nombre_Entidad_Fisica = .Item("Nombre_Entidad_Fisica").ToString
                _Contacto_Ent = NuloPorNro(.Item("Contacto_Ent").ToString, "")
                _NomFuncionario = .Item("NomFuncionario").ToString
                _Es_Electronico = 1 '(CInt(.Item("Es_Electronico")) * -1)
                _TipoMoneda = .Item("TipoMoneda").ToString
                _Vizado = CInt(.Item("Vizado")) * -1
                '_FechaRecepcion_Linea = Format(.Item("FechaRecepcion"), "yyyyMMdd")

            End With

            '------------------------------------------------------------------------------------------------------------

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "BakApp")
            Return 0
        End Try


        Try

            myTrans = cn2.BeginTransaction()


            Consulta_sql = "INSERT INTO " & _Global_BaseBk & "Zw_Casi_DocEnc (Empresa,TipoDoc,NroDocumento,CodEntidad,CodSucEntidad )" & vbCrLf & _
                           "VALUES ( '" & _Empresa & "','" & _TipoDoc & "','" & _NroDocumento & _
                           "','" & _CodEntidad & "','" & _CodSucEntidad & "')"

            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            Comando = New SqlCommand("SELECT @@IDENTITY AS 'Identity'", cn2)
            Comando.Transaction = myTrans
            Dim dfd1 As SqlDataReader = Comando.ExecuteReader()
            While dfd1.Read()
                _Id_DocEnc = dfd1("Identity")
            End While
            dfd1.Close()

            Bd_Documento.Tables("Detalle_Doc").Dispose()
            Dim Tbl_Detalle As DataTable = Bd_Documento.Tables("Detalle_Doc")

            Dim Contador As Integer = 1

            For Each FDetalle As DataRow In Tbl_Detalle.Rows
                Dim Estado As DataRowState = FDetalle.RowState
                If Not Estado = DataRowState.Deleted Then

                    With FDetalle



                        Id_Linea = .Item("Id")


                        _NroLinea = numero_(Contador, 5)

                        _Bodega = .Item("Bodega")
                        _Codigo = .Item("Codigo")
                        _Descripcion = .Item("Descripcion")
                        _Rtu = De_Num_a_Tx_01(.Item("Rtu"), False, 5)
                        _Sucursal_Li = .Item("Sucursal")
                        _CodVendedor = _Funcionario 'FUNCIONARIO '.Item("CodVendedor") ' FUNCIONARIO ' Codigo de funcionario

                        _Tict = .Item("Tict")
                        _Prct = .Item("Prct")
                        _Tipr = .Item("Tipr")

                        _Cantidad = De_Num_a_Tx_01(.Item("Cantidad"), False, 5)
                        _Precio = De_Num_a_Tx_01(.Item("Precio"), False, 5)

                        _CantUd1 = De_Num_a_Tx_01(.Item("CantUd1"), False, 5) ' Cantidad de la primera unidad
                        _CantUd2 = De_Num_a_Tx_01(.Item("CantUd2"), False, 5) ' Cantidad de la segunda unidad

                        _Lincondest = CInt(.Item("Lincondest")) * -1

                        'CantidadTotal = CantidadTotal + Val(CAPRCO1)

                        ' _CDespUd1, _CDespUd2 

                        If CBool(_Lincondesp) Then
                            _CDespUd1 = _CantUd1 ' Cantidad que mueve Stock Fisico, según el producto.
                            _CDespUd2 = _CantUd2 ' Cantidad que mueve Stock Fisico, según el producto.
                        Else
                            _CDespUd1 = De_Num_a_Tx_01(NuloPorNro(.Item("CDespUd1"), 0), False, 5) ' Cantidad que mueve Stock Fisico, según el producto.
                            _CDespUd2 = De_Num_a_Tx_01(NuloPorNro(.Item("CDespUd2"), 0), False, 5) ' Cantidad que mueve Stock Fisico, según el producto.
                        End If

                        _Estado = NuloPorNro(.Item("Estado"), "")

                        Dim _CaprexUd1 = 0 ' Cantidad  
                        Dim _CaprexUd2 = 0
                        Dim _CaprncUd1 = 0 ' Cantidad de Nota de credito
                        Dim _CaprncUd2 = 0

                        _UnTrans = .Item("UnTrans")  ' Unidad de la transaccion
                        _UdTrans = .Item("UdTrans")  ' Unidad de la transaccion

                        _Ud01PR = .Item("Ud01PR")
                        _Ud02PR = .Item("Ud02PR")
                        _CodLista = .Item("CodLista") 'LISTADEPRECIO
                        _Moneda = .Item("Moneda") 'trae_dato(tb, cn1, "KOMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        _Tipo_Moneda = .Item("Tipo_Moneda") 'trae_dato(tb, cn1, "TIMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        _Tipo_Cambio = De_Num_a_Tx_01(.Item("Tipo_Cambio"), False, 5) 'De_Num_a_Tx_01(trae_dato(tb, cn1, "VAMO", "TABMO", "NOKOMO LIKE '%PESO%'"), False, 5)
                        ' _PrecioNetoUd  = De_Num_a_Tx_01(.Item("PrecioNetoUd"), False, 5)
                        _DescuentoPorc = De_Num_a_Tx_01(.Item("DescuentoPorc"), False, 5)
                        _DescuentoValor = De_Num_a_Tx_01(.Item("DescuentoValor"), False, 5)

                        _PorIla = De_Num_a_Tx_01(.Item("PorIla"), False, 5)

                        _Operacion = .Item("Operacion")
                        _Potencia = De_Num_a_Tx_01(.Item("Potencia"), False, 5)

                        Dim Campo = "Precio"

                        _PrecioNetoUd = De_Num_a_Tx_01(.Item("PrecioNetoUd"), False, 5)
                        _PrecioBrutoUd = De_Num_a_Tx_01(.Item("PrecioBrutoUd"), False, 5)
                        _PrecioNetoUdLista = De_Num_a_Tx_01(NuloPorNro(Of Double)(.Item("PrecioNetoUdLista"), 0), False, 5)
                        _PrecioBrutoUdLista = De_Num_a_Tx_01(.Item("PrecioBrutoUdLista"), False, 0) ' PRECIO BRUTO DE LA LISTA

                        _PorIva = De_Num_a_Tx_01(.Item("PorIva"), True)
                        _NroDscto = De_Num_a_Tx_01(.Item("NroDscto"), True)

                        _NroImpuestos = De_Num_a_Tx_01(.Item("NroImpuestos"), True)

                        _DsctoRealPorc = De_Num_a_Tx_01(.Item("DsctoRealPorc"), False, 5)
                        _DsctoRealValor = De_Num_a_Tx_01(.Item("DsctoRealValor"), False, 5)
                        _DsctoNeto = De_Num_a_Tx_01(.Item("DsctoNeto"), False, 5)
                        _DsctoBruto = De_Num_a_Tx_01(.Item("DsctoBruto"), False, 5) 'ValDscto
                        _ValSubNetoLinea = 0 'De_Num_a_Tx_01(.Item("ValSubNetoLinea"), False, 5)

                        _StockBodega = De_Num_a_Tx_01(.Item("StockBodega"), False, 5)
                        _UbicacionBod = NuloPorNro(.Item("UbicacionBod"), "")

                        _SubTotal = De_Num_a_Tx_01(.Item("SubTotal"), False, 5)

                        _ValNetoLinea = De_Num_a_Tx_01(.Item("ValNetoLinea"), False, 5)
                        _ValIlaLinea = De_Num_a_Tx_01(.Item("ValIlaLinea"), False, 5)
                        _ValIvaLinea = De_Num_a_Tx_01(.Item("ValIvaLinea"), False, 5)
                        _ValBrutoLinea = De_Num_a_Tx_01(.Item("ValBrutoLinea"), False, 5)

                        _FechaEmision_Linea = _Feemdo 'Format(Now.Date, "yyyyMMdd") '""20121127"
                        _FechaRecepcion_Linea = _Feemdo 'Format(Now.Date, "yyyyMMdd")

                        _Observa = NuloPorNro(.Item("Observa"), "")

                        _PrecioNetoRealUd1 = De_Num_a_Tx_01(.Item("PrecioNetoRealUd1"), False, 5)
                        _PrecioNetoRealUd2 = De_Num_a_Tx_01(.Item("PrecioNetoRealUd2"), False, 5)
                        _PmLinea = De_Num_a_Tx_01(NuloPorNro(.Item("PmLinea"), 0), False, 5)
                        _PmSucLinea = De_Num_a_Tx_01(NuloPorNro(.Item("PmSucLinea"), 0), False, 5)

                        _CodigoProv = NuloPorNro(.Item("CodigoProv"), "")
                        _TipoValor = NuloPorNro(.Item("TipoValor"), "")

                        _ValVtaDescMax = CInt(.Item("ValVtaDescMax")) * -1
                        _ValVtaStockInf = CInt(.Item("ValVtaStockInf")) * -1

                        _DescMaximo = De_Num_a_Tx_01(NuloPorNro(.Item("DescMaximo"), 0), False, 5)

                        _Idmaeedo_Dori = .Item("Idmaeedo_Dori")
                        _Idmaeddo_Dori = .Item("Idmaeddo_Dori")

                        Dim CantUd1_Dori As Double = .Item("CantUd1_Dori")
                        Dim CantUd2_Dori As Double = .Item("CantUd2_Dori")

                        _CantUd1_Dori = De_Num_a_Tx_01(CantUd1_Dori, False, 5)
                        _CantUd2_Dori = De_Num_a_Tx_01(CantUd2_Dori, False, 5)

                        If String.IsNullOrEmpty(_Idmaeddo_Dori) Then _Idmaeddo_Dori = 0

                        _Tidopa = NuloPorNro(.Item("Tidopa"), "")
                        _NudoPa = NuloPorNro(.Item("NudoPa"), "")


                        _DsctoGlobalSuperado = 0 'CInt(.Item("DsctoGlobalSuperado")) * -1
                        _Tiene_Dscto = CInt(.Item("Tiene_Dscto")) * -1
                        _CantidadCalculo = De_Num_a_Tx_01(NuloPorNro(.Item("CantidadCalculo"), 0), False, 5)
                        _PrecioCalculo = De_Num_a_Tx_01(NuloPorNro(.Item("PrecioCalculo"), 0), False, 5)
                        _OCCGenerada = CInt(.Item("Tiene_Dscto")) * -1
                        _Bloqueapr = NuloPorNro(.Item("Bloqueapr"), "")
                        _CodFunAutoriza = NuloPorNro(.Item("CodFunAutoriza"), "")
                        _CodPermiso = NuloPorNro(.Item("CodPermiso"), "")
                        _Nuevo_Producto = CInt(.Item("Nuevo_Producto")) * -1
                        _Solicitado_bodega = CInt(.Item("Nuevo_Producto")) * -1



                        If Not String.IsNullOrEmpty(Trim(_Tict)) Then
                            'Dim TipoValor = _TipoValor 'trae_dato(tb, cn1, "TipoValor", "ZW_Bkp_Configuracion")

                            If _TipoValor = "N" Then
                                _CantUd1 = _ValNetoLinea * -1
                                '_ValBrutoLinea = De_Txt_a_Num_01((_Vabrli), 0) * -1
                            Else
                                '_ValNetoLinea = De_Num_a_Tx_01(De_Txt_a_Num_01((_ValNetoLinea), 5) * -1, False, 5)
                                _CantUd1 = _ValBrutoLinea * -1
                            End If

                            _CantUd2 = 0
                            _Caprad2 = 0
                            _Cafaco = 0
                            _PrecioNetoUdLista = 0
                            _PrecioNetoUd = 0
                            _PrecioBrutoUdLista = 0
                            _PrecioBrutoUd = 0
                            _Prct = 1
                            _PmLinea = 0
                            _PmSucLinea = 0
                            _Lincondesp = 1
                            _NroDscto = 0
                            _Estado = "C"
                        Else
                            _Cafaco = _CantUd1
                        End If




                        If _TipoDoc <> "COV" Then

                            If _TipoDoc = "OCC" Then

                                'Consulta_sql = "UPDATE MAEST SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," & _
                                '                                "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf & _
                                '                                "WHERE EMPRESA='" & _Empresa & _
                                '                                "' AND KOSU='" & _Sulido & _
                                '                                "' AND KOBO='" & _Bosulido & _
                                '                                "' AND KOPR='" & _Koprct & "'" & vbCrLf & _
                                '               "UPDATE MAEPREM SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," & _
                                '                                "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf & _
                                '                                "WHERE EMPRESA='" & _Empresa & _
                                '                                "' AND KOPR='" & _Koprct & "'" & vbCrLf & _
                                '               "UPDATE MAEPR SET STOCNV1C = STOCNV1C +" & _Caprco1 & "," & _
                                '                                "STOCNV2C = STOCNV2C + " & _Caprco2 & vbCrLf & _
                                '                                "WHERE KOPR='" & _Koprct & "'"

                                'Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                'Comando.Transaction = myTrans
                                'Comando.ExecuteNonQuery()

                            ElseIf _Tido = "NVV" Then

                                'Consulta_sql = "UPDATE MAEST SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," & _
                                '                               "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf & _
                                '                               "WHERE EMPRESA='" & _Empresa & _
                                '                               "' AND KOSU='" & _Sulido & _
                                '                               "' AND KOBO='" & _Bosulido & _
                                '                               "' AND KOPR='" & _Koprct & "'" & vbCrLf & _
                                '              "UPDATE MAEPREM SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," & _
                                '                               "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf & _
                                '                               "WHERE EMPRESA='" & _Empresa & _
                                '                               "' AND KOPR='" & _Koprct & "'" & vbCrLf & _
                                '              "UPDATE MAEPR SET STOCNV1 = STOCNV1 +" & _Caprco1 & "," & _
                                '                               "STOCNV2 = STOCNV2 + " & _Caprco2 & vbCrLf & _
                                '                               "WHERE KOPR='" & _Koprct & "'"

                                'Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                'Comando.Transaction = myTrans
                                'Comando.ExecuteNonQuery()

                            Else

                                If _Lincondesp Then

                                    'Consulta_sql = "UPDATE MAEPREM SET" & vbCrLf & _
                                    '              "STFI1 = STFI1 - " & _Caprco1 & ",STFI2 =  - " & _Caprco2 & vbCrLf & _
                                    '              "WHERE EMPRESA = '" & _Empresa & "' AND KOPR = '" & _Koprct & "'" & vbCrLf & _
                                    '              "UPDATE MAEPR SET  STFI1 = STFI1 - " & _Caprco1 & ",STFI2 = - " & _Caprco2 & vbCrLf & _
                                    '              "WHERE KOPR = '" & _Koprct & "'" & vbCrLf & _
                                    '              "UPDATE MAEST SET STFI1 = STFI1 - " & _Caprco1 & ",STFI2 =  - " & _Caprco2 & vbCrLf & _
                                    '              "WHERE EMPRESA='" & _Empresa & "' AND KOSU='" & _Sudo & _
                                    '              "' AND KOBO='" & _Bosulido & "' AND KOPR='" & _Koprct & "'"

                                    '_Caprad1 = _Caprco1
                                    '_Caprad2 = _Caprco2


                                Else

                                    'Consulta_sql = "UPDATE MAEPREM SET" & vbCrLf & _
                                    '               "STDV1 = STDV1 + " & _Caprco1 & ",STDV2 =  + " & _Caprco2 & vbCrLf & _
                                    '               "WHERE EMPRESA = '" & _Empresa & "' AND KOPR = '" & _Koprct & "'" & vbCrLf & _
                                    '               "UPDATE MAEPR SET  STDV1 = STDV1 + " & _Caprco1 & ",STDV2 = + " & _Caprco2 & vbCrLf & _
                                    '               "WHERE KOPR = '" & _Koprct & "'" & vbCrLf & _
                                    '               "UPDATE MAEST SET STDV1 = STDV1 + " & _Caprco1 & ",STDV2 =  + " & _Caprco2 & vbCrLf & _
                                    '               "WHERE EMPRESA='" & _Empresa & "' AND KOSU='" & _Sudo & _
                                    '               "' AND KOBO='" & _Bosulido & "' AND KOPR='" & _Koprct & "'"

                                    '_Caprad1 = 0
                                    '_Caprad2 = 0


                                End If

                                'Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                                'Comando.Transaction = myTrans
                                'Comando.ExecuteNonQuery()

                            End If
                        End If


                        '_Mopppr = .Item("Moneda") 'trae_dato(tb, cn1, "KOMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        '_Timopppr = .Item("Tipo_Moneda") 'trae_dato(tb, cn1, "TIMO", "TABMO", "NOKOMO LIKE '%PESO%'")
                        '_Tamopppr = .Item("Tipo_Cambio") 'De_Num_a_Tx_01(trae_dato(tb, cn1, "VAMO", "TABMO", "NOKOMO LIKE '%PESO%'"), False, 5)

                        Consulta_sql = _
                              "INSERT INTO MAEDDO (Id_DocEnc,Empresa,TIDO,NUDO,ENDO,SUENDO,LILG,NULIDO," & vbCrLf & _
                              "SULIDO,BOSULIDO,LUVTLIDO,KOFULIDO,TIPR,TICT,PRCT,KOPRCT,UDTRPR,RLUDPR,CAPRCO1," & vbCrLf & _
                              "UD01PR,CAPRCO2,UD02PR,CAPRAD1,CAPRAD2,KOLTPR,MOPPPR,TIMOPPPR,TAMOPPPR,NUIMLI,NUDTLI," & vbCrLf & _
                              "PODTGLLI,POIMGLLI,VAIMLI,VADTNELI,VADTBRLI,POIVLI,VAIVLI,VANELI,VABRLI,TIGELI," & vbCrLf & _
                              "FEEMLI,FEERLI,PPPRNELT,PPPRNE,PPPRBRLT,PPPRBR,PPPRPM,PPPRNERE1,PPPRNERE2,CAFACO," & vbCrLf & _
                              "FVENLOTE,FCRELOTE,NOKOPR,ALTERNAT,TASADORIG,CUOGASDIF,LINCONDESP,OPERACION,POTENCIA,ESLIDO,OBSERVA)" & vbCrLf & _
                       "VALUES (" & _Idmaeedo & ",'',0,'" & _Empresa & "','" & _Tido & "','" & _Nudo & "','" & _Endo & _
                              "','" & _Suendo & "','SI','" & _Nulido & "','" & _Sudo & "','" & _Bosulido & _
                              "','" & _Luvtven & "','" & _Kofudo & "','" & _Tipr & "','" & _Tict & "'," & CInt(_Prct) & ",'" & _Koprct & _
                              "'," & _Udtrpr & "," & _Rludpr & "," & _Caprco1 & ",'" & _Ud01PR & "'," & _Caprco2 & _
                              ",'" & _Ud02PR & "'," & _Caprad1 & "," & _Caprad2 & ",'TABPP" & _Koltpr & "'" & _
                              ",'" & _Mopppr & "','" & _Timopppr & "'," & _Tamopppr & _
                              "," & _Nuimli & "," & _Nudtli & "," & _Podtglli & "," & _Poimglli & "," & _Vaimli & _
                              "," & _Vadtneli & "," & _Vadtbrli & "," & _Poivli & "," & _Vaivli & "," & _Vaneli & _
                              "," & _Vabrli & ",'I','" & _Feemli & "','" & _Feerli & "'," & _Ppprnelt & "," & _Ppprne & _
                              "," & _Ppprbrlt & "," & _Ppprbr & "," & _Ppprpm & "," & _Ppprnere1 & "," & _Ppprnere2 & _
                              "," & _Cafaco & ",NULL,NULL,'" & _Nokopr & "','" & _Alternat & "',1.00000,0," & CInt(_Lincondesp) & _
                              ",'" & _Operacion & "'," & _Potencia & ",'" & _Eslido & "',' " & _Observa & "')"


                        Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Casi_DocDet " & vbCrLf & _
                                       "(Id_DocEnc,Sucursal,Bodega,UnTrans,Lincondest,NroLinea,Codigo,CodigoProv," & _
                                       "UdTrans,Cantidad,TipoValor,Precio,DescuentoPorc,DescuentoValor,Descripcion," & _
                                       "PrecioNetoUd,PrecioNetoUdLista,PrecioBrutoUd,PrecioBrutoUdLista,Rtu,Ud01PR,CantUd1," & _
                                       "CDespUd1,CaprexUd1,CaprncUd1,CodVendedor,Prct,Tict,Tipr,DsctoNeto,DsctoBruto,Ud02PR," & _
                                       "CantUd2,CDespUd2,CaprexUd2,CaprncUd12,ValVtaDescMax,ValVtaStockInf,CodLista," & _
                                       "DescMaximo,NroDscto,NroImpuestos,PorIva,PorIla,ValIvaLinea,ValIlaLinea,ValSubNetoLinea," & _
                                       "ValNetoLinea,ValBrutoLinea,PmLinea,PmSucLinea,PrecioNetoRealUd1,PrecioNetoRealUd2," & _
                                       "FechaEmision_Linea,FechaRecepcion_Linea," & _
                                       "Idmaeedo_Dori,Idmaeddo_Dori,CantUd1_Dori,CantUd2_Dori,Estado,Tidopa,NudoPa,SubTotal," & _
                                       "StockBodega,UbicacionBod,DsctoRealPorc,DsctoRealValor,DsctoGlobalSuperado,Tiene_Dscto,CantidadCalculo," & _
                                       "Operacion,Potencia,PrecioCalculo,OCCGenerada,Bloqueapr,Observa,CodFunAutoriza," & _
                                       "CodPermiso,Nuevo_Producto,Solicitado_bodega,Moneda,Tipo_Moneda,Tipo_Cambio) Values" & vbCrLf & _
                                       "(" & _Id_DocEnc & ",'" & _Sucursal & "','" & _Bodega & "'," & _UnTrans & "," & _Lincondest & _
                                       ",'" & _NroLinea & "','" & _Codigo & "','" & _CodigoProv & "','" & _UdTrans & _
                                       "'," & _Cantidad & ",'" & _TipoValor & "'," & _Precio & "," & _DescuentoPorc & _
                                       "," & _DescuentoValor & ",'" & _Descripcion & "'," & _PrecioNetoUd & _
                                       "," & _PrecioNetoUdLista & "," & _PrecioBrutoUd & "," & _PrecioBrutoUdLista & _
                                       "," & _Rtu & ",'" & _Ud01PR & "'," & _CantUd1 & "," & _CDespUd1 & "," & _CaprexUd1 & _
                                       "," & _CaprncUd1 & ",'" & _CodVendedor & "'," & _Prct & ",'" & _Tict & "','" & _Tipr & _
                                       "'," & _DsctoNeto & "," & _DsctoBruto & ",'" & _Ud02PR & "'," & _CantUd2 & _
                                       "," & _CDespUd2 & "," & _CaprexUd2 & "," & _CaprncUd2 & "," & _ValVtaDescMax & _
                                       "," & _ValVtaStockInf & ",'" & _CodLista & "'," & _DescMaximo & "," & _NroDscto & _
                                       "," & _NroImpuestos & "," & _PorIva & "," & _PorIla & "," & _ValIvaLinea & _
                                       "," & _ValIlaLinea & "," & _ValSubNetoLinea & "," & _ValNetoLinea & _
                                       "," & _ValBrutoLinea & "," & _PmLinea & "," & _PmSucLinea & "," & _PrecioNetoRealUd1 & _
                                       "," & _PrecioNetoRealUd2 & ",'" & _FechaEmision & "','" & _FechaRecepcion & _
                                       "'," & _Idmaeedo_Dori & "," & _Idmaeddo_Dori & "," & _CantUd1_Dori & "," & _CantUd2_Dori & _
                                       ",'" & _Estado & "','" & _Tidopa & "','" & _NudoPa & _
                                       "'," & _SubTotal & "," & _StockBodega & ",'" & Trim(_UbicacionBod) & "'," & _DsctoRealPorc & _
                                       "," & _DsctoRealValor & "," & _DsctoGlobalSuperado & "," & _Tiene_Dscto & "," & _CantidadCalculo & _
                                       ",'" & _Operacion & "'," & _Potencia & "," & _PrecioCalculo & "," & _OCCGenerada & _
                                       ",'" & _Bloqueapr & "','" & _Observa & "','" & _CodFunAutoriza & "','" & _CodPermiso & _
                                       "'," & _Nuevo_Producto & "," & _Solicitado_bodega & ",'" & _Moneda & "','" & _Tipo_Moneda & _
                                       "'," & _Tipo_Cambio & ")"


                        Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                        Comando.Transaction = myTrans
                        Comando.ExecuteNonQuery()

                        'Dim _Reg As Integer = Cuenta_registros("INFORMATION_SCHEMA.COLUMNS", _
                        '                                      "COLUMN_NAME LIKE 'PPPRPMSUC' AND TABLE_NAME = 'MAEDDO'")

                        'If CBool(_Reg) Then
                        'Comando = New SqlCommand("SELECT @@IDENTITY AS 'Identity'", cn2)
                        'Comando.Transaction = myTrans
                        'dfd1 = Comando.ExecuteReader()
                        'Dim _Idmaeddo As Integer
                        'While dfd1.Read()
                        '_Idmaeddo = dfd1("Identity")
                        'End While
                        'dfd1.Close()

                        'Consulta_sql = "UPDATE MAEDDO SET PPPRPMSUC = " & _Ppprmsuc & vbCrLf & _
                        '               "WHERE IDMAEDDO = " & _Idmaeddo

                        'Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                        'Comando.Transaction = myTrans
                        'Comando.ExecuteNonQuery()
                        'End If

                    End With


                    ' TABLA DE IMPUESTOS

                    Dim Tbl_Impuestos As DataTable = Bd_Documento.Tables("Impuestos_Doc")

                    If Val(_NroImpuestos) > 0 And String.IsNullOrEmpty(Trim(_Tict)) Then

                        For Each FImpto As DataRow In Tbl_Impuestos.Select("Id = " & Id_Linea)

                            Dim _Poimli As String = De_Num_a_Tx_01(FImpto.Item("Poimli").ToString, False, 5)
                            Dim _Koimli As String = FImpto.Item("Koimli").ToString
                            Dim _Vaimli = De_Num_a_Tx_01(FImpto.Item("Vaimli").ToString, False, 5)

                            Consulta_sql = "INSERT INTO " & _Global_BaseBk & "Zw_DocImp (Id_DocEnc,Nulido,Koimli,Poimli,Vaimli,Lilg) VALUES " & vbCrLf & _
                                           "(" & _Id_DocEnc & ",'" & _Nulido & "','" & _Koimli & "'," & _Poimli & "," & _Vaimli & ",'')"

                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                            '-- 3RA TRANSACCION--------------------------------------------------------------------
                        Next
                    End If



                    ' TABLA DE DESCUENTOS
                    Dim Tbl_Descuentos As DataTable = Bd_Documento.Tables("Descuentos_Doc")
                    _Nudtli = Tbl_Descuentos.Rows.Count
                    If Val(_Nudtli) > 0 And String.IsNullOrEmpty(Trim(_Tict)) Then

                        For Each FDscto As DataRow In Tbl_Descuentos.Select("Id = " & Id_Linea)

                            Dim _Podt = De_Num_a_Tx_01(FDscto.Item("Podt").ToString, False, 5)
                            Dim _Vadt = De_Num_a_Tx_01(FDscto.Item("Vadt").ToString, False, 5)

                            Consulta_sql = "INSERT INTO " & _Global_BaseBk & "Zw_Casi_DocDsc (Id_DocEnc,Nulido,Kodt,Podt,Vadt)" & vbCrLf & _
                                   "values (" & _Id_DocEnc & ",'" & _NroLinea & "','D_SIN_TIPO'," & _Podt & "," & _Vadt & ")"


                            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
                            Comando.Transaction = myTrans
                            Comando.ExecuteNonQuery()

                            '-- 4TA TRANSACCION--------------------------------------------------------------------
                        Next
                    End If

                    Contador += 1
                End If
            Next



            'TABLA DE VENCIMIENTOS

            'Dim Tbl_Vencimientos As DataTable = Bd_Documento.Tables("Vencimientos_Doc")

            '_Nuvedo = Tbl_Vencimientos.Rows.Count

            'For Each FVencimientos As DataRow In Tbl_Vencimientos.Rows

            'Dim _Feve As String = Format(FVencimientos.Item("Fecha_Vencimiento"), "yyyyMMdd")
            'Dim _Espgve As String = String.Empty 'FilaX.Item("Estado_Pago").ToString
            'Dim _Vave As String = De_Num_a_Tx_01(FVencimientos.Item("Valor_Vencimiento").ToString, False, 5)
            'Dim _Vaabve As String = De_Num_a_Tx_01(FVencimientos.Item("Valor_Abonado").ToString, False, 5)
            'Dim _Archirst As String = String.Empty 'FilaX.Item("Archirst").ToString
            'Dim _porestpag As String = 0 'De_Num_a_Tx_01(FilaX.Item("Porestpag").ToString, False, 5)
            'Dim __Observa As String = String.Empty 'FilaX.Item("Archirst").ToString

            'Consulta_sql = "INSERT INTO MAEVEN (IDMAEEDO,FEVE,ESPGVE,VAVE,VAABVE,ARCHIRST,PORESTPAG,OBSERVA)" & vbCrLf & _
            '               "values (" & _Idmaeedo & ",'" & _Feve & "','" & _Espgve & "'," & _Vave & "," & _Vaabve & _
            '               ",'" & _Archirst & "'," & _porestpag & ",'" & __Observa & "')"

            'Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            'Comando.Transaction = myTrans
            'Comando.ExecuteNonQuery()

            'Next




            If _Nuvedo = 0 Then _Nuvedo = 1

            Dim _HoraGrab As String
            'Dim _HH_sistem As Date

            '_HH_sistem = FechaDelServidor()
            '_HoraGrab = _HH_sistem.Hour

            'Dim _HH, _MM, _SS As Double

            '_HH = _HH_sistem.Hour
            '_MM = _HH_sistem.Minute
            '_SS = _HH_sistem.Second

            If EsAjuste Then
                _Marca = 1 ' Generalmente se marcan las GRI o GDI que son por ajuste
                _Subtido = "AJU" ' Generalmente se Marcan las Guias de despacho o recibo
                '_HH = 23 : _MM = 59 : _SS = 59
            Else
                _Marca = String.Empty
                _Subtido = String.Empty
            End If

            _HoraGrab = Hora_Grab_fx(EsAjuste, _FechaEmision) 'Math.Round((_HH * 3600) + (_MM * 60) + _SS, 0)

            'Consulta_sql = "Declare @HoraGrab Int" & vbCrLf & _
            '               "set @HoraGrab = convert(money,substring(convert(varchar(20),getdate(),114),1,2)) * 3600 +" & vbCrLf & _
            '               "convert(money,substring(convert(varchar(20),getdate(),114),4,2)) * 60 + " & vbCrLf & _
            '               "Convert(money, substring(Convert(varchar(20), getdate(), 114), 7, 2))" & vbCrLf & vbCrLf & _


            Dim _Espgdo As String = "P"
            If _Tido = "OCC" Then _Espgdo = "S"
            ' HAY QUE PONER EL CAMPO TIPO DE MONEDA  "TIMODO"
            Consulta_sql = "UPDATE " & _Global_BaseBk & "Zw_DocEnc SET Sucursal = '" & _Sucursal & "',TIGEDO='I',SUDO='" & _Sudo & _
                           "',FechaEmision='" & _FechaEmision & "',CodFuncionario='" & _CodFuncionario & "',ESPGDO='" & _Espgdo & "',CAPRCO=" & _Caprco & _
                           ",CAPRAD=" & _Caprad & ",MEARDO = '" & _Meardo & "',MODO = '" & _Modo & _
                           "',TIMODO = '" & _Timodo & "',TAMODO = " & _Tamodo & ",VAIVDO = " & _Vaivdo & ",VAIMDO = " & _Vaimdo & vbCrLf & _
                           ",VANEDO = " & _Vanedo & ",VABRDO = " & _Vabrdo & ",FE01VEDO = '" & _Fe01vedo & _
                           "',FEULVEDO = '" & _Feulvedo & "',NUVEDO = " & _Nuvedo & ",FEER = '" & _Feer & _
                           "',KOTU = '1',LCLV = NULL,LAHORA = GETDATE(), DESPACHO = 1,HORAGRAB = " & _HoraGrab & _
                           ",FECHATRIB = NULL,NUMOPERVEN = 0,FLIQUIFCV = '" & _Feemdo & "',SUBTIDO = '" & _Subtido & _
                           "',MARCA = '" & _Marca & "',ESDO = '',NUDONODEFI = " & CInt(_Es_ValeTransitorio) & _
                           ",TIDOELEC = " & CInt(_Es_Documento_Electronico) & ",LUVTDO = '" & _Luvtven & "'" & vbCrLf & _
                           "WHERE IDMAEEDO=" & _Idmaeedo
            'Empresa,TipoDoc,NroDocumento,CodEntidad,CodSucEntidad

            Consulta_sql = "Update " & _Global_BaseBk & "Zw_Casi_DocEnc SET" & vbCrLf & _
                           "Modalidad = '" & _Modalidad & "'" & vbCrLf & _
                           ",Sucursal = '" & _Sucursal & "'" & vbCrLf & _
                           ",Nombre_Entidad = '" & _Nombre_Entidad & "'" & vbCrLf & _
                           ",FechaEmision = '" & _FechaEmision & "'" & vbCrLf & _
                           ",Fecha_1er_Vencimiento = '" & _Fecha_1er_Vencimiento & "'" & vbCrLf & _
                           ",FechaUltVencimiento = '" & _FechaUltVencimiento & "'" & vbCrLf & _
                           ",FechaRecepcion = '" & _FechaRecepcion & "'" & vbCrLf & _
                           ",Cuotas = '" & _Cuotas & "'" & vbCrLf & _
                           ",Dias_1er_Vencimiento = '" & _Dias_1er_Vencimiento & "'" & vbCrLf & _
                           ",Dias_Vencimiento = '" & _Dias_Vencimiento & "'" & vbCrLf & _
                           ",ListaPrecios = '" & _ListaPrecios & "'" & vbCrLf & _
                           ",CodEntidadFisica = '" & _CodEntidadFisica & "'" & vbCrLf & _
                           ",CodSucEntidadFisica = '" & _CodSucEntidadFisica & "'" & vbCrLf & _
                           ",Nombre_Entidad_Fisica = '" & _Nombre_Entidad_Fisica & "'" & vbCrLf & _
                           ",Contacto_Ent = '" & _Contacto_Ent & "'" & vbCrLf & _
                           ",CodFuncionario = '" & _CodFuncionario & "'" & vbCrLf & _
                           ",NomFuncionario = '" & _NomFuncionario & "'" & vbCrLf & _
                           ",Centro_Costo = '" & _Centro_Costo & "'" & vbCrLf & _
                           ",Moneda_Doc = '" & _Moneda_Doc & "'" & vbCrLf & _
                           ",Valor_Dolar = " & _Valor_Dolar & vbCrLf & _
                           ",TotalNetoDoc = " & _TotalNetoDoc & vbCrLf & _
                           ",TotalIvaDoc = " & _TotalIvaDoc & vbCrLf & _
                           ",TotalIlaDoc = " & _TotalIlaDoc & vbCrLf & _
                           ",TotalBrutoDoc = " & _TotalBrutoDoc & vbCrLf & _
                           ",CantTotal = " & _CantTotal & vbCrLf & _
                           ",CantDesp = " & _CantDesp & vbCrLf & _
                           ",DocEn_Neto_Bruto = '" & _DocEn_Neto_Bruto & "'" & vbCrLf & _
                           ",Es_ValeTransitorio = " & CInt(_Es_ValeTransitorio) & vbCrLf & _
                           ",Es_Electronico = " & _Es_Electronico & vbCrLf & _
                           ",TipoMoneda = '" & _TipoMoneda & "'" & vbCrLf & _
                           ",Vizado = '" & _Vizado & "'" & vbCrLf & _
                           "Where Id_DocEnc = " & _Id_DocEnc



            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()



            'Dim Reg As Integer = Cuenta_registros("INFORMATION_SCHEMA.COLUMNS", _
            '                                             "COLUMN_NAME LIKE 'LISACTIVA' AND TABLE_NAME = 'MAEEDO'")

            'If CBool(Reg) Then

            'Consulta_sql = "UPDATE MAEEDO SET LISACTIVA = 'TABPP" & _Lisactiva & "'" & vbCrLf & _
            '               "WHERE IDMAEEDO=" & _Idmaeedo

            'Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            'Comando.Transaction = myTrans
            'Comando.ExecuteNonQuery()

            'End If

            '=========================================== OBSERVACIONES ==================================================================

            Dim Tbl_Observaciones As DataTable = Bd_Documento.Tables("Observaciones_Doc")

            With Tbl_Observaciones

                _Obdo = Trim(.Rows(0).Item("Observaciones"))
                _Cpdo = Trim(.Rows(0).Item("Forma_pago"))
                _Ocdo = Trim(.Rows(0).Item("Orden_compra"))

                For i = 0 To 34

                    Dim Campo As String = "Obs" & i + 1
                    Obs(i) = .Rows(0).Item(Campo)

                Next

            End With

            Consulta_sql = "INSERT INTO MAEEDOOB (IDMAEEDO,OBDO,CPDO,OCDO,DIENDESP,TEXTO1,TEXTO2,TEXTO3,TEXTO4,TEXTO5,TEXTO6," & vbCrLf & _
                           "TEXTO7,TEXTO8,TEXTO9,TEXTO10,TEXTO11,TEXTO12,TEXTO13,TEXTO14,TEXTO15,CARRIER,BOOKING,LADING,AGENTE," & vbCrLf & _
                           "MEDIOPAGO,TIPOTRANS,KOPAE,KOCIE,KOCME,FECHAE,HORAE,KOPAD,KOCID,KOCMD,FECHAD,HORAD,OBDOEXPO,MOTIVO," & vbCrLf & _
                           "TEXTO16,TEXTO17,TEXTO18,TEXTO19,TEXTO20,TEXTO21,TEXTO22,TEXTO23,TEXTO24,TEXTO25,TEXTO26,TEXTO27," & vbCrLf & _
                           "TEXTO28,TEXTO29,TEXTO30,TEXTO31,TEXTO32,TEXTO33,TEXTO34,TEXTO35,PLACAPAT) VALUES " & vbCrLf & _
                           "(" & _Idmaeedo & ",'" & _Obdo & "','" & _Cpdo & "','" & _Ocdo & "','','" & Obs(0) & "','" & Obs(1) & _
                           "','" & Obs(2) & "','" & Obs(3) & "','" & Obs(4) & "','" & Obs(5) & "','" & Obs(6) & "','" & Obs(7) & _
                           "','" & Obs(8) & "','" & Obs(9) & "','" & Obs(10) & "','" & Obs(11) & "','" & Obs(12) & "','" & Obs(13) & _
                           "','" & Obs(14) & "','','','','','','','','','',GETDATE(),'','','','',GETDATE(),'','','','" & Obs(15) & _
                           "','" & Obs(16) & "','" & Obs(17) & "','" & Obs(18) & "','" & Obs(19) & _
                           "','" & Obs(20) & "','" & Obs(21) & "','" & Obs(22) & "','" & Obs(23) & "','" & Obs(24) & _
                           "','" & Obs(25) & "','" & Obs(26) & "','" & Obs(27) & "','" & Obs(28) & "','" & Obs(29) & _
                           "','" & Obs(30) & "','" & Obs(31) & "','" & Obs(32) & "','" & Obs(33) & "','" & Obs(34) & _
                           "','')"

            Consulta_sql = "Insert Into " & _Global_BaseBk & "Zw_Casi_DocObs (Id_DocEnc,Observaciones,Forma_pago,Orden_compra,Obs1," & _
                           "Obs2,Obs3,Obs4,Obs5,Obs6,Obs7,Obs8,Obs9,Obs10," & _
                           "Obs11,Obs12,Obs13,Obs14,Obs15,Obs16,Obs17,Obs18,Obs19,Obs20,Obs21,Obs22,Obs23,Obs24,Obs25,Obs26," & _
                           "Obs27,Obs28,Obs29,Obs30,Obs31,Obs32,Obs33,Obs34,Obs35) Values " & vbCrLf & _
                           "(" & _Id_DocEnc & ",'" & _Obdo & "','" & _Cpdo & "','" & _Ocdo & _
                           "','" & Obs(0) & "','" & Obs(1) & "','" & Obs(2) & "','" & Obs(3) & "','" & Obs(4) & "','" & Obs(5) & _
                           "','" & Obs(6) & "','" & Obs(7) & "','" & Obs(8) & "','" & Obs(9) & "','" & Obs(10) & _
                           "','" & Obs(11) & "','" & Obs(12) & "','" & Obs(13) & "','" & Obs(14) & "','" & Obs(15) & _
                           "','" & Obs(16) & "','" & Obs(17) & "','" & Obs(18) & "','" & Obs(19) & "','" & Obs(20) & _
                           "','" & Obs(21) & "','" & Obs(22) & "','" & Obs(23) & "','" & Obs(24) & "','" & Obs(25) & _
                           "','" & Obs(26) & "','" & Obs(27) & "','" & Obs(28) & "','" & Obs(29) & "','" & Obs(30) & _
                           "','" & Obs(31) & "','" & Obs(32) & "','" & Obs(33) & "','" & Obs(34) & "')"


            Comando = New SqlClient.SqlCommand(Consulta_sql, cn2)
            Comando.Transaction = myTrans
            Comando.ExecuteNonQuery()

            ' ====================================================================================================================================

            myTrans.Commit()


            Return _Id_DocEnc

        Catch ex As Exception

            'Dim _Erro_VB As String = ex.Message & vbCrLf & ex.StackTrace & vbCrLf & _
            '                         "Código: " & _Koprct
            'Clipboard.SetText(_Erro_VB)

            'My.Computer.FileSystem.WriteAllText("ArchivoSalida", ex.Message & vbCrLf & ex.StackTrace, False)
            'MessageBoxEx.Show(ex.Message, "Error", _
            '                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
            myTrans.Rollback()

            'MessageBoxEx.Show("Transaccion desecha", "Problema", _
            '                 Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Stop)
            'SQL_ServerClass.CerrarConexion(cn2)
            Return 0
        Finally
            _Sql.Sb_Cerrar_Conexion(cn2)
        End Try

    End Function

#End Region

#Region "EDITAR DOCUMENTO"

    Function Fx_Editar_Documento(ByVal _Idmaeedo As Integer, _
                                 ByVal _Cod_Func_Eliminado As String, _
                                 ByVal Bd_Documento As DataSet) As Integer

        ' Obtengo el tipo y numero de documento que hay que modificar
        Dim _Tipo_de_documento As String = _Sql.Fx_Trae_Dato("TIDO", "MAEEDO", "IDMAEEDO = " & _Idmaeedo)
        Dim _Numero_de_documento As String = _Sql.Fx_Trae_Dato("NUDO", "MAEEDO", "IDMAEEDO = " & _Idmaeedo)

        ' Obtengo la fecha del servidor para poner la fecha de eliminación del documento
        Dim _FechaEliminacion = FechaDelServidor()

        Dim _New_Idmaeedo As Integer = Fx_Crear_Documento(_Tipo_de_documento, _
                                                          _Numero_de_documento, _
                                                          False, _
                                                          False, _
                                                          Bd_Documento, _
                                                          False, _
                                                          False)

        If CBool(_New_Idmaeedo) Then

            Dim _Class_E As New Clase_EliminarAnular_Documento

            Dim _Eliminado As Boolean = _Class_E.Fx_EliminarAnular_Doc(_Idmaeedo, _
                                                                       _Cod_Func_Eliminado, _
                                                                       Clase_EliminarAnular_Documento._Accion_EA.Modificar, _
                                                                       False)

            If _Eliminado Then

                Consulta_sql = "Update MAEEDO Set NUDO = '" & _Numero_de_documento & "' Where IDMAEEDO = " & _New_Idmaeedo & vbCrLf & _
                               "Update MAEDDO Set NUDO = '" & _Numero_de_documento & "' Where IDMAEEDO = " & _New_Idmaeedo
                If _Sql.Fx_Ej_consulta_IDU(Consulta_sql) Then 'Ej_consulta_IDU(Consulta_sql, cn1) Then
                    Return _New_Idmaeedo
                End If
            Else
                _Class_E.Fx_EliminarAnular_Doc(_New_Idmaeedo, _
                                               _Cod_Func_Eliminado, _
                                               Clase_EliminarAnular_Documento._Accion_EA.Modificar, _
                                               False)
            End If

        Else
            Return False
        End If

    End Function

#End Region


End Class
