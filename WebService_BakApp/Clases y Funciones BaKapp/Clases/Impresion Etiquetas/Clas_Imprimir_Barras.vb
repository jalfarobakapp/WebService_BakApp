Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Public Class Clas_Imprimir_Barras

    Dim _Sql As New Class_SQL '(Cadena_ConexionSQL_Server)
    Dim Consulta_sql As String

    Public _Empresa, _Sucursal, _Bodega As String
    Public _TblProductos As DataTable
    Private prtSettings As PrinterSettings
    Public _Filas_X_Documento As Integer

    Dim _Item = 1


    Private Sub Sb_Print_PrintPage_Codigos_Barra(sender As Object,
                                                  e As PrintPageEventArgs)
        ' Este evento se producirá cada vez que se imprima una nueva página
        ' imprimir HOLA MUNDO en Arial tamaño 24 y negrita

        Try

            ' imprimimos la cadena en el margen izquierdo
            Dim xPos As Single = 3 'e.MarginBounds.Left
            ' La fuente a usar


            Dim DtFont As New Font("Arial", 9, FontStyle.Regular) ' Fuente del detalle
            Dim prFont As New Font("Arial", 9, FontStyle.Bold)
            Dim FontNro As New Font("Times New Roman", 14, FontStyle.Bold)
            Dim FontCon As New Font("Times New Roman", 11, FontStyle.Bold)

            Dim FteCourier_New_C_4 As New Font("Courier New", 4, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_6 As New Font("Courier New", 6, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_7 As New Font("Courier New", 7, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_8 As New Font("Courier New", 8, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_9 As New Font("Courier New", 9, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_10 As New Font("Courier New", 10, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_11 As New Font("Courier New", 11, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_12 As New Font("Courier New", 12, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_13 As New Font("Courier New", 13, FontStyle.Bold) ' Crea la fuente
            Dim FteCourier_New_C_14 As New Font("Courier New", 13, FontStyle.Bold) ' Crea la fuente


            ' la posición superior
            Dim yPos As Single = prFont.GetHeight(e.Graphics) - 10


            e.Graphics.DrawString("PRODUCTOS", FontNro, Brushes.Black, xPos, yPos)
            yPos = yPos + 25
            e.Graphics.DrawString("_____________________________________________", DtFont, Brushes.Black, xPos, yPos)
            yPos = yPos + 30

            xPos = 40

            Dim _Contador = 0

            ' imprimimos la cadena
            For Each _Fila As DataRow In _TblProductos.Rows


                If Not _Fila.Item("Impreso") Then

                    Dim _Codigo As String = _Fila.Item("KOPR")
                    Dim _Codigo_Tecnico As String = _Fila.Item("KOPRTE")
                    Dim _Descripcion As String = _Fila.Item("NOKOPR")


                    Consulta_sql = "Select * From TABBOPR" & vbCrLf &
                                   "Where EMPRESA = '" & _Empresa &
                                   "' And KOSU = '" & _Sucursal & "' And KOBO = '" & _Bodega & "' And KOPR = '" & _Codigo & "'"

                    Dim _TblBodega_producto As DataTable = _Sql.Fx_Get_DataTable(Consulta_sql)

                    Dim _Ubicacion = String.Empty

                    If CBool(_TblBodega_producto.Rows.Count) Then
                        _Ubicacion = Trim(_TblBodega_producto.Rows(0).Item("DATOSUBIC"))
                    End If


                    'Vale-BkPost
                    Dim bm As Bitmap = Nothing
                    Dim CodBarras As New PictureBox
                    Dim Imagen As New PictureBox

                    Dim iType As BarCode.Code128SubTypes =
                    DirectCast([Enum].Parse(GetType(BarCode.Code128SubTypes), "CODE128"), BarCode.Code128SubTypes)

                    'Dim iType As BarCode.Code128SubTypes = _
                    ' DirectCast([Enum].Parse(GetType(BarCode.Code128SubTypes), "CODE128"), BarCode.Code128SubTypes)


                    bm = BarCode.Code128(_Codigo, iType, False)

                    'bm = BarCode.Code128(_Codigo, iType, False) ' Imprime solo el código
                    'bm = BarCode.Code128(_Codigo, iType, True) 'Incluye el TEXTO código en la barra
                    If Not IsNothing(bm) Then
                        CodBarras.Image = bm
                    End If

                    e.Graphics.DrawString("Item: " & _Item, FteCourier_New_C_8, Brushes.Black, xPos + 300, yPos)
                    e.Graphics.DrawString(_Descripcion, FteCourier_New_C_6, Brushes.Black, xPos, yPos)
                    yPos += 12
                    e.Graphics.DrawImage(CodBarras.Image, xPos + 10, yPos, 200, 30)
                    yPos += 35
                    e.Graphics.DrawString(_Codigo_Tecnico, FteCourier_New_C_11, Brushes.Black, xPos + 10, yPos)
                    yPos += 15
                    e.Graphics.DrawString(_Ubicacion, FteCourier_New_C_8, Brushes.Black, xPos + 150, yPos)
                    yPos += 15
                    e.Graphics.DrawString("_____________________________________________", DtFont, Brushes.Black, xPos - 40, yPos)
                    yPos += 20

                    _Fila.Item("Impreso") = True
                    _Contador += 1
                    _Item += 1

                    If _Contador = _Filas_X_Documento Then
                        Exit For
                    End If
                End If

            Next


            ' indicamos que ya no hay nada más que imprimir
            ' (el valor predeterminado de esta propiedad es False)

            Dim _Saldo_Registros As Integer = NuloPorNro(_TblProductos.Compute("Sum(Contador)", "Impreso = 0"), 0)

            If CBool(_Saldo_Registros) Then
                e.HasMorePages = True
            Else
                e.Graphics.DrawString("FIN IMPRESION", FontNro, Brushes.Black, xPos, yPos)

                e.HasMorePages = False
            End If


        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText("Log_Errores.log", ex.Message & vbCrLf & ex.StackTrace, False)
            MsgBox(ex.Message)
            MsgBox("Error lo puesde ver en archivo Log de errores")
        End Try

    End Sub

End Class
