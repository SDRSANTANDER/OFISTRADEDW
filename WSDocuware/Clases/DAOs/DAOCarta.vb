Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOCarta
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getSolicitudesPedidosCompraIP(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntCarta)

        Dim retVal As New List(Of EntCarta)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaSolicitudesPedidosCompraIP(FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oCarta As EntCarta = DataRowToEntidadSolicitudesPedidosCompraIP(row)
                retVal.Add(oCarta)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaSolicitudesPedidosCompraIP(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = "  SELECT * " & vbCrLf
            SQL &= " FROM " & putQuotes("SEI_VIEW_DW_SOLICITUD_PEDIDOS_COMPRA_IP") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And " & putQuotes("FECHACONTABLE") & ">=N'" & FechaInicio & "'" & vbCrLf
            SQL &= " And " & putQuotes("FECHACONTABLE") & "<=N'" & FechaFin & "'" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadSolicitudesPedidosCompraIP(DR As DataRow) As EntCarta

        Dim oCarta As New EntCarta

        With oCarta

            .InvestigadorPrincipalId = CInt(DR.Item("INVESTIGADORPRINCIPALID").ToString)
            .InvestigadorPrincipalNombre = DR.Item("INVESTIGADORPRINCIPALNOMBRE").ToString
            .InvestigadorPrincipalApellido = DR.Item("INVESTIGADORPRINCIPALAPELLIDO").ToString
            .InvestigadorPrincipalMail = DR.Item("INVESTIGADORPRINCIPALMAIL").ToString

            .InvestigadorId = CInt(DR.Item("INVESTIGADORID").ToString)
            .InvestigadorNombre = DR.Item("INVESTIGADORNOMBRE").ToString
            .InvestigadorApellido = DR.Item("INVESTIGADORAPELLIDO").ToString
            .InvestigadorMail = DR.Item("INVESTIGADORMAIL").ToString

            .Numerador = DR.Item("NUMERADOR").ToString
            .DocumentoNumero = CInt(DR.Item("DOCUMENTONUMERO").ToString)

            .InterlocutorId = DR.Item("INTERLOCUTORID").ToString
            .InterLocutorRazonSocial = DR.Item("INTERLOCUTORRAZONSOCIAL").ToString
            .InterlocutorNumeroReferencia = DR.Item("INTERLOCUTORNUMEROREFERENCIA").ToString

            .FechaContable = CDate(DR.Item("FECHACONTABLE")).ToString("yyyyMMdd")
            .FechaEntrega = CDate(DR.Item("FECHAENTREGA")).ToString("yyyyMMdd")

            .LineaNumero = CInt(DR.Item("LINEANUMERO").ToString)
            .LineaEstado = DR.Item("LINEAESTADO").ToString

            .ArticuloId = DR.Item("ARTICULOID").ToString
            .ArticuloDescripcion = DR.Item("ARTICULODESCRIPCION").ToString
            .ArticuloReferencia = DR.Item("ARTICULOREFERENCIA").ToString
            .ArticuloGenericoLaboratorio = DR.Item("ARTICULOGENERICOLABORATORIO").ToString
            .ArticuloTipoContrato = DR.Item("ARTICULOTIPOCONTRATO").ToString

            .Cantidad = CDbl(DR.Item("CANTIDAD"))
            .Precio = CDbl(DR.Item("PRECIO"))
            .LineaTotal = CDbl(DR.Item("LINEATOTAL"))

            .ProyectoId = DR.Item("PROYECTOID").ToString
            .ProyectoDescripcion = DR.Item("PROYECTODESCRIPCION").ToString
            .Hito = DR.Item("HITO").ToString
            .NormaReparto = DR.Item("NORMAREPARTO").ToString

            .TextoLibre = DR.Item("TEXTOLIBRE").ToString
            .CASS = DR.Item("CASS").ToString

        End With

        Return oCarta

    End Function

    Public Function getRelacionesDocumentosCompra(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntCarta)

        Dim retVal As New List(Of EntCarta)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaRelacionesDocumentosCompra(FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oCarta As EntCarta = DataRowToEntidadRelacionesDocumentosCompra(row)
                retVal.Add(oCarta)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaRelacionesDocumentosCompra(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = "  SELECT * " & vbCrLf
            SQL &= " FROM " & putQuotes("SEI_VIEW_DW_RELACIONES_DOCUMENTOS_COMPRA") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And " & putQuotes("FECHACONTABLE") & ">=N'" & FechaInicio & "'" & vbCrLf
            SQL &= " And " & putQuotes("FECHACONTABLE") & "<=N'" & FechaFin & "'" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadRelacionesDocumentosCompra(DR As DataRow) As EntCarta

        Dim oCarta As New EntCarta

        With oCarta

            .SolicitudCompra = CInt(DR.Item("SOLICITUDCOMPRA").ToString)
            .PedidoVenta = DR.Item("PEDIDOVENTA").ToString
            .PedidoCompra = DR.Item("PEDIDOCOMPRA").ToString

            .InterlocutorId = DR.Item("INTERLOCUTORID").ToString
            .InterLocutorRazonSocial = DR.Item("INTERLOCUTORRAZONSOCIAL").ToString

            .FechaContable = CDate(DR.Item("FECHACONTABLE")).ToString("yyyyMMdd")

            .LineaNumero = CInt(DR.Item("LINEANUMERO").ToString)

            .ArticuloId = DR.Item("ARTICULOID").ToString
            .ArticuloDescripcion = DR.Item("ARTICULODESCRIPCION").ToString
            .ArticuloTipoContrato = DR.Item("ARTICULOTIPOCONTRATO").ToString

            .Cantidad = CDbl(DR.Item("CANTIDAD"))
            .Precio = CDbl(DR.Item("PRECIO"))

        End With

        Return oCarta

    End Function


End Class
