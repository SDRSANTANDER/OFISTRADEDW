Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOPrecioEntrega
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

#Region "Abiertos"

    Public Function getPreciosEntregaAbiertos(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntPrecioEntrega)

        Dim retVal As New List(Of EntPrecioEntrega)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaPreciosEntrega(FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oPrecioEntrega As EntPrecioEntrega = DataRowToEntidadPrecioEntrega(row)
                retVal.Add(oPrecioEntrega)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaPreciosEntrega(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""

            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_PRECIOS_ENTREGA") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf
            SQL &= " AND " & putQuotes("DOCSTATUS") & " <> '" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " AND " & putQuotes("CANCELED") & " = '" & SN.No & "'" & vbCrLf
            SQL &= " AND " & putQuotes("DOCDATE") & " >= '" & FechaInicio & "'" & vbCrLf
            SQL &= " AND " & putQuotes("DOCDATE") & " <= '" & FechaFin & "'" & vbCrLf
            SQL &= " ORDER BY " & putQuotes("DOCNUM") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadPrecioEntrega(DR As DataRow) As EntPrecioEntrega

        Dim oPrecioEntrega As New EntPrecioEntrega

        With oPrecioEntrega

            .Interlocutor = DR.Item("CARDCODE").ToString
            .RazonSocial = DR.Item("CARDNAME").ToString
            .NIFTercero = DR.Item("LICTRADNUM").ToString

            .AduanaInterlocutor = DR.Item("AGENTCARDCODE").ToString
            .AduanaRazonSocial = DR.Item("AGENTCARDNAME").ToString
            .AduanaNIFTercero = DR.Item("AGENTLICTRADNUM").ToString

            .DocEntry = CInt(DR.Item("DOCENTRY"))
            .DocNum = CInt(DR.Item("DOCNUM"))
            .DocDate = CInt(DR.Item("DOCDATE"))
            .DocDueDate = CInt(DR.Item("DOCDUEDATE"))
            .NumAtCard = DR.Item("REF1").ToString
            .DocTotal = CDbl(DR.Item("DOCTOTAL"))
            .DocMoneda = DR.Item("DOCCUR").ToString

            .AduanaPrevista = CDbl(DR.Item("EXPCUSTOM"))
            .AduanaReal = CDbl(DR.Item("ACTCUSTOM"))
            .AduanaFecha = CInt(DR.Item("CUSTDATE"))
            .GastosTotal = CDbl(DR.Item("COSTSUM"))

        End With

        Return oPrecioEntrega

    End Function

#End Region

End Class
