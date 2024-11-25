Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOFacturaAnticipo
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getFacturasAnticipo(ByVal Tipo As String) As List(Of EntFacturaAnticipo)

        Dim retVal As New List(Of EntFacturaAnticipo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaFacturasAnticipo(Tipo)

            For Each row As DataRow In DT.Rows

                Dim oFacturaAnticipo As EntFacturaAnticipo = DataRowToEntidadFacturaAnticipo(row)
                retVal.Add(oFacturaAnticipo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaFacturasAnticipo(ByVal Tipo As String) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_FACTURAS_ANTICIPO") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            If Not String.IsNullOrEmpty(Tipo) Then
                SQL &= " And " & Utilidades.putQuotes("CARDTYPE") & " = N'" & Tipo & "'" & vbCrLf
            End If

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadFacturaAnticipo(DR As DataRow) As EntFacturaAnticipo

        Dim oFacturaAnticipo As New EntFacturaAnticipo

        With oFacturaAnticipo

            .Interlocutor = DR.Item("CARDCODE").ToString
            .Tipo = DR.Item("CARDTYPE").ToString
            .RazonSocial = DR.Item("CARDNAME").ToString
            .NIFTercero = DR.Item("LICTRADNUM").ToString

            .DocEntry = CInt(DR.Item("DOCENTRY"))
            .DocNum = CInt(DR.Item("DOCNUM"))
            .DocDate = CInt(DR.Item("DOCDATE"))
            .NumAtCard = DR.Item("NUMATCARD").ToString
            .DocTotal = CDbl(DR.Item("DOCTOTAL"))
            .DocMoneda = DR.Item("DOCCUR").ToString

        End With

        Return oFacturaAnticipo

    End Function

End Class

