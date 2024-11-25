Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOBancoPropio
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getBancosPropios() As List(Of EntBancoPropio)

        Dim retVal As New List(Of EntBancoPropio)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaBancosPropios()

            For Each row As DataRow In DT.Rows

                Dim oBancoPropio As EntBancoPropio = DataRowToEntidadBancoPropio(row)
                retVal.Add(oBancoPropio)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaBancosPropios() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_BANCOS_PROPIOS") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadBancoPropio(DR As DataRow) As EntBancoPropio

        Dim oBancoPropio As New EntBancoPropio

        With oBancoPropio

            .IBAN = DR.Item("IBAN").ToString
            .Banco = DR.Item("BANCO").ToString
            .Sucursal = DR.Item("SUCURSAL").ToString
            .DigitosControl = DR.Item("DIGITOSCONTROL").ToString
            .CuentaNumero = DR.Item("CUENTANUMERO").ToString
            .CuentaNombre = DR.Item("CUENTANOMBRE").ToString
            .Pais = DR.Item("PAIS").ToString
            .CuentaContable = DR.Item("CUENTACONTABLE").ToString

        End With

        Return oBancoPropio

    End Function

End Class

