Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOViaPago
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getViasPago() As List(Of EntViaPago)

        Dim retVal As New List(Of EntViaPago)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaViasPago()

            For Each row As DataRow In DT.Rows

                Dim oViaPago As EntViaPago = DataRowToEntidadViaPago(row)
                retVal.Add(oViaPago)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaViasPago() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_VIAS_PAGO") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadViaPago(DR As DataRow) As EntViaPago

        Dim oViaPago As New EntViaPago

        With oViaPago

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Tipo = DR.Item("TIPO").ToString
            .Medio = DR.Item("MEDIO").ToString
            .CuentaBanco = DR.Item("CUENTABANCO").ToString
            .CuentaContable = DR.Item("CUENTACONTABLE").ToString

        End With

        Return oViaPago

    End Function

End Class
