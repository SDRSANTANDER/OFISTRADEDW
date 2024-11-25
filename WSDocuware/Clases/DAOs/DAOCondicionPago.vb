Imports System.Reflection

Public Class DAOCondicionPago
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getCondicionesPago() As List(Of EntCondicionPago)

        Dim retVal As New List(Of EntCondicionPago)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaCondicionesPago()

            For Each row As DataRow In DT.Rows

                Dim oCondicionPago As EntCondicionPago = DataRowToEntidadCondicionPago(row)
                retVal.Add(oCondicionPago)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaCondicionesPago() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & Utilidades.putQuotes("SEI_VIEW_DW_CONDICIONES_PAGO") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadCondicionPago(DR As DataRow) As EntCondicionPago

        Dim oCondicionPago As New EntCondicionPago

        With oCondicionPago

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oCondicionPago

    End Function

End Class
