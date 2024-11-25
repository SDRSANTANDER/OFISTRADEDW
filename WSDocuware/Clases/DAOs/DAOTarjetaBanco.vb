Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOTarjetaBanco
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getTarjetasBanco() As List(Of EntTarjetaBanco)

        Dim retVal As New List(Of EntTarjetaBanco)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaTarjetasBanco()

            For Each row As DataRow In DT.Rows

                Dim oTarjetaBanco As EntTarjetaBanco = DataRowToEntidadTarjetaBanco(row)
                retVal.Add(oTarjetaBanco)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaTarjetasBanco() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_TARJETAS_BANCO") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadTarjetaBanco(DR As DataRow) As EntTarjetaBanco

        Dim oTarjetaBanco As New EntTarjetaBanco

        With oTarjetaBanco

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString
            .Cuenta = DR.Item("CUENTA").ToString

        End With

        Return oTarjetaBanco

    End Function

End Class
