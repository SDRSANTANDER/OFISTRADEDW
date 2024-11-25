Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOTasaCambio
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getTasasCambio(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntTasaCambio)

        Dim retVal As New List(Of EntTasaCambio)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaTasasCambio(FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oTasaCambio As EntTasaCambio = DataRowToEntidadTasaCambio(row)
                retVal.Add(oTasaCambio)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaTasasCambio(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_TASAS_CAMBIO") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf
            SQL &= " And " & putQuotes("FECHA") & ">=N'" & FechaInicio & "'" & vbCrLf
            SQL &= " And " & putQuotes("FECHA") & "<=N'" & FechaFin & "'" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadTasaCambio(DR As DataRow) As EntTasaCambio

        Dim oTasaCambio As New EntTasaCambio

        With oTasaCambio

            .Fecha = CDate(DR.Item("FECHA").ToString).ToString("yyyyMMdd")
            .Moneda = DR.Item("MONEDA").ToString
            .Tasa = CDbl(DR.Item("TASA").ToString)

        End With

        Return oTasaCambio

    End Function

End Class
