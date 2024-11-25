Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAONormaReparto
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getNormasReparto() As List(Of EntNormaReparto)

        Dim retVal As New List(Of EntNormaReparto)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaNormasReparto()

            For Each row As DataRow In DT.Rows

                Dim oNormaReparto As EntNormaReparto = DataRowToEntidadNormaReparto(row)
                retVal.Add(oNormaReparto)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaNormasReparto() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_NORMAS_REPARTO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadNormaReparto(DR As DataRow) As EntNormaReparto

        Dim oNormaReparto As New EntNormaReparto

        With oNormaReparto

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Dimension = CInt(DR.Item("DIMENSION").ToString)
            .Total = CDbl(DR.Item("TOTAL").ToString)

        End With

        Return oNormaReparto

    End Function

End Class
