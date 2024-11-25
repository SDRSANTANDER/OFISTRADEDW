Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOIdioma
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getIdiomas() As List(Of EntIdioma)

        Dim retVal As New List(Of EntIdioma)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaIdiomas()

            For Each row As DataRow In DT.Rows

                Dim oIdioma As EntIdioma = DataRowToEntidadIdioma(row)
                retVal.Add(oIdioma)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaIdiomas() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_IDIOMAS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadIdioma(DR As DataRow) As EntIdioma

        Dim oIdioma As New EntIdioma

        With oIdioma

            .ID = CInt(DR.Item("ID").ToString)
            .NombreCorto = DR.Item("NOMBRECORTO").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oIdioma

    End Function

End Class
