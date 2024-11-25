Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOResponsable
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getResponsables() As List(Of EntResponsable)

        Dim retVal As New List(Of EntResponsable)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaResponsables()

            For Each row As DataRow In DT.Rows

                Dim oResponsable As EntResponsable = DataRowToEntidadResponsable(row)
                retVal.Add(oResponsable)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaResponsables() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_RESPONSABLES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadResponsable(DR As DataRow) As EntResponsable

        Dim oResponsable As New EntResponsable

        With oResponsable

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oResponsable

    End Function

End Class
