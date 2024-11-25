Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOProyecto
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getProyectos() As List(Of EntProyecto)

        Dim retVal As New List(Of EntProyecto)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaProyectos()

            For Each row As DataRow In DT.Rows

                Dim oProyecto As EntProyecto = DataRowToEntidadProyecto(row)
                retVal.Add(oProyecto)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaProyectos() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_PROYECTOS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadProyecto(DR As DataRow) As EntProyecto

        Dim oProyecto As New EntProyecto

        With oProyecto

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .CentrosCoste = DR.Item("CENTROCOSTE").ToString
            .ResponsableID = DR.Item("RESPONSABLEID").ToString
            .ResponsableNombre = DR.Item("RESPONSABLENOMBRE").ToString
            .ResponsableMail = DR.Item("RESPONSABLEMAIL").ToString
            .Generico1 = DR.Item("GENERICO1").ToString
            .Generico2 = DR.Item("GENERICO2").ToString
            .Generico3 = DR.Item("GENERICO3").ToString
            .Generico4 = DR.Item("GENERICO4").ToString
            .Generico5 = DR.Item("GENERICO5").ToString

        End With

        Return oProyecto

    End Function

End Class
