Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOGrupoImpositivo
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getGruposImpositivo() As List(Of EntGrupoImpositivo)

        Dim retVal As New List(Of EntGrupoImpositivo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaGruposImpositivo()

            For Each row As DataRow In DT.Rows

                Dim oGrupoImpositivo As EntGrupoImpositivo = DataRowToEntidadGrupoImpositivo(row)
                retVal.Add(oGrupoImpositivo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaGruposImpositivo() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_GRUPOS_IMPOSITIVOS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadGrupoImpositivo(DR As DataRow) As EntGrupoImpositivo

        Dim oGrupoImpositivo As New EntGrupoImpositivo

        With oGrupoImpositivo

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Porcentaje = CDbl(DR.Item("PORCENTAJE").ToString)
            .Categoria = DR.Item("CATEGORIA").ToString
            .Intracomunitario = DR.Item("INTRACOMUNITARIO").ToString

        End With

        Return oGrupoImpositivo

    End Function

End Class
