Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOGrupoUnidadMedida
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getGruposUnidadMedida() As List(Of EntGrupoUnidadMedida)

        Dim retVal As New List(Of EntGrupoUnidadMedida)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaGruposUnidadMedida()

            For Each row As DataRow In DT.Rows

                Dim oGrupoUnidadMedida As EntGrupoUnidadMedida = DataRowToEntidadGrupoUnidadMedida(row)
                retVal.Add(oGrupoUnidadMedida)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaGruposUnidadMedida() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_GRUPOS_UNIDAD_MEDIDA") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadGrupoUnidadMedida(DR As DataRow) As EntGrupoUnidadMedida

        Dim oGrupoUnidadMedida As New EntGrupoUnidadMedida

        With oGrupoUnidadMedida

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oGrupoUnidadMedida

    End Function

End Class
