Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOActivoFijoGrupo
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getActivosFijosGrupos() As List(Of EntActivoFijoGrupo)

        Dim retVal As New List(Of EntActivoFijoGrupo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActivosFijosGrupos()

            For Each row As DataRow In DT.Rows

                Dim oActivoFijoGrupo As EntActivoFijoGrupo = DataRowToEntidadActivoFijoGrupo(row)
                retVal.Add(oActivoFijoGrupo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActivosFijosGrupos() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVOS_FIJOS_GRUPOS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActivoFijoGrupo(DR As DataRow) As EntActivoFijoGrupo

        Dim oActivoFijoGrupo As New EntActivoFijoGrupo

        With oActivoFijoGrupo

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oActivoFijoGrupo

    End Function

End Class
