Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOActivoFijoGrupoAmortizacion
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getActivosFijosGruposAmortizacion() As List(Of EntActivoFijoGrupoAmortizacion)

        Dim retVal As New List(Of EntActivoFijoGrupoAmortizacion)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActivosFijosGruposAmortizacion()

            For Each row As DataRow In DT.Rows

                Dim oActivoFijoGrupoAmortizacion As EntActivoFijoGrupoAmortizacion = DataRowToEntidadActivoFijoGrupoAmortizacion(row)
                retVal.Add(oActivoFijoGrupoAmortizacion)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActivosFijosGruposAmortizacion() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVOS_FIJOS_GRUPOS_AMORTIZACION") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActivoFijoGrupoAmortizacion(DR As DataRow) As EntActivoFijoGrupoAmortizacion

        Dim oActivoFijoGrupoAmortizacion As New EntActivoFijoGrupoAmortizacion

        With oActivoFijoGrupoAmortizacion

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Grupo = DR.Item("GRUPO").ToString

        End With

        Return oActivoFijoGrupoAmortizacion

    End Function

End Class
