Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOActivoFijoAreaValoracion
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getActivosFijosAreasValoracion() As List(Of EntActivoFijoAreaValoracion)

        Dim retVal As New List(Of EntActivoFijoAreaValoracion)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActivosFijosAreasValoracion()

            For Each row As DataRow In DT.Rows

                Dim oActivoFijoAreaValoracion As EntActivoFijoAreaValoracion = DataRowToEntidadActivoFijoAreaValoracion(row)
                retVal.Add(oActivoFijoAreaValoracion)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActivosFijosAreasValoracion() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVOS_FIJOS_AREAS_VALORACION") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActivoFijoAreaValoracion(DR As DataRow) As EntActivoFijoAreaValoracion

        Dim oActivoFijoAreaValoracion As New EntActivoFijoAreaValoracion

        With oActivoFijoAreaValoracion

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Tipo = DR.Item("TIPO").ToString

        End With

        Return oActivoFijoAreaValoracion

    End Function

End Class
