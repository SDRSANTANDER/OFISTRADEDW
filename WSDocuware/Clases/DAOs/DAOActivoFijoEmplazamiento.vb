Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOActivoFijoEmplazamiento
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getActivosFijosEmplazamientos() As List(Of EntActivoFijoEmplazamiento)

        Dim retVal As New List(Of EntActivoFijoEmplazamiento)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActivosFijosEmplazamientos()

            For Each row As DataRow In DT.Rows

                Dim oActivoFijoEmplazamiento As EntActivoFijoEmplazamiento = DataRowToEntidadActivoFijoEmplazamiento(row)
                retVal.Add(oActivoFijoEmplazamiento)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActivosFijosEmplazamientos() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVOS_FIJOS_EMPLAZAMIENTOS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActivoFijoEmplazamiento(DR As DataRow) As EntActivoFijoEmplazamiento

        Dim oActivoFijoEmplazamiento As New EntActivoFijoEmplazamiento

        With oActivoFijoEmplazamiento

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oActivoFijoEmplazamiento

    End Function

End Class
