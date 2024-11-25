Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOActivoFijoClase
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getActivosFijosClases() As List(Of EntActivoFijoClase)

        Dim retVal As New List(Of EntActivoFijoClase)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActivosFijosClases()

            For Each row As DataRow In DT.Rows

                Dim oActivoFijoClase As EntActivoFijoClase = DataRowToEntidadActivoFijoClase(row)
                retVal.Add(oActivoFijoClase)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActivosFijosClases() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVOS_FIJOS_CLASES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActivoFijoClase(DR As DataRow) As EntActivoFijoClase

        Dim oActivoFijoClase As New EntActivoFijoClase

        With oActivoFijoClase

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Tipo = DR.Item("TIPO").ToString

        End With

        Return oActivoFijoClase

    End Function

End Class
