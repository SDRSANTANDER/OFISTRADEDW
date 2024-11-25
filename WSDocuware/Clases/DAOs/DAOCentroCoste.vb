Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOCentroCoste
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getCentrosCoste() As List(Of EntCentroCoste)

        Dim retVal As New List(Of EntCentroCoste)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaCentrosCoste()

            For Each row As DataRow In DT.Rows

                Dim oCentroCoste As EntCentroCoste = DataRowToEntidadCentroCoste(row)
                retVal.Add(oCentroCoste)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaCentrosCoste() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_CENTROS_COSTE") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadCentroCoste(DR As DataRow) As EntCentroCoste

        Dim oCentroCoste As New EntCentroCoste

        With oCentroCoste

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Responsable = CInt(DR.Item("RESPONSABLE").ToString)
            .Dimension = CInt(DR.Item("DIMENSION").ToString)
            .Importe = CDbl(DR.Item("IMPORTE").ToString)

        End With

        Return oCentroCoste

    End Function

End Class
