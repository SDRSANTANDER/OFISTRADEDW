Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOMoneda
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getMonedas() As List(Of EntMoneda)

        Dim retVal As New List(Of EntMoneda)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaMonedas()

            For Each row As DataRow In DT.Rows

                Dim oMoneda As EntMoneda = DataRowToEntidadMoneda(row)
                retVal.Add(oMoneda)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaMonedas() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_MONEDAS") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadMoneda(DR As DataRow) As EntMoneda

        Dim oMoneda As New EntMoneda

        With oMoneda

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Internacional = DR.Item("INTERNACIONAL").ToString

        End With

        Return oMoneda

    End Function

End Class
