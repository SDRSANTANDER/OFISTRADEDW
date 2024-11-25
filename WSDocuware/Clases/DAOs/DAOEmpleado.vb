Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOEmpleado
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getEmpleados() As List(Of EntEmpleado)

        Dim retVal As New List(Of EntEmpleado)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaEmpleados()

            For Each row As DataRow In DT.Rows

                Dim oEmpleado As EntEmpleado = DataRowToEntidadEmpleado(row)
                retVal.Add(oEmpleado)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaEmpleados() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_EMPLEADOS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadEmpleado(DR As DataRow) As EntEmpleado

        Dim oEmpleado As New EntEmpleado

        With oEmpleado

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .CorreoE = DR.Item("CORREOE").ToString
            .CentroCoste = DR.Item("CENTROCOSTE").ToString

        End With

        Return oEmpleado

    End Function

End Class
