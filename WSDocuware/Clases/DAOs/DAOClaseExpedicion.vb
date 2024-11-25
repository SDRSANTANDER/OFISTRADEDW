Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOClaseExpedicion
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getClasesExpedicion() As List(Of EntClaseExpedicion)

        Dim retVal As New List(Of EntClaseExpedicion)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaClasesExpedicion()

            For Each row As DataRow In DT.Rows

                Dim oClaseExpedicion As EntClaseExpedicion = DataRowToEntidadClaseExpedicion(row)
                retVal.Add(oClaseExpedicion)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaClasesExpedicion() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_CLASES_EXPEDICION") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadClaseExpedicion(DR As DataRow) As EntClaseExpedicion

        Dim oClaseExpedicion As New EntClaseExpedicion

        With oClaseExpedicion

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oClaseExpedicion

    End Function

End Class
