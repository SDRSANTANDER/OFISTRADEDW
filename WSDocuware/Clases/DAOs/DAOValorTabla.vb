Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOValorTabla
    Inherits clsConexion

    Sub New(ByVal Sociedad As eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getValoresTablas(ByVal Tabla As String) As List(Of EntValorTabla)

        Dim retVal As New List(Of EntValorTabla)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaValoresTablas(Tabla)

            For Each row As DataRow In DT.Rows

                Dim oValorTabla As EntValorTabla = DataRowToEntidadValorTabla(row)
                retVal.Add(oValorTabla)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaValoresTablas(ByVal Tabla As String) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes(Tabla) & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadValorTabla(DR As DataRow) As EntValorTabla

        Dim oValorTabla As New EntValorTabla

        With oValorTabla

            .Codigo = DR.Item("Code").ToString
            .Descripcion = DR.Item("Name").ToString

        End With

        Return oValorTabla

    End Function

End Class
