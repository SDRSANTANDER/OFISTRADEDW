Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOAlmacen
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getAlmacenes() As List(Of EntAlmacen)

        Dim retVal As New List(Of EntAlmacen)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaAlmacenes()

            For Each row As DataRow In DT.Rows

                Dim oAlmacen As EntAlmacen = DataRowToEntidadAlmacen(row)
                retVal.Add(oAlmacen)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaAlmacenes() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ALMACENES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadAlmacen(DR As DataRow) As EntAlmacen

        Dim oAlmacen As New EntAlmacen

        With oAlmacen

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oAlmacen

    End Function

End Class
