Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOFabricante
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getFabricantes() As List(Of EntFabricante)

        Dim retVal As New List(Of EntFabricante)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaFabricantes()

            For Each row As DataRow In DT.Rows

                Dim oFabricante As EntFabricante = DataRowToEntidadFabricante(row)
                retVal.Add(oFabricante)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaFabricantes() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_FABRICANTES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadFabricante(DR As DataRow) As EntFabricante

        Dim oFabricante As New EntFabricante

        With oFabricante

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oFabricante

    End Function

End Class
