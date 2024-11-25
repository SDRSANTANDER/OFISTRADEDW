Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOSucursal
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getSucursales() As List(Of EntSucursal)

        Dim retVal As New List(Of EntSucursal)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaSucursales()

            For Each row As DataRow In DT.Rows

                Dim oSucursal As EntSucursal = DataRowToEntidadSucursal(row)
                retVal.Add(oSucursal)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaSucursales() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_SUCURSALES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadSucursal(DR As DataRow) As EntSucursal

        Dim oSucursal As New EntSucursal

        With oSucursal

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oSucursal

    End Function

End Class
