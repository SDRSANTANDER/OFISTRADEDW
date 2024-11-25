Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOArticulo
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getArticulos() As List(Of EntArticulo)

        Dim retVal As New List(Of EntArticulo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaArticulos()

            For Each row As DataRow In DT.Rows

                Dim oArticulo As EntArticulo = DataRowToEntidadArticulo(row)
                retVal.Add(oArticulo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaArticulos() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ARTICULOS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadArticulo(DR As DataRow) As EntArticulo

        Dim retval As New EntArticulo

        With retval

            .ID = DR.Item("ID").ToString
            .Descripcion = DR.Item("DESCRIPCION").ToString
            .Activo = DR.Item("ACTIVO").ToString
            .Proveedor = DR.Item("PROVEEDORCODIGO").ToString
			.ReferenciaExterna = DR.Item("REFERENCIAEXTERNA").ToString
            .Tipo = DR.Item("TIPO").ToString
            .Grupo = DR.Item("GRUPO").ToString
            .Venta = DR.Item("VENTA").ToString
            .Compra = DR.Item("COMPRA").ToString
            .Inventario = DR.Item("INVENTARIO").ToString
            .UltimaCompraPrecio = CDbl(DR.Item("ULTIMACOMPRAPRECIO"))
            .UltimaCompraMoneda = DR.Item("ULTIMACOMPRAMONEDA").ToString

        End With

        Return retval

    End Function

End Class
