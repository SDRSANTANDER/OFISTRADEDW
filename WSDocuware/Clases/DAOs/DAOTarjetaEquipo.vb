Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOTarjetaEquipo
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getTarjetasEquipo() As List(Of EntTarjetaEquipo)

        Dim retVal As New List(Of EntTarjetaEquipo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaTarjetasEquipo()

            For Each row As DataRow In DT.Rows

                Dim oTarjetaEquipo As EntTarjetaEquipo = DataRowToEntidadTarjetaEquipo(row)
                retVal.Add(oTarjetaEquipo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaTarjetasEquipo() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_TARJETAS_EQUIPOS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadTarjetaEquipo(DR As DataRow) As EntTarjetaEquipo

        Dim retval As New EntTarjetaEquipo

        With retval

            .ID = CInt(DR.Item("ID").ToString)
            .Tipo = DR.Item("TIPO").ToString
            .NumFabricante = DR.Item("NUMFABRICANTE").ToString
            .NumSerie = DR.Item("NUMSERIE").ToString
            .Articulo = DR.Item("ARTICULO").ToString
            .Descripcion = DR.Item("DESCRIPCION").ToString
            .Interlocutor = DR.Item("INTERLOCUTOR").ToString

        End With

        Return retval

    End Function

End Class
