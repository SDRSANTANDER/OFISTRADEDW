Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOPorte
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getPortes() As List(Of EntPorte)

        Dim retVal As New List(Of EntPorte)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaPortes()

            For Each row As DataRow In DT.Rows

                Dim oPorte As EntPorte = DataRowToEntidadPorte(row)
                retVal.Add(oPorte)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaPortes() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_PORTES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadPorte(DR As DataRow) As EntPorte

        Dim oPorte As New EntPorte

        With oPorte

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oPorte

    End Function

End Class
