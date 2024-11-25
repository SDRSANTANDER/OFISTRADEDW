Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOGrupoArticulo
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getGruposArticulo() As List(Of EntGrupoArticulo)

        Dim retVal As New List(Of EntGrupoArticulo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaGruposArticulo()

            For Each row As DataRow In DT.Rows

                Dim oGrupoArticulo As EntGrupoArticulo = DataRowToEntidadGrupoArticulo(row)
                retVal.Add(oGrupoArticulo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaGruposArticulo() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_GRUPOS_ARTICULO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadGrupoArticulo(DR As DataRow) As EntGrupoArticulo

        Dim oGrupoArticulo As New EntGrupoArticulo

        With oGrupoArticulo

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oGrupoArticulo

    End Function

End Class
