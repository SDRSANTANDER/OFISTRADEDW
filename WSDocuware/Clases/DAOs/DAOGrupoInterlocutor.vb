Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOGrupoInterlocutor
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getGruposInterlocutor() As List(Of EntGrupoInterlocutor)

        Dim retVal As New List(Of EntGrupoInterlocutor)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaGruposInterlocutor()

            For Each row As DataRow In DT.Rows

                Dim oGrupoInterlocutor As EntGrupoInterlocutor = DataRowToEntidadGrupoInterlocutor(row)
                retVal.Add(oGrupoInterlocutor)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaGruposInterlocutor() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_GRUPOS_INTERLOCUTOR") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadGrupoInterlocutor(DR As DataRow) As EntGrupoInterlocutor

        Dim oGrupoInterlocutor As New EntGrupoInterlocutor

        With oGrupoInterlocutor

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oGrupoInterlocutor

    End Function

End Class
