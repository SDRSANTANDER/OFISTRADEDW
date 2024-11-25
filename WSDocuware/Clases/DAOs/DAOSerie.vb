Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOSerie
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getSeries(ByVal ObjType As Integer) As List(Of EntSerie)

        Dim retVal As New List(Of EntSerie)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaSeries(ObjType)

            For Each row As DataRow In DT.Rows

                Dim oSerie As EntSerie = DataRowToEntidadSerie(row)
                retVal.Add(oSerie)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaSeries(ByVal ObjType As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_SERIES") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            If ObjType > 0 Then
                SQL &= " AND " & putQuotes("OBJTYPE") & " = N'" & ObjType & "'" & vbCrLf
            End If

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadSerie(DR As DataRow) As EntSerie

        Dim oSerie As New EntSerie

        With oSerie

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString
            .ObjType = CInt(DR.Item("OBJTYPE").ToString)
            .Manual = DR.Item("MANUAL").ToString

        End With

        Return oSerie

    End Function

End Class
