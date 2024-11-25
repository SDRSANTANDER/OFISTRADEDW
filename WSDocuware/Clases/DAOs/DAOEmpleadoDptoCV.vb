Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOEmpleadoDptoCV
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getEmpleadosDptoCVs() As List(Of EntEmpleadoDptoCV)

        Dim retVal As New List(Of EntEmpleadoDptoCV)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaEmpleadosDptoCVs()

            For Each row As DataRow In DT.Rows

                Dim oEmpleadoDptoCV As EntEmpleadoDptoCV = DataRowToEntidadEmpleadoDptoCV(row)
                retVal.Add(oEmpleadoDptoCV)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaEmpleadosDptoCVs() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_EMPLEADOS_DPTO_CV") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadEmpleadoDptoCV(DR As DataRow) As EntEmpleadoDptoCV

        Dim oEmpleadoDptoCV As New EntEmpleadoDptoCV

        With oEmpleadoDptoCV

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .EmpID = DR.Item("EMPID").ToString
            .Telefono = DR.Item("TELEFONO").ToString
            .CorreoE = DR.Item("CORREOE").ToString

        End With

        Return oEmpleadoDptoCV

    End Function

End Class
