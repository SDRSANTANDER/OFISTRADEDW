Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOCuentaContable
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getCuentasContables() As List(Of EntCuentaContable)

        Dim retVal As New List(Of EntCuentaContable)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaCuentasContables()

            For Each row As DataRow In DT.Rows

                Dim oCuentaContable As EntCuentaContable = DataRowToEntidadCuentaContable(row)
                retVal.Add(oCuentaContable)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaCuentasContables() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_CUENTAS_CONTABLES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadCuentaContable(DR As DataRow) As EntCuentaContable

        Dim oCuentaContable As New EntCuentaContable

        With oCuentaContable

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .Padre = DR.Item("PADRE").ToString
            .Nivel = CInt(DR.Item("NIVEL").ToString)

        End With

        Return oCuentaContable

    End Function

End Class
