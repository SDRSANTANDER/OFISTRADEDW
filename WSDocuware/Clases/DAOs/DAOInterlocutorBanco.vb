Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOInterlocutorBanco
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getInterlocutoresBancos(ByVal Tipo As String) As List(Of EntInterlocutorBanco)

        Dim retVal As New List(Of EntInterlocutorBanco)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaInterlocutoresBancos(Tipo)

            For Each row As DataRow In DT.Rows

                Dim oInterlocutorBanco As EntInterlocutorBanco = DataRowToEntidadInterlocutorBanco(row)
                retVal.Add(oInterlocutorBanco)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaInterlocutoresBancos(ByVal Tipo As String) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_INTERLOCUTORES_BANCOS") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            If Not String.IsNullOrEmpty(Tipo) Then
                SQL &= " And " & Utilidades.putQuotes("TIPO") & " = N'" & Tipo & "'" & vbCrLf
            End If

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadInterlocutorBanco(DR As DataRow) As EntInterlocutorBanco

        Dim oInterlocutorBanco As New EntInterlocutorBanco

        With oInterlocutorBanco

            .Interlocutor = DR.Item("INTERLOCUTOR").ToString
            .Tipo = DR.Item("TIPO").ToString
            .IBAN = DR.Item("IBAN").ToString
            .Banco = DR.Item("BANCO").ToString
            .Sucursal = DR.Item("SUCURSAL").ToString
            .DigitosControl = DR.Item("DIGITOSCONTROL").ToString
            .CuentaNumero = DR.Item("CUENTANUMERO").ToString
            .CuentaNombre = DR.Item("CUENTANOMBRE").ToString
            .Pais = DR.Item("PAIS").ToString
            .CuentaContable = DR.Item("CUENTACONTABLE").ToString

        End With

        Return oInterlocutorBanco

    End Function

End Class

