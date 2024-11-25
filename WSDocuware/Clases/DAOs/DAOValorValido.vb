Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOValorValido
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getValoresValidos(ByVal TablaID As String, ByVal AliasID As String) As List(Of EntValorValido)

        Dim retVal As New List(Of EntValorValido)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaValoresValidos(TablaID, AliasID)

            For Each row As DataRow In DT.Rows

                Dim oValorValido As EntValorValido = DataRowToEntidadValorValido(row)
                retVal.Add(oValorValido)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaValoresValidos(ByVal TablaID As String, ByVal AliasID As String) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_VALORES_VALIDOS") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TablaID) Then
                SQL &= " AND " & putQuotes("TABLAID") & " = N'" & TablaID & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(AliasID) Then
                SQL &= " AND " & putQuotes("ALIASID") & " = N'" & AliasID & "'" & vbCrLf
            End If

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadValorValido(DR As DataRow) As EntValorValido

        Dim oValorValido As New EntValorValido

        With oValorValido

            .TablaID = DR.Item("TABLAID").ToString
            .CampoID = DR.Item("CAMPOID").ToString
            .AliasID = DR.Item("ALIASID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .ValorValidoDefecto = DR.Item("VALORVALIDODEFECTO").ToString

            .ValoresValidosID = DR.Item("VALORESVALIDOSID").ToString
            If .ValoresValidosID.Length > 0 AndAlso .ValoresValidosID.Contains("||") Then
                .ValoresValidosID = .ValoresValidosID.Substring(0, .ValoresValidosID.Length - 2)
            End If

            .ValoresValidosNombre = DR.Item("VALORESVALIDOSNOMBRE").ToString
            If .ValoresValidosNombre.Length > 0 AndAlso .ValoresValidosNombre.Contains("||") Then
                .ValoresValidosNombre = .ValoresValidosNombre.Substring(0, .ValoresValidosNombre.Length - 2)
            End If

        End With

        Return oValorValido

    End Function

End Class
