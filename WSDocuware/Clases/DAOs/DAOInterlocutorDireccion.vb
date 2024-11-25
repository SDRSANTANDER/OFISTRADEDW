Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOInterlocutorDireccion
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getInterlocutoresDirecciones(ByVal Tipo As String) As List(Of EntInterlocutorDireccion)

        Dim retVal As New List(Of EntInterlocutorDireccion)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaInterlocutoresDirecciones(Tipo)

            For Each row As DataRow In DT.Rows

                Dim oInterlocutorDireccion As EntInterlocutorDireccion = DataRowToEntidadInterlocutorDireccion(row)
                retVal.Add(oInterlocutorDireccion)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaInterlocutoresDirecciones(ByVal Tipo As String) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_INTERLOCUTORES_DIRECCIONES") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf

            If Not String.IsNullOrEmpty(Tipo) Then
                SQL &= " And " & Utilidades.putQuotes("INTERLOCUTORTIPO") & " = N'" & Tipo & "'" & vbCrLf
            End If

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadInterlocutorDireccion(DR As DataRow) As EntInterlocutorDireccion

        Dim oInterlocutorDireccion As New EntInterlocutorDireccion

        With oInterlocutorDireccion

            .Interlocutor = DR.Item("INTERLOCUTOR").ToString
            .InterlocutorTipo = DR.Item("INTERLOCUTORTIPO").ToString
            .InterlocutorNIF = DR.Item("INTERLOCUTORNIF").ToString
            .Tipo = DR.Item("TIPO").ToString
            .ID = DR.Item("ID").ToString
            .Nombre2 = DR.Item("NOMBRE2").ToString
            .Nombre3 = DR.Item("NOMBRE3").ToString
            .Calle = DR.Item("CALLE").ToString
            .Bloque = DR.Item("BLOQUE").ToString
            .Ciudad = DR.Item("CIUDAD").ToString
            .CodigoPostal = DR.Item("CODIGOPOSTAL").ToString
            .Provincia = DR.Item("PROVINCIA").ToString
            .Estado = DR.Item("ESTADO").ToString
            .Pais = DR.Item("PAIS").ToString
            .NIF = DR.Item("NIF").ToString
            .NumeroCalle = DR.Item("NUMEROCALLE").ToString
            .Edificio = DR.Item("EDIFICIO").ToString
            .DelegacionHacienda = DR.Item("DELEGACIONHACIENDA").ToString
            .GLN = DR.Item("GLN").ToString
            .PorDefecto = DR.Item("PORDEFECTO").ToString

        End With

        Return oInterlocutorDireccion

    End Function

End Class

