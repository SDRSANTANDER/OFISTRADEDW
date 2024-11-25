Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOInterlocutorContacto
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getInterlocutoresContactos(ByVal Tipo As String) As List(Of EntInterlocutorContacto)

        Dim retVal As New List(Of EntInterlocutorContacto)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaInterlocutoresContactos(Tipo)

            For Each row As DataRow In DT.Rows

                Dim oInterlocutorContacto As EntInterlocutorContacto = DataRowToEntidadInterlocutorContacto(row)
                retVal.Add(oInterlocutorContacto)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaInterlocutoresContactos(ByVal Tipo As String) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_INTERLOCUTORES_CONTACTOS") & vbCrLf
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

    Public Function DataRowToEntidadInterlocutorContacto(DR As DataRow) As EntInterlocutorContacto

        Dim oInterlocutorContacto As New EntInterlocutorContacto

        With oInterlocutorContacto

            .Interlocutor = DR.Item("INTERLOCUTOR").ToString
            .InterlocutorTipo = DR.Item("INTERLOCUTORTIPO").ToString
            .InterlocutorNIF = DR.Item("INTERLOCUTORNIF").ToString
            .Codigo = CInt(DR.Item("CODIGO"))
            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString
            .SegundoNombre = DR.Item("SEGUNDONOMBRE").ToString
            .Apellido = DR.Item("APELLIDO").ToString
            .Titulo = DR.Item("TITULO").ToString
            .Posicion = DR.Item("POSICION").ToString
            .Direccion = DR.Item("DIRECCION").ToString
            .Telefono1 = DR.Item("TELEFONO1").ToString
            .Telefono2 = DR.Item("TELEFONO2").ToString
            .TelefonoMovil = DR.Item("TELEFONOMOVIL").ToString
            .Fax = DR.Item("FAX").ToString
            .Email = DR.Item("EMAIL").ToString
            '.EmailGrupo = DR.Item("EMAILGRUPO").ToString
            '.Busca = DR.Item("BUSCA").ToString
            '.Observaciones1 = DR.Item("OBSERVACIONES1").ToString
            '.Observaciones2 = DR.Item("OBSERVACIONES2").ToString
            '.ClaveAcceso = DR.Item("CLAVEACCESO").ToString
            '.NacimientoCiudad = DR.Item("NACIMIENTOCIUDAD").ToString
            '.NacimientoPais = DR.Item("NACIMIENTOPAIS").ToString
            '.NacimientoFecha = CInt(DR.Item("NACIMIENTOFECHA"))
            '.Sexo = DR.Item("SEXO").ToString
            .Profesion = DR.Item("PROFESION").ToString
            .PorDefecto = DR.Item("PORDEFECTO").ToString
            .ContactoFirma = DR.Item("CONTACTOFIRMA").ToString

        End With

        Return oInterlocutorContacto

    End Function

End Class

