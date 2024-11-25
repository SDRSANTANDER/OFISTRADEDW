Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOInterlocutor
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getInterlocutores(ByVal Tipo As String) As List(Of EntInterlocutor)

        Dim retVal As New List(Of EntInterlocutor)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaInterlocutores(Tipo)

            For Each row As DataRow In DT.Rows

                Dim oInterlocutor As EntInterlocutor = DataRowToEntidadInterlocutor(row)
                retVal.Add(oInterlocutor)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaInterlocutores(ByVal Tipo As String) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""

            If Not String.IsNullOrEmpty(Tipo) Then

                Select Case Tipo

                    Case CardType.Proveedor, CardType.Cliente, CardType.Lead

                        'Proveedores, clientes, leads
                        SQL = "  SELECT * " & vbCrLf
                        SQL &= " FROM " & Utilidades.putQuotes("SEI_VIEW_DW_INTERLOCUTORES") & vbCrLf
                        SQL &= " WHERE 1=1 " & vbCrLf
                        SQL &= " And " & Utilidades.putQuotes("TIPO") & " = N'" & Tipo & "'" & vbCrLf

                        DT = ObtenerDT(SQL)

                    Case CardType.ClienteConPedidosAbiertos

                        'Clientes con pedidos abiertos
                        SQL = "  SELECT * " & vbCrLf
                        SQL &= " FROM " & Utilidades.putQuotes("SEI_VIEW_DW_INTERLOCUTORES") & " T0 " & vbCrLf
                        SQL &= " INNER JOIN " & Utilidades.putQuotes("ORDR") & " T1 " & Utilidades.getWithNoLock() & " ON T0." & Utilidades.putQuotes("ID") & " = T1." & Utilidades.putQuotes("CardCode") & " " & vbCrLf
                        SQL &= "                                                                                       AND T0." & Utilidades.putQuotes("TIPO") & " = N'" & CardType.Cliente & "' " & vbCrLf
                        SQL &= "                                                                                       AND T1." & Utilidades.putQuotes("DocStatus") & " = N'" & DocStatus.Abierto & "'" & vbCrLf

                        DT = ObtenerDT(SQL)

                End Select

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadInterlocutor(DR As DataRow) As EntInterlocutor

        Dim oInterlocutor As New EntInterlocutor

        With oInterlocutor

            .ID = DR.Item("ID").ToString
            .Tipo = DR.Item("TIPO").ToString
            .RazonSocial = DR.Item("RAZONSOCIAL").ToString
            .NombreExtranjero = DR.Item("NOMBREEXTRANJERO").ToString
            .NIF = DR.Item("NIF").ToString
            .ICDW = DR.Item("ICDW").ToString
            .Idioma = CInt(DR.Item("IDIOMA"))
            .Encargado = CInt(DR.Item("ENCARGADO").ToString)
            .CorreoE = DR.Item("CORREOE").ToString
			.IVAGrupo = DR.Item("IVAGRUPO").ToString
            .IVAPorcentaje = CDbl(DR.Item("IVAPORCENTAJE"))
            .FacturaCiudad = DR.Item("FACTURACIUDAD").ToString
            .ViaPago = DR.Item("VIAPAGO").ToString
            .CondicionPago = CInt(DR.Item("CONDICIONPAGO").ToString)
            .Recurrente = DR.Item("RECURRENTE").ToString
            .Importe = CDbl(DR.Item("IMPORTE"))
            .Auxiliar = DR.Item("AUXILIAR").ToString
            .Activo = DR.Item("ACTIVO").ToString

        End With

        Return oInterlocutor

    End Function

End Class
