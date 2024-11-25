Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOLlamadaServicio
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub


#Region "Llamada servicio"

    Public Function getLlamadasServicio(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntLlamadaServicio)

        Dim retVal As New List(Of EntLlamadaServicio)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaLlamadasServicio(FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oLlamadaServicio As EntLlamadaServicio = DataRowToEntidadLlamadaServicio(row)
                retVal.Add(oLlamadaServicio)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaLlamadasServicio(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_LLAMADAS_SERVICIO") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf
            SQL &= " And " & Utilidades.putQuotes("CREATEDATE") & ">='" & FechaInicio & "'" & vbCrLf
            SQL &= " And " & Utilidades.putQuotes("CREATEDATE") & "<='" & FechaFin & "'" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadLlamadaServicio(DR As DataRow) As EntLlamadaServicio

        Dim oLlamadaServicio As New EntLlamadaServicio

        With oLlamadaServicio

            .Tipo = DR.Item("TIPO").ToString
            .DocNum = CInt(DR.Item("DOCNUM").ToString)
            .IdLlamada = CInt(DR.Item("IDLLAMADA").ToString)
            .CreateDate = CDate(DR.Item("CREATEDATE")).ToString("yyyyMMdd")
            .CreateTime = CInt(DR.Item("CREATETIME").ToString)
            .IdInterlocutor = DR.Item("IDINTERLOCUTOR").ToString
            .RazonSocial = DR.Item("RAZONSOCIAL").ToString
            .PersonaContacto = DR.Item("PERSONACONTACTO").ToString
            .Asunto = DR.Item("ASUNTO").ToString
            .NumAtCard = DR.Item("NUMATCARD").ToString
            .NumeroSerieFabricante = DR.Item("NUMEROSERIEFABRICANTE").ToString
            .NumeroSerie = DR.Item("NUMEROSERIE").ToString
            .Articulo = DR.Item("ARTICULO").ToString
            .Descripcion = DR.Item("DESCRIPCION").ToString
            .ReferenciaExterna = DR.Item("REFERENCIAEXTERNA").ToString
            .IdInterlocutorDefecto = DR.Item("IDINTERLOCUTORDEFECTO").ToString
            .RazonSocialDefecto = DR.Item("RAZONSOCIALDEFECTO").ToString
            .Estado = DR.Item("ESTADO").ToString
            .Prioridad = DR.Item("PRIORIDAD").ToString
            .Origen = DR.Item("ORIGEN").ToString
            .TipoProblema = DR.Item("TIPOPROBLEMA").ToString
            .SubtipoProblema = DR.Item("SUBTIPOPROBLEMA").ToString
            .TipoLlamada = DR.Item("TIPOLLAMADA").ToString
            .Tecnico = DR.Item("TECNICO").ToString.Trim
            .EsCola = DR.Item("ESCOLA").ToString
            .Cola = DR.Item("COLA").ToString
            .Comentarios = DR.Item("COMENTARIOS").ToString

        End With

        Return oLlamadaServicio

    End Function

#End Region


#Region "Llamada servicio estado"

    Public Function getLlamadasServicioEstado() As List(Of EntLlamadaServicioEstado)

        Dim retVal As New List(Of EntLlamadaServicioEstado)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaLlamadasServicioEstado()

            For Each row As DataRow In DT.Rows

                Dim oLlamadaServicioEstado As EntLlamadaServicioEstado = DataRowToEntidadLlamadaServicioEstado(row)
                retVal.Add(oLlamadaServicioEstado)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaLlamadasServicioEstado() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_LLAMADAS_ESTADO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadLlamadaServicioEstado(DR As DataRow) As EntLlamadaServicioEstado

        Dim oLlamadaServicioEstado As New EntLlamadaServicioEstado

        With oLlamadaServicioEstado

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString
            .Descripcion = DR.Item("DESCRIPCION").ToString

        End With

        Return oLlamadaServicioEstado

    End Function

#End Region

#Region "Llamada servicio tipo"

    Public Function getLlamadasServicioTipo() As List(Of EntLlamadaServicioTipo)

        Dim retVal As New List(Of EntLlamadaServicioTipo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaLlamadasServicioTipo()

            For Each row As DataRow In DT.Rows

                Dim oLlamadaServicioTipo As EntLlamadaServicioTipo = DataRowToEntidadLlamadaServicioTipo(row)
                retVal.Add(oLlamadaServicioTipo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaLlamadasServicioTipo() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_LLAMADAS_TIPO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadLlamadaServicioTipo(DR As DataRow) As EntLlamadaServicioTipo

        Dim oLlamadaServicioTipo As New EntLlamadaServicioTipo

        With oLlamadaServicioTipo

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString
            .Descripcion = DR.Item("DESCRIPCION").ToString

        End With

        Return oLlamadaServicioTipo

    End Function

#End Region

#Region "Llamada servicio origen"

    Public Function getLlamadasServicioOrigen() As List(Of EntLlamadaServicioOrigen)

        Dim retVal As New List(Of EntLlamadaServicioOrigen)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaLlamadasServicioOrigen()

            For Each row As DataRow In DT.Rows

                Dim oLlamadaServicioOrigen As EntLlamadaServicioOrigen = DataRowToEntidadLlamadaServicioOrigen(row)
                retVal.Add(oLlamadaServicioOrigen)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaLlamadasServicioOrigen() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_LLAMADAS_ORIGEN") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadLlamadaServicioOrigen(DR As DataRow) As EntLlamadaServicioOrigen

        Dim oLlamadaServicioOrigen As New EntLlamadaServicioOrigen

        With oLlamadaServicioOrigen

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString
            .Descripcion = DR.Item("DESCRIPCION").ToString

        End With

        Return oLlamadaServicioOrigen

    End Function

#End Region

#Region "Llamada servicio problema tipo"

    Public Function getLlamadasServicioProblemaTipo() As List(Of EntLlamadaServicioProblemaTipo)

        Dim retVal As New List(Of EntLlamadaServicioProblemaTipo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaLlamadasServicioProblemaTipo()

            For Each row As DataRow In DT.Rows

                Dim oLlamadaServicioProblemaTipo As EntLlamadaServicioProblemaTipo = DataRowToEntidadLlamadaServicioProblemaTipo(row)
                retVal.Add(oLlamadaServicioProblemaTipo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaLlamadasServicioProblemaTipo() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_LLAMADAS_PROBLEMA_TIPO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadLlamadaServicioProblemaTipo(DR As DataRow) As EntLlamadaServicioProblemaTipo

        Dim oLlamadaServicioProblemaTipo As New EntLlamadaServicioProblemaTipo

        With oLlamadaServicioProblemaTipo

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString
            .Descripcion = DR.Item("DESCRIPCION").ToString

        End With

        Return oLlamadaServicioProblemaTipo

    End Function

#End Region

#Region "Llamada servicio problema subtipo"

    Public Function getLlamadasServicioProblemaSubtipo() As List(Of EntLlamadaServicioProblemaSubtipo)

        Dim retVal As New List(Of EntLlamadaServicioProblemaSubtipo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaLlamadasServicioProblemaSubtipo()

            For Each row As DataRow In DT.Rows

                Dim oLlamadaServicioProblemaSubtipo As EntLlamadaServicioProblemaSubtipo = DataRowToEntidadLlamadaServicioProblemaSubtipo(row)
                retVal.Add(oLlamadaServicioProblemaSubtipo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaLlamadasServicioProblemaSubtipo() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_LLAMADAS_PROBLEMA_SUBTIPO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadLlamadaServicioProblemaSubtipo(DR As DataRow) As EntLlamadaServicioProblemaSubtipo

        Dim oLlamadaServicioProblemaSubtipo As New EntLlamadaServicioProblemaSubtipo

        With oLlamadaServicioProblemaSubtipo

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString
            .Descripcion = DR.Item("DESCRIPCION").ToString

        End With

        Return oLlamadaServicioProblemaSubtipo

    End Function

#End Region

#Region "Llamada servicio cola"

    Public Function getLlamadasServicioCola() As List(Of EntLlamadaServicioCola)

        Dim retVal As New List(Of EntLlamadaServicioCola)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaLlamadasServicioCola()

            For Each row As DataRow In DT.Rows

                Dim oLlamadaServicioCola As EntLlamadaServicioCola = DataRowToEntidadLlamadaServicioCola(row)
                retVal.Add(oLlamadaServicioCola)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaLlamadasServicioCola() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_LLAMADAS_COLA") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadLlamadaServicioCola(DR As DataRow) As EntLlamadaServicioCola

        Dim oLlamadaServicioCola As New EntLlamadaServicioCola

        With oLlamadaServicioCola

            .ID = CInt(DR.Item("ID").ToString)
            .Nombre = DR.Item("NOMBRE").ToString
            .Responsable = CInt(DR.Item("RESPONSABLE").ToString)
            .Email = DR.Item("EMAIL").ToString

        End With

        Return oLlamadaServicioCola

    End Function

#End Region

End Class
