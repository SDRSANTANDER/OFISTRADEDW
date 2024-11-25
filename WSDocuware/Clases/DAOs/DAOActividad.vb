Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOActividad
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

#Region "Actividad "

    Public Function getActividades(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntActividad)

        Dim retVal As New List(Of EntActividad)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActividades(FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oActividad As EntActividad = DataRowToEntidadActividad(row)
                retVal.Add(oActividad)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActividades(ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVIDADES") & vbCrLf
            SQL &= " WHERE 1=1" & vbCrLf
            SQL &= " And " & Utilidades.putQuotes("FECHAINICIO") & ">='" & FechaInicio & "'" & vbCrLf
            SQL &= " And " & Utilidades.putQuotes("FECHAINICIO") & "<='" & FechaFin & "'" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActividad(DR As DataRow) As EntActividad

        Dim oActividad As New EntActividad

        With oActividad

            .IdActividad = CInt(DR.Item("IDACTIVIDAD").ToString)
            .IdInterlocutor = DR.Item("IDINTERLOCUTOR").ToString
            .RazonSocial = DR.Item("RAZONSOCIAL").ToString
            .PersonaContacto = DR.Item("PERSONACONTACTO").ToString
            .Actividad = DR.Item("ACTIVIDAD").ToString
            .Tipo = DR.Item("TIPO").ToString
            .Asunto = DR.Item("ASUNTO").ToString
            .EmpleadoAsignado = DR.Item("EMPLEADOASIGNADO").ToString
            .UsuarioAsignado = DR.Item("USUARIOASIGNADO").ToString
            .Comentarios = DR.Item("COMENTARIOS").ToString
            .FechaInicio = CDate(DR.Item("FECHAINICIO")).ToString("yyyyMMdd")
            .HoraInicio = CInt(DR.Item("HORAINICIO").ToString)
            .FechaFin = CDate(DR.Item("FECHAFIN")).ToString("yyyyMMdd")
            .HoraFin = CInt(DR.Item("HORAFIN").ToString)
            .Prioridad = DR.Item("PRIORIDAD").ToString
            .Emplazamiento = DR.Item("EMPLAZAMIENTO").ToString
            .Repeticion = DR.Item("REPETICION").ToString

        End With

        Return oActividad

    End Function

#End Region

#Region "Actividad tipo"

    Public Function getActividadesTipo() As List(Of EntActividadTipo)

        Dim retVal As New List(Of EntActividadTipo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActividadesTipo()

            For Each row As DataRow In DT.Rows

                Dim oActividadTipo As EntActividadTipo = DataRowToEntidadActividadTipo(row)
                retVal.Add(oActividadTipo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActividadesTipo() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVIDADES_TIPO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActividadTipo(DR As DataRow) As EntActividadTipo

        Dim oActividadTipo As New EntActividadTipo

        With oActividadTipo

            .ID = CInt(DR.Item("ID").ToString)
            .Descripcion = DR.Item("DESCRIPCION").ToString

        End With

        Return oActividadTipo

    End Function

#End Region

#Region "Actividad asunto"

    Public Function getActividadesAsunto() As List(Of EntActividadAsunto)

        Dim retVal As New List(Of EntActividadAsunto)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActividadesAsunto()

            For Each row As DataRow In DT.Rows

                Dim oActividadAsunto As EntActividadAsunto = DataRowToEntidadActividadAsunto(row)
                retVal.Add(oActividadAsunto)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActividadesAsunto() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVIDADES_ASUNTO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActividadAsunto(DR As DataRow) As EntActividadAsunto

        Dim oActividadAsunto As New EntActividadAsunto

        With oActividadAsunto

            .ID = CInt(DR.Item("ID").ToString)
            .Descripcion = DR.Item("DESCRIPCION").ToString
            .Tipo = CInt(DR.Item("TIPO").ToString)

        End With

        Return oActividadAsunto

    End Function

#End Region

#Region "Actividad emplazamiento"

    Public Function getActividadesEmplazamiento() As List(Of EntActividadEmplazamiento)

        Dim retVal As New List(Of EntActividadEmplazamiento)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaActividadesEmplazamiento()

            For Each row As DataRow In DT.Rows

                Dim oActividadEmplazamiento As EntActividadEmplazamiento = DataRowToEntidadActividadEmplazamiento(row)
                retVal.Add(oActividadEmplazamiento)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaActividadesEmplazamiento() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ACTIVIDADES_EMPLAZAMIENTO") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadActividadEmplazamiento(DR As DataRow) As EntActividadEmplazamiento

        Dim oActividadEmplazamiento As New EntActividadEmplazamiento

        With oActividadEmplazamiento

            .ID = CInt(DR.Item("ID").ToString)
            .Descripcion = DR.Item("DESCRIPCION").ToString

        End With

        Return oActividadEmplazamiento

    End Function

#End Region

End Class
