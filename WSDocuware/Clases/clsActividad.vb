Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsActividad

#Region "Públicas"

    Public Function NuevaActividad(ByVal objActividad As EntActividadSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company

        Dim oCompanyService As CompanyService = Nothing
        Dim oActivityService As ActivitiesService = Nothing
        Dim oActivity As Activity = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVA ACTIVIDAD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de actividad para " & objActividad.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objActividad.UserSAP, objActividad.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objActividad.NIFTercero) Then Throw New Exception("NIF no suministrado")

            If objActividad.FechaHoraInicio > 0 AndAlso Not DateTime.TryParseExact(objActividad.FechaHoraInicio.ToString, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha/Hora inicio no suministrada o incorrecta")
            If objActividad.FechaHoraFin > 0 AndAlso Not DateTime.TryParseExact(objActividad.FechaHoraFin.ToString, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha/Hora fin no suministrada o incorrecta")

            'If String.IsNullOrEmpty(objActividad.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            'If String.IsNullOrEmpty(objActividad.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objActividad.NIFTercero, objActividad.RazonSocial, objActividad.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objActividad.NIFTercero & ", Razón social: " & objActividad.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Objeto actividad
            oCompanyService = oCompany.GetCompanyService
            oActivityService = oCompanyService.GetBusinessService(ServiceTypes.ActivitiesService)
            oActivity = oActivityService.GetDataInterface(ActivitiesServiceDataInterfaces.asActivity)

            'Rellena la actividad
            With oActivity

                'Actividad
                If objActividad.Actividad > -1 Then .Activity = objActividad.Actividad

                'Tipo
                If objActividad.Tipo <> 0 AndAlso objActividad.Tipo > -2 Then .ActivityType = objActividad.Tipo

                'CardCode
                .CardCode = CardCode

                'Asunto
                If objActividad.Asunto > 0 Then .Subject = objActividad.Asunto

                'Asignado a
                If objActividad.AsignadoAEmpleado > 0 Then
                    .HandledByEmployee = objActividad.AsignadoAEmpleado
                ElseIf objActividad.AsignadoAUsuario > 0 Then
                    .HandledBy = objActividad.AsignadoAUsuario
                End If

                'Persona contacto
                If objActividad.PersonaContacto > 0 Then .ContactPersonCode = objActividad.PersonaContacto

                'Telefono
                If Not String.IsNullOrEmpty(objActividad.Telefono) Then .Phone = objActividad.Telefono

                'Comentarios
                If Not String.IsNullOrEmpty(objActividad.Comentarios) Then .Details = objActividad.Comentarios

                'Contenido
                If Not String.IsNullOrEmpty(objActividad.Contenido) Then .Notes = objActividad.Contenido

                'Fechas
                If objActividad.FechaHoraInicio > 0 AndAlso objActividad.FechaHoraFin > 0 Then

                    Dim FechaHoraInicio As DateTime = DateTime.ParseExact(objActividad.FechaHoraInicio.ToString, "yyyyMMddHHmmss", CultureInfo.CurrentCulture)
                    Dim FechaHoraFin As DateTime = DateTime.ParseExact(objActividad.FechaHoraFin.ToString, "yyyyMMddHHmmss", CultureInfo.CurrentCulture)

                    .StartDate = FechaHoraInicio
                    .StartTime = FechaHoraInicio
                    .EndDuedate = FechaHoraFin
                    .EndTime = FechaHoraFin
                    .DurationType = BoDurations.du_Minuts

                End If

                'Prioridad
                If objActividad.Prioridad > -1 Then .Priority = objActividad.Prioridad

                'Emplazamiento
                If objActividad.Emplazamiento = -2 OrElse objActividad.Emplazamiento > 0 Then .Location = objActividad.Emplazamiento

                'Recordatorio
                .Reminder = BoYesNoEnum.tNO

                'ID DOCUWARE
                If Not String.IsNullOrEmpty(objActividad.DOCIDDW) Then .UserFields.Item("U_SEIIDDW").Value = objActividad.DOCIDDW

                'URL DOCUWARE
                If Not String.IsNullOrEmpty(objActividad.DOCURLDW) Then .UserFields.Item("U_SEIURLDW").Value = objActividad.DOCURLDW

                If objActividad.Cerrar = SN.Si Then .Closed = BoYesNoEnum.tYES

            End With

            'Añadimos la actividad
            oActivityService.AddActivity(oActivity)

            'Obtiene el código de la actividad añadida
            Dim NumeroActividad As String = getNumeroDeCardCode(CardCode, Sociedad)

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Actividad creada con éxito"
            retVal.MENSAJEAUX = NumeroActividad

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oCompanyService)
            oComun.LiberarObjCOM(oActivityService)
            oComun.LiberarObjCOM(oActivity)
        End Try

        Return retVal

    End Function

    Public Function ActualizarActividad(ByVal objActividad As EntActividadSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company

        Dim oCompanyService As CompanyService = Nothing
        Dim oActivityService As ActivitiesService = Nothing
        Dim oActivity As Activity = Nothing
        Dim oParams As ActivityParams = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR ACTIVIDAD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de actividad: " & objActividad.Numero)

            oCompany = ConexionSAP.getCompany(objActividad.UserSAP, objActividad.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objActividad.Numero) Then Throw New Exception("Número no suministrado")

            If objActividad.FechaHoraInicio > 0 AndAlso Not DateTime.TryParseExact(objActividad.FechaHoraInicio.ToString, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha/Hora inicio no suministrada o incorrecta")

            If objActividad.FechaHoraFin > 0 AndAlso Not DateTime.TryParseExact(objActividad.FechaHoraFin.ToString, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha/Hora fin no suministrada o incorrecta")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objActividad.NIFTercero, objActividad.RazonSocial, objActividad.Ambito, Sociedad)
            If Not String.IsNullOrEmpty(CardCode) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Objeto actividad
            oCompanyService = oCompany.GetCompanyService
            oActivityService = oCompanyService.GetBusinessService(ServiceTypes.ActivitiesService)
            oParams = oActivityService.GetDataInterface(ActivitiesServiceDataInterfaces.asActivityParams)
            oParams.ActivityCode = CInt(objActividad.Numero)
            oActivity = oActivityService.GetActivity(oParams)

            'Rellena la actividad
            With oActivity

                'Actividad
                If objActividad.Actividad > -1 Then .Activity = objActividad.Actividad

                'Tipo
                If objActividad.Tipo <> 0 AndAlso objActividad.Tipo > -2 Then .ActivityType = objActividad.Tipo

                'CardCode
                If Not String.IsNullOrEmpty(CardCode) Then .CardCode = CardCode

                'Asunto
                If objActividad.Asunto > 0 Then .Subject = objActividad.Asunto

                'Asignado a
                If objActividad.AsignadoAEmpleado > 0 Then
                    .HandledByEmployee = objActividad.AsignadoAEmpleado
                ElseIf objActividad.AsignadoAUsuario > 0 Then
                    .HandledBy = objActividad.AsignadoAUsuario
                End If

                'Persona contacto
                If objActividad.PersonaContacto > 0 Then .ContactPersonCode = objActividad.PersonaContacto

                'Telefono
                If Not String.IsNullOrEmpty(objActividad.Telefono) Then .Phone = objActividad.Telefono

                'Comentarios
                If Not String.IsNullOrEmpty(objActividad.Comentarios) Then .Details = objActividad.Comentarios

                'Contenido
                If Not String.IsNullOrEmpty(objActividad.Contenido) Then .Notes = objActividad.Contenido

                'Fechas
                If objActividad.FechaHoraInicio > 0 AndAlso objActividad.FechaHoraFin > 0 Then

                    Dim FechaHoraInicio As DateTime = DateTime.ParseExact(objActividad.FechaHoraInicio.ToString, "yyyyMMddHHmmss", CultureInfo.CurrentCulture)
                    Dim FechaHoraFin As DateTime = DateTime.ParseExact(objActividad.FechaHoraFin.ToString, "yyyyMMddHHmmss", CultureInfo.CurrentCulture)

                    .StartDate = FechaHoraInicio
                    .StartTime = FechaHoraInicio

                    .EndDuedate = FechaHoraFin
                    .EndTime = FechaHoraFin

                    Dim duracion As Double = (FechaHoraFin - FechaHoraInicio).TotalSeconds
                    oActivity.Duration = duracion

                End If

                'Prioridad
                If objActividad.Prioridad > -1 Then .Priority = objActividad.Prioridad

                'Emplazamiento
                If objActividad.Emplazamiento = -2 OrElse objActividad.Emplazamiento > 0 Then .Location = objActividad.Emplazamiento

                'Recordatorio
                .Reminder = BoYesNoEnum.tNO

                'ID DOCUWARE
                If Not String.IsNullOrEmpty(objActividad.DOCIDDW) Then .UserFields.Item("U_SEIIDDW").Value = objActividad.DOCIDDW

                'URL DOCUWARE
                If Not String.IsNullOrEmpty(objActividad.DOCURLDW) Then .UserFields.Item("U_SEIURLDW").Value = objActividad.DOCURLDW

                If objActividad.Cerrar = SN.Si Then .Closed = BoYesNoEnum.tYES Else .Closed = BoYesNoEnum.tNO

            End With

            'Actualiza la actividad
            oActivityService.UpdateActivity(oActivity)

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Actividad actualizada con éxito"
            retVal.MENSAJEAUX = objActividad.Numero

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oCompanyService)
            oComun.LiberarObjCOM(oActivityService)
            oComun.LiberarObjCOM(oActivity)
            oComun.LiberarObjCOM(oParams)
        End Try

        Return retVal

    End Function

#End Region

#Region "Privadas"

    Private Function getNumeroDeCardCode(ByVal CardCode As String,
                                         ByVal Sociedad As eSociedad) As String

        'Devuelve el número de actividad

        Dim retVal As String = ""

        Try

            'Buscamos por DocNum, DocEntry o NumAtCard
            Dim SQL As String = ""

            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("ClgCode") & " As ClgCode " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OCLG", Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf

            SQL &= " ORDER BY " & vbCrLf
            SQL &= " T0." & putQuotes("ClgCode") & " DESC " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

#End Region

End Class
