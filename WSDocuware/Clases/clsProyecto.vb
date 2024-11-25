Imports SAPbobsCOM
Imports System.Globalization
Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class clsProyecto

#Region "Públicas"

    Public Function CrearProyecto(ByVal objProyecto As EntProyectoSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oService As IProjectsService = Nothing
        Dim oProject As Project = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO PROYECTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & ") Inicio de creación de proyecto para " & objProyecto.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objProyecto.UserSAP, objProyecto.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objProyecto.Codigo) Then Throw New Exception("Código no suministrado")
            If String.IsNullOrEmpty(objProyecto.Nombre) Then Throw New Exception("Nombre no suministrado")

            If objProyecto.FechaDesde > 0 AndAlso Not DateTime.TryParseExact(objProyecto.FechaDesde.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha desde no suministrada o incorrecta")
            If objProyecto.FechaHasta > 0 AndAlso Not DateTime.TryParseExact(objProyecto.FechaHasta.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha hasta no suministrada o incorrecta")

            If String.IsNullOrEmpty(objProyecto.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objProyecto.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            'Comprueba si existe el proyecto
            If Not String.IsNullOrEmpty(getExisteProyecto(objProyecto.Codigo, Sociedad)) Then Throw New Exception("Ya existe el proyecto con código: " & objProyecto.Codigo)

            'Objeto proyecto
            oService = oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
            oProject = oService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject)

            'Rellena los campos
            With oProject

                .Code =  objProyecto.Codigo
                .Name =  objProyecto.Nombre

                If DateTime.TryParseExact(objProyecto.FechaDesde.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                    .ValidFrom = Date.ParseExact(objProyecto.FechaDesde.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                If DateTime.TryParseExact(objProyecto.FechaHasta.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                    .ValidTo = Date.ParseExact(objProyecto.FechaHasta.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                .Active = IIf(objProyecto.Activo = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                .UserFields.Item("U_SEIIDDW").Value = objProyecto.DOCIDDW
                .UserFields.Item("U_SEIURLDW").Value = objProyecto.DOCURLDW

            End With

            oService.AddProject(oProject)

            'Añadimos proyecto
            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Proyecto creado con éxito"
            retVal.MENSAJEAUX = oProject.Code

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oProject)
            oComun.LiberarObjCOM(oService)
        End Try

        Return retVal

    End Function

    Public Function ActualizarProyecto(ByVal objProyecto As EntProyectoSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR PROYECTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & ") Inicio de actualización de proyecto: " & objProyecto.Codigo & " para IDDW " & objProyecto.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objProyecto.Codigo) Then Throw New Exception("Código no suministrado")
            If String.IsNullOrEmpty(objProyecto.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objProyecto.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            'Comprueba si existe el proyecto
            If String.IsNullOrEmpty(getExisteProyecto(objProyecto.Codigo, Sociedad)) Then Throw New Exception("No existe el proyecto con código: " & objProyecto.Codigo)

            'Se actualizan los campos de usuario de DW
            setDatosDWProyecto(objProyecto.Codigo, objProyecto.DOCIDDW, objProyecto.DOCURLDW, Sociedad)

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Proyecto actualizado con éxito"
            retVal.MENSAJEAUX = objProyecto.Codigo

            clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX & " - " & retVal.MENSAJEAUX & " - " & retVal.MENSAJEAUX)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

#End Region

#Region "Privadas"

    Private Function getExisteProyecto(ByVal Codigo As String,
                                       ByVal Sociedad As eSociedad) As String

        'Devuelve si existe el proyecto
        Dim retVal As String = ""

        Try

            Dim SQL As String = ""

            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("PrjCode") & " As PrjCode " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OPRJ", Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("PrjCode") & "  = N'" & Codigo & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWProyecto(ByVal Codigo As String,
                                   ByVal IDDW As String,
                                   ByVal URLDW As String,
                                   ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el documento

        Try

            'Actualizamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef("OPRJ", Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("PrjCode") & ",'') = N'" & Codigo & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el proyecto: OPRJ para PrjCode: " & Codigo)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

End Class
