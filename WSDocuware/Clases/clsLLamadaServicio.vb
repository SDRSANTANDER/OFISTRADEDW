Imports SAPbobsCOM
Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class clsLLamadaServicio

#Region "Públicas"

    Public Function NuevaLlamadaServicio(ByVal objLlamadaServicio As EntLlamadaServicioSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company

        Dim oCall As ServiceCalls = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVA LLAMADA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de llamada de servicio para " & objLlamadaServicio.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objLlamadaServicio.UserSAP, objLlamadaServicio.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objLlamadaServicio.NIFTercero) Then Throw New Exception("NIF no suministrado")
            If String.IsNullOrEmpty(objLlamadaServicio.Asunto) Then Throw New Exception("Asunto no suministrado")
            'If String.IsNullOrEmpty(objLlamadaServicio.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            'If String.IsNullOrEmpty(objLlamadaServicio.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objLlamadaServicio.NIFTercero, objLlamadaServicio.RazonSocial, objLlamadaServicio.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objLlamadaServicio.NIFTercero & ", Razón social: " & objLlamadaServicio.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Objeto llamada
            oCall = CType(oCompany.GetBusinessObject(BoObjectTypes.oServiceCalls), ServiceCalls)

            'Rellena la llamada servicio
            With oCall

                'Serie
                If objLlamadaServicio.Serie > 0 Then .Series = objLlamadaServicio.Serie

                'Status
                If objLlamadaServicio.Estado > -4 AndAlso objLlamadaServicio.Estado <> 0 Then .Status = objLlamadaServicio.Estado

                'Prioridad
                If objLlamadaServicio.Prioridad > -1 Then .Priority = objLlamadaServicio.Prioridad

                'Cliente
                .CustomerCode = CardCode

                'PersonaContacto
                If objLlamadaServicio.PersonaContacto > 0 Then .ContactCode = objLlamadaServicio.PersonaContacto

                'Telefono
                If Not String.IsNullOrEmpty(objLlamadaServicio.Telefono) Then .Telephone = objLlamadaServicio.Telefono

                'NumAtCard
                If Not String.IsNullOrEmpty(objLlamadaServicio.NumAtCard) Then .CustomerRefNo = objLlamadaServicio.NumAtCard

                'NumeroSerieFabricante
                If Not String.IsNullOrEmpty(objLlamadaServicio.NumeroSerieFabricante) Then .ManufacturerSerialNum = objLlamadaServicio.NumeroSerieFabricante

                'NumeroSerie
                If Not String.IsNullOrEmpty(objLlamadaServicio.NumeroSerie) Then .InternalSerialNum = objLlamadaServicio.NumeroSerie

                'ItemCode
                If Not String.IsNullOrEmpty(objLlamadaServicio.ItemCode) Then .ItemCode = objLlamadaServicio.ItemCode

                'Asunto
                If Not String.IsNullOrEmpty(objLlamadaServicio.Asunto) Then .Subject = objLlamadaServicio.Asunto

                'Origen
                If objLlamadaServicio.Origen > -4 AndAlso objLlamadaServicio.Origen <> 0 Then .Origin = objLlamadaServicio.Origen

                'ProblemaTipo
                If objLlamadaServicio.ProblemaTipo > 0 Then .ProblemType = objLlamadaServicio.ProblemaTipo

                'ProblemaSubtipo
                If objLlamadaServicio.ProblemaSubtipo > 0 Then .ProblemSubType = objLlamadaServicio.ProblemaSubtipo

                'TipoLlamada
                If objLlamadaServicio.LlamadaTipo > 0 Then .CallType = objLlamadaServicio.LlamadaTipo

                'Tecnico
                If objLlamadaServicio.Tecnico > 0 Then .TechnicianCode = objLlamadaServicio.Tecnico

                'TratadoPor
                If objLlamadaServicio.TratadoPor > 0 Then .AssigneeCode = objLlamadaServicio.TratadoPor

                'Resolucion
                If Not String.IsNullOrEmpty(objLlamadaServicio.Resolucion) Then .Resolution = objLlamadaServicio.Resolucion

                'ID DOCUWARE
                If Not String.IsNullOrEmpty(objLlamadaServicio.DOCIDDW) Then .UserFields.Item("U_SEIIDDW").Value = objLlamadaServicio.DOCIDDW

                'URL DOCUWARE
                If Not String.IsNullOrEmpty(objLlamadaServicio.DOCURLDW) Then .UserFields.Item("U_SEIURLDW").Value = objLlamadaServicio.DOCURLDW

            End With

            'Añadimos llamada de servicio
            If oCall.Add() = 0 Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Llamada servicio creada con éxito"
                retVal.MENSAJEAUX = getNumeroDeCardCode(CardCode, Sociedad)

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = oCompany.GetLastErrorCode & "#" & oCompany.GetLastErrorDescription
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oCall)
        End Try

        Return retVal

    End Function

    Public Function ActualizarLlamadaServicio(ByVal objLlamadaServicio As EntLlamadaServicioSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company

        Dim oCall As ServiceCalls = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR LLAMADA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de llamada de servicio para " & objLlamadaServicio.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objLlamadaServicio.UserSAP, objLlamadaServicio.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objLlamadaServicio.Numero) Then Throw New Exception("Número no suministrado")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objLlamadaServicio.NIFTercero, objLlamadaServicio.RazonSocial, objLlamadaServicio.Ambito, Sociedad)
            If Not String.IsNullOrEmpty(CardCode) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Objeto llamada
            oCall = CType(oCompany.GetBusinessObject(BoObjectTypes.oServiceCalls), ServiceCalls)

            'Comprueba si existe la llamada
            If oCall.GetByKey(CInt(objLlamadaServicio.Numero)) <> True Then

                Throw New Exception("No existe la llamada de servicio " & objLlamadaServicio.Numero)

            Else

                'Rellena la llamada servicio
                With oCall

                    'Serie
                    'If objLlamadaServicio.Serie > 0 Then .Series = objLlamadaServicio.Serie

                    'Status
                    If objLlamadaServicio.Estado > -4 AndAlso objLlamadaServicio.Estado <> 0 Then .Status = objLlamadaServicio.Estado

                    'Prioridad
                    If objLlamadaServicio.Prioridad > -1 Then .Priority = objLlamadaServicio.Prioridad

                    'Cliente
                    .CustomerCode = CardCode

                    'PersonaContacto
                    If objLlamadaServicio.PersonaContacto > 0 Then .ContactCode = objLlamadaServicio.PersonaContacto

                    'Telefono
                    If Not String.IsNullOrEmpty(objLlamadaServicio.Telefono) Then .Telephone = objLlamadaServicio.Telefono

                    'NumAtCard
                    If Not String.IsNullOrEmpty(objLlamadaServicio.NumAtCard) Then .CustomerRefNo = objLlamadaServicio.NumAtCard

                    'NumeroSerieFabricante
                    If Not String.IsNullOrEmpty(objLlamadaServicio.NumeroSerieFabricante) Then .ManufacturerSerialNum = objLlamadaServicio.NumeroSerieFabricante

                    'NumeroSerie
                    If Not String.IsNullOrEmpty(objLlamadaServicio.NumeroSerie) Then .InternalSerialNum = objLlamadaServicio.NumeroSerie

                    'ItemCode
                    If Not String.IsNullOrEmpty(objLlamadaServicio.ItemCode) Then .ItemCode = objLlamadaServicio.ItemCode

                    'Asunto
                    If Not String.IsNullOrEmpty(objLlamadaServicio.Asunto) Then .Subject = objLlamadaServicio.Asunto

                    'Origen
                    If objLlamadaServicio.Origen > -4 AndAlso objLlamadaServicio.Origen <> 0 Then .Origin = objLlamadaServicio.Origen

                    'ProblemaTipo
                    If objLlamadaServicio.ProblemaTipo > 0 Then .ProblemType = objLlamadaServicio.ProblemaTipo

                    'ProblemaSubtipo
                    If objLlamadaServicio.ProblemaSubtipo > 0 Then .ProblemSubType = objLlamadaServicio.ProblemaSubtipo

                    'TipoLlamada
                    If objLlamadaServicio.LlamadaTipo > 0 Then .CallType = objLlamadaServicio.LlamadaTipo

                    'Tecnico
                    If objLlamadaServicio.Tecnico > 0 Then .TechnicianCode = objLlamadaServicio.Tecnico

                    'TratadoPor
                    If objLlamadaServicio.TratadoPor > 0 Then .AssigneeCode = objLlamadaServicio.TratadoPor

                    'Resolucion
                    If Not String.IsNullOrEmpty(objLlamadaServicio.Resolucion) Then .Resolution = objLlamadaServicio.Resolucion

                    'ID DOCUWARE
                    If Not String.IsNullOrEmpty(objLlamadaServicio.DOCIDDW) Then .UserFields.Item("U_SEIIDDW").Value = objLlamadaServicio.DOCIDDW

                    'URL DOCUWARE
                    If Not String.IsNullOrEmpty(objLlamadaServicio.DOCURLDW) Then .UserFields.Item("U_SEIURLDW").Value = objLlamadaServicio.DOCURLDW

                End With

                'Actualizamos llamada de servicio
                If oCall.Update() = 0 Then

                    retVal.CODIGO = Respuesta.Ok
                    retVal.MENSAJE = "Llamada servicio actualizada con éxito"
                    retVal.MENSAJEAUX = objLlamadaServicio.Numero

                Else

                    retVal.CODIGO = Respuesta.Ko
                    retVal.MENSAJE = oCompany.GetLastErrorCode & "#" & oCompany.GetLastErrorDescription
                    retVal.MENSAJEAUX = ""

                End If

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oCall)
        End Try

        Return retVal

    End Function

    Public Function seDocumentoRelacionadoLlamada(ByVal objDocumentoLLamada As EntDocumentoLLamada, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado
        Dim oCompany As Company

        Dim oComun As New clsComun

        Dim sLogInfo As String = "RELACIONAR LLAMADA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar llamada servicio para DOCIDDW: " & objDocumentoLLamada.DOCIDDW & ", ObjType: " & objDocumentoLLamada.ObjType & ", DocNum: " & objDocumentoLLamada.DocNum & ", NumAtCard: " & objDocumentoLLamada.NumAtCard & ", CallID: " & objDocumentoLLamada.CallID)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoLLamada.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoLLamada.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoLLamada.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            oCompany = ConexionSAP.getCompany(objDocumentoLLamada.UserSAP, objDocumentoLLamada.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoLLamada.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoLLamada.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Buscamos DocEntry del documento definitivo
            Dim DocEntry As String = oComun.getDocEntryDocumentoDefinitivo(Tabla, "", objDocumentoLLamada.DOCIDDW, objDocumentoLLamada.DocNum, objDocumentoLLamada.NumAtCard, True, True, Sociedad)
            If String.IsNullOrEmpty(DocEntry) Then Throw New Exception("No encuentro DocEntry para el documento con IDDW: " & objDocumentoLLamada.DOCIDDW & ", DocNum: " & objDocumentoLLamada.DocNum & " y NumAtCard: " & objDocumentoLLamada.NumAtCard)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocEntry: " & DocEntry)

            'Busca el DocNum si no se ha proporcionado
            If String.IsNullOrEmpty(objDocumentoLLamada.DocNum) Then objDocumentoLLamada.DocNum = oComun.getDocNumDeDocEntry(Tabla, DocEntry, Sociedad)

            'Relaciona el documento con la llamada de servicio
            If RelacionarDocumentoConLlamadaServicio(oCompany, objDocumentoLLamada.CallID, DocEntry, objDocumentoLLamada.ObjType) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento relacionado con la llamada"
                retVal.MENSAJEAUX = objDocumentoLLamada.DocNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "No se puede relacionar el documento con la llamada"
                retVal.MENSAJEAUX = ""

                clsLog.Log.Fatal(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

            End If

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

    Private Function getNumeroDeCardCode(ByVal CardCode As String,
                                         ByVal Sociedad As eSociedad) As String

        'Devuelve el número de actividad

        Dim retVal As String = ""

        Try

            'Buscamos por DocNum, DocEntry o NumAtCard
            Dim SQL As String = ""

            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("callID") & " As callID " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OSCL", Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("customer") & "  = N'" & CardCode & "'" & vbCrLf

            SQL &= " ORDER BY " & vbCrLf
            SQL &= " T0." & putQuotes("callID") & " DESC " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function RelacionarDocumentoConLlamadaServicio(ByVal oCompany As Company,
                                                           ByVal CallID As Integer,
                                                           ByVal DocEntry As String,
                                                           ByVal ObjType As String) As Boolean

        'Anexa la entrega a la llamada de servicio
        Dim retVal As Boolean = False

        Try

            Dim oLlamada As ServiceCalls = Nothing
            oLlamada = oCompany.GetBusinessObject(BoObjectTypes.oServiceCalls)

            'Comprueba si existe la llamada
            If oLlamada.GetByKey(CallID) <> True Then
                Throw New Exception("No existe la llamada " & CallID)
            Else

                'Comprueba si tiene que añadir un nuevo documento
                If oLlamada.Expenses.DocEntry <> 0 Then oLlamada.Expenses.Add()

                'Establece el DocEntry y el tipo de documento
                oLlamada.Expenses.DocEntry = DocEntry
                oLlamada.Expenses.DocumentType = ObjType

                'Actualizamos la llamada
                If oLlamada.Update() <> 0 Then
                    Throw New Exception("No se ha podido actualizar la llamada " & CallID & ": " & oCompany.GetLastErrorDescription())
                Else
                    retVal = True
                End If

                Dim oComun As New clsComun
                oComun.LiberarObjCOM(CType(oLlamada, Object))

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

#End Region

End Class
