Option Strict On

Imports Newtonsoft.Json
Imports System.Reflection
Imports System.Web.Services
Imports System.ComponentModel
Imports System.Globalization
Imports WSDocuware.Utilidades


' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="https://www.docuware.es/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class WSDocuware
    Inherits System.Web.Services.WebService

    Public Sub New()

    End Sub

#Region "Conexión"

    <WebMethod()>
    Public Function TestSAP(ByVal NIFEmpresa As String, ByVal UserSAP As String, ByVal PassSAP As String) As String

        Dim retVal As New EntResultado

        Try

            Dim Conectado As Boolean = True

            clsLog.Log.Info("Test SAP - NIF " & NIFEmpresa)

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(NIFEmpresa)

            Conectado = ConexionSAP.TestCompany(UserSAP, PassSAP, Sociedad)

            If Conectado Then
                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Conexión SAP correcta - NIF " & NIFEmpresa
                retVal.MENSAJEAUX = ""
            Else
                Throw New Exception("ERROR conexión SAP - NIF " & NIFEmpresa)
            End If

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        Return JsonConvert.SerializeObject(retVal)

    End Function

    <WebMethod()>
    Public Function TestDB(ByVal NIFEmpresa As String) As String

        Dim retVal As New EntResultado

        Try

            Dim Conectado As Boolean = False

            clsLog.Log.Info("Test SQL/HANA - NIF " & NIFEmpresa)

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(NIFEmpresa)

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            oCon.AbrirConexion()
            oCon.CerrarConexion()
            Conectado = True

            If Conectado Then
                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Conexión SQL/HANA correcta - NIF " & NIFEmpresa
                retVal.MENSAJEAUX = ""
            Else
                Throw New Exception("ERROR conexión SQL/HANA - NIF " & NIFEmpresa)
            End If

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

#End Region

#Region "Docuware"


    <WebMethod()>
    Public Function ActualizarDocuwareDatos(ByVal DocuwareJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(DOCUWARE) DocuwareJSON: " & DocuwareJSON)

            'Obtiene las entidades
            Dim objDocuware As New EntDocuware
            objDocuware = JsonConvert.DeserializeObject(Of EntDocuware)(DocuwareJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocuware.UserDW, objDocuware.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocuware.NIFEmpresa)

            'Comprueba si se ha creado el documento de forma definitiva o si existe el documento en firme 
            Dim oDocumento As New clsDocuware
            retVal = oDocumento.setDocuware(objDocuware, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Documento"

    <WebMethod()>
    Public Function IntegrarDocumentoJSON(ByVal DocumentoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(DOCUMENTO) DocumentoJSON: " & DocumentoJSON)

            'Obtiene las entidades
            Dim objDocumento As New EntDocumentoCab
            objDocumento = JsonConvert.DeserializeObject(Of EntDocumentoCab)(DocumentoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocumento.UserDW, objDocumento.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocumento.NIFEmpresa)

            'Integración
            Dim oDocumento As New clsDocumento

            'Crear documento o actualizar DW
            Select Case objDocumento.Opcion

                Case Opcion.Crear
                    'Crear documento
                    retVal = oDocumento.CrearDocumento(objDocumento, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar campos DW
                    retVal = oDocumento.ActualizarDocumentoDW(objDocumento, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

    <WebMethod()>
    Public Function ActualizarDocumentoJSON(ByVal DocumentoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(DOCUMENTO) DocumentoJSON: " & DocumentoJSON)

            'Obtiene las entidades
            Dim objDocumento As New EntDocumentoCab
            objDocumento = JsonConvert.DeserializeObject(Of EntDocumentoCab)(DocumentoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocumento.UserDW, objDocumento.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocumento.NIFEmpresa)

            'Integración
            Dim oDocumento As New clsDocumento
            retVal = oDocumento.ActualizarDocumento(objDocumento, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

    <WebMethod()>
    Public Function ModificarDocumentoJSON(ByVal DocumentoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(DOCUMENTO) DocumentoJSON: " & DocumentoJSON)

            'Obtiene las entidades
            Dim objDocumento As New EntDocumentoCab
            objDocumento = JsonConvert.DeserializeObject(Of EntDocumentoCab)(DocumentoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocumento.UserDW, objDocumento.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocumento.NIFEmpresa)

            'Integración
            Dim oDocumento As New clsDocumento
            retVal = oDocumento.ModificarDocumento(objDocumento, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

    <WebMethod()>
    Public Function ActualizarBloqueoPago(ByVal DocumentoBloqueoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(BLOQUEO PAGO) DocumentoBloqueoJSON: " & DocumentoBloqueoJSON)

            'Obtiene las entidades
            Dim objDocumentoBloqueo As New EntDocumentoBloqueo
            objDocumentoBloqueo = JsonConvert.DeserializeObject(Of EntDocumentoBloqueo)(DocumentoBloqueoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocumentoBloqueo.UserDW, objDocumentoBloqueo.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocumentoBloqueo.NIFEmpresa)

            'Comprueba si se ha creado el documento de forma definitiva o si existe el documento en firme con el IDDW
            Dim oDocumento As New clsDocumento
            retVal = oDocumento.setBloqueoPago(objDocumentoBloqueo, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

    <WebMethod()>
    Public Function ActualizarDocumentoTratado(ByVal DocumentoTratadoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(TRATADO) DocumentoTratadoJSON: " & DocumentoTratadoJSON)

            'Obtiene las entidades
            Dim objDocumentoTratado As New EntDocumentoTratado
            objDocumentoTratado = JsonConvert.DeserializeObject(Of EntDocumentoTratado)(DocumentoTratadoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocumentoTratado.UserDW, objDocumentoTratado.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocumentoTratado.NIFEmpresa)

            'Comprueba si se ha creado el documento de forma definitiva o si existe el documento en firme con el IDDW
            Dim oDocumento As New clsDocumento
            retVal = oDocumento.setDocumentoTratado(objDocumentoTratado, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

    <WebMethod()>
    Public Function ActualizarDocumentoEstado(ByVal DocumentoEstadoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(ESTADO) DocumentoEstadoJSON: " & DocumentoEstadoJSON)

            'Obtiene las entidades
            Dim objDocumentoEstado As New EntDocumentoEstado
            objDocumentoEstado = JsonConvert.DeserializeObject(Of EntDocumentoEstado)(DocumentoEstadoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocumentoEstado.UserDW, objDocumentoEstado.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocumentoEstado.NIFEmpresa)

            'Comprueba si se ha creado el documento de forma definitiva o si existe el documento en firme 
            Dim oDocumento As New clsDocumento
            retVal = oDocumento.setDocumentoEstado(objDocumentoEstado, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Inventario"

    <WebMethod()>
    Public Function IntegrarInventarioJSON(ByVal InventarioJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(INVENTARIO) InventarioJSON: " & InventarioJSON)

            'Obtiene las entidades
            Dim objInventario As New EntInventarioCab
            objInventario = JsonConvert.DeserializeObject(Of EntInventarioCab)(InventarioJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objInventario.UserDW, objInventario.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objInventario.NIFEmpresa)

            'Integración
            Dim oInventarioo As New clsInventario

            'Crear documento o actualizar DW
            Select Case objInventario.Opcion

                Case Opcion.Crear
                    'Crear inventario
                    retVal = oInventarioo.CrearInventario(objInventario, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar campos DW
                    retVal = oInventarioo.ActualizarInventario(objInventario, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

#End Region

#Region "Asiento"

    <WebMethod()>
    Public Function IntegrarAsientoJSON(ByVal AsientoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(ASIENTO) AsientoJSON: " & AsientoJSON)

            'Obtiene las entidades
            Dim objAsiento As New EntAsientoCab
            objAsiento = JsonConvert.DeserializeObject(Of EntAsientoCab)(AsientoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objAsiento.UserDW, objAsiento.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objAsiento.NIFEmpresa)

            'Integración
            Dim oAsiento As New clsAsiento

            'Crear Asiento o actualizar DW
            Select Case objAsiento.Opcion

                Case Opcion.Crear
                    'Crear Asiento
                    retVal = oAsiento.CrearAsiento(objAsiento, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar campos DW
                    retVal = oAsiento.ActualizarAsiento(objAsiento, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

    <WebMethod()>
    Public Function ActualizarAsientoTratado(ByVal AsientoTratadoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(TRATADO) AsientoTratadoJSON: " & AsientoTratadoJSON)

            'Obtiene las entidades
            Dim objAsientoTratado As New EntAsientoTratado
            objAsientoTratado = JsonConvert.DeserializeObject(Of EntAsientoTratado)(AsientoTratadoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objAsientoTratado.UserDW, objAsientoTratado.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objAsientoTratado.NIFEmpresa)

            'Comprueba si se ha creado el Asiento de forma definitiva o si existe el Asiento en firme con el IDDW
            Dim oAsiento As New clsAsiento
            retVal = oAsiento.setAsientoTratado(objAsientoTratado, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Anexo"

    <WebMethod()>
    Public Function IntegrarAnexoJSON(ByVal AnexoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(ANEXO) AnexoJSON: " & AnexoJSON)

            'Obtiene las entidades
            Dim objAnexo As New EntAnexoSAP
            objAnexo = JsonConvert.DeserializeObject(Of EntAnexoSAP)(AnexoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objAnexo.UserDW, objAnexo.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objAnexo.NIFEmpresa)

            'Integración
            Dim oAnexo As New clsAnexo
            retVal = oAnexo.CrearAnexo(objAnexo, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

#End Region

#Region "Interlocutor"

    <WebMethod()>
    Public Function IntegrarInterlocutorJSON(ByVal InterlocutorJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(INTERLOCUTOR) InterlocutorJSON: " & InterlocutorJSON)

            'Obtiene las entidades
            Dim objInterlocutor As New EntInterlocutorSAP
            objInterlocutor = JsonConvert.DeserializeObject(Of EntInterlocutorSAP)(InterlocutorJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objInterlocutor.UserDW, objInterlocutor.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objInterlocutor.NIFEmpresa)

            'Crear o actualizar interlocutor
            Dim oInterlocutor As New clsInterlocutor

            Select Case objInterlocutor.Opcion

                Case Opcion.Crear
                    'Crear interlocutor
                    retVal = oInterlocutor.NuevoInterlocutor(objInterlocutor, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar interlocutor
                    retVal = oInterlocutor.ActualizarInterlocutor(objInterlocutor, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

#End Region

#Region "Artículo"

    <WebMethod()>
    Public Function IntegrarArticuloJSON(ByVal ArticuloJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(ARTICULO) ArticuloJSON: " & ArticuloJSON)

            'Obtiene las entidades
            Dim objArticulo As New EntArticuloSAP
            objArticulo = JsonConvert.DeserializeObject(Of EntArticuloSAP)(ArticuloJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objArticulo.UserDW, objArticulo.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objArticulo.NIFEmpresa)

            'Crear o actualizar artículo
            Dim oArticulo As New clsArticulo

            Select Case objArticulo.Opcion

                Case Opcion.Crear
                    'Crear artículo
                    retVal = oArticulo.NuevoArticulo(objArticulo, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar artículo
                    retVal = oArticulo.ActualizarArticulo(objArticulo, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

#End Region

#Region "Proyecto"

    <WebMethod()>
    Public Function IntegrarProyectoJSON(ByVal ProyectoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(PROYECTO) ProyectoJSON: " & ProyectoJSON)

            'Obtiene las entidades
            Dim objProyecto As New EntProyectoSAP
            objProyecto = JsonConvert.DeserializeObject(Of EntProyectoSAP)(ProyectoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objProyecto.UserDW, objProyecto.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objProyecto.NIFEmpresa)

            'Crear o actualizar proyecto
            Dim oProyecto As New clsProyecto

            Select Case objProyecto.Opcion

                Case Opcion.Crear
                    'Crear proyecto
                    retVal = oProyecto.CrearProyecto(objProyecto, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar proyecto
                    retVal = oProyecto.ActualizarProyecto(objProyecto, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

#End Region

#Region "Actividad"

    <WebMethod()>
    Public Function IntegrarActividadJSON(ByVal ActividadJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(ACTIVIDAD) ActividadJSON: " & ActividadJSON)

            'Obtiene las entidades
            Dim objActividad As New EntActividadSAP
            objActividad = JsonConvert.DeserializeObject(Of EntActividadSAP)(ActividadJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objActividad.UserDW, objActividad.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objActividad.NIFEmpresa)

            'Crear o actualizar actividad
            Dim oActividad As New clsActividad

            Select Case objActividad.Opcion

                Case Opcion.Crear
                    'Crear actividad
                    retVal = oActividad.NuevaActividad(objActividad, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar actividad
                    retVal = oActividad.ActualizarActividad(objActividad, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

#End Region

#Region "Llamada servicio"

    <WebMethod()>
    Public Function IntegrarLlamadaServicioJSON(ByVal LlamadaServicioJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(LLAMADA SERVICIO) LlamadaServicioJSON: " & LlamadaServicioJSON)

            'Obtiene las entidades
            Dim objLlamadaServicio As New EntLlamadaServicioSAP
            objLlamadaServicio = JsonConvert.DeserializeObject(Of EntLlamadaServicioSAP)(LlamadaServicioJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objLlamadaServicio.UserDW, objLlamadaServicio.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objLlamadaServicio.NIFEmpresa)

            'Crear o actualizar llamada servicio
            Dim oLlamadaServicio As New clsLLamadaServicio

            Select Case objLlamadaServicio.Opcion

                Case Opcion.Crear
                    'Crear llamada
                    retVal = oLlamadaServicio.NuevaLlamadaServicio(objLlamadaServicio, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar llamada
                    retVal = oLlamadaServicio.ActualizarLlamadaServicio(objLlamadaServicio, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

    <WebMethod()>
    Public Function RelacionDocumentoIDLlamada(ByVal DocumentoLlamadaJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(RELACIONAR LLAMADA) DocumentoLlamadaJSON: " & DocumentoLlamadaJSON)

            'Obtiene las entidades
            Dim objDocumentoLLamada As New EntDocumentoLLamada
            objDocumentoLLamada = JsonConvert.DeserializeObject(Of EntDocumentoLLamada)(DocumentoLlamadaJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocumentoLLamada.UserDW, objDocumentoLLamada.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocumentoLLamada.NIFEmpresa)

            'Comprueba si se ha creado el documento de forma definitiva o si existe el documento en firme con el IDDW
            Dim oLlamada As New clsLLamadaServicio
            retVal = oLlamada.seDocumentoRelacionadoLlamada(objDocumentoLLamada, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Acuerdo"

    <WebMethod()>
    Public Function IntegrarAcuerdoJSON(ByVal AcuerdoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(ACUERDO) AcuerdoJSON: " & AcuerdoJSON)

            'Obtiene las entidades
            Dim objAcuerdo As New EntAcuerdoCab
            objAcuerdo = JsonConvert.DeserializeObject(Of EntAcuerdoCab)(AcuerdoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objAcuerdo.UserDW, objAcuerdo.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objAcuerdo.NIFEmpresa)

            'Crear o actualizar acuerdo
            Dim oAcuerdo As New clsAcuerdo

            Select Case objAcuerdo.Opcion

                Case Opcion.Crear
                    'Crear acuerdo
                    retVal = oAcuerdo.NuevoAcuerdo(objAcuerdo, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar acuerdo
                    retVal = oAcuerdo.ActualizarAcuerdo(objAcuerdo, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

    <WebMethod()>
    Public Function ComprobarAcuerdo(ByVal AcuerdoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            'Obtiene las entidades
            Dim objAcuerdo As New EntAcuerdoCab
            objAcuerdo = JsonConvert.DeserializeObject(Of EntAcuerdoCab)(AcuerdoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objAcuerdo.UserDW, objAcuerdo.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objAcuerdo.NIFEmpresa)

            'Comprueba si se ha creado el acuerdo o si existe el acuerdo con el IDDW
            Dim oAcuerdo As New clsAcuerdo
            retVal = oAcuerdo.getComprobarAcuerdo(objAcuerdo, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

    <WebMethod()>
    Public Function ActualizarAcuerdoEstado(ByVal AcuerdoEstadoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(ACUERDO ESTADO) AcuerdoEstadoJSON: " & AcuerdoEstadoJSON)

            'Obtiene las entidades
            Dim objAcuerdoEstado As New EntAcuerdoEstado
            objAcuerdoEstado = JsonConvert.DeserializeObject(Of EntAcuerdoEstado)(AcuerdoEstadoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objAcuerdoEstado.UserDW, objAcuerdoEstado.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objAcuerdoEstado.NIFEmpresa)

            'Comprueba si se ha creado el Acuerdo de forma definitiva o si existe el Acuerdo en firme 
            Dim oAcuerdo As New clsAcuerdo
            retVal = oAcuerdo.setAcuerdoEstado(objAcuerdoEstado, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Oportunidad"

    <WebMethod()>
    Public Function IntegrarOportunidadJSON(ByVal OportunidadJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(OPORTUNIDAD) OportunidadJSON: " & OportunidadJSON)

            'Obtiene las entidades
            Dim objOportunidad As New EntOportunidadCab
            objOportunidad = JsonConvert.DeserializeObject(Of EntOportunidadCab)(OportunidadJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objOportunidad.UserDW, objOportunidad.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objOportunidad.NIFEmpresa)

            'Crear o actualizar Oportunidad
            Dim oOportunidad As New clsOportunidad

            Select Case objOportunidad.Opcion

                Case Opcion.Crear
                    'Crear oportunidad
                    retVal = oOportunidad.NuevaOportunidad(objOportunidad, Sociedad)

                Case Opcion.Actualizar
                    'Actualizar oportunidad
                    retVal = oOportunidad.ActualizarOportunidad(objOportunidad, Sociedad)

                Case Else
                    Throw New Exception("Opción no definida (crear/actualizar)")

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta JSON
        Return JsonConvert.SerializeObject(retVal)

    End Function

    <WebMethod()>
    Public Function ComprobarOportunidad(ByVal OportunidadJSON As String) As String

        Dim retVal As New EntResultado

        Try

            'Obtiene las entidades
            Dim objOportunidad As New EntOportunidadCab
            objOportunidad = JsonConvert.DeserializeObject(Of EntOportunidadCab)(OportunidadJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objOportunidad.UserDW, objOportunidad.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objOportunidad.NIFEmpresa)

            'Comprueba si se ha creado el Oportunidad o si existe el Oportunidad con el IDDW
            Dim oOportunidad As New clsOportunidad
            retVal = oOportunidad.getComprobarOportunidad(objOportunidad, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Lote"

    <WebMethod()>
    Public Function ActualizarLoteJSON(ByVal LoteJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(LOTE) LoteJSON: " & LoteJSON)

            'Obtiene las entidades
            Dim objLote As New EntLoteSAP
            objLote = JsonConvert.DeserializeObject(Of EntLoteSAP)(LoteJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objLote.UserDW, objLote.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objLote.NIFEmpresa)

            'Actualizar lote
            Dim oLote As New clsLote

            'Actualizar lote
            retVal = oLote.ActualizarLote(objLote, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Informe"

    <WebMethod()>
    Public Function GenerarInfome(ByVal InformeJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(GENERAR INFORME) InformeJSON: " & InformeJSON)

            'Obtiene las entidades
            Dim objInforme As New EntInforme
            objInforme = JsonConvert.DeserializeObject(Of EntInforme)(InformeJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objInforme.UserDW, objInforme.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objInforme.NIFEmpresa)

            'Comprueba si existe el fichero
            Dim oDocumento As New clsDocumento
            retVal = oDocumento.getGenerarInforme(objInforme, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Fichero"

    <WebMethod()>
    Public Function GuardarFichero(ByVal FicheroJSON As String) As String

        Dim retVal As New EntResultado

        Try

            'clsLog.Log.Info("(GUARDAR FICHERO) FicheroJSON: " & FicheroJSON)

            'Obtiene las entidades
            Dim objFichero As New EntFichero
            objFichero = JsonConvert.DeserializeObject(Of EntFichero)(FicheroJSON)

            clsLog.Log.Info("(GUARDAR FICHERO) FicheroNombre: " & objFichero.Nombre)

            'Credenciales fijas
            If Not ComprobarCredenciales(objFichero.UserDW, objFichero.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objFichero.NIFEmpresa)

            'Dim FicheroBinario As Byte() = IO.File.ReadAllBytes("C:\Users\mlanza\OneDrive - SEIDOR SA\Escritorio\Adjuntos\1_AEAT.pdf")
            'objFichero.Nombre = "1_AEAT.pdf"
            'objFichero.Base64 = Convert.ToBase64String(FicheroBinario)

            'FicheroJSON = JsonConvert.SerializeObject(objFichero)

            'Comprueba si existe el fichero
            Dim oFichero As New clsFichero
            retVal = oFichero.getGuardarFichero(objFichero, Sociedad)

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        'Solo guarda el binario si no ha podido crear el fichero
        If retVal.CODIGO = Respuesta.Ko Then clsLog.Log.Info("(GUARDAR FICHERO) FicheroJSON: " & FicheroJSON)

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

#End Region

#Region "Obtener datos"

    <WebMethod()>
    Public Function ObtenerDocumentoCampo(ByVal DocumentoCampoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(DOCUMENTO CAMPO) DocumentoCampoJSON: " & DocumentoCampoJSON)

            'Obtiene las entidades
            Dim objDocumentoCampo As New EntDocumentoCampo
            objDocumentoCampo = JsonConvert.DeserializeObject(Of EntDocumentoCampo)(DocumentoCampoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDocumentoCampo.UserDW, objDocumentoCampo.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDocumentoCampo.NIFEmpresa)

            'Obtiene el campo a consultar del documento definitivo
            Select Case objDocumentoCampo.CampoTipo
                Case Campos.NumDocumento
                    'DocNum 
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarNumDocumento(objDocumentoCampo, Sociedad)

                Case Campos.Importe
                    'Importe
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarImporte(objDocumentoCampo, Sociedad)

                Case Campos.FechaVencimiento
                    'Fecha vencimiento
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarFechaVencimiento(objDocumentoCampo, Sociedad)

                Case Campos.ResponsableMail
                    'Responsable
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarResponsableMail(objDocumentoCampo, Sociedad)

                Case Campos.Comentarios
                    'Comentarios
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarComentarios(objDocumentoCampo, Sociedad)

                Case Campos.FechaImporteVencimiento
                    'Fecha vencimiento
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarVencimientoFechaImporte(objDocumentoCampo, Sociedad)

                Case Campos.PagadoVencimiento
                    'Pagado vencimiento
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarVencimientoPagado(objDocumentoCampo, Sociedad)

                Case Campos.ImporteSinIVA
                    'Importe sin IVA
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarImporteSinIVA(objDocumentoCampo, Sociedad)

                Case Campos.ViaPago
                    'Via de pago
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarViaPago(objDocumentoCampo, Sociedad)

                Case Campos.Moneda
                    'Via de pago
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarMoneda(objDocumentoCampo, Sociedad)

                Case Campos.NumTransaccion
                    'Número de transacción 
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarNumTransaccion(objDocumentoCampo, Sociedad)

                Case Campos.Proyecto
                    'Proyecto
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarProyecto(objDocumentoCampo, Sociedad)

                Case Campos.CampoUsuario
                    'Campo de usuario
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarCampoUsuario(objDocumentoCampo, Sociedad)

                Case Campos.Sucursal
                    'Sucursal
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarSucursal(objDocumentoCampo, Sociedad)

                Case Campos.CentroCoste
                    'Centro de coste
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarCentroCoste(objDocumentoCampo, Sociedad)

                Case Campos.NumEnvio
                    'Número de envío
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarNumEnvio(objDocumentoCampo, Sociedad)

                Case Campos.NumDestino
                    'Número de documento destino
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarNumDestino(objDocumentoCampo, Sociedad)

                Case Campos.Titular
                    'Titular
                    Dim oDocumento As New clsDocumento
                    retVal = oDocumento.getComprobarTitular(objDocumentoCampo, Sociedad)

                Case Else
                    retVal.CODIGO = Respuesta.Ko
                    retVal.MENSAJE = "Tipo de campo no permitido"
                    retVal.MENSAJEAUX = ""

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

    <WebMethod()>
    Public Function ObtenerAsientoCampo(ByVal AsientoCampoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(ASIENTO CAMPO) AsientoCampoJSON: " & AsientoCampoJSON)

            'Obtiene las entidades
            Dim objAsientoCampo As New EntAsientoCampo
            objAsientoCampo = JsonConvert.DeserializeObject(Of EntAsientoCampo)(AsientoCampoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objAsientoCampo.UserDW, objAsientoCampo.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objAsientoCampo.NIFEmpresa)

            'Obtiene el campo a consultar del Asiento definitivo
            Select Case objAsientoCampo.CampoTipo
                Case Campos.NumDocumento
                    'DocNum 
                    Dim oAsiento As New clsAsiento
                    retVal = oAsiento.getComprobarNumAsiento(objAsientoCampo, Sociedad)

                Case Campos.Importe
                    'Importe
                    Dim oAsiento As New clsAsiento
                    retVal = oAsiento.getComprobarImporte(objAsientoCampo, Sociedad)

                Case Campos.FechaVencimiento
                    'Fecha vencimiento
                    Dim oAsiento As New clsAsiento
                    retVal = oAsiento.getComprobarFechaVencimiento(objAsientoCampo, Sociedad)

                Case Campos.Comentarios
                    'Comentarios
                    Dim oAsiento As New clsAsiento
                    retVal = oAsiento.getComprobarComentarios(objAsientoCampo, Sociedad)

                Case Campos.NumTransaccion
                    'Número de transacción 
                    Dim oAsiento As New clsAsiento
                    retVal = oAsiento.getComprobarNumTransaccion(objAsientoCampo, Sociedad)

                Case Campos.Proyecto
                    'Proyecto
                    Dim oAsiento As New clsAsiento
                    retVal = oAsiento.getComprobarProyecto(objAsientoCampo, Sociedad)

                Case Campos.CampoUsuario
                    'Campo de usuario
                    Dim oAsiento As New clsAsiento
                    retVal = oAsiento.getComprobarCampoUsuario(objAsientoCampo, Sociedad)

                Case Else
                    retVal.CODIGO = Respuesta.Ko
                    retVal.MENSAJE = "Tipo de campo no permitido"
                    retVal.MENSAJEAUX = ""

            End Select

        Catch ex As Exception
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

        'Respuesta string
        Return MensajeSalida(JsonConvert.SerializeObject(retVal))

    End Function

    <WebMethod()>
    Public Function ObtenerDatos(ByVal DatoJSON As String) As String

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(DATOS) DatoJSON: " & DatoJSON)

            'Obtiene las entidades
            Dim objDato As New EntDato
            objDato = JsonConvert.DeserializeObject(Of EntDato)(DatoJSON)

            'Credenciales fijas
            If Not ComprobarCredenciales(objDato.UserDW, objDato.PassDW) Then Throw New Exception("Credenciales DW incorrectas")

            'Sociedad por NIF suministrado
            Dim Sociedad As eSociedad
            Sociedad = SOCIEDADPORNIF(objDato.NIFEmpresa)

            'Obtiene el listado
            Select Case objDato.PeticionTipo

                Case Peticion.Empleados
                    Dim oDAOEmpleado As New DAOEmpleado(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOEmpleado.getEmpleados())

                Case Peticion.Interlocutores

                    If String.IsNullOrEmpty(objDato.InterlocutorTipo) Then Throw New Exception("Tipo interlocutor no suministrado")

                    Dim oDAOInterlocutor As New DAOInterlocutor(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOInterlocutor.getInterlocutores(objDato.InterlocutorTipo))

                Case Peticion.Articulos
                    Dim oDAOArticulo As New DAOArticulo(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOArticulo.getArticulos())

                Case Peticion.CentrosCoste
                    Dim oDAOCentroCoste As New DAOCentroCoste(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOCentroCoste.getCentrosCoste())

                Case Peticion.Proyectos
                    Dim oDAOProyecto As New DAOProyecto(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOProyecto.getProyectos())

                Case Peticion.DocumentosAbiertos

                    If objDato.ObjType <= 0 Then Throw New Exception("ObtType no suministrado")
                    If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha inicio no suministrada")
                    If Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha fin no suministrada")

                    Dim oDAODocumento As New DAODocumento(Sociedad)
                    Return JsonConvert.SerializeObject(oDAODocumento.getDocumentos(objDato.ObjType, objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.ICIQ_SolicitudesPedidosCompraIP
                    Return "-1 Tipo de petición no permitida en " & MethodBase.GetCurrentMethod().Name

                    'If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                    '    Throw New Exception("Fecha inicio no suministrada")
                    'ElseIf Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                    '    Throw New Exception("Fecha fin no suministrada")
                    'End If

                    'Dim oDAOCarta As New DAOCarta(Sociedad)
                    'Return JsonConvert.SerializeObject(oDAOCarta.getSolicitudesPedidosCompraIP(objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.ICIQ_RelacionesDocumentosCompra
                    Return "-1 Tipo de petición no permitida en " & MethodBase.GetCurrentMethod().Name

                    'If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                    '    Throw New Exception("Fecha inicio no suministrada")
                    'ElseIf Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                    '    Throw New Exception("Fecha fin no suministrada")
                    'End If

                    'Dim oDAOCarta As New DAOCarta(Sociedad)
                    'Return JsonConvert.SerializeObject(oDAOCarta.getRelacionesDocumentosCompra(objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.Sucursales
                    Dim oDAOSucursal As New DAOSucursal(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOSucursal.getSucursales())

                Case Peticion.OrigenDocumentos

                    If objDato.ObjType <= 0 Then Throw New Exception("ObtType no suministrado")
                    If String.IsNullOrEmpty(objDato.NIFTercero) AndAlso String.IsNullOrEmpty(objDato.RazonSocial) Then Throw New Exception("NIF o razón social no suministrado")
                    If String.IsNullOrEmpty(objDato.DOCIDDW) AndAlso String.IsNullOrEmpty(objDato.DocNum) AndAlso String.IsNullOrEmpty(objDato.NumAtCard) Then _
                        Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

                    Dim oDocumento As New clsDocumento
                    Return JsonConvert.SerializeObject(oDocumento.getComprobarDocumentoOrigen(objDato, Sociedad))

                Case Peticion.Series
                    Dim oDAOSerie As New DAOSerie(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOSerie.getSeries(objDato.ObjType))

                Case Peticion.GruposArticulo
                    Dim oDAOGrupoArticulo As New DAOGrupoArticulo(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOGrupoArticulo.getGruposArticulo())

                Case Peticion.GruposUnidadesMedida
                    Dim oDAOGrupoUnidadMedida As New DAOGrupoUnidadMedida(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOGrupoUnidadMedida.getGruposUnidadMedida())

                Case Peticion.Fabricantes
                    Dim oDAOFabricante As New DAOFabricante(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOFabricante.getFabricantes())

                Case Peticion.ClasesExpedicion
                    Dim oDAOClaseExpedicion As New DAOClaseExpedicion(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOClaseExpedicion.getClasesExpedicion())

                Case Peticion.GruposImpositivos
                    Dim oDAOGrupoImpositivo As New DAOGrupoImpositivo(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOGrupoImpositivo.getGruposImpositivo())

                Case Peticion.ValoresValidos

                    Dim TablaID As String = getTablaDeObjType(objDato.ObjType)

                    Dim oDAOValorValido As New DAOValorValido(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOValorValido.getValoresValidos(TablaID, ""))

                Case Peticion.TarjetasBanco
                    Dim oDAOTarjeta As New DAOTarjetaBanco(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOTarjeta.getTarjetasBanco())

                Case Peticion.ViasPago
                    Dim oDAOViaPago As New DAOViaPago(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOViaPago.getViasPago())

                Case Peticion.Monedas
                    Dim oDAOMoneda As New DAOMoneda(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOMoneda.getMonedas())

                Case Peticion.TasasCambio

                    If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha inicio no suministrada")
                    If Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha fin no suministrada")

                    Dim oDAOTasaCambio As New DAOTasaCambio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOTasaCambio.getTasasCambio(objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.ActividadesTipo
                    Dim oDAOActividad As New DAOActividad(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActividad.getActividadesTipo())

                Case Peticion.ActividadesAsunto
                    Dim oDAOActividad As New DAOActividad(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActividad.getActividadesAsunto())

                Case Peticion.ActividadesEmplazamiento
                    Dim oDAOActividad As New DAOActividad(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActividad.getActividadesEmplazamiento())

                Case Peticion.LlamadasServicioEstado
                    Dim oDAOLlamadaServicio As New DAOLlamadaServicio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOLlamadaServicio.getLlamadasServicioEstado())

                Case Peticion.LlamadasServicioTipo
                    Dim oDAOLlamadaServicio As New DAOLlamadaServicio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOLlamadaServicio.getLlamadasServicioTipo())

                Case Peticion.LlamadasServicioOrigen
                    Dim oDAOLlamadaServicio As New DAOLlamadaServicio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOLlamadaServicio.getLlamadasServicioOrigen())

                Case Peticion.LlamadasServicioProblemaTipo
                    Dim oDAOLlamadaServicio As New DAOLlamadaServicio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOLlamadaServicio.getLlamadasServicioProblemaTipo())

                Case Peticion.LlamadasServicioProblemaSubtipo
                    Dim oDAOLlamadaServicio As New DAOLlamadaServicio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOLlamadaServicio.getLlamadasServicioProblemaSubtipo())

                Case Peticion.Almacenes
                    Dim oDAOAlmacen As New DAOAlmacen(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOAlmacen.getAlmacenes())

                Case Peticion.GruposInterlocutor
                    Dim oDAOGrupoInterlocutor As New DAOGrupoInterlocutor(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOGrupoInterlocutor.getGruposInterlocutor())

                Case Peticion.AURIA_CentrosCosteRIC
                    Dim oDAORic As New DAORic(Sociedad)
                    Return JsonConvert.SerializeObject(oDAORic.getCentrosCosteRIC())

                Case Peticion.CondicionesPago
                    Dim oDAOCondicionPago As New DAOCondicionPago(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOCondicionPago.getCondicionesPago())

                Case Peticion.Responsables
                    Dim oDAOResponsable As New DAOResponsable(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOResponsable.getResponsables())

                Case Peticion.EmpleadosDptoCV
                    Dim oDAOEmpleadoDptoCV As New DAOEmpleadoDptoCV(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOEmpleadoDptoCV.getEmpleadosDptoCVs())

                Case Peticion.BancosPropios
                    Dim oDAOBancoPropio As New DAOBancoPropio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOBancoPropio.getBancosPropios())

                Case Peticion.InterlocutoresBancos
                    Dim oDAOInterlocutorBanco As New DAOInterlocutorBanco(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOInterlocutorBanco.getInterlocutoresBancos(objDato.InterlocutorTipo))

                Case Peticion.Portes
                    Dim oDAOPorte As New DAOPorte(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOPorte.getPortes())

                Case Peticion.GrupoLaPuente_PedidosCompraLineas
                    Return "-1 Tipo de petición no permitida en " & MethodBase.GetCurrentMethod().Name

                Case Peticion.InterlocutoresDirecciones

                    Dim oDAOInterlocutorDireccion As New DAOInterlocutorDireccion(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOInterlocutorDireccion.getInterlocutoresDirecciones(objDato.InterlocutorTipo))

                Case Peticion.InterlocutoresContactos

                    Dim oDAOInterlocutorContacto As New DAOInterlocutorContacto(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOInterlocutorContacto.getInterlocutoresContactos(objDato.InterlocutorTipo))

                Case Peticion.FacturasAnticipo

                    Dim oDAOFacturaAnticipo As New DAOFacturaAnticipo(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOFacturaAnticipo.getFacturasAnticipo(objDato.InterlocutorTipo))

                Case Peticion.ValoresTablas

                    If String.IsNullOrEmpty(objDato.Tabla) Then Throw New Exception("Tabla no suministrada")

                    Dim oDAOValorTabla As New DAOValorTabla(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOValorTabla.getValoresTablas(objDato.Tabla))

                Case Peticion.DocumentosExtendida

                    If objDato.ObjType <= 0 Then Throw New Exception("ObtType no suministrado")
                    If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha inicio no suministrada")
                    If Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha fin no suministrada")

                    Dim oDAODocumento As New DAODocumento(Sociedad)
                    Return JsonConvert.SerializeObject(oDAODocumento.getDocumentosExtendido(objDato.ObjType, objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.TarjetasEquipos
                    Dim oDAOTarjetaEquipo As New DAOTarjetaEquipo(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOTarjetaEquipo.getTarjetasEquipo())

                Case Peticion.AsientosModelos
                    Dim oDAOAsiento As New DAOAsiento(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOAsiento.getAsientosModelo())

                Case Peticion.AsientosIndicadores
                    Dim oDAOAsiento As New DAOAsiento(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOAsiento.getAsientosIndicador())

                Case Peticion.AsientosOperaciones
                    Dim oDAOAsiento As New DAOAsiento(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOAsiento.getAsientosOperacion())

                Case Peticion.Idiomas
                    Dim oDAOIdioma As New DAOIdioma(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOIdioma.getIdiomas())

                Case Peticion.CuentasContables
                    Dim oDAOCuentaContable As New DAOCuentaContable(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOCuentaContable.getCuentasContables())

                Case Peticion.AcuerdosGlobales
                    If objDato.ObjType <= 0 Then Throw New Exception("ObtType no suministrado")
                    If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha inicio no suministrada")
                    If Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha fin no suministrada")

                    Dim oDAOAcuerdo As New DAOAcuerdo(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOAcuerdo.getAcuerdos(objDato.ObjType, objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.NormasReparto
                    Dim oDAONormaReparto As New DAONormaReparto(Sociedad)
                    Return JsonConvert.SerializeObject(oDAONormaReparto.getNormasReparto())

                Case Peticion.ActivosFijosClases
                    Dim oDAOActivoFijoClase As New DAOActivoFijoClase(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActivoFijoClase.getActivosFijosClases())

                Case Peticion.ActivosFijosGrupos
                    Dim oDAOActivoFijoGrupo As New DAOActivoFijoGrupo(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActivoFijoGrupo.getActivosFijosGrupos())

                Case Peticion.ActivosFijosGruposAmortizacion
                    Dim oDAOActivoFijoGrupoAmortizacion As New DAOActivoFijoGrupoAmortizacion(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActivoFijoGrupoAmortizacion.getActivosFijosGruposAmortizacion())

                Case Peticion.ActivosFijosEmplazamientos
                    Dim oDAOActivoFijoEmplazamiento As New DAOActivoFijoEmplazamiento(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActivoFijoEmplazamiento.getActivosFijosEmplazamientos())

                Case Peticion.ActivosFijosAreasValoracion
                    Dim oDAOActivoFijoAreaValoracion As New DAOActivoFijoAreaValoracion(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActivoFijoAreaValoracion.getActivosFijosAreasValoracion())

                Case Peticion.LlamadasServicio
                    If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                        Throw New Exception("Fecha inicio no suministrada")
                    ElseIf Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                        Throw New Exception("Fecha fin no suministrada")
                    End If

                    Dim oDAOLlamadaServicio As New DAOLlamadaServicio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOLlamadaServicio.getLlamadasServicio(objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.Actividades
                    If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                        Throw New Exception("Fecha inicio no suministrada")
                    ElseIf Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                        Throw New Exception("Fecha fin no suministrada")
                    End If

                    Dim oDAOActividad As New DAOActividad(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOActividad.getActividades(objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.LlamadasServicioCola
                    Dim oDAOLlamadaServicio As New DAOLlamadaServicio(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOLlamadaServicio.getLlamadasServicioCola())

                Case Peticion.PreciosEntregaAbiertos

                    If Not DateTime.TryParseExact(objDato.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha inicio no suministrada")
                    If Not DateTime.TryParseExact(objDato.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha fin no suministrada")

                    Dim oDAOPrecioEntrega As New DAOPrecioEntrega(Sociedad)
                    Return JsonConvert.SerializeObject(oDAOPrecioEntrega.getPreciosEntregaAbiertos(objDato.FechaInicio, objDato.FechaFin))

                Case Peticion.Ceamsa_FacturacionConceptoEstandar
                    Return "-1 Tipo de petición no permitida en " & MethodBase.GetCurrentMethod().Name

                Case Peticion.Ceamsa_EmpleadosExtendido
                    Return "-1 Tipo de petición no permitida en " & MethodBase.GetCurrentMethod().Name

                Case Peticion.VHIO_Documentos
                    Return "-1 Tipo de petición no permitida en " & MethodBase.GetCurrentMethod().Name

                Case Peticion.ICIQ_DocumentosProyectos
                    Return "-1 Tipo de petición no permitida en " & MethodBase.GetCurrentMethod().Name

                Case Else
                    Return "-1 Tipo de petición no permitida en " & MethodBase.GetCurrentMethod().Name

            End Select

        Catch ex As Exception
            Return "-1 " & ex.Message & " en " & MethodBase.GetCurrentMethod().Name
        End Try

    End Function

#End Region

End Class