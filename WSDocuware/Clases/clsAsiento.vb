Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsAsiento

#Region "Públicas"

    Public Function CrearAsiento(ByVal objAsiento As EntAsientoCab, ByVal Sociedad As eSociedad) As EntResultado

        'Crea un asiento
        Dim retVal As New EntResultado

        Try

            retVal = NuevoAsiento(objAsiento, Sociedad)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function ActualizarAsiento(ByVal objAsiento As EntAsientoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR ASIENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de asiento: " & objAsiento.RefOrigen & " para IDDW " & objAsiento.DOCIDDW)

            'Obligatorios
            If Not DateTime.TryParseExact(objAsiento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha asiento no suministrada")

            If String.IsNullOrEmpty(objAsiento.RefOrigen) OrElse String.IsNullOrEmpty(objAsiento.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objAsiento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objAsiento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAsiento.ObjTypeDestino)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsiento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objAsiento.RefOrigen.Split("#")

            'Recorre cada uno de los asientos origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el asiento origen y que se puede acceder a él
                'Pueden ser varios asientos origen con la misma referencia
                Dim TransIds As List(Of String) = getTransIdDeRefOrigen(Tabla, RefOrigen, objAsiento.TipoRefOrigen, Sociedad)
                If TransIds Is Nothing OrElse TransIds.Count = 0 Then Throw New Exception("No existe asiento con referencia: " & RefOrigen)

                'Se actualizan los campos de usuario de DW
                setDatosDWAsiento(Tabla, objAsiento.DocDate, TransIds, objAsiento.DOCIDDW, objAsiento.DOCURLDW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Asiento actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", TransIds)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarAsientoID(ByVal IDDW As String,
                                          ByVal ObjType As Integer,
                                          ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un asiento en firme

        Dim retVal As New EntResultado

        Dim sLogInfo As String = "ASIENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar asiento para IDDW: " & IDDW)

            'Obligatorios
            If String.IsNullOrEmpty(IDDW) Then Throw New Exception("ID docuware no suministrado")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el TransId del asiento definitivo
            Dim TransNum As String = getTransNumAsientoPorDWID(Tabla, IDDW, Sociedad)

            If Not String.IsNullOrEmpty(TransNum) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Asiento en firme encontrado"
                retVal.MENSAJEAUX = TransNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "No se encuentra el asiento en firme"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function setAsientoTratado(ByVal objAsientoTratado As EntAsientoTratado, ByVal Sociedad As eSociedad) As EntResultado

        'Actualiza el campo tratado DW

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "TRATADO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de asiento tratado para ObjType: " & objAsientoTratado.ObjType & ", DOCIDDW: " & objAsientoTratado.DOCIDDW & ", TransNum: " & objAsientoTratado.TransNum & ", Ref1: " & objAsientoTratado.Ref1 & ", Ref2: " & objAsientoTratado.Ref2 & ", Ref3: " & objAsientoTratado.Ref3)

            'Obligatorios
            If String.IsNullOrEmpty(objAsientoTratado.DOCIDDW) AndAlso String.IsNullOrEmpty(objAsientoTratado.TransNum) _
                AndAlso String.IsNullOrEmpty(objAsientoTratado.Ref1) AndAlso String.IsNullOrEmpty(objAsientoTratado.Ref2) _
                AndAlso String.IsNullOrEmpty(objAsientoTratado.Ref3) Then _
                Throw New Exception("ID Docuware, número de asiento o referencias origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAsientoTratado.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsientoTratado.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Buscamos TransId por DOCIDDW, TransId o NumAtCard
            Dim TransId As String = oComun.getTransIdAsientoDefinitivo(Tabla, objAsientoTratado.DOCIDDW, objAsientoTratado.TransNum, objAsientoTratado.Ref1, objAsientoTratado.Ref2, objAsientoTratado.Ref3, Sociedad)
            If String.IsNullOrEmpty(TransId) Then Throw New Exception("No encuentro asiento con DOCIDDW - '" & objAsientoTratado.DOCIDDW & "', TransId - '" & objAsientoTratado.TransNum & "', Ref1 - '" & objAsientoTratado.Ref1 & "', Ref2 - '" & objAsientoTratado.Ref2 & "' y Ref3 - '" & objAsientoTratado.Ref3 & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado TransId: " & TransId)

            'Buscamos TransNum por TransId
            Dim sTransNum As String = oComun.getTransNumDeTransId(Tabla, TransId, Sociedad)
            If String.IsNullOrEmpty(sTransNum) Then Throw New Exception("No encuentro asiento con TransId: '" & TransId & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado TransNum: " & sTransNum)

            'Actualiza el campo tratado DW 
            Dim Actualizado As Boolean = setAsientoDWTratado(Tabla, TransId, Sociedad)

            If Actualizado Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Asiento actualizado con éxito"
                retVal.MENSAJEAUX = sTransNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "Asiento no actualizado"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

#Region "Públicas: Obtener campo"

    Public Function getComprobarNumAsiento(ByVal objAsientoCampo As EntAsientoCampo,
                                             ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un asiento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUMERO ASIENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar asiento para ObjType: " & objAsientoCampo.ObjType & ", DOCIDDW: " & objAsientoCampo.DOCIDDW & ", TransNum: " & objAsientoCampo.TransNum & ", Ref1: " & objAsientoCampo.Ref1 & ", Ref2: " & objAsientoCampo.Ref2 & ", Ref3: " & objAsientoCampo.Ref3)

            'Obligatorios
            If String.IsNullOrEmpty(objAsientoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAsientoCampo.TransNum) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref1) AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref2) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref3) Then _
               Throw New Exception("ID Docuware, número de asiento o referencia origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAsientoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsientoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el TransId del asiento definitivo
            Dim sTransNum As String = getTransNumAsientoDefinitivo(Tabla, objAsientoCampo.DOCIDDW, objAsientoCampo.TransNum, objAsientoCampo.Ref1, objAsientoCampo.Ref2, objAsientoCampo.Ref3, Sociedad)

            If Not String.IsNullOrEmpty(sTransNum) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Asiento en firme encontrado"
                retVal.MENSAJEAUX = sTransNum

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el asiento en firme"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarImporte(ByVal objAsientoCampo As EntAsientoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba el importe de un pedido

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "IMPORTE"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar importe para ObjType: " & objAsientoCampo.ObjType & ", DOCIDDW: " & objAsientoCampo.DOCIDDW & ", TransNum: " & objAsientoCampo.TransNum & ", Ref1: " & objAsientoCampo.Ref1 & ", Ref2: " & objAsientoCampo.Ref2 & ", Ref3: " & objAsientoCampo.Ref3)

            'Obligatorios
            If String.IsNullOrEmpty(objAsientoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAsientoCampo.TransNum) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref1) AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref2) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref3) Then _
                Throw New Exception("ID Docuware, número de asiento o referencia origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAsientoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsientoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocTotal del asiento definitivo
            Dim DocTotal As String = getTransTotalAsientoDefinitivo(Tabla, objAsientoCampo.DOCIDDW, objAsientoCampo.TransNum, objAsientoCampo.Ref1, objAsientoCampo.Ref2, objAsientoCampo.Ref3, Sociedad)

            If Not String.IsNullOrEmpty(DocTotal) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Importe encontrado de este asiento"
                retVal.MENSAJEAUX = DocTotal

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el importe de este asiento"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarFechaVencimiento(ByVal objAsientoCampo As EntAsientoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un Asiento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "VENCIMIENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar fecha vencimiento para ObjType: " & objAsientoCampo.ObjType & ", DOCIDDW: " & objAsientoCampo.DOCIDDW & ", TransNum: " & objAsientoCampo.TransNum & ", Ref1: " & objAsientoCampo.Ref1 & ", Ref2: " & objAsientoCampo.Ref2 & ", Ref3: " & objAsientoCampo.Ref3)

            'Obligatorios
            If String.IsNullOrEmpty(objAsientoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAsientoCampo.TransNum) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref1) AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref2) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref3) Then _
                Throw New Exception("ID Docuware, número de asiento o referencia origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAsientoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsientoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el TransId del asiento definitivo
            Dim DocDueDate As String = getTransDueDateAsientoDefinitivo(Tabla, objAsientoCampo.DOCIDDW, objAsientoCampo.TransNum, objAsientoCampo.Ref1, objAsientoCampo.Ref2, objAsientoCampo.Ref3, Sociedad)

            If Not String.IsNullOrEmpty(DocDueDate) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Fecha de vencimiento de este asiento encontrada"
                retVal.MENSAJEAUX = DocDueDate

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra la fecha de vencimiento de este asiento"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarComentarios(ByVal objAsientoCampo As EntAsientoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un Asiento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "COMENTARIOS"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar comentarios para ObjType: " & objAsientoCampo.ObjType & ", DOCIDDW: " & objAsientoCampo.DOCIDDW & ", TransNum: " & objAsientoCampo.TransNum & ", Ref1: " & objAsientoCampo.Ref1 & ", Ref2: " & objAsientoCampo.Ref2 & ", Ref3: " & objAsientoCampo.Ref3)

            'Obligatorios
            If String.IsNullOrEmpty(objAsientoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAsientoCampo.TransNum) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref1) AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref2) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref3) Then _
                Throw New Exception("ID Docuware, número de asiento o referencia origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAsientoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsientoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el TransId del asiento definitivo
            Dim Comentarios As String = getComentariosAsientoDefinitivo(Tabla, objAsientoCampo.DOCIDDW, objAsientoCampo.TransNum, objAsientoCampo.Ref1, objAsientoCampo.Ref2, objAsientoCampo.Ref3, Sociedad)

            If Not String.IsNullOrEmpty(Comentarios) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Comentarios encontrados de este asiento"
                retVal.MENSAJEAUX = Comentarios

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentran los comentarios de este asiento"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarProyecto(ByVal objAsientoCampo As EntAsientoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un Asiento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "PROYECTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar proyecto para ObjType: " & objAsientoCampo.ObjType & ", DOCIDDW: " & objAsientoCampo.DOCIDDW & ", TransNum: " & objAsientoCampo.TransNum & ", Ref1: " & objAsientoCampo.Ref1 & ", Ref2: " & objAsientoCampo.Ref2 & ", Ref3: " & objAsientoCampo.Ref3)

            'Obligatorios
            If String.IsNullOrEmpty(objAsientoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAsientoCampo.TransNum) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref1) AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref2) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref3) Then _
                Throw New Exception("ID Docuware, número de asiento o referencia origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAsientoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsientoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el TransId del asiento definitivo
            Dim sProyecto As String = getProyectoAsientoDefinitivo(Tabla, objAsientoCampo.DOCIDDW, objAsientoCampo.TransNum, objAsientoCampo.Ref1, objAsientoCampo.Ref2, objAsientoCampo.Ref3, Sociedad)

            If Not String.IsNullOrEmpty(sProyecto) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Proyecto encontrado"
                retVal.MENSAJEAUX = sProyecto

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el proyecto"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarNumTransaccion(ByVal objAsientoCampo As EntAsientoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un Asiento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUMERO TRANSACCION"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar número transacción para ObjType: " & objAsientoCampo.ObjType & ", DOCIDDW: " & objAsientoCampo.DOCIDDW & ", TransNum: " & objAsientoCampo.TransNum & ", Ref1: " & objAsientoCampo.Ref1 & ", Ref2: " & objAsientoCampo.Ref2 & ", Ref3: " & objAsientoCampo.Ref3)

            'Obligatorios
            If String.IsNullOrEmpty(objAsientoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAsientoCampo.TransNum) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref1) AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref2) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref3) Then _
                Throw New Exception("ID Docuware, número de asiento o referencia origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAsientoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsientoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el TransId del asiento definitivo
            Dim NumTransaccion As String = getNumTransaccionAsientoDefinitivo(Tabla, objAsientoCampo.DOCIDDW, objAsientoCampo.TransNum, objAsientoCampo.Ref1, objAsientoCampo.Ref2, objAsientoCampo.Ref3, Sociedad)

            If Not String.IsNullOrEmpty(NumTransaccion) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Número de transacción encontrado de este asiento"
                retVal.MENSAJEAUX = NumTransaccion

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el número de transacción de este asiento"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarCampoUsuario(ByVal objAsientoCampo As EntAsientoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un Asiento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "CAMPO USUARIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar campo usuario para ObjType: " & objAsientoCampo.ObjType & ", DOCIDDW: " & objAsientoCampo.DOCIDDW & ", TransNum: " & objAsientoCampo.TransNum & ", Ref1: " & objAsientoCampo.Ref1 & ", Ref2: " & objAsientoCampo.Ref2 & ", Ref3: " & objAsientoCampo.Ref3)

            'Obligatorios
            If String.IsNullOrEmpty(objAsientoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAsientoCampo.TransNum) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref1) AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref2) _
                AndAlso String.IsNullOrEmpty(objAsientoCampo.Ref3) Then _
                Throw New Exception("ID Docuware, número de asiento o referencia origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = Utilidades.getTablaDeObjType(objAsientoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAsientoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el TransId del asiento definitivo
            Dim CampoUsuario As String = getCampoUsuarioAsientoDefinitivo(Tabla, objAsientoCampo.DOCIDDW, objAsientoCampo.TransNum, objAsientoCampo.Ref1, objAsientoCampo.Ref2, objAsientoCampo.Ref3, objAsientoCampo.CampoUsuario, Sociedad)

            If Not String.IsNullOrEmpty(CampoUsuario) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Campo de usuario encontrado"
                retVal.MENSAJEAUX = CampoUsuario

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el campo de usuario"
                retVal.MENSAJEAUX = ""

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

#End Region

#Region "Comunes"

    Private Function getPeriodoAbiertoDeTaxDate(ByVal TaxDate As String,
                                                ByVal Sociedad As eSociedad) As Boolean

        'Comprueba si el periodo contable está abierto

        Dim retVal As Boolean = False

        Try

            'Buscamos por fecha contable
            Dim SQL As String = ""
            SQL = "  SELECT COUNT(*) " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OFPR", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("PeriodStat") & "<>N'" & DocStatus.Cerrado & "' " & vbCrLf
            SQL &= " And '" & TaxDate & "' BETWEEN T0." & putQuotes("F_TaxDate") & " And T0." & putQuotes("T_TaxDate") & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = CInt(oObj) > 0

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    'Private Function getVatGroupDeVatTotal(ByVal LineTotal As Double,
    '                                       ByVal VatTotal As Double,
    '                                       ByVal VatPorc As Double,
    '                                       ByVal Intracomunitario As String,
    '                                       ByVal Sociedad As eSociedad) As String

    '    'Grupo de IVA

    '    Dim retVal As String = ""

    '    Try

    '        'Porcentaje de IVA (cuidado divisiones por 0)
    '        Dim VatPerc As Double = 0
    '        If LineTotal > 0 Then VatPerc = Math.Round(VatTotal / LineTotal * 100.0, 2)

    '        'Buscamos por porcentaje de IVA entre Sx
    '        Dim SQL As String = ""

    '        SQL = "  SELECT " & vbCrLf
    '        SQL &= " T." & putQuotes("CODE") & vbCrLf
    '        SQL &= " FROM ( " & vbCrLf
    '        SQL &= "    SELECT " & vbCrLf
    '        SQL &= "	MAX(COALESCE(T1." & putQuotes("Rate") & ", 0)) As RATE," & vbCrLf
    '        SQL &= "	T1." & putQuotes("Code") & " As CODE" & vbCrLf
    '        SQL &= "	FROM " & getDataBaseRef("OVTG", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
    '        SQL &= "	JOIN " & getDataBaseRef("VTG1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("Code") & "=T0." & putQuotes("Code") & vbCrLf
    '        SQL &= "	WHERE 1 = 1" & vbCrLf

    '        'Activo
    '        SQL &= "    And T0." & putQuotes("Inactive") & "<>N'" & SN.Yes & "'" & vbCrLf

    '        'Intracomunitario
    '        SQL &= "    And T0." & putQuotes("IsEC") & "=N'" & IIf(Intracomunitario = SN.Si, SN.Yes, SN.No) & "'" & vbCrLf

    '        'IVA soportado/repercutivo
    '        If Ambito = Utilidades.Ambito.Ventas Then

    '            'Categoría
    '            SQL &= "    And T0." & putQuotes("Category") & "=N'" & IVA.Repercutido & "'" & vbCrLf

    '            'Grupo de IVA del IC
    '            If Not String.IsNullOrEmpty(ICGrupoIVA) Then

    '                'MLD (2022/11/07): Grupo de IVA del IC cambiando los números
    '                SQL &= "	And T0." & putQuotes("Code") & " IN ('x1_y2_z3'"
    '                For i = 0 To 9
    '                    SQL &= ",N'" & ObtenerGrupoIVA(ICGrupoIVA, i.ToString) & "'"
    '                Next
    '                SQL &= ") " & vbCrLf

    '            ElseIf Intracomunitario = SN.Si Then

    '                'Códigos de IVA según el nombre
    '                SQL &= "    And T0." & putQuotes("Name") & " like '%intracomunitari%'" & vbCrLf

    '            Else

    '                'Códigos de IVA por defecto
    '                SQL &= "	And T0." & putQuotes("Code") & " IN ('R0','R1','R2','R3') " & vbCrLf
    '                'SQL &= "    And T0." & putQuotes("Name") & " like 'IVA repercutido al%'" & vbCrLf

    '            End If

    '        Else

    '            'Categoría
    '            SQL &= "    And T0." & putQuotes("Category") & "=N'" & IVA.Soportado & "'" & vbCrLf

    '            'Grupo de IVA del IC
    '            If Not String.IsNullOrEmpty(ICGrupoIVA) Then

    '                'MLD (2022/11/07): Grupo de IVA del IC cambiando los números
    '                SQL &= "	And T0." & putQuotes("Code") & " IN ('x1_y2_z3'"
    '                For i = 0 To 9
    '                    SQL &= ",N'" & ObtenerGrupoIVA(ICGrupoIVA, i.ToString) & "'"
    '                Next
    '                SQL &= ") " & vbCrLf

    '            ElseIf Intracomunitario = SN.Si Then

    '                'Códigos de IVA según el nombre
    '                SQL &= "    And T0." & putQuotes("Name") & " like 'Adquisiciones Intracomunitarias de Bienes%'" & vbCrLf

    '            Else

    '                'Códigos de IVA por defecto
    '                SQL &= "	And T0." & putQuotes("Code") & " IN ('S0','S1','S2','S3') " & vbCrLf
    '                'SQL &= "    And T0." & putQuotes("Name") & " like 'IVA soportado al%'" & vbCrLf

    '            End If

    '        End If

    '        SQL &= "	GROUP BY " & vbCrLf
    '        SQL &= "	T1." & putQuotes("Code") & ", " & vbCrLf
    '        SQL &= "	T0." & putQuotes("Code") & vbCrLf

    '        SQL &= " ) As T " & vbCrLf

    '        SQL &= " WHERE 1=1 " & vbCrLf

    '        'MLD (2021/02/03): Pasan el % de IVA, aplicar corrección al calculado por nosotros
    '        If VatPorc > 0 Then
    '            SQL &= " And T." & putQuotes("RATE") & "=" & VatPorc.ToString.Replace(",", ".") & vbCrLf
    '        Else
    '            SQL &= " And T." & putQuotes("RATE") & "<=" & (VatPerc * 1.03).ToString.Replace(",", ".") & vbCrLf
    '            SQL &= " And T." & putQuotes("RATE") & ">=" & (VatPerc * 0.97).ToString.Replace(",", ".") & vbCrLf
    '        End If

    '        Dim oCon As clsConexion = New clsConexion(Sociedad)

    '        Dim oObj As Object = oCon.ExecuteScalar(SQL)

    '        If Not String.IsNullOrEmpty(ICGrupoIVA) AndAlso String.IsNullOrEmpty(oObj) Then _
    '            Throw New Exception("No se encuentra grupo IVA al " & VatPerc & "% tomando como base " & ICGrupoIVA)

    '        If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

    '    Catch ex As Exception
    '        clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
    '        If ex.Message.Contains("No se encuentra grupo IVA al ") Then Throw New Exception(ex.Message)
    '    End Try

    '    Return retVal

    'End Function

#End Region

#Region "Privadas"

    Private Function NuevoAsiento(ByVal objAsiento As EntAsientoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As JournalEntries = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO ASIENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de asiento para " & objAsiento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objAsiento.UserSAP, objAsiento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If Not DateTime.TryParseExact(objAsiento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha asiento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objAsiento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha asiento auxiliar no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objAsiento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If String.IsNullOrEmpty(objAsiento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objAsiento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            If objAsiento.Lineas Is Nothing Then Throw New Exception("Líneas no suministradas")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objAsiento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objAsiento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Objeto asiento
            oDocDestino = oCompany.GetBusinessObject(objAsiento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Asiento en firme")

            'Serie
            If objAsiento.Serie > 0 Then oDocDestino.Series = objAsiento.Serie

            'Referencias
            If Not String.IsNullOrEmpty(objAsiento.Ref1) Then oDocDestino.Reference = objAsiento.Ref1
            If Not String.IsNullOrEmpty(objAsiento.Ref2) Then oDocDestino.Reference2 = objAsiento.Ref2
            If Not String.IsNullOrEmpty(objAsiento.Ref3) Then oDocDestino.Reference3 = objAsiento.Ref3

            'Fecha contable 
            oDocDestino.ReferenceDate = Date.ParseExact(objAsiento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha vencimiento
            If DateTime.TryParseExact(objAsiento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                oDocDestino.DueDate = Date.ParseExact(objAsiento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha Asiento 
            'Comprueba si el periodo contable está abierto para la fecha DocDate de entrada. Si no lo está, coge la fecha auxiliar DocDateAux
            Dim PeriodoAbierto As Boolean = getPeriodoAbiertoDeTaxDate(objAsiento.DocDate.ToString, Sociedad)
            If PeriodoAbierto Then
                oDocDestino.TaxDate = Date.ParseExact(objAsiento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            Else
                oDocDestino.TaxDate = Date.ParseExact(objAsiento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            End If

            'Indicador
            If Not String.IsNullOrEmpty(objAsiento.Indicador) Then oDocDestino.Indicator = objAsiento.Indicador

            'Operacion
            If Not String.IsNullOrEmpty(objAsiento.Operacion) Then oDocDestino.TransactionCode = objAsiento.Operacion

            'Factura
            If Not String.IsNullOrEmpty(objAsiento.Factura) Then oDocDestino.UserFields.Fields.Item("U_B1SYS_INV_TYPE").Value = objAsiento.Factura

            'Proyecto
            If Not String.IsNullOrEmpty(objAsiento.Proyecto) Then oDocDestino.ProjectCode = objAsiento.Proyecto

            'Comentarios
            If objAsiento.Comments.Length > 254 Then objAsiento.Comments = objAsiento.Comments.Substring(0, 254)
            oDocDestino.Memo = objAsiento.Comments

            'Recorre cada línea 
            For Each objLinea In objAsiento.Lineas

                'Índice
                Dim indLinea As Integer = objAsiento.Lineas.IndexOf(objLinea) + 1

                'Comprobar campos
                If String.IsNullOrEmpty(objLinea.CuentaContable) AndAlso String.IsNullOrEmpty(objLinea.ShortName) Then _
                    Throw New Exception("Cuenta contable o código IC de la línea " & indLinea.ToString & " no suministrado")

                If String.IsNullOrEmpty(objLinea.LineTipo) Then Throw New Exception("Tipo de importe (D/H) de la línea " & indLinea.ToString & " no suministrado")

                If objLinea.LineTotal = 0 Then Throw New Exception("Importe de la línea " & indLinea.ToString & " no suministrado")

                'Se posiciona en la línea
                oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                'Cuenta contable
                If Not String.IsNullOrEmpty(objLinea.CuentaContable) Then oDocDestino.Lines.AccountCode = objLinea.CuentaContable

                'Codigo IC
                If Not String.IsNullOrEmpty(objLinea.ShortName) Then oDocDestino.Lines.ShortName = objLinea.ShortName

                'Importe
                If objLinea.LineTipo = AsientoLineaTipo.Debe Then
                    oDocDestino.Lines.Debit = objLinea.LineTotal
                Else
                    oDocDestino.Lines.Credit = objLinea.LineTotal
                End If

                'VATGroup
                If Not String.IsNullOrEmpty(objLinea.VATGroup) Then oDocDestino.Lines.VatGroup = objLinea.VATGroup

                'LineMemo
                If objLinea.LineMemo.Length > 254 Then objLinea.LineMemo = objLinea.LineMemo.Substring(0, 254)
                If Not String.IsNullOrEmpty(objLinea.LineMemo) Then oDocDestino.Lines.LineMemo = objLinea.LineMemo

                'Campos de usuario
                If Not objLinea.CamposUsuario Is Nothing AndAlso objLinea.CamposUsuario.Count > 0 Then

                    For Each oCampoUsuario In objLinea.CamposUsuario

                        Dim oUserField As Field = oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo)

                        Select Case oUserField.Type

                            Case BoFieldTypes.db_Numeric
                                If oUserField.SubType = BoFldSubTypes.st_Time Then
                                    If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                            oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                                Else
                                    If IsNumeric(oCampoUsuario.Valor) Then _
                                            oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                                End If

                            Case BoFieldTypes.db_Float
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                        oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                            Case BoFieldTypes.db_Date
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                        oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                            Case Else
                                If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                        oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                        End Select

                    Next

                End If

                'Añade línea
                oDocDestino.Lines.Add()

            Next

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objAsiento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objAsiento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objAsiento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = "S"

            'Campos de usuario
            If Not objAsiento.CamposUsuario Is Nothing AndAlso objAsiento.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objAsiento.CamposUsuario

                    Dim oUserField As Field = oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                    oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

            'Añadimos Asiento
            If oDocDestino.Add() = 0 Then

                'Obtiene el TransId y TransId del asiento añadido
                Dim TransId As String = oCompany.GetNewObjectKey

                Dim TransNum As String = ""
                TransNum = oComun.getTransNumDeTransId(TablaDestino, TransId, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Asiento creado con éxito"
                retVal.MENSAJEAUX = TransNum

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
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Private Function getTransNumAsientoPorDWID(ByVal Tabla As String,
                                               ByVal IDDW As String,
                                               ByVal Sociedad As eSociedad) As String

        'Devuelve el TransId del asiento en firme

        Dim retVal As String = ""

        Try

            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("Number") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(T0." & putQuotes("U_SEIIDDW") & ",'')  = N'" & IDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWAsiento(ByVal Tabla As String,
                                  ByVal DocDate As Integer,
                                  ByVal TransIds As List(Of String),
                                  ByVal IDDW As String,
                                  ByVal URLDW As String,
                                  ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el asiento

        Try

            'Actualizamos por CardCode, DocDate y TransId
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIIMPDW") & " = " & putQuotes("SysTotal") & " " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            'SQL &= " And COALESCE(" & putQuotes("DocDate") & "," & getDefaultDate & ") = N'" & DocDate & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("TransId") & ",-1)  IN ('-1' " & vbCrLf
            For Each TransId In TransIds
                SQL &= ", N'" & TransId & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el asiento: " & putQuotes(Tabla) & " para TransId: " & String.Join("#", TransIds))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Function setAsientoDWTratado(ByVal Tabla As String,
                                         ByVal TransId As String,
                                         ByVal Sociedad As eSociedad) As Boolean

        'Actualiza los campos de DW en el asiento

        Dim retval As Boolean = False

        Try

            'Actualizamos por CardCode, DocDate y TransId
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEITRADW") & " =N'" & SN.Si & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("TransId") & ",'') = N'" & TransId & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("Actualización nula al ejecutar sentencia " & SQL)

            retval = True

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retval

    End Function

    Private Function getTransIdDeRefOrigen(ByVal Tabla As String,
                                           ByVal RefOrigen As String,
                                           ByVal TipoRefOrigen As String,
                                           ByVal Sociedad As eSociedad) As List(Of String)

        'Devuelve el TransId del asiento

        Dim retVal As New List(Of String)

        Try

            'Buscamos por TransId, TransId o NumAtCard
            'El asiento no esté cerrado o cancelado para que no cree borradores sin líneas
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("TransId") & " As TransId " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.TransId
                    SQL &= " And T0." & putQuotes("TransId") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.TransNum
                    SQL &= " And T0." & putQuotes("Number") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.TransRef1
                    SQL &= " And T0." & putQuotes("Ref1") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
                Case Utilidades.RefOrigen.TransRef2
                    SQL &= " And T0." & putQuotes("Ref2") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("Ref3") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(DR.Item("TransId").ToString)
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getTransIdDeRefOrigenUnica(ByVal Tabla As String,
                                                ByVal RefOrigen As String,
                                                ByVal TipoRefOrigen As String,
                                                ByVal Sociedad As eSociedad) As String

        'Devuelve el TransId del asiento

        Dim retVal As String = ""

        Try

            'Buscamos por TransId, TransId o NumAtCard
            'El asiento no esté cerrado o cancelado para que no cree borradores sin líneas
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("TransId") & " As TransId " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.TransId
                    SQL &= " And T0." & putQuotes("TransId") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.TransNum
                    SQL &= " And T0." & putQuotes("Number") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.TransRef1
                    SQL &= " And T0." & putQuotes("Ref1") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
                Case Utilidades.RefOrigen.TransRef2
                    SQL &= " And T0." & putQuotes("Ref2") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("Ref3") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

#Region "Privadas: Obtener campo"

    Private Function getTransNumAsientoDefinitivo(ByVal Tabla As String,
                                                  ByVal DOCIDDW As String,
                                                  ByVal TransNum As String,
                                                  ByVal Ref1 As String,
                                                  ByVal Ref2 As String,
                                                  ByVal Ref3 As String,
                                                  ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un asiento definitivo

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("Number") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TransNum) Then SQL &= " And T0." & putQuotes("Number") & " = N'" & TransNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref1) Then SQL &= " And T0." & putQuotes("Ref1") & " = N'" & Ref1 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref3) Then SQL &= " And T0." & putQuotes("Ref3") & " = N'" & Ref3 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getTransTotalAsientoDefinitivo(ByVal Tabla As String,
                                                    ByVal DOCIDDW As String,
                                                    ByVal TransNum As String,
                                                    ByVal Ref1 As String,
                                                    ByVal Ref2 As String,
                                                    ByVal Ref3 As String,
                                                    ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un asiento definitivo

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("SysTotal") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TransNum) Then SQL &= " And T0." & putQuotes("Number") & " = N'" & TransNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref1) Then SQL &= " And T0." & putQuotes("Ref1") & " = N'" & Ref1 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref3) Then SQL &= " And T0." & putQuotes("Ref3") & " = N'" & Ref3 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf


            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString.Replace(",", ".")

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getTransDueDateAsientoDefinitivo(ByVal Tabla As String,
                                                      ByVal DOCIDDW As String,
                                                      ByVal TransNum As String,
                                                      ByVal Ref1 As String,
                                                      ByVal Ref2 As String,
                                                      ByVal Ref3 As String,
                                                      ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un asiento definitivo

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " " & getDateAsString_yyyyMMdd("T0.", "DueDate") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TransNum) Then SQL &= " And T0." & putQuotes("Number") & " = N'" & TransNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref1) Then SQL &= " And T0." & putQuotes("Ref1") & " = N'" & Ref1 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref3) Then SQL &= " And T0." & putQuotes("Ref3") & " = N'" & Ref3 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getComentariosAsientoDefinitivo(ByVal Tabla As String,
                                                     ByVal DOCIDDW As String,
                                                     ByVal TransNum As String,
                                                     ByVal Ref1 As String,
                                                     ByVal Ref2 As String,
                                                     ByVal Ref3 As String,
                                                     ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un asiento definitivo

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("Memo") & ",'') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TransNum) Then SQL &= " And T0." & putQuotes("Number") & " = N'" & TransNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref1) Then SQL &= " And T0." & putQuotes("Ref1") & " = N'" & Ref1 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref3) Then SQL &= " And T0." & putQuotes("Ref3") & " = N'" & Ref3 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getNumTransaccionAsientoDefinitivo(ByVal Tabla As String,
                                                        ByVal DOCIDDW As String,
                                                        ByVal TransNum As String,
                                                        ByVal Ref1 As String,
                                                        ByVal Ref2 As String,
                                                        ByVal Ref3 As String,
                                                        ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un asiento definitivo

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("TransId") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TransNum) Then SQL &= " And T0." & putQuotes("Number") & " = N'" & TransNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref1) Then SQL &= " And T0." & putQuotes("Ref1") & " = N'" & Ref1 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref3) Then SQL &= " And T0." & putQuotes("Ref3") & " = N'" & Ref3 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getProyectoAsientoDefinitivo(ByVal Tabla As String,
                                                  ByVal DOCIDDW As String,
                                                  ByVal TransNum As String,
                                                  ByVal Ref1 As String,
                                                  ByVal Ref2 As String,
                                                  ByVal Ref3 As String,
                                                  ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un asiento definitivo

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("Project") & vbCrLf
            SQL &= " FROM " & Utilidades.getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            'SQL &= " JOIN " & Utilidades.getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("TransId") & " = T0." & putQuotes("TransId") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TransNum) Then SQL &= " And T0." & putQuotes("Number") & " = N'" & TransNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref1) Then SQL &= " And T0." & putQuotes("Ref1") & " = N'" & Ref1 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref3) Then SQL &= " And T0." & putQuotes("Ref3") & " = N'" & Ref3 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            'SQL &= " ORDER BY T1." & putQuotes("Line_ID") & " ASC" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            Dim oObj As Object = oCon.ExecuteScalar(SQL)

            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then
                retVal = oObj.ToString
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getCampoUsuarioAsientoDefinitivo(ByVal Tabla As String,
                                                      ByVal DOCIDDW As String,
                                                      ByVal TransNum As String,
                                                      ByVal Ref1 As String,
                                                      ByVal Ref2 As String,
                                                      ByVal Ref3 As String,
                                                      ByVal CampoUsuario As String,
                                                      ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un asiento definitivo

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes(CampoUsuario) & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TransNum) Then SQL &= " And T0." & putQuotes("Number") & " = N'" & TransNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref1) Then SQL &= " And T0." & putQuotes("Ref1") & " = N'" & Ref1 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref3) Then SQL &= " And T0." & putQuotes("Ref3") & " = N'" & Ref3 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            Dim oObj As Object = oCon.ExecuteScalar(SQL)

            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then
                retVal = oObj.ToString
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & System.Reflection.MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

#End Region

#End Region

End Class
