Imports SAPbobsCOM
Imports System.IO
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsAnexo

#Region "Públicas"

    Public Function CrearAnexo(ByVal objAnexo As EntAnexoSAP, ByVal Sociedad As eSociedad) As EntResultado

        'Adjuntar anexo

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As Documents = Nothing
        Dim oDocAnexo As Attachments2 = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ANEXO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de anexo para ObjType: " & objAnexo.ObjType & ", Ambito: " & objAnexo.Ambito & ", NIFTercero: " & objAnexo.NIFTercero & ", Razón social: " & objAnexo.RazonSocial & ", DOCIDDW: " & objAnexo.DOCIDDW & ", DocNum: " & objAnexo.DocNum & ", NumAtCard: " & objAnexo.NumAtCard & ", Nombre: " & objAnexo.Nombre)

            oCompany = ConexionSAP.getCompany(objAnexo.UserSAP, objAnexo.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objAnexo.NIFTercero) AndAlso String.IsNullOrEmpty(objAnexo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objAnexo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAnexo.DocNum) AndAlso String.IsNullOrEmpty(objAnexo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            If String.IsNullOrEmpty(objAnexo.Nombre) Then Throw New Exception("Nombre no suministrado")
            If String.IsNullOrEmpty(Path.GetExtension(objAnexo.Nombre)) Then Throw New Exception("Nombre sin extensión")
            If String.IsNullOrEmpty(objAnexo.Base64) Then Throw New Exception("Base64 no suministrada")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objAnexo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAnexo.DocNum) Then
                CardCode = oComun.getCardCode(objAnexo.NIFTercero, objAnexo.RazonSocial, objAnexo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objAnexo.NIFTercero & ", Razón social: " & objAnexo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAnexo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAnexo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Buscamos DocEntry por DOCIDDW, DocNum o NumAtCard
            Dim DocEntry As String = oComun.getDocEntryDocumentoDefinitivo(Tabla, CardCode, objAnexo.DOCIDDW, objAnexo.DocNum, objAnexo.NumAtCard, False, False, Sociedad)
            If String.IsNullOrEmpty(DocEntry) Then Throw New Exception("No encuentro documento con DOCIDDW - '" & objAnexo.DOCIDDW & "', DocNum - '" & objAnexo.DocNum & "' y NumAtCard - '" & objAnexo.NumAtCard & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocEntry: " & DocEntry)

            'Ruta
            Dim RutaAnexo As String = GuardarAnexo(objAnexo)
            If Not File.Exists(RutaAnexo) Then Throw New Exception("Error al guardar el anexo en base64")

            'Objeto destino
            oDocDestino = oCompany.GetBusinessObject(objAnexo.ObjType)
            If Not oDocDestino.GetByKey(DocEntry) Then Throw New Exception("No puedo recuperar el documento destino con DocEntry: " & DocEntry)

            'Objeto anexo
            oDocAnexo = CType(oCompany.GetBusinessObject(BoObjectTypes.oAttachments2), Attachments2)

            If Not oDocAnexo.GetByKey(oDocDestino.AttachmentEntry) Then

                'Nuevo 
                retVal = NuevoAnexo(oCompany, oDocDestino, oDocAnexo, RutaAnexo)

            Else

                'Actualizar
                retVal = ActualizarAnexo(oCompany, oDocDestino, oDocAnexo, RutaAnexo)

            End If

            'Borrar anexo
            BorrarAnexo(RutaAnexo)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oDocAnexo)
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

#End Region

#Region "Privadas"

    Private Function NuevoAnexo(ByRef oCompany As Company, ByRef oDocDestino As Documents, ByRef oDocAnexo As Attachments2,
                                ByVal RutaAnexo As String) As EntResultado

        'Adjuntar anexo

        Dim retVal As New EntResultado

        Try

            'Nuevo 
            oDocAnexo.Lines.SetCurrentLine(oDocAnexo.Lines.Count - 1)
            oDocAnexo.Lines.SourcePath = Path.GetDirectoryName(RutaAnexo)
            oDocAnexo.Lines.FileExtension = Path.GetExtension(RutaAnexo).Replace(".", "")
            oDocAnexo.Lines.FileName = Path.GetFileNameWithoutExtension(RutaAnexo)
            oDocAnexo.Lines.Add()

            'Añadimos anexo
            If oDocAnexo.Add() = 0 Then

                'Obtiene el AttachEntry del anexo añadido
                Dim AttachEntry As String = oCompany.GetNewObjectKey

                'Anexo destino
                If Not oDocAnexo.GetByKey(CInt(AttachEntry)) Then Throw New Exception("No puedo recuperar el anexo con AttachEntry: " & AttachEntry)

                'Crear anexo
                oDocDestino.AttachmentEntry = oDocAnexo.AbsoluteEntry

                'Actualizar documento
                If oDocDestino.Update() = 0 Then

                    retVal.CODIGO = Respuesta.Ok
                    retVal.MENSAJE = "Anexo añadido con éxito"
                    retVal.MENSAJEAUX = oDocDestino.DocNum

                Else

                    retVal.CODIGO = Respuesta.Ko
                    retVal.MENSAJE = oCompany.GetLastErrorCode & "#" & oCompany.GetLastErrorDescription
                    retVal.MENSAJEAUX = ""

                End If

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
        End Try

        Return retVal

    End Function

    Private Function ActualizarAnexo(ByRef oCompany As Company, ByRef oDocDestino As Documents, ByRef oDocAnexo As Attachments2,
                                     ByVal RutaAnexo As String) As EntResultado

        'Adjuntar anexo
        Dim retVal As New EntResultado

        Try

            'Actualizar
            oDocAnexo.Lines.SetCurrentLine(oDocAnexo.Lines.Count - 1)
            If Not String.IsNullOrEmpty(oDocAnexo.Lines.FileName) Then
                oDocAnexo.Lines.Add()
                oDocAnexo.Lines.SetCurrentLine(oDocAnexo.Lines.Count - 1)
            End If

            oDocAnexo.Lines.SourcePath = Path.GetDirectoryName(RutaAnexo)
            oDocAnexo.Lines.FileExtension = Path.GetExtension(RutaAnexo).Replace(".", "")
            oDocAnexo.Lines.FileName = Path.GetFileNameWithoutExtension(RutaAnexo)
            oDocAnexo.Lines.Override = BoYesNoEnum.tNO

            'Actualizar anexo
            If Not oDocAnexo.Update() = 0 Then Throw New Exception(oCompany.GetLastErrorCode & "#" & oCompany.GetLastErrorDescription)

            'Actualizar documento
            If oDocDestino.Update() = 0 Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Anexo añadido con éxito"
                retVal.MENSAJEAUX = oDocDestino.DocNum

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
        End Try

        Return retVal

    End Function

    Private Function GuardarAnexo(ByVal objAnexo As EntAnexoSAP) As String

        'Crea el anexo en temporal
        Dim retVal As String = ""

        Try

            'TEST
            If objAnexo.Base64 = "xyz" Then
                Dim Binario As Byte() = File.ReadAllBytes("C:\Users\mlanza\OneDrive - SEIDOR SA\Escritorio\Adjuntos\Fotos.pdf")
                objAnexo.Base64 = Convert.ToBase64String(Binario)
            End If

            'Formatea la base64 
            Dim Base64 As String = objAnexo.Base64.Replace("\n", "")

            'Elimina los caracteres inválidos del nombre
            Dim Nombre As String = objAnexo.Nombre.Replace(Path.GetInvalidFileNameChars(), "").Replace(vbLf, "").Replace(vbCrLf, "")

            'Compone la ruta
            Dim Ruta As String = Path.GetTempPath() & Nombre

            'Guardar en disco
            File.WriteAllBytes(Ruta, Convert.FromBase64String(Base64))

            retVal = ruta

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function BorrarAnexo(ByVal RutaAnexo As String) As String

        'Borra el anexo en temporal
        Dim retVal As String = ""

        Try

            If File.Exists(RutaAnexo) Then File.Delete(RutaAnexo)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

#End Region

End Class
