Imports System.IO
Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class clsFichero


#Region "Públicas"

    Public Function getGuardarFichero(ByVal objFichero As EntFichero, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "GUARDAR FICHERO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de fichero " & objFichero.Nombre)

            'Obligatorios
            If String.IsNullOrEmpty(objFichero.Ruta) Then objFichero.Ruta = ObtenerRutaPorDefecto(Sociedad)

            If String.IsNullOrEmpty(objFichero.Ruta) Then Throw New Exception("Ruta no suministrada")
            If Not Directory.Exists(objFichero.Ruta) Then Throw New Exception("Ruta no existente")

            If String.IsNullOrEmpty(objFichero.Nombre) Then Throw New Exception("Nombre no suministrado")

            Dim NombreSinExtension As String = Path.GetFileNameWithoutExtension(objFichero.Nombre)
            If String.IsNullOrEmpty(NombreSinExtension) Then Throw New Exception("Nombre no suministrado o incorrecto")

            Dim Extension As String = Path.GetExtension(objFichero.Nombre)
            If String.IsNullOrEmpty(Extension) Then Throw New Exception("Extensión no suministrada")

            If String.IsNullOrEmpty(objFichero.Base64) Then Throw New Exception("Base64 no suministrada")

            'Base64 a binario
            Dim Base64 As String = FormatearBase64(objFichero.Base64)
            Dim Binario As Byte() = Convert.FromBase64String(Base64)

            'Ruta
            Dim Ruta As String = ObtenerRuta(objFichero.Ruta, objFichero.Nombre, Sociedad)
            If File.Exists(Ruta) Then Throw New Exception("Ya existe fichero " & objFichero.Nombre)

            'Guarda en disco
            GuardarFichero(Binario, Ruta)

            'Ruta del fichero
            If File.Exists(Ruta) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Fichero creado con éxito"
                retVal.MENSAJEAUX = objFichero.Nombre

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "Fichero no creado"
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

#Region "Privadas"

    Public Function FormatearBase64(ByVal Base64 As String) As String

        Return Base64.Replace("\n", "")

    End Function

    Public Shared Sub GuardarFichero(ByVal Binario As Byte(), ByVal Ruta As String)

        Try

            'Dim sDirectorio As String = Path.GetDirectoryName(sRuta)
            'Dim sExtension As String = Path.GetExtension(sRuta)
            'If (Not Directory.Exists(sDirectorio)) Then Directory.CreateDirectory(sDirectorio)

            'Se asegura de que el fichero destino no existe
            If File.Exists(Ruta) Then File.Delete(Ruta)

            File.WriteAllBytes(Ruta, Binario)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Public Shared Function ObtenerRutaPorDefecto(ByVal Sociedad As eSociedad) As String

        Dim ruta As String = ""

        Try

            'Por sociedad
            Dim SociedadNombre As String = Utilidades.NOMBRESOCIEDAD(Sociedad)
            ruta = ConfigurationManager.AppSettings.Get("ruta_" & SociedadNombre).ToString

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return ruta

    End Function

    Public Shared Function ObtenerRuta(ByVal Directorio As String, ByVal Nombre As String, ByVal Sociedad As eSociedad) As String

        Dim ruta As String = ""

        Try

            If Not Directorio.EndsWith("\") Then Directorio = Directorio & "\"

            ruta = Directorio & Nombre

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return ruta

    End Function

#End Region

End Class
