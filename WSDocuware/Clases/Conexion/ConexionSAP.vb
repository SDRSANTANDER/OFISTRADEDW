Imports System.Reflection
Public Class ConexionSAP

    Public Shared Function TestCompany(ByVal userSAP As String, ByVal passSAP As String, ByVal Sociedad As Utilidades.eSociedad) As Boolean

        Dim retVal As Boolean

        If Not getCompany(userSAP, passSAP, Sociedad) Is Nothing Then

            If getCompany(userSAP, passSAP, Sociedad).Connected Then
                retVal = True
                getCompany(userSAP, passSAP, Sociedad).Disconnect()
            Else
                retVal = False
            End If

        Else
            retVal = False
        End If

        Return retVal

    End Function

    Public Shared Function getCompany(ByVal userSAP As String, ByVal passSAP As String, ByVal Sociedad As Utilidades.eSociedad) As SAPbobsCOM.Company

        Try

            Dim oCompany As SAPbobsCOM.Company
            oCompany = New SAPbobsCOM.Company

            If loginSAP(oCompany, userSAP, passSAP, Sociedad) Then
                Return oCompany
            Else
                Throw New Exception("Imposible conectar con SAP")
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return Nothing

    End Function

    Private Shared Function loginSAP(ByRef oCompany As SAPbobsCOM.Company, ByVal userSAP As String, ByVal passSAP As String, ByVal Sociedad As Utilidades.eSociedad) As Boolean

        Try

            If oCompany Is Nothing Then
                oCompany = New SAPbobsCOM.Company
            End If

            If oCompany.Connected Then
                oCompany.Disconnect()
            End If

            'Por sociedad
            Dim SociedadNombre As String = Utilidades.NOMBRESOCIEDAD(Sociedad)

            Dim BD As String = ConfigurationManager.AppSettings.Get("bd_" & SociedadNombre).ToString

            'Si no informa del usuario SAP, usa el de por defecto
            Dim SAPUser As String = userSAP
            Dim SAPPass As String = passSAP

            If String.IsNullOrEmpty(SAPUser) OrElse String.IsNullOrEmpty(SAPPass) Then
                SAPUser = ConfigurationManager.AppSettings.Get("userSAP").ToString()
                SAPPass = ConfigurationManager.AppSettings.Get("passSAP").ToString()
            End If

            'Por servidor (común)
            Dim Server As String = ConfigurationManager.AppSettings.Get("Server").ToString
            Dim License As String = ConfigurationManager.AppSettings.Get("LicenseServer").ToString
            Dim DBType As Integer = CInt(ConfigurationManager.AppSettings.Get("DBType").ToString)
            Dim DBUser As String = ConfigurationManager.AppSettings.Get("DBUser").ToString
            Dim DBPass As String = ConfigurationManager.AppSettings.Get("DBPass").ToString()
            Dim SLD As String = ConfigurationManager.AppSettings.Get("SLD").ToString()

            oCompany.CompanyDB = BD
            oCompany.UserName = SAPUser
            oCompany.Password = SAPPass
            oCompany.LicenseServer = License
            oCompany.Server = Server

            oCompany.DbUserName = DBUser
            oCompany.DbPassword = DBPass
            oCompany.DbServerType = CType(DBType, SAPbobsCOM.BoDataServerTypes)

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
            oCompany.UseTrusted = False

            If Not String.IsNullOrEmpty(SLD) Then
                oCompany.SLDServer = SLD
            End If

            Dim iEstado As Integer = oCompany.Connect
            Dim err As String = oCompany.GetLastErrorDescription
            'clsLog.Log.Info(iEstado & "-" & err)

            If iEstado = 0 AndAlso String.IsNullOrEmpty(err) Then
                Return True
            Else
                clsLog.Log.Fatal("El estado en connect es " & iEstado & ":" & err & " en " & MethodBase.GetCurrentMethod().Name)
                Return False
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Return False
        End Try

    End Function

End Class



