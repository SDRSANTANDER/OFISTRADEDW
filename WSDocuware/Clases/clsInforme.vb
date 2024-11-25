Imports System.Reflection
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportAppServer.ClientDoc
Imports WSDocuware.Utilidades

Public Class clsInforme

#Region "Públicas"

    Public Function GenerarInformeBase64(ByVal InfomeTipo As String,
                                         ByVal DocEntry As String,
                                         ByVal DocNum As String,
                                         ByRef FicheroRuta As String,
                                         ByVal UserSAP As String,
                                         ByVal PassSAP As String,
                                         ByVal Sociedad As eSociedad) As String

        Dim FicheroBase64 As String = ""
        Dim oCompany As SAPbobsCOM.Company

        Dim oComun As New clsComun

        Dim sLogInfo As String = "GENERAR INFORME"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de generación de fichero para InfomeTipo: " & InfomeTipo & ", DocEntry: " & DocEntry & ", DocNum: " & DocNum)

            oCompany = ConexionSAP.getCompany(UserSAP, PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Nombre del fichero
            FicheroRuta = getFicheroRuta(InfomeTipo, DocNum)

            Dim rptDoc As ReportDocument = Nothing

            'ESPECIFICOS PARA CADA EMPRESA
            Dim SAP_BD As String = ConfigurationManager.AppSettings.Get("bd_" & NOMBRESOCIEDAD(Sociedad)).ToString
            Dim SAP_Server As String = ConfigurationManager.AppSettings.Get("server").ToString
            Dim SQL_UserName As String = ConfigurationManager.AppSettings.Get("DBUser").ToString
            Dim SQL_Password As String = ConfigurationManager.AppSettings.Get("DBPass").ToString
            Dim CrystalRuta As String = getCrystalRuta(InfomeTipo)

            rptDoc = New ReportDocument()
            rptDoc.Load(CrystalRuta)

            Dim rptClientDoc As ISCDReportClientDocument
            rptClientDoc = rptDoc.ReportClientDocument

            Dim crConnectionInfo As New ConnectionInfo
            With crConnectionInfo
                .ServerName = SAP_Server
                .DatabaseName = SAP_BD
                .UserID = SQL_UserName
                .Password = SQL_Password
            End With

            Dim CrTables As Tables
            Dim CrTable As Table
            Dim crTableLogoninfo As New TableLogOnInfo

            CrTables = rptDoc.Database.Tables
            For Each CrTable In CrTables
                crTableLogoninfo = CrTable.LogOnInfo
                crTableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crTableLogoninfo)
            Next

            rptDoc.SetDatabaseLogon(SQL_UserName, SQL_Password, SAP_Server, SAP_BD)

            rptDoc.SetParameterValue("Dockey@", DocEntry)

            rptDoc.ExportToDisk(ExportFormatType.PortableDocFormat, FicheroRuta)
            rptDoc.Dispose()
            rptDoc.Close()

            Dim FicheroBinario As Byte() = IO.File.ReadAllBytes(FicheroRuta)
            FicheroBase64 = Convert.ToBase64String(FicheroBinario)

            oComun.EliminarFichero(FicheroRuta)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return FicheroBase64

    End Function

#End Region

#Region "Privadas"

    Private Function getFicheroRuta(ByVal InfomeTipo As String, ByVal DocNum As String) As String

        'Ruta del fichero
        Dim FicheroRuta As String = ""

        Try

            Select Case InfomeTipo
                Case Generar.CONTRATOMENOR
                    FicheroRuta = IO.Path.GetTempPath & "ContratoMenor_" & DocNum & ".pdf"
                Case Else
                    Throw New Exception("Tipo de fichero no permitido")
            End Select

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return FicheroRuta

    End Function

    Private Function getCrystalRuta(ByVal InfomeTipo As String) As String

        'Ruta del crystal
        Dim CrystalRuta As String = ""

        Try

            Select Case InfomeTipo
                Case Generar.CONTRATOMENOR
                    CrystalRuta = ConfigurationManager.AppSettings.Get("RutaCrystalContratoMenor").ToString
                Case Else
                    Throw New Exception("Tipo de informe no permitido")
            End Select

            'Comprueba que exista el crystal
            If Not IO.File.Exists(CrystalRuta) Then
                Throw New Exception("No se encuentra el informe de crystal")
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return CrystalRuta

    End Function

#End Region

End Class
