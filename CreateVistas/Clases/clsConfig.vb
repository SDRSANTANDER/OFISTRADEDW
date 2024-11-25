Imports System.Configuration

Public Class clsConfig

#Region "Web config"

    'SOCIEDAD
    Shared Function wcSociedad() As String
        Return ConfigurationManager.AppSettings("Sociedad").ToString
    End Function

    'RUTAS
    Shared Function wcVISRuta() As String
        Return ConfigurationManager.AppSettings("VISRuta").ToString
    End Function

#End Region

End Class

