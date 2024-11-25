Imports Microsoft.VisualBasic

Public Class clsLog
    Public Shared Log As log4net.ILog = log4net.LogManager.GetLogger("root")

    Shared Sub New()
        log4net.Config.XmlConfigurator.Configure()
    End Sub

End Class
