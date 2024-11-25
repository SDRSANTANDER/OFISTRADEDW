Imports Microsoft.VisualBasic

Public Class clsLog

    Private Shared mLog As log4net.ILog = log4net.LogManager.GetLogger("root")

    Public Enum eLevel
        isDebug = 1
        isInfo = 2
        isWarning = 3
        isError = 4
        isFatal = 5
    End Enum

    Shared Sub New()
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    Public Shared Sub Log(ByVal Mensaje As String, ByVal excepcion As Exception, ByVal Sev As eLevel)

        Try

            Select Case Sev

                Case eLevel.isDebug
                    mLog.Debug(Mensaje, excepcion)
                Case eLevel.isInfo
                    mLog.Info(Mensaje, excepcion)
                Case eLevel.isWarning
                    mLog.Warn(Mensaje, excepcion)
                Case eLevel.isError
                    mLog.Error(Mensaje, excepcion)
                Case eLevel.isFatal
                    mLog.Fatal(Mensaje, excepcion)

            End Select

        Catch ex As Exception

        End Try

    End Sub

    Public Shared Sub Log(ByVal Mensaje As String, ByVal Sev As eLevel)

        Try

            Select Case Sev

                Case eLevel.isDebug
                    mLog.Debug(Mensaje)
                Case eLevel.isInfo
                    mLog.Info(Mensaje)
                Case eLevel.isWarning
                    mLog.Warn(Mensaje)
                Case eLevel.isError
                    mLog.Error(Mensaje)
                Case eLevel.isFatal
                    mLog.Fatal(Mensaje)

            End Select

        Catch ex As Exception

        End Try

    End Sub


End Class
