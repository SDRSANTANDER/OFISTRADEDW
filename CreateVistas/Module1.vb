Imports System.IO
Imports System.Reflection
Imports CrearVistas.clsConfig

Module Module1

    Sub Main()

        Try

            clsLog.Log("INICIO: Creación de scripts en " & wcSociedad(), clsLog.eLevel.isInfo)

            'Leer todos los scripts a crear
            Dim sqlScripts As String() = Directory.GetFiles(wcVISRuta, "*.sql", SearchOption.AllDirectories)

            'Crear los scripts
            For Each sqlScript In sqlScripts

                If File.Exists(sqlScript) Then

                    Dim sql As String = File.ReadAllText(sqlScript)

                    Dim oDAOScript As New DAOScript
                    Dim bCreado As Boolean = oDAOScript.setEjecutarScript(sql, wcSociedad)

                    If bCreado Then
                        clsLog.Log("Creado script " & Path.GetFileName(sqlScript) & " en " & MethodBase.GetCurrentMethod().Name, clsLog.eLevel.isInfo)
                    Else
                        clsLog.Log("No creado script " & Path.GetFileName(sqlScript) & " en " & MethodBase.GetCurrentMethod().Name, clsLog.eLevel.isFatal)
                    End If

                Else
                    clsLog.Log("No encuentro script " & sqlScript & " en " & MethodBase.GetCurrentMethod().Name, clsLog.eLevel.isError)
                End If

            Next

            clsLog.Log("FIN: Creación de scripts en " & wcSociedad(), clsLog.eLevel.isInfo)

        Catch ex As Exception
            clsLog.Log(ex.Message & " en [" & wcSociedad() & "] " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
        End Try

    End Sub

End Module
