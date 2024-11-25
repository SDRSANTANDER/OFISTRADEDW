Option Strict On

Imports System.Reflection


Public Class DAOScript
    Inherits clsConexion

    Sub New()
        MyBase.New()
    End Sub

#Region "Script"

    Public Function setEjecutarScript(ByVal sql As String, ByVal sSociedad As String) As Boolean

        'Devuelve creado S/N
        Dim retval As Boolean = False

        Dim oConnection = getConexionTipo()

        Try

            'Abre la conexión 
            oConnection = LogInSQL(sSociedad)

            If Not oConnection Is Nothing AndAlso oConnection.State = ConnectionState.Open Then

                'Actualiza los datos del envío en el pedido SAP
                retval = EjecutarSQL(sql, sSociedad) > 0

            End If

            'Acepta los cambios
            retval = True

        Catch ex As Exception
            clsLog.Log(ex.Message & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
        Finally
            If Not oConnection Is Nothing AndAlso oConnection.State = ConnectionState.Open Then
                oConnection.Close()
            End If
        End Try

        Return retval

    End Function


#End Region

End Class
