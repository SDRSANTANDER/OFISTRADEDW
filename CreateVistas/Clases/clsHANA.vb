Imports System.Reflection
Imports System.Data.Odbc
Imports System.Configuration

Public Class clsHANA
    'Public Class clsConexion

    Public Shared Function LogInSQL(ByVal CompanyDB As String) As OdbcConnection

            Try

                Dim oConnection As OdbcConnection = New OdbcConnection

                'HANA
                ' Para apps de 32 bits sería HDBODBC32
                ' Para apps de 64 bits sería HDBODBC
                oConnection.ConnectionString = ConfigurationManager.ConnectionStrings.Item("Conexion").ConnectionString
                oConnection.Open()

                If oConnection.State = ConnectionState.Open Then
                    Return oConnection
                Else
                    Throw New Exception("No se puede abrir la base de datos.")
                End If

            Catch ex As Exception
            clsLog.Log(ex.Message & " en [" & CompanyDB & "] " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            Throw New Exception(ex.Message)
        End Try
    End Function

#Region "Sin transacción"

    Public Shared Function ObtenerDT(ByVal SQL As String, ByVal SociedadSAP As String) As DataTable

        Dim retVal As New DataSet("DS")
        Dim oAdapter As OdbcDataAdapter
        Dim oConnection As OdbcConnection = Nothing

        Try

            oConnection = LogInSQL(SociedadSAP)

            oAdapter = New OdbcDataAdapter(SQL, oConnection)
            oAdapter.Fill(retVal)

        Catch ex As Exception
            clsLog.Log(ex.Message & " en [" & SociedadSAP & "] " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            Return Nothing
        Finally
            oAdapter = Nothing
            If Not oConnection Is Nothing Then
                oConnection.Close()
            End If
        End Try

        If retVal.Tables.Count > 0 Then
            Return retVal.Tables(0)
        Else
            Return Nothing
        End If

    End Function

    Public Shared Function ObtenerDS(ByVal SQL As String, ByVal NombreDS As String, SociedadSAP As String) As DataSet

        Dim retVal As New DataSet(NombreDS)
        Dim oAdapter As OdbcDataAdapter
        Dim oConnection As OdbcConnection = Nothing

        Try

            oConnection = LogInSQL(SociedadSAP)
            oAdapter = New OdbcDataAdapter(SQL, oConnection)
            oAdapter.Fill(retVal)

        Catch ex As Exception
            clsLog.Log(ex.Message & " (SQL:" & SQL & ") en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isError)
            Return Nothing
        Finally
            oAdapter = Nothing
            If Not oConnection Is Nothing Then
                oConnection.Close()
            End If
        End Try

        Return retVal

    End Function

    Friend Function ObtenerOBJ(ByVal SQL As String, ByVal SociedadSAP As String) As Object

        Dim retVal As Object = Nothing
        Dim oCommand As OdbcCommand
        Dim oConnection As OdbcConnection = Nothing

        Try

            oConnection = LogInSQL(SociedadSAP)

            oCommand = oConnection.CreateCommand
            oCommand.Parameters.Clear()
            oCommand.CommandText = SQL

            retVal = oCommand.ExecuteScalar

        Catch ex As Exception
            clsLog.Log(ex.Message & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            clsLog.Log("SQL:" & SQL & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            Throw ex
        Finally
            oCommand = Nothing
            If Not oConnection Is Nothing Then
                oConnection.Close()
            End If
        End Try

        Return retVal

    End Function

    Friend Function EjecutarSQL(ByVal sql As String, ByVal SociedadSAP As String) As Integer

        Dim retVal As Integer = 0
        Dim oCommand As OdbcCommand
        Dim oConnection As OdbcConnection = Nothing

        Try

            If Not String.IsNullOrEmpty(sql) Then

                oConnection = LogInSQL(SociedadSAP)

                oCommand = New OdbcCommand(sql, oConnection)
                oCommand.Parameters.Clear()
                oCommand.CommandText = sql
                oCommand.CommandTimeout = 300

                retVal = oCommand.ExecuteNonQuery()

            Else

                retVal = 1

            End If

        Catch ex As Exception
            clsLog.Log(ex.Message & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            clsLog.Log("SQL:" & sql & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            Throw ex
        Finally
            oCommand = Nothing
            If Not oConnection Is Nothing Then
                oConnection.Close()
            End If
        End Try

        Return retVal

    End Function

#End Region

#Region "Con transacción"

    Public Shared Function getConexionTipo() As OdbcConnection

        Dim oConnection As OdbcConnection = Nothing
        Return oConnection

    End Function

    Public Shared Function getTransaccionTipo() As OdbcTransaction

        Dim oTransaction As OdbcTransaction = Nothing
        Return oTransaction

    End Function

    Public Shared Function AbrirTransacion(oConnection As OdbcConnection, SociedadSAP As String) As OdbcTransaction

        Dim oTransaction As OdbcTransaction = Nothing

        Try

            oTransaction = oConnection.BeginTransaction(IsolationLevel.ReadUncommitted)

        Catch ex As Exception
            clsLog.Log(ex.Message & " en [" & SociedadSAP & "] " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            Throw New Exception(ex.Message)
        End Try

        Return oTransaction

    End Function

    Public Shared Function ObtenerDT(ByVal oConnection As OdbcConnection, ByVal oTransaction As OdbcTransaction, ByVal SQL As String) As DataTable

        Dim retVal As New DataSet("DS")
        Dim oAdapter As OdbcDataAdapter

        Try

            oAdapter = New OdbcDataAdapter(SQL, oConnection)
            oAdapter.SelectCommand.Transaction = oTransaction
            oAdapter.Fill(retVal)

        Catch ex As Exception
            clsLog.Log(ex.Message & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            clsLog.Log("SQL:" & SQL & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            Return Nothing
        Finally
            oAdapter = Nothing
        End Try

        If retVal.Tables.Count > 0 Then
            Return retVal.Tables(0)
        Else
            Return Nothing
        End If

    End Function

    Friend Function ObtenerOBJ(ByVal oConnection As OdbcConnection, ByVal oTransaction As OdbcTransaction, ByVal SQL As String) As Object

        Dim retVal As Object = Nothing
        Dim oCommand As OdbcCommand

        Try

            oCommand = oConnection.CreateCommand
            oCommand.Parameters.Clear()
            oCommand.Transaction = oTransaction
            oCommand.CommandText = SQL

            retVal = oCommand.ExecuteScalar

        Catch ex As Exception
            clsLog.Log(ex.Message & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            clsLog.Log("SQL:" & SQL & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            Throw ex
        Finally
            oCommand = Nothing
        End Try

        Return retVal

    End Function

    Friend Function EjecutarSQL(ByVal oConnection As OdbcConnection, ByVal oTransaction As OdbcTransaction, ByVal SQL As String) As Integer

        Dim retVal As Integer = 0
        Dim oCommand As OdbcCommand

        Try

            If Not String.IsNullOrEmpty(SQL) Then

                oCommand = New OdbcCommand(SQL, oConnection)
                oCommand.Parameters.Clear()
                oCommand.Transaction = oTransaction
                oCommand.CommandText = SQL
                oCommand.CommandTimeout = 300

                retVal = oCommand.ExecuteNonQuery()

            Else

                retVal = 1

            End If

        Catch ex As Exception
            clsLog.Log(ex.Message & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            clsLog.Log("SQL:" & SQL & " en " & MethodBase.GetCurrentMethod().Name, ex, clsLog.eLevel.isFatal)
            Throw ex
            Finally
                oCommand = Nothing
            End Try

            Return retVal

        End Function

#End Region

        Public Shared Function getDataBaseRef(ByVal TABName As String, ByVal CompanyDB As String) As String

            Dim dataBaseName As String = putQuotes(CompanyDB)
            Dim quotedTabName As String = putQuotes(TABName)

            Return dataBaseName + "." + quotedTabName

        End Function

        Public Shared Function putQuotes(ByVal string2Quote As String, Optional ByVal bToUpper As Boolean = True) As String

            If bToUpper Then
                Return ControlChars.Quote + string2Quote.ToUpper + ControlChars.Quote
            Else
                Return ControlChars.Quote + string2Quote + ControlChars.Quote
            End If

        End Function

        Public Shared Function getEmptyDate() As String

            Return "TO_DATE('1990-01-01')"

        End Function

        Public Shared Function getCurrentDate() As String

            Return "CURRENT_DATE"

        End Function

    End Class
