Imports System.Reflection
Imports System.Data.Odbc
Imports System.Data.SqlClient

Public Class clsConexion

    Private ReadOnly _conexionSQL As SqlConnection
    Private ReadOnly _conexionODBC As OdbcConnection
    Private ReadOnly _NombreBBDD As String

    Sub New(ByVal Sociedad As Utilidades.eSociedad)

        Dim SociedadNombre As String = Utilidades.NOMBRESOCIEDAD(Sociedad)

        Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

            Case Utilidades.DBTypeHANA
                'HANA
                _conexionODBC = New OdbcConnection()
                _conexionODBC.ConnectionString = ConfigurationManager.ConnectionStrings.Item("Conexion_" & SociedadNombre).ConnectionString
                _NombreBBDD = ConfigurationManager.AppSettings.Item("BD_" & SociedadNombre).ToString

            Case Else
                'SQL
                _conexionSQL = New SqlConnection()
                _conexionSQL.ConnectionString = ConfigurationManager.ConnectionStrings.Item("Conexion_" & SociedadNombre).ConnectionString
                _NombreBBDD = ConfigurationManager.AppSettings.Item("BD_" & SociedadNombre).ToString

        End Select

    End Sub

#Region "Propiedades"

    Public ReadOnly Property ConexionSQL As SqlConnection
        Get
            Return _conexionSQL
        End Get
    End Property

    Public ReadOnly Property ConexionODBC As OdbcConnection
        Get
            Return _conexionODBC
        End Get
    End Property

#End Region

#Region "Abrir conexión"

    Public Sub AbrirConexion()

        Try

            Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

                Case Utilidades.DBTypeHANA
                    'HANA
                    AbrirConexionODBC()

                Case Else
                    'SQL
                    AbrirConexionSQL()

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible abrir conexion")
        End Try

    End Sub

    Public Sub AbrirConexionSQL()

        Try

            If ConexionSQL.State <> ConnectionState.Open Then
                ConexionSQL.Open()
            End If

            If ConexionSQL Is Nothing Then
                Throw New Exception("Error. conexion nula!")
            End If

            If ConexionSQL.State <> ConnectionState.Open Then
                Throw New Exception("Error. conexion cerrada!")
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible abrir conexion")
        End Try

    End Sub

    Public Sub AbrirConexionODBC()

        Try

            If ConexionODBC.State <> ConnectionState.Open Then
                ConexionODBC.Open()
            End If

            If ConexionODBC Is Nothing Then
                Throw New Exception("Error. conexion nula!")
            End If

            If ConexionODBC.State <> ConnectionState.Open Then
                Throw New Exception("Error. conexion cerrada!")
            End If

            EstablecerEsquemaODBC()

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible abrir conexion")
        End Try

    End Sub

#End Region

#Region "Cerrar conexión"

    Public Sub CerrarConexion()

        Try

            Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

                Case Utilidades.DBTypeHANA
                    'HANA
                    CerrarConexionODBC()

                Case Else
                    'SQL
                    CerrarConexionSQL()

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible abrir conexion")
        End Try

    End Sub

    Public Sub CerrarConexionSQL()

        Try

            If Not IsNothing(ConexionSQL) Then
                If ConexionSQL.State = ConnectionState.Open Then
                    ConexionSQL.Close()
                End If
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible cerrar conexion")
        End Try

    End Sub

    Public Sub CerrarConexionODBC()

        Try

            If Not IsNothing(ConexionODBC) Then
                If ConexionODBC.State = ConnectionState.Open Then
                    ConexionODBC.Close()
                End If
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible cerrar conexion")
        End Try

    End Sub

#End Region

#Region "Establecer esquema"

    Private Sub EstablecerEsquemaODBC()

        Try

            Dim SQL As String = ""
            SQL = "SET SCHEMA " & _NombreBBDD

            Dim oCommand As OdbcCommand
            oCommand = ConexionODBC.CreateCommand
            oCommand.CommandText = SQL
            oCommand.ExecuteNonQuery()

        Catch ex As Exception
            clsLog.Log.Error(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible establecer SCHEMA")
        End Try

    End Sub

#End Region

#Region "Ejecutar consultas"

    Public Function ExecuteScalar(ByVal SQL As String) As Object

        Try

            Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

                Case Utilidades.DBTypeHANA
                    'HANA
                    Return ExecuteScalarODBC(SQL)

                Case Else
                    'SQL
                    Return ExecuteScalarSQL(SQL)

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        End Try

    End Function

    Public Function ExecuteScalarSQL(ByVal SQL As String) As Object

        Try

            AbrirConexion()

            Dim oCommand As SqlCommand
            oCommand = _conexionSQL.CreateCommand()
            oCommand.CommandText = SQL

            Return oCommand.ExecuteScalar

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        Finally
            CerrarConexion()
        End Try

    End Function

    Public Function ExecuteScalarODBC(ByVal SQL As String) As Object

        Try

            AbrirConexion()

            Dim oCommand As OdbcCommand
            oCommand = _conexionODBC.CreateCommand()
            oCommand.CommandText = SQL

            Return oCommand.ExecuteScalar

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        Finally
            CerrarConexion()
        End Try

    End Function

    Public Function ExecuteQuery(ByVal SQL As String) As Object

        Try

            Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

                Case Utilidades.DBTypeHANA
                    'HANA
                    Return ExecuteQueryODBC(SQL)

                Case Else
                    'SQL
                    Return ExecuteQuerySQL(SQL)

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        End Try

    End Function

    Friend Function ExecuteQuerySQL(ByVal SQL As String) As Integer

        Dim retVal As Integer = 0

        Try

            AbrirConexion()

            Dim oCommand As SqlCommand
            oCommand = _conexionSQL.CreateCommand()
            oCommand.CommandText = SQL
            oCommand.CommandTimeout = 300

            retVal = oCommand.ExecuteNonQuery()

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        Finally
            CerrarConexion()
        End Try

        Return retVal

    End Function

    Friend Function ExecuteQueryODBC(ByVal SQL As String) As Integer

        Dim retVal As Integer = 0

        Try

            AbrirConexion()

            Dim oCommand As OdbcCommand
            oCommand = _conexionODBC.CreateCommand()
            oCommand.CommandText = SQL
            oCommand.CommandTimeout = 300

            retVal = oCommand.ExecuteNonQuery()

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        Finally
            CerrarConexion()
        End Try

        Return retVal

    End Function

#End Region

#Region "Obtener datos"

    Public Function ObtenerDT(ByVal SQL As String) As DataTable

        Dim DS As New DataSet("DS")
        Dim retVal As DataTable = Nothing

        Try

            Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

                Case Utilidades.DBTypeHANA
                    'HANA
                    DS = ObtenerDTODBC(SQL)

                Case Else
                    'SQL
                    DS = ObtenerDTSQL(SQL)

            End Select

            If DS.Tables.Count > 0 Then
                retVal = DS.Tables(0)
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        End Try

        Return retVal

    End Function

    Friend Function ObtenerDTSQL(ByVal SQL As String) As DataSet

        Dim retVal As New DataSet("DS")

        Try

            AbrirConexion()

            Dim oAdapter As SqlDataAdapter
            oAdapter = New SqlDataAdapter(SQL, _conexionSQL)
            oAdapter.Fill(retVal)

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        Finally
            CerrarConexion()
        End Try

        Return retVal

    End Function

    Friend Function ObtenerDTODBC(ByVal SQL As String) As DataSet

        Dim retVal As New DataSet("DS")

        Try

            AbrirConexion()

            Dim oAdapter As OdbcDataAdapter
            oAdapter = New OdbcDataAdapter(SQL, _conexionODBC)
            oAdapter.Fill(retVal)

        Catch ex As Exception
            clsLog.Log.Fatal(SQL)
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception("Imposible ejecutar consulta")
        Finally
            CerrarConexion()
        End Try

        Return retVal

    End Function

#End Region

End Class



