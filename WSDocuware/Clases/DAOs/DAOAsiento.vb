Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOAsiento
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

#Region "Asientos modelos"

    Public Function getAsientosModelo() As List(Of EntAsientoModelo)

        Dim retVal As New List(Of EntAsientoModelo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaAsientosModelo()

            For Each row As DataRow In DT.Rows

                Dim oAsientoModelo As EntAsientoModelo = DataRowToEntidadAsientoModelo(row)
                retVal.Add(oAsientoModelo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaAsientosModelo() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ASIENTOS_MODELOS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadAsientoModelo(DR As DataRow) As EntAsientoModelo

        Dim oAsientoModelo As New EntAsientoModelo

        With oAsientoModelo

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oAsientoModelo

    End Function

#End Region

#Region "Asientos indicadores"

    Public Function getAsientosIndicador() As List(Of EntAsientoIndicador)

        Dim retVal As New List(Of EntAsientoIndicador)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaAsientosIndicador()

            For Each row As DataRow In DT.Rows

                Dim oAsientoIndicador As EntAsientoIndicador = DataRowToEntidadAsientoIndicador(row)
                retVal.Add(oAsientoIndicador)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaAsientosIndicador() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ASIENTOS_INDICADORES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadAsientoIndicador(DR As DataRow) As EntAsientoIndicador

        Dim oAsientoIndicador As New EntAsientoIndicador

        With oAsientoIndicador

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oAsientoIndicador

    End Function

#End Region

#Region "Asientos códigos operaciones"

    Public Function getAsientosOperacion() As List(Of EntAsientoOperacion)

        Dim retVal As New List(Of EntAsientoOperacion)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaAsientosOperacion()

            For Each row As DataRow In DT.Rows

                Dim oAsientoOperacion As EntAsientoOperacion = DataRowToEntidadAsientoOperacion(row)
                retVal.Add(oAsientoOperacion)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaAsientosOperacion() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("SEI_VIEW_DW_ASIENTOS_OPERACIONES") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadAsientoOperacion(DR As DataRow) As EntAsientoOperacion

        Dim oAsientoOperacion As New EntAsientoOperacion

        With oAsientoOperacion

            .ID = DR.Item("ID").ToString
            .Nombre = DR.Item("NOMBRE").ToString

        End With

        Return oAsientoOperacion

    End Function

#End Region

End Class
