Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAORic
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

    Public Function getCentrosCosteRIC() As List(Of EntCentroCosteRIC)

        Dim retVal As New List(Of EntCentroCosteRIC)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaCentrosCosteRIC()

            For Each row As DataRow In DT.Rows

                Dim oCentroCosteRIC As EntCentroCosteRIC = DataRowToEntidadCentroCosteRIC(row)
                retVal.Add(oCentroCosteRIC)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaCentrosCosteRIC() As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""
            SQL = " SELECT * FROM " & putQuotes("RIC_DW_COSTCENTERS") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadCentroCosteRIC(DR As DataRow) As EntCentroCosteRIC

        Dim oCentroCosteRIC As New EntCentroCosteRIC

        With oCentroCosteRIC

            .NumContrato = CInt(DR.Item("Contract").ToString)
            .ProyectoID = DR.Item("Project").ToString
            .ProyectoNombre = DR.Item("Descript").ToString
            .CentroCoste1ID = DR.Item("OcrCode").ToString
            .CentroCoste1Nombre = DR.Item("OcrName").ToString
            .CentroCoste2ID = DR.Item("OcrCode2").ToString
            .CentroCoste2Nombre = DR.Item("OcrName2").ToString
            .CentroCoste3ID = DR.Item("OcrCode3").ToString
            .CentroCoste3Nombre = DR.Item("OcrName3").ToString
            .Generico = DR.Item("Generic").ToString

        End With

        Return oCentroCosteRIC

    End Function

End Class
