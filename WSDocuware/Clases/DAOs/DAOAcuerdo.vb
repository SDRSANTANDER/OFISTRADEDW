Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAOAcuerdo
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

#Region "Acuerdo"

    Public Function getAcuerdos(ByVal ObjType As Integer, ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntAcuerdo)

        Dim retVal As New List(Of EntAcuerdo)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaAcuerdos(ObjType, FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oAcuerdo As EntAcuerdo = DataRowToEntidadAcuerdo(row)
                retVal.Add(oAcuerdo)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaAcuerdos(ByVal ObjType As Integer, ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""

            Dim Tabla As String = getTablaDeObjType(ObjType)

            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("ObjType") & "," & vbCrLf
            SQL &= " T0." & putQuotes("AbsID") & "," & vbCrLf
            SQL &= " T0." & putQuotes("BpCode") & "," & vbCrLf
            SQL &= " T0." & putQuotes("BpName") & "," & vbCrLf
            SQL &= " T1." & putQuotes("LicTradNum") & "," & vbCrLf
            SQL &= " T0." & putQuotes("NumAtCard") & "," & vbCrLf
            SQL &= " T0." & putQuotes("Number") & "," & vbCrLf
            SQL &= " T0." & putQuotes("Project") & "," & vbCrLf
            SQL &= " T0." & putQuotes("StartDate") & "," & vbCrLf
            SQL &= " T0." & putQuotes("EndDate") & "," & vbCrLf
            SQL &= " T0." & putQuotes("SignDate") & "," & vbCrLf
            SQL &= " T0." & putQuotes("Descript") & "," & vbCrLf
            SQL &= " T0." & putQuotes("Type") & "," & vbCrLf
            SQL &= " T0." & putQuotes("Method") & "," & vbCrLf
            SQL &= " T0." & putQuotes("Status") & "," & vbCrLf
            SQL &= " T0." & putQuotes("Cancelled") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & putQuotes("OCRD") & " T1 " & getWithNoLock() & " ON T1." & putQuotes("CardCode") & " = T0." & putQuotes("BpCode") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            'SQL &= " AND T0." & putQuotes("Cancelled") & " = N'" & SN.No & "'" & vbCrLf
            SQL &= " AND COALESCE(T0." & putQuotes("StartDate") & "," & getDefaultDateWithoutTime() & ") >= N'" & FechaInicio & "'"
            SQL &= " AND COALESCE(T0." & putQuotes("StartDate") & "," & getDefaultDateWithoutTime() & ") <= N'" & FechaFin & "'"
            SQL &= " ORDER BY T0." & putQuotes("AbsID")

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadAcuerdo(DR As DataRow) As EntAcuerdo

        Dim oAcuerdo As New EntAcuerdo

        With oAcuerdo

            .ObjectType = DR.Item("ObjType").ToString
            .AbsID = CInt(DR.Item("AbsID").ToString)
            .IdInterlocutor = DR.Item("BpCode").ToString
            .RazonSocial = DR.Item("BpName").ToString
            .NIFTercero = DR.Item("LicTradNum").ToString
            .NumAtCard = DR.Item("NumAtCard").ToString
            .Numero = CInt(DR.Item("Number").ToString)
            .Proyecto = DR.Item("Project").ToString
            .FechaInicio = CDate(DR.Item("StartDate")).ToString("yyyyMMdd")
            .FechaFin = CDate(DR.Item("EndDate")).ToString("yyyyMMdd")
            .FechaFirma = CDate(DR.Item("SignDate")).ToString("yyyyMMdd")
            .Descripcion = DR.Item("Descript").ToString
            .Tipo = DR.Item("Type").ToString
            .Metodo = DR.Item("Method").ToString
            .Status = DR.Item("Status").ToString
            .Cancelado = DR.Item("Cancelled").ToString

        End With

        Return oAcuerdo

    End Function

#End Region

End Class
