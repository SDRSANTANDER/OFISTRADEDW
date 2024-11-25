Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class DAODocumento
    Inherits clsConexion

    Sub New(ByVal Sociedad As Utilidades.eSociedad)
        MyBase.New(Sociedad)
    End Sub

#Region "Abiertos"

    Public Function getDocumentos(ByVal ObjType As Integer, ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntDocumento)

        Dim retVal As New List(Of EntDocumento)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaDocumentos(ObjType, FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oDocumento As EntDocumento = DataRowToEntidadDocumento(row)
                retVal.Add(oDocumento)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaDocumentos(ByVal ObjType As Integer, ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""

            Select Case ObjType

                Case Utilidades.ObjType.OrdenProduccion

                    SQL = "  SELECT " & vbCrLf
                    SQL &= " T0." & putQuotes("ObjType") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocEntry") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocNum") & "," & vbCrLf
                    SQL &= " T1." & putQuotes("DocDate") & "," & vbCrLf
                    SQL &= " T1." & putQuotes("DocStatus") & "," & vbCrLf
                    SQL &= " T1." & putQuotes("DocCur") & "," & vbCrLf
                    SQL &= " T1." & putQuotes("CANCELED") & "," & vbCrLf
                    SQL &= " 'N' As " & putQuotes("TratadoDW") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("CardCode") & "," & vbCrLf
                    SQL &= " T1." & putQuotes("CardName") & "," & vbCrLf
                    SQL &= " T1." & putQuotes("LicTradNum") & "," & vbCrLf
                    SQL &= " T1." & putQuotes("NumAtCard") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("CmpltQty") & " As " & putQuotes("DocTotal") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("CmpltQty") & " As " & putQuotes("DocTotalFC") & "," & vbCrLf
                    SQL &= " 1.0 As " & putQuotes("DocRate") & "," & vbCrLf
                    SQL &= " 0 As " & putQuotes("OwnerCode") & "," & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_01") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_02") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_03") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_04") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_05") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_06") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_07") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_08") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_09") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_10") & " " & vbCrLf
                    SQL &= " FROM " & getTablaDeObjType(ObjType) & " T0 " & getWithNoLock() & vbCrLf
                    SQL &= " INNER JOIN " & putQuotes("ORDR") & " T1 " & getWithNoLock() & " ON T0." & putQuotes("OriginNum") & " = T1." & putQuotes("DocNum") & vbCrLf
                    SQL &= " INNER JOIN " & putQuotes("OPRJ") & " T2 " & getWithNoLock() & " ON T0." & putQuotes("Project") & " = T2." & putQuotes("PrjCode") & vbCrLf
                    SQL &= " WHERE 1=1 " & vbCrLf
                    'SQL &= " And T1." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
                    'SQL &= " AND T1." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf
                    SQL &= " AND COALESCE(T1." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") >= N'" & FechaInicio & "'" & vbCrLf
                    SQL &= " AND COALESCE(T1." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") <= N'" & FechaFin & "'" & vbCrLf
                    SQL &= " ORDER BY T0." & putQuotes("DocNum")

                Case Utilidades.ObjType.SolicitudCompra

                    Dim Tabla As String = getTablaDeObjType(ObjType)

                    SQL = "  SELECT " & vbCrLf
                    SQL &= " T0." & putQuotes("ObjType") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocEntry") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocNum") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocDate") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocStatus") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocCur") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("CANCELED") & "," & vbCrLf
                    SQL &= " COALESCE(T0." & putQuotes("U_SEITRADW") & ",'N') " & " As " & putQuotes("TratadoDW") & "," & vbCrLf
                    SQL &= " T2." & putQuotes("CardCode") & "," & vbCrLf
                    SQL &= " T2." & putQuotes("CardName") & "," & vbCrLf
                    SQL &= " T2." & putQuotes("LicTradNum") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("NumAtCard") & "," & vbCrLf
                    SQL &= " COALESCE(T0." & putQuotes("DocTotal") & ",0) " & " As " & putQuotes("DocTotal") & "," & vbCrLf
                    SQL &= " COALESCE(T0." & putQuotes("DocTotalFC") & ",0) " & " As " & putQuotes("DocTotalFC") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocRate") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("OwnerCode") & "," & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_01") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_02") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_03") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_04") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_05") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_06") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_07") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_08") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_09") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_10") & " " & vbCrLf
                    SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
                    SQL &= " JOIN " & putQuotes(Tabla.Substring(1, 3) & "1") & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
                    SQL &= " LEFT JOIN " & putQuotes("OCRD") & " T2 " & getWithNoLock() & " ON T2." & putQuotes("CardCode") & " = T1." & putQuotes("LineVendor") & vbCrLf
                    SQL &= " WHERE 1=1 " & vbCrLf
                    'SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
                    'SQL &= " AND T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf
                    'SQL &= " AND COALESCE(T0." & putQuotes("U_SEITRADW") & ",N'" & SN.No & "') <> N'" & SN.Si & "'" & vbCrLf
                    SQL &= " AND COALESCE(T0." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") >= N'" & FechaInicio & "'" & vbCrLf
                    SQL &= " AND COALESCE(T0." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") <= N'" & FechaFin & "'" & vbCrLf
                    SQL &= " GROUP BY " & vbCrLf
                    SQL &= " T0." & putQuotes("ObjType") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocEntry") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocNum") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocDate") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocStatus") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("CANCELED") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("U_SEITRADW") & "," & vbCrLf
                    SQL &= " T2." & putQuotes("CardCode") & "," & vbCrLf
                    SQL &= " T2." & putQuotes("CardName") & "," & vbCrLf
                    SQL &= " T2." & putQuotes("LicTradNum") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("NumAtCard") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocTotal") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocTotalFC") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocRate") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("OwnerCode") & " " & vbCrLf
                    SQL &= " ORDER BY T0." & putQuotes("DocNum") & vbCrLf

                Case Else

                    SQL = "  SELECT " & vbCrLf
                    SQL &= " T0." & putQuotes("ObjType") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocEntry") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocNum") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocDate") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocStatus") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocCur") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("CANCELED") & "," & vbCrLf
                    SQL &= " COALESCE(T0." & putQuotes("U_SEITRADW") & ",'N') " & " As " & putQuotes("TratadoDW") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("CardCode") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("CardName") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("LicTradNum") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("NumAtCard") & "," & vbCrLf
                    SQL &= " COALESCE(T0." & putQuotes("DocTotal") & ",0) " & " As " & putQuotes("DocTotal") & "," & vbCrLf
                    SQL &= " COALESCE(T0." & putQuotes("DocTotalFC") & ",0) " & " As " & putQuotes("DocTotalFC") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("DocRate") & "," & vbCrLf
                    SQL &= " T0." & putQuotes("OwnerCode") & "," & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_01") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_02") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_03") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_04") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_05") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_06") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_07") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_08") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_09") & ", " & vbCrLf
                    SQL &= " '' " & " As " & putQuotes("CU_10") & " " & vbCrLf
                    SQL &= " FROM " & getTablaDeObjType(ObjType) & " T0 " & getWithNoLock() & vbCrLf
                    SQL &= " WHERE 1=1 " & vbCrLf
                    'SQL &= " AND T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
                    'SQL &= " AND T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf
                    'SQL &= " AND COALESCE(T0." & putQuotes("U_SEITRADW") & ",N'" & SN.No & "') <> N'" & SN.Si & "'" & vbCrLf
                    SQL &= " AND COALESCE(T0." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") >= N'" & FechaInicio & "'"
                    SQL &= " AND COALESCE(T0." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") <= N'" & FechaFin & "'"
                    SQL &= " ORDER BY T0." & putQuotes("DocNum")

            End Select

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadDocumento(DR As DataRow) As EntDocumento

        Dim oDocumento As New EntDocumento

        With oDocumento

            .ObjectType = DR.Item("ObjType").ToString    'ObjType
            .DocEntry = CInt(DR.Item("DocEntry").ToString)
            .DocNum = CInt(DR.Item("DocNum").ToString)
            .DocDate = CDate(DR.Item("DocDate")).ToString("yyyyMMdd")
            .DocMoneda = DR.Item("DocCur").ToString
            .IdInterlocutor = DR.Item("CardCode").ToString    'CardCode
            .RazonSocial = DR.Item("CardName").ToString     'CardName
            .NIFTercero = DR.Item("LicTradNum").ToString    'LicTradNum
            .NumAtCard = DR.Item("NumAtCard").ToString
            .DocTotal = CDbl(DR.Item("DocTotal").ToString)  'En moneda de documento
            If (CDbl(DR.Item("DocTotalFC").ToString) > 0) Then .DocTotal = CDbl(DR.Item("DocTotalFC").ToString)
            .DocTotalEUR = CDbl(DR.Item("DocTotal").ToString)
            .DocRate = CDbl(DR.Item("DocRate").ToString)
            .DocStatus = DR.Item("DocStatus").ToString
            .Cancelado = DR.Item("CANCELED").ToString
            .TratadoDW = DR.Item("TratadoDW").ToString
			.IdAutor = DR.Item("OwnerCode").ToString
			
			.CU_01 = ""
            .CU_02 = ""
            .CU_03 = ""
            .CU_04 = ""
            .CU_05 = ""
            .CU_06 = ""
            .CU_07 = ""
            .CU_08 = ""
            .CU_09 = ""
            .CU_10 = ""

        End With

        Return oDocumento

    End Function

#End Region

#Region "Extendido"

    Public Function getDocumentosExtendido(ByVal ObjType As Integer, ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntDocumentoExtendido)

        Dim retVal As New List(Of EntDocumentoExtendido)

        Try

            'Los resultados de la query en un datatable
            Dim DTCab As DataTable = getConsultaDocumentosExtendido(ObjType, FechaInicio, FechaFin)
            Dim lLineas As List(Of EntDocumentoExtendidoLin) = getDocumentosExtendidoLin(ObjType, FechaInicio, FechaFin)

            For Each row As DataRow In DTCab.Rows

                Dim oDocumento As EntDocumentoExtendido = DataRowToEntidadDocumentoExtendido(row, lLineas)
                retVal.Add(oDocumento)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaDocumentosExtendido(ByVal ObjType As Integer, ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        'Tablas en donde el CardCode va a nivel de línea
        Dim Tabla = getTablaDeObjType(ObjType)
        Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(Utilidades.ObjType.SolicitudCompra))

        Try

            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("ObjType") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocType") & "," & vbCrLf
            SQL &= " T0." & putQuotes("Series") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocEntry") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocNum") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocStatus") & "," & vbCrLf
            SQL &= " T0." & putQuotes("CANCELED") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocDate") & "," & vbCrLf
            SQL &= " T0." & putQuotes("TaxDate") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocDueDate") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("U_SEITRADW") & ",'N') " & " As " & putQuotes("TratadoDW") & "," & vbCrLf
            SQL &= " T0." & putQuotes("CardCode") & "," & vbCrLf
            SQL &= " COALESCE(T2." & putQuotes("CardName") & ",'') " & " As " & putQuotes("CardName") & "," & vbCrLf
            SQL &= " COALESCE(T2." & putQuotes("LicTradNum") & ",'') " & " As " & putQuotes("LicTradNum") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("CardName") & ",'') " & " As " & putQuotes("DocCardName") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("LicTradNum") & ",'') " & " As " & putQuotes("DocLicTradNum") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("NumAtCard") & ",'') " & " As " & putQuotes("NumAtCard") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DocTotal") & ",0) " & " As " & putQuotes("DocTotal") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DocTotalFC") & ",0) " & " As " & putQuotes("DocTotalFC") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("VatSum") & ",0) " & " As " & putQuotes("VatSum") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("VatSumFC") & ",0) " & " As " & putQuotes("VatSumFC") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DocTotal") & ",0) - COALESCE(T0." & putQuotes("VatSum") & ",0)" & " As " & putQuotes("BaseTotal") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DocTotalFC") & ",0) - COALESCE(T0." & putQuotes("VatSumFC") & ",0)" & " As " & putQuotes("BaseTotalFC") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DiscPrcnt") & ",0) " & " As " & putQuotes("DiscPrcnt") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DiscSum") & ",0) " & " As " & putQuotes("DiscSum") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DiscSumFC") & ",0) " & " As " & putQuotes("DiscSumFC") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DocRate") & ",0) " & " As " & putQuotes("DocRate") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DocCur") & ",'') " & " As " & putQuotes("DocCur") & ", " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("GroupNum") & ",0) " & " As " & putQuotes("GroupNum") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("PeyMethod") & ",'') " & " As " & putQuotes("PeyMethod") & ", " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("SlpCode") & ",0) " & " As " & putQuotes("SlpCode") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("OwnerCode") & ",0) " & " As " & putQuotes("OwnerCode") & ", " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("U_SEIIDDW") & ",'') " & " As " & putQuotes("IDDW") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("U_SEIURLDW") & ",'') " & " As " & putQuotes("URLDW") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("Comments") & ",'') " & " As " & putQuotes("Comments") & " " & vbCrLf

            SQL &= " FROM " & Tabla & " T0 " & getWithNoLock() & vbCrLf

            If Not bICLinea Then
                SQL &= " JOIN " & putQuotes("OCRD") & " T2 " & getWithNoLock() & " ON T2." & putQuotes("CardCode") & " = T0." & putQuotes("CardCode") & vbCrLf
            Else
                SQL &= " JOIN " & putQuotes(Tabla.Substring(1, 3) & "1") & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
                SQL &= " JOIN " & putQuotes("OCRD") & " T2 " & getWithNoLock() & " ON T2." & putQuotes("CardCode") & " = T1." & putQuotes("LineVendor") & vbCrLf
            End If

            SQL &= " WHERE 1=1 " & vbCrLf
            'SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            'SQL &= " AND T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf
            'SQL &= " AND COALESCE(T0." & putQuotes("U_SEITRADW") & ",N'" & SN.No & "') <> N'" & SN.Si & "'" & vbCrLf
            SQL &= " AND COALESCE(T0." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") >= N'" & FechaInicio & "'"
            SQL &= " AND COALESCE(T0." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") <= N'" & FechaFin & "'"
            SQL &= " ORDER BY T0." & putQuotes("DocNum")

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadDocumentoExtendido(DR As DataRow, lLineas As List(Of EntDocumentoExtendidoLin)) As EntDocumentoExtendido

        Dim oDocumento As New EntDocumentoExtendido

        With oDocumento

            .ObjectType = DR.Item("ObjType").ToString
            .DocType = DR.Item("DocType").ToString
            .DocSerie = CInt(DR.Item("Series").ToString)
            .DocEntry = CInt(DR.Item("DocEntry").ToString)
            .DocNum = CInt(DR.Item("DocNum").ToString)
            .DocStatus = DR.Item("DocStatus").ToString
            .Cancelado = DR.Item("CANCELED").ToString
            .DocDate = CDate(DR.Item("DocDate")).ToString("yyyyMMdd")
            .TaxDate = CDate(DR.Item("TaxDate")).ToString("yyyyMMdd")
            .DocDueDate = CDate(DR.Item("DocDueDate")).ToString("yyyyMMdd")
            .TratadoDW = DR.Item("TratadoDW").ToString
            .IdInterlocutor = DR.Item("CardCode").ToString
            .RazonSocial = DR.Item("CardName").ToString
            .NIFTercero = DR.Item("LicTradNum").ToString
            .DocRazonSocial = DR.Item("DocCardName").ToString
            .DocNIFTercero = DR.Item("DocLicTradNum").ToString
            .NumAtCard = DR.Item("NumAtCard").ToString
            .DocTotal = CDbl(DR.Item("DocTotal").ToString)
            .DocTotalFC = CDbl(DR.Item("DocTotalFC").ToString)
            .VatSum = CDbl(DR.Item("VatSum").ToString)
            .VatSumFC = CDbl(DR.Item("VatSumFC").ToString)
            .BaseTotal = CDbl(DR.Item("BaseTotal").ToString)
            .BaseTotalFC = CDbl(DR.Item("BaseTotalFC").ToString)
            .DiscPorc = CDbl(DR.Item("DiscPrcnt").ToString)
            .DiscSum = CDbl(DR.Item("DiscSum").ToString)
            .DiscSumFC = CDbl(DR.Item("DiscSumFC").ToString)
            .DocRate = CDbl(DR.Item("DocRate").ToString)
            .DocMoneda = DR.Item("DocCur").ToString
            .CondicionPago = CInt(DR.Item("GroupNum").ToString)
            .ViaPago = DR.Item("PeyMethod").ToString
            .EncargadoCV = CInt(DR.Item("SlpCode").ToString)
            .Titular = CInt(DR.Item("OwnerCode").ToString)
            .IDDW = DR.Item("IDDW").ToString
            .URLDW = DR.Item("URLDW").ToString
            .Comentarios = DR.Item("Comments").ToString

            'Líneas
            Dim lDocLin As List(Of EntDocumentoExtendidoLin) = (From p In lLineas
                                                                Order By p.LineNum
                                                                Where p.DocEntry = oDocumento.DocEntry).Distinct.ToList

            .Lineas = New List(Of EntDocumentoExtendidoLin)
            If Not lDocLin Is Nothing AndAlso lDocLin.Count > 0 Then .Lineas = lDocLin

        End With

        Return oDocumento

    End Function

    Public Function getDocumentosExtendidoLin(ByVal ObjType As Integer, ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As List(Of EntDocumentoExtendidoLin)

        Dim retVal As New List(Of EntDocumentoExtendidoLin)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaDocumentosExtendidoLin(ObjType, FechaInicio, FechaFin)

            For Each row As DataRow In DT.Rows

                Dim oDocumento As EntDocumentoExtendidoLin = DataRowToEntidadDocumentoExtendidoLin(row)
                retVal.Add(oDocumento)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaDocumentosExtendidoLin(ByVal ObjType As Integer, ByVal FechaInicio As Integer, ByVal FechaFin As Integer) As DataTable

        Dim DT As New DataTable

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim Tabla = getTablaDeObjType(ObjType)
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(Utilidades.ObjType.SolicitudCompra))

            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("ObjType") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocEntry") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocNum") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("NumAtCard") & ",'') " & " As " & putQuotes("NumAtCard") & "," & vbCrLf
            SQL &= " T1." & putQuotes("LineNum") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("BaseType") & ",-1) " & " As " & putQuotes("BaseType") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("BaseEntry") & ",-1) " & " As " & putQuotes("BaseEntry") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("BaseLine") & ",-1) " & " As " & putQuotes("BaseLine") & "," & vbCrLf
            SQL &= " T1." & putQuotes("LineStatus") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("ItemCode") & ",'') " & " As " & putQuotes("ItemCode") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("Dscription") & ",'') " & " As " & putQuotes("Dscription") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("SubCatNum") & ",'') " & " As " & putQuotes("SubCatNum") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("AcctCode") & ",'') " & " As " & putQuotes("AcctCode") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("Quantity") & ",0) " & " As " & putQuotes("Quantity") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("OpenQty") & ",0) " & " As " & putQuotes("OpenQty") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("PriceBefDi") & ",0) " & " As " & putQuotes("PriceBefDi") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("DiscPrcnt") & ",0) " & " As " & putQuotes("DiscPrcnt") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("Price") & ",0) " & " As " & putQuotes("Price") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("LineTotal") & ",0) " & " As " & putQuotes("LineTotal") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("TotalFrgn") & ",0) " & " As " & putQuotes("TotalFrgn") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("OpenSum") & ",0) " & " As " & putQuotes("OpenSum") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("OpenSumFC") & ",0) " & " As " & putQuotes("OpenSumFC") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("VatGroup") & ",'N') " & " As " & putQuotes("VatGroup") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("VatPrcnt") & ",0) " & " As " & putQuotes("VatPrcnt") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("Currency") & ",'') " & " As " & putQuotes("Currency") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("Rate") & ",0) " & " As " & putQuotes("Rate") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("TaxOnly") & ",'N') " & " As " & putQuotes("TaxOnly") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("TaxCode") & ",'') " & " As " & putQuotes("TaxCode") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("WtLiable") & ",'N') " & " As " & putQuotes("WtLiable") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("WhsCode") & ",'') " & " As " & putQuotes("WhsCode") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("Project") & ",'') " & " As " & putQuotes("Project") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("OcrCode") & ",'') " & " As " & putQuotes("OcrCode") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("OcrCode2") & ",'') " & " As " & putQuotes("OcrCode2") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("OcrCode3") & ",'') " & " As " & putQuotes("OcrCode3") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("OcrCode4") & ",'') " & " As " & putQuotes("OcrCode4") & "," & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("OcrCode5") & ",'') " & " As " & putQuotes("OcrCode5") & " " & vbCrLf


            SQL &= " FROM " & Tabla & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & putQuotes(Tabla.Substring(1, 3) & "1") & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
            If Not bICLinea Then
                SQL &= " JOIN " & putQuotes("OCRD") & " T2 " & getWithNoLock() & " ON T2." & putQuotes("CardCode") & " = T0." & putQuotes("CardCode") & vbCrLf
            Else
                SQL &= " JOIN " & putQuotes("OCRD") & " T2 " & getWithNoLock() & " ON T2." & putQuotes("CardCode") & " = T1." & putQuotes("LineVendor") & vbCrLf
            End If

            SQL &= " WHERE 1=1 " & vbCrLf
            'SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            'SQL &= " AND T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf
            'SQL &= " AND COALESCE(T0." & putQuotes("U_SEITRADW") & ",N'" & SN.No & "') <> N'" & SN.Si & "'" & vbCrLf
            SQL &= " AND COALESCE(T0." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") >= N'" & FechaInicio & "'"
            SQL &= " AND COALESCE(T0." & putQuotes("DocDate") & "," & getDefaultDateWithoutTime() & ") <= N'" & FechaFin & "'"
            SQL &= " ORDER BY T0." & putQuotes("DocNum")

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function DataRowToEntidadDocumentoExtendidoLin(DR As DataRow) As EntDocumentoExtendidoLin

        Dim oDocumentoLin As New EntDocumentoExtendidoLin

        With oDocumentoLin

            .ObjectType = DR.Item("ObjType").ToString
            .DocEntry = CInt(DR.Item("DocEntry").ToString)
            .DocNum = CInt(DR.Item("DocNum").ToString)
            .NumAtCard = DR.Item("NumAtCard").ToString

            .LineNum = CInt(DR.Item("LineNum").ToString)
            .BaseType = CInt(DR.Item("BaseType").ToString)
            .BaseEntry = CInt(DR.Item("BaseEntry").ToString)
            .BaseLine = CInt(DR.Item("BaseLine").ToString)

            .LineStatus = DR.Item("LineStatus").ToString
            .Articulo = DR.Item("ItemCode").ToString
            .Descripcion = DR.Item("Dscription").ToString
            .RefExt = DR.Item("SubCatNum").ToString
            .CuentaContable = DR.Item("AcctCode").ToString
            .Cantidad = CDbl(DR.Item("Quantity").ToString)
            .CantidadPte = CDbl(DR.Item("OpenQty").ToString)
            .Precio = CDbl(DR.Item("PriceBefDi").ToString)
            .PorcDto = CDbl(DR.Item("DiscPrcnt").ToString)
            .PrecioFinal = CDbl(DR.Item("Price").ToString)
            .LineTotal = CDbl(DR.Item("LineTotal").ToString)
            .LineTotalFC = CDbl(DR.Item("TotalFrgn").ToString)
            .LineTotalPte = CDbl(DR.Item("OpenSum").ToString)
            .LineTotalPteFC = CDbl(DR.Item("OpenSumFC").ToString)
            .VatGroup = DR.Item("VatGroup").ToString
            .VatPorc = CDbl(DR.Item("VatPrcnt").ToString)
            .Moneda = DR.Item("Currency").ToString
            .Rate = CDbl(DR.Item("Rate").ToString)

            .TaxOnly = DR.Item("TaxOnly").ToString
            .TaxCode = DR.Item("TaxCode").ToString
            .WTLiable = DR.Item("WtLiable").ToString

            .Almacen = DR.Item("WhsCode").ToString
            .Proyecto = DR.Item("Project").ToString
            .Centro1Coste = DR.Item("OcrCode").ToString
            .Centro2Coste = DR.Item("OcrCode2").ToString
            .Centro3Coste = DR.Item("OcrCode3").ToString
            .Centro4Coste = DR.Item("OcrCode4").ToString
            .Centro5Coste = DR.Item("OcrCode5").ToString

        End With

        Return oDocumentoLin

    End Function


#End Region

#Region "Documentos origen"

    Public Function getDocumentosOrigen(ByVal TablaOrigen As String,
                                        ByVal CardCode As String,
                                        ByVal DOCIDDW As String,
                                        ByVal DocNum As String,
                                        ByVal NumAtCard As String) As List(Of EntDocumento)

        Dim retVal As New List(Of EntDocumento)

        Try

            'Los resultados de la query en un datatable
            Dim DT As DataTable = getConsultaDocumentosOrigen(TablaOrigen, CardCode, DOCIDDW, DocNum, NumAtCard)

            For Each row As DataRow In DT.Rows

                Dim oDocumento As EntDocumento = getEntidadDocumentoOrigen(row.Item("BaseType").ToString, CInt(row.Item("BaseEntry").ToString))
                retVal.Add(oDocumento)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getConsultaDocumentosOrigen(ByVal Tabla As String,
                                                ByVal CardCode As String,
                                                ByVal DOCIDDW As String,
                                                ByVal DocNum As String,
                                                ByVal NumAtCard As String) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " T1." & putQuotes("BaseType") & "," & vbCrLf
            SQL &= " T1." & putQuotes("BaseEntry") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " INNER JOIN " & putQuotes(Tabla.Substring(1, 3) & "1") & " T1 " & getWithNoLock() & " ON T0." & putQuotes("DocEntry") & " = T1." & putQuotes("DocEntry") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf
            SQL &= " And T1." & putQuotes("BaseType") & "<>N'-1'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then
                SQL &= " And T0." & putQuotes("U_SEIIDDW") & "=N'" & DOCIDDW & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(DocNum) Then
                SQL &= " And T0." & putQuotes("DocNum") & "=N'" & DocNum & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(NumAtCard) Then
                SQL &= " And T0." & putQuotes("NumAtCard") & "=N'" & NumAtCard & "'" & vbCrLf
            End If

            SQL &= " GROUP BY " & vbCrLf
            SQL &= " T1." & putQuotes("BaseType") & "," & vbCrLf
            SQL &= " T1." & putQuotes("BaseEntry") & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

    Public Function getEntidadDocumentoOrigen(ByVal ObjectType As String, ByVal DocEntry As Integer) As EntDocumento

        Dim retval As New EntDocumento

        Try

            Dim DT As DataTable = getConsultaEntidadDocumentoOrigen(ObjectType, DocEntry)

            If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then
                retval = DataRowToEntidadDocumento(DT.Rows(0))
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retval

    End Function

    Public Function getConsultaEntidadDocumentoOrigen(ByVal ObjectType As String, ByVal DocEntry As Integer) As DataTable

        Dim DT As New DataTable

        Try

            Dim SQL As String = ""

            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("ObjType") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocEntry") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocNum") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocDate") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocStatus") & "," & vbCrLf
            SQL &= " T0." & putQuotes("CANCELED") & "," & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("U_SEITRADW") & ",'N') " & " As " & putQuotes("TratadoDW") & "," & vbCrLf
            SQL &= " T0." & putQuotes("CardCode") & "," & vbCrLf
            SQL &= " T0." & putQuotes("CardName") & "," & vbCrLf
            SQL &= " T0." & putQuotes("LicTradNum") & "," & vbCrLf
            SQL &= " T0." & putQuotes("NumAtCard") & "," & vbCrLf
            SQL &= " T0." & putQuotes("DocTotal") & ","
            SQL &= " T0." & putQuotes("DocCur")
            SQL &= " FROM " & getTablaDeObjType(ObjectType) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " AND T0." & putQuotes("DocEntry") & " = N'" & DocEntry & "'" & vbCrLf

            DT = ObtenerDT(SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return DT

    End Function

#End Region

End Class
