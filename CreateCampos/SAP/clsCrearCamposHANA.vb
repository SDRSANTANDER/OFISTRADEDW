Option Strict Off
Option Explicit On

Imports SAPbobsCOM.BoObjectTypes
Imports SAPbobsCOM.BoUTBTableType
Imports SAPbobsCOM.BoFieldTypes
Imports SAPbobsCOM.BoFldSubTypes
Imports SAPbobsCOM.BoYesNoEnum
Imports SAPbobsCOM.BoUDOObjType
Imports System
'
'
Public MustInherit Class clsCreaCamposHANA
    'comentario

    Private mComp As SAPbobsCOM.Company
    Private mPrefijoHANA As String = ""

    Const cPermisCampUsuari As String = "261"
    Const cPermisUDO As String = "500"

    Public Sub New(ByRef oCompany As SAPbobsCOM.Company)
        mComp = oCompany
        mPrefijoHANA = " " & mComp.CompanyDB & "."

    End Sub



#Region "Funcions Generals"

    Public Sub AddUserKey(ByVal sTableName As String, ByVal sKeyName As String,
                            ByVal iUnica As SAPbobsCOM.BoYesNoEnum, ByVal Array_Fields() As String)
        '
        Dim oClau As SAPbobsCOM.UserKeysMD = Nothing
        Dim iField As Integer
        '
        sTableName = Replace(sTableName, "@", "")
        '
        If Not Me.ExisteixClau(sTableName, sKeyName) Then
            oClau = mComp.GetBusinessObject(oUserKeys)
            With oClau
                .KeyName = sKeyName
                .TableName = sTableName
                .Unique = iUnica
                With .Elements
                    For iField = LBound(Array_Fields) To UBound(Array_Fields)
                        If iField <> LBound(Array_Fields) Then .Add()
                        .ColumnAlias = Replace(Array_Fields(iField), "U_", "")
                    Next
                End With
                If .Add <> 0 Then Throw New Exception(mComp.GetLastErrorDescription)
            End With
        End If
        '
        clsUtil.LiberarObjetoCOM(oClau)
        '
    End Sub

    Public Sub AddUserUDO(ByVal sName As String, ByVal sDescription As String, ByVal sLogTable As String,
                        ByVal iType As SAPbobsCOM.BoUDOObjType, ByVal iCanCancel As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanClose As SAPbobsCOM.BoYesNoEnum, ByVal iDefaultForm As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanDelete As SAPbobsCOM.BoYesNoEnum, ByVal iCanFind As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanLog As SAPbobsCOM.BoYesNoEnum, ByVal iCanYearTransfer As SAPbobsCOM.BoYesNoEnum,
                        ByVal iManageSeries As SAPbobsCOM.BoYesNoEnum, Optional ByVal Array_ChildTables() As String = Nothing)
        '
        Dim oUDO As SAPbobsCOM.UserObjectsMD
        Dim iChild As Integer
        '
        sName = Replace(sName, "@", "")
        '
        oUDO = mComp.GetBusinessObject(oUserObjectsMD)
        If Not oUDO.GetByKey(sName) Then
            With oUDO
                .CanCancel = iCanCancel
                .CanClose = iCanClose
                .CanCreateDefaultForm = iDefaultForm
                .CanDelete = iCanDelete
                .CanFind = iCanFind
                .CanLog = iCanLog
                .LogTableName = sLogTable
                .CanYearTransfer = iCanYearTransfer
                .Code = sName
                .TableName = sName
                .ObjectType = iType
                .Name = sDescription
                .ManageSeries = iManageSeries
                If Not IsNothing(Array_ChildTables) Then
                    With .ChildTables
                        For iChild = LBound(Array_ChildTables) To UBound(Array_ChildTables)
                            If iChild <> LBound(Array_ChildTables) Then .Add()
                            .TableName = Array_ChildTables(iChild)
                        Next
                    End With
                End If
                If .Add <> 0 Then Throw New Exception(mComp.GetLastErrorDescription)
            End With
        End If
        '
        clsUtil.LiberarObjetoCOM(oUDO)
        '
    End Sub

    Public Sub AddUserField(ByVal sTable As String, ByVal sField As String, ByVal sDescription As String,
                            ByVal iType As SAPbobsCOM.BoFieldTypes, ByVal iSubType As SAPbobsCOM.BoFldSubTypes,
                            ByVal iSize As Integer, ByVal sLinkedTable As String,
                            ByVal iMandatory As SAPbobsCOM.BoYesNoEnum, ByVal sDefaultValue As String,
                            Optional ByVal aCodesValidValues() As String = Nothing,
                            Optional ByVal aNamesValidValues() As String = Nothing)
        '
        Dim oCamp As SAPbobsCOM.UserFieldsMD
        Dim i As Long
        '
        If Not ExisteixCamp(sTable, sField) Then
            oCamp = mComp.GetBusinessObject(oUserFields)
            With oCamp
                .TableName = sTable
                .Name = sField
                .Description = sDescription
                .Type = iType
                If iSubType <> st_None Then .SubType = iSubType
                If iMandatory = tYES Then
                    If sDefaultValue = "" Then
                        Throw New Exception("Debe definirse un valor por defecto para el campo '" & sField & "-" & sDescription & "'")
                        Exit Sub
                    Else
                        .Mandatory = iMandatory
                    End If
                End If
                If iSize <> 0 Then .EditSize = iSize
                If sLinkedTable <> "" Then .LinkedTable = sLinkedTable
                If sDefaultValue <> "" Then .DefaultValue = sDefaultValue
                If Not IsNothing(aCodesValidValues) Then
                    With .ValidValues
                        For i = LBound(aCodesValidValues) To UBound(aCodesValidValues)
                            If i <> LBound(aCodesValidValues) Then .Add()
                            .Value = aCodesValidValues(i)
                            .Description = aNamesValidValues(i)
                        Next
                    End With
                End If
                If .Add <> 0 Then Throw New Exception(mComp.GetLastErrorDescription)
            End With
            '
            clsUtil.LiberarObjetoCOM(oCamp)
            '
        End If
        '
    End Sub

    Public Sub AddUserTable(ByVal sName As String, ByVal sDescription As String, ByVal iType As SAPbobsCOM.BoUTBTableType)
        '
        Dim oTaula As SAPbobsCOM.UserTablesMD
        '
        sName = Replace(sName, "@", "")
        '
        oTaula = mComp.GetBusinessObject(oUserTables)
        If Not oTaula.GetByKey(sName) Then
            With oTaula
                .TableName = sName
                .TableDescription = sDescription
                .TableType = iType
                If .Add <> 0 Then
                    Throw New Exception(mComp.GetLastErrorDescription)
                End If
            End With
        End If
        '
        clsUtil.LiberarObjetoCOM(oTaula)
        '
    End Sub

    Private Function ExisteixCamp(ByVal sTableName As String, ByVal sFieldName As String) As Boolean
        '
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        '
        oTmpRecordset = mComp.GetBusinessObject(BoRecordset)
        If Mid(sTableName, 1, 1) <> "@" And Len(sTableName) > 4 Then
            sTableName = "@" & sTableName
        End If
        '

        oTmpRecordset.DoQuery("Select Count(*) From " & mPrefijoHANA & "CUFD Where ""TableID"" = '" & sTableName & "' And ""AliasID"" = '" & sFieldName & "'")

        '
        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            ExisteixCamp = True
        Else
            ExisteixCamp = False
        End If
        '
        clsUtil.LiberarObjetoCOM(oTmpRecordset)
        '
    End Function

    Private Function ExisteixClau(ByVal sTableName As String, ByVal sKeyName As String) As Boolean
        '
        Dim ls As String
        Dim oRecordset As SAPbobsCOM.Recordset = Nothing
        '
        If Left(sTableName, 1) <> "@" And Len(sTableName) > 4 Then sTableName = "@" & sTableName
        '

        ls = "SELECT TOP 1 KeyId "
        ls = ls & " FROM " & mPrefijoHANA & "OUKD "
        ls = ls & " WHERE ""TableName"" = '" & sTableName & "' "
        ls = ls & " AND ""KeyName"" = '" & sKeyName & "' "



        clsUtil.LiberarObjetoCOM(oRecordset)
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)
        If oRecordset.EoF Then
            ExisteixClau = False
        Else
            ExisteixClau = True
        End If
        '
        clsUtil.LiberarObjetoCOM(oRecordset)
        '
    End Function

    Private Function IDCamp(ByVal sTableName As String, ByVal sFieldName As String) As Integer
        '
        Dim ls As String
        Dim oRecordset As SAPbobsCOM.Recordset = Nothing
        '
        If Mid(sTableName, 1, 1) <> "@" And Len(sTableName) > 4 Then
            sTableName = "@" & sTableName
        End If
        '

        ls = "SELECT TOP 1 FieldId "
        ls = ls & " FROM " & mPrefijoHANA & "CUFD "
        ls = ls & " WHERE ""TableID"" = '" & sTableName & "' "
        ls = ls & " AND ""AliasId"" = '" & sFieldName & "' "



        clsUtil.LiberarObjetoCOM(oRecordset)
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)
        If oRecordset.EoF Then
            IDCamp = -1
        Else
            IDCamp = oRecordset.Fields.Item("FieldId").Value
        End If
        '
        clsUtil.LiberarObjetoCOM(oRecordset)
        '
    End Function

    Private Sub AddUserValidValue(ByVal sTable As String, ByVal sField As String, ByVal sValue As String, ByVal sDescription As String)
        '
        Dim ls As String
        Dim oRecordset As SAPbobsCOM.Recordset = Nothing
        Dim iField As Integer
        Dim iIndex As Integer
        '
        If Left(sTable, 1) <> "@" And Len(sTable) > 4 Then sTable = "@" & sTable
        '
        iField = Me.IDCamp(sTable, sField)
        If iField < 0 Then Exit Sub

        ls = "SELECT TOP 1 IndexId "
        ls = ls & " FROM " & mPrefijoHANA & "UFD1 "
        ls = ls & " WHERE ""TableID"" = '" & sTable & "' "
        ls = ls & " AND ""FieldId"" = " & iField.ToString
        ls = ls & " AND ""FldValue"" = '" & sValue.Replace("'", "''").Trim & "' "


        '

        clsUtil.LiberarObjetoCOM(oRecordset)
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)
        If oRecordset.EoF Then


            ls = "SELECT MAX(IndexId) AS Valor "
            ls = ls & " FROM " & mPrefijoHANA & "UFD1 "
            ls = ls & " WHERE ""TableID"" = '" & sTable & "' "
            ls = ls & " AND ""FieldId"" = " & iField.ToString



            clsUtil.LiberarObjetoCOM(oRecordset)
            oRecordset = mComp.GetBusinessObject(BoRecordset)
            oRecordset.DoQuery(ls)
            If oRecordset.EoF Then
                iIndex = 0
            Else
                iIndex = oRecordset.Fields.Item("Valor").Value + 1
            End If


            ls = "INSERT INTO " & mPrefijoHANA & "UFD1 (TableID, FieldID, IndexID, FldValue, Descr) VALUES ("
            ls = ls & " '" & sTable & "', "
            ls = ls & iField.ToString & ", "
            ls = ls & iIndex.ToString & ", "
            ls = ls & " '" & sValue.Replace("'", "''").Trim & "', "
            ls = ls & " '" & sDescription.Replace("'", "''").Trim & "') "



        Else


            iIndex = oRecordset.Fields.Item("IndexId").Value
            ls = "UPDATE " & mPrefijoHANA & "UFD1 "
            ls = ls & " SET Descr = '" & sDescription.Replace("'", "''").Trim & "' "
            ls = ls & " WHERE ""TableID"" = '" & sTable & "' "
            ls = ls & " AND ""FieldID"" = " & iField.ToString
            ls = ls & " AND ""IndexID"" = " & iIndex.ToString
            ls = ls & " AND ""FldValue"" = '" & sValue.Replace("'", "''").Trim & "' "



        End If
        clsUtil.LiberarObjetoCOM(oRecordset)
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)
        '
        clsUtil.LiberarObjetoCOM(oRecordset)
        '
    End Sub

    Private Sub EliminarCamp(ByVal sTable As String, ByVal sField As String)
        '
        Dim oCamp As SAPbobsCOM.UserFieldsMD = Nothing
        Dim iCamp As Integer
        '
        iCamp = Me.IDCamp(sTable, sField)
        If iCamp >= 0 Then
            clsUtil.LiberarObjetoCOM(oCamp)
            oCamp = mComp.GetBusinessObject(oUserFields)
            If oCamp.GetByKey(sTable, iCamp) Then
                If oCamp.Remove <> 0 Then
                    Throw New Exception(mComp.GetLastErrorDescription)
                End If
            End If
        End If
        '
        clsUtil.LiberarObjetoCOM(oCamp)
        '
    End Sub

#End Region

#Region "Contadores"
    '
    Public Sub CrearContador(ByVal sCode As String, ByVal sName As String, Optional ByVal lNum As Long = 0, Optional ByVal sDesc As String = "")
        '
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim ls As String

        ls = ""
        ls = ls & "SELECT Code, Name, U_SEIconta"
        ls = ls & " FROM " & mPrefijoHANA & "[@SEICONTA] "
        ls = ls & " WHERE ""Code"" = '" & sCode & "'"


        '
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)

        'introduce el valor predeterminado en el registro seleccionado en la otra instrucción
        If oRecordset.EoF Then
            '


            ls = ""
            ls = ls & "INSERT INTO " & mPrefijoHANA & "[@SEICONTA] (Code, Name, U_SEIconta,U_SEIdescr)"
            ls = ls & " VALUES ( "
            ls = ls & "'" & sCode & "',"
            ls = ls & "'" & sName & "',"
            ls = ls & lNum.ToString & ","
            ls = ls & "'" & sDesc & "'"
            ls = ls & ")"



            '
            oRecordset = mComp.GetBusinessObject(BoRecordset)
            oRecordset.DoQuery(ls)
            '
        End If
        '
        clsUtil.LiberarObjetoCOM(oRecordset)
        '
    End Sub

    Public Sub CrearConstante(ByVal sCode As String, ByVal sName As String, ByVal sValorInicial As String)
        '
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim ls2 As String
        Dim ls As String


        ls = ""
        ls = ls & "SELECT Code, Name, U_SEIconta FROM " & mPrefijoHANA & "[@SEICONTA] "
        ls = ls & " WHERE ""Code"" = '" & sCode & "'"



        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)

        'introduce el valor predeterminado en el registro seleccionado en la otra instrucción
        If oRecordset.EoF Then
            '


            ls2 = ""
            ls2 = ls2 & "INSERT INTO " & mPrefijoHANA & "[@SEICONTA] (Code, Name, U_SEIdescr, U_SEIconta)"
            ls2 = ls2 & " VALUES ( "
            ls2 = ls2 & "'" & sCode & "',"
            ls2 = ls2 & "'" & sName & "',"
            ls2 = ls2 & "'" & sValorInicial & "',"
            ls2 = ls2 & "0)"


            '
            oRecordset = mComp.GetBusinessObject(BoRecordset)
            oRecordset.DoQuery(ls2)
            '
        End If
        '
        clsUtil.LiberarObjetoCOM(oRecordset)
        '
    End Sub


#End Region

End Class
