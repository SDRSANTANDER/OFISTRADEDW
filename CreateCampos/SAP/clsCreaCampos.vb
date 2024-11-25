Option Strict Off
Option Explicit On
'
Imports SAPbobsCOM.BoObjectTypes
Imports SAPbobsCOM.BoUTBTableType
Imports SAPbobsCOM.BoFieldTypes
Imports SAPbobsCOM.BoFldSubTypes
Imports SAPbobsCOM.BoYesNoEnum
Imports SAPbobsCOM.BoUDOObjType
Imports System
'
Public MustInherit Class clsCreaCampos


    Private mComp As SAPbobsCOM.Company
    Const cPermisoCampoUsuario As String = "261"
    Const cPermisoUDO As String = "500"

    Public Sub New(ByRef oCompany As SAPbobsCOM.Company)
        mComp = oCompany
    End Sub

    Protected Sub AddUserKey(ByVal sTableName As String, ByVal sKeyName As String,
                            ByVal iUnica As SAPbobsCOM.BoYesNoEnum, ByVal Array_Fields() As String)
        '
        Dim oClau As SAPbobsCOM.UserKeysMD
        Dim iField As Integer
        '
        sTableName = Replace(sTableName, "@", "")
        '
        If Not Me.ExisteClave(sTableName, sKeyName) Then
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
                If .Add <> 0 Then
                    Throw New Exception(mComp.GetLastErrorDescription)
                End If

            End With
        End If
        '
        clsUtil.LiberarObjetoCOM(oClau)
        '
    End Sub

    Protected Sub AddUserUDO(ByVal sName As String, ByVal sDescription As String, ByVal sLogTable As String,
                        ByVal iType As SAPbobsCOM.BoUDOObjType, ByVal iCanCancel As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanClose As SAPbobsCOM.BoYesNoEnum, ByVal iDefaultForm As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanDelete As SAPbobsCOM.BoYesNoEnum, ByVal iCanFind As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanLog As SAPbobsCOM.BoYesNoEnum, ByVal iCanYearTransfer As SAPbobsCOM.BoYesNoEnum,
                        ByVal iManageSeries As SAPbobsCOM.BoYesNoEnum,
                        Optional ByVal Array_ChildTables() As String = Nothing,
                        Optional ByVal Array_FindFields() As String = Nothing,
                        Optional ByVal Array_FormFields() As String = Nothing
                        )
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
                .UseUniqueFormType = tYES
                If Not IsNothing(Array_ChildTables) Then
                    With .ChildTables
                        For iChild = LBound(Array_ChildTables) To UBound(Array_ChildTables)
                            If iChild <> LBound(Array_ChildTables) Then .Add()
                            .TableName = Array_ChildTables(iChild)
                        Next
                    End With
                End If
                If iCanFind = tYES Then
                    If Not IsNothing(Array_FindFields) Then
                        With .FindColumns
                            For iChild = LBound(Array_FindFields) To UBound(Array_FindFields)
                                If iChild <> LBound(Array_FindFields) Then .Add()
                                .ColumnAlias = Array_FindFields(iChild)
                            Next
                        End With
                    End If
                End If
                If iDefaultForm = tYES Then
                    If Not IsNothing(Array_FormFields) Then
                        With .FormColumns
                            .FormColumnAlias = "Code" 'El Code es obligatorio
                            For iChild = LBound(Array_FormFields) To UBound(Array_FormFields)
                                If Array_FormFields(iChild).Trim.ToUpper <> "CODE" Then
                                    .Add()
                                    .FormColumnAlias = Array_FormFields(iChild)
                                End If
                            Next
                        End With
                    End If

                End If
                If .Add <> 0 Then
                    Throw New Exception(mComp.GetLastErrorDescription)
                End If

            End With
        End If
        '
        clsUtil.LiberarObjetoCOM(oUDO)
        '
    End Sub

    Protected Sub AddUserDefaultUDO(ByVal sName As String, ByVal sDescription As String, ByVal sLogTable As String,
                        ByVal iType As SAPbobsCOM.BoUDOObjType, ByVal iCanCancel As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanClose As SAPbobsCOM.BoYesNoEnum, ByVal iDefaultForm As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanDelete As SAPbobsCOM.BoYesNoEnum, ByVal iCanFind As SAPbobsCOM.BoYesNoEnum,
                        ByVal iCanLog As SAPbobsCOM.BoYesNoEnum, ByVal iCanYearTransfer As SAPbobsCOM.BoYesNoEnum,
                        ByVal iManageSeries As SAPbobsCOM.BoYesNoEnum,
                        Optional ByVal Array_ChildTables() As String = Nothing,
                        Optional ByVal Array_FindFields() As String = Nothing,
                        Optional ByVal Array_FormFields() As String = Nothing,
                        Optional ByVal Array_FormChildFields() As String = Nothing)
        '
        Dim oRecordSet As SAPbobsCOM.Recordset

        Dim oUDO As SAPbobsCOM.UserObjectsMD
        Dim iChild As Integer
        '
        Dim oUDOFind As SAPbobsCOM.UserObjectMD_FindColumns = Nothing
        Dim oUDOForm As SAPbobsCOM.UserObjectMD_FormColumns = Nothing
        Dim oUDOEnhancedForm As SAPbobsCOM.UserObjectMD_EnhancedFormColumns = Nothing
        oUDO = mComp.GetBusinessObject(oUserObjectsMD)

        oUDOFind = oUDO.FindColumns
        oUDOForm = oUDO.FormColumns
        oUDOEnhancedForm = oUDO.EnhancedFormColumns



        sName = Replace(sName, "@", "")
        '

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
                .UseUniqueFormType = tYES
                If Not IsNothing(Array_ChildTables) Then
                    With .ChildTables
                        For iChild = LBound(Array_ChildTables) To UBound(Array_ChildTables)
                            If iChild <> LBound(Array_ChildTables) Then .Add()
                            .TableName = Array_ChildTables(iChild)
                        Next
                    End With
                End If
                If iCanFind = tYES Then
                    If Not IsNothing(Array_FindFields) Then
                        With .FindColumns
                            For iChild = LBound(Array_FindFields) To UBound(Array_FindFields)
                                If iChild <> LBound(Array_FindFields) Then .Add()
                                .ColumnAlias = Array_FindFields(iChild)
                            Next
                        End With
                    End If
                End If
                If iDefaultForm = tYES Then
                    If Not IsNothing(Array_FormFields) Then
                        oUDOForm.FormColumnAlias = "Code" 'El Code es obligatorio
                        oUDOForm.FormColumnDescription = "Code"
                        oUDOForm.Editable = False
                        'oUDOForm.Editable = tNO
                        oUDOForm.Add()
                        For iChild = LBound(Array_FormFields) To UBound(Array_FormFields)
                            If Array_FormFields(iChild).Trim.ToUpper <> "CODE" Then
                                oUDOForm.FormColumnAlias = Array_FormFields(iChild)
                                oRecordSet = mComp.GetBusinessObject(BoRecordset)
                                Dim sql As String = "SELECT Descr FROM CUFD WHERE TableId = '@" & oUDO.Code & "' and AliasId = '" & Replace(Array_FormFields(iChild), "U_", "") & "'"
                                oRecordSet.DoQuery(sql)
                                oUDOForm.FormColumnDescription = oRecordSet.Fields.Item("Descr").Value.ToString
                                oUDOForm.Editable = tYES
                                oUDOForm.Add()
                                clsUtil.LiberarObjetoCOM(oRecordSet)
                            End If
                        Next
                    End If
                    If Not IsNothing(Array_FormChildFields) Then
                        For iChild = LBound(Array_FormChildFields) To UBound(Array_FormChildFields)

                            oUDOEnhancedForm.ColumnAlias = Array_FormChildFields(iChild)
                            oRecordSet = mComp.GetBusinessObject(BoRecordset)
                            Dim sql As String = "SELECT Descr FROM CUFD WHERE TableId = '@" & oUDO.ChildTables.TableName & "' and AliasId = '" & Replace(Array_FormChildFields(iChild), "U_", "") & "'"
                            oRecordSet.DoQuery(sql)
                            oUDOEnhancedForm.ColumnDescription = oRecordSet.Fields.Item("Descr").Value.ToString
                            oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                            oUDOEnhancedForm.Editable = tYES
                            oUDOEnhancedForm.ColumnNumber = iChild + 1
                            oUDOEnhancedForm.ChildNumber = 1
                            oUDOEnhancedForm.Add()
                            clsUtil.LiberarObjetoCOM(oRecordSet)
                        Next
                    End If
                End If

                If .Add <> 0 Then
                    Throw New Exception(mComp.GetLastErrorDescription)
                End If

            End With
        End If
        '
        clsUtil.LiberarObjetoCOM(oUDO)
        '
    End Sub

    Protected Sub AddUserField(ByVal sTable As String, ByVal sField As String, ByVal sDescription As String,
                            ByVal iType As SAPbobsCOM.BoFieldTypes, ByVal iSubType As SAPbobsCOM.BoFldSubTypes,
                            ByVal iSize As Integer, ByVal sLinkedTable As String,
                            ByVal iMandatory As SAPbobsCOM.BoYesNoEnum, ByVal sDefaultValue As String,
                            Optional ByVal aCodesValidValues() As String = Nothing,
                            Optional ByVal aNamesValidValues() As String = Nothing)
        '
        Dim oCamp As SAPbobsCOM.UserFieldsMD
        Dim i As Long
        '
        If Not ExisteCampo(sTable, sField) Then
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
                If .Add <> 0 Then
                    Throw New Exception(mComp.GetLastErrorDescription)
                End If

            End With
            '
            clsUtil.LiberarObjetoCOM(oCamp)
            '
        End If
        '
    End Sub

    'Protected Sub AddUserField(ByVal sTable As String, ByVal sField As String, ByVal sDescription As String,
    '                        ByVal iType As SAPbobsCOM.BoFieldTypes, ByVal iSubType As SAPbobsCOM.BoFldSubTypes,
    '                        ByVal iSize As Integer, ByVal sLinkedTable As String,
    '                        ByVal iMandatory As SAPbobsCOM.BoYesNoEnum, ByVal sDefaultValue As String,
    '                        Optional ByVal aCodesValidValues() As String = Nothing,
    '                        Optional ByVal aNamesValidValues() As String = Nothing,
    '                        Optional ByVal iLinkedSystemObject As SAPbobsCOM.UDFLinkedSystemObjectTypesEnum = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone)
    '    '
    '    Dim oCamp As SAPbobsCOM.UserFieldsMD
    '    Dim i As Long
    '    '
    '    If Not ExisteCampo(sTable, sField) Then
    '        oCamp = mComp.GetBusinessObject(oUserFields)
    '        With oCamp
    '            .TableName = sTable
    '            .Name = sField
    '            .Description = sDescription
    '            .Type = iType
    '            If iSubType <> st_None Then .SubType = iSubType
    '            If iMandatory = tYES Then
    '                If sDefaultValue = "" Then
    '                    Throw New Exception("Debe definirse un valor por defecto para el campo '" & sField & "-" & sDescription & "'")
    '                    Exit Sub
    '                Else
    '                    .Mandatory = iMandatory
    '                End If
    '            End If
    '            If iSize <> 0 Then .EditSize = iSize
    '            If sLinkedTable <> "" Then .LinkedTable = sLinkedTable
    '            If sDefaultValue <> "" Then .DefaultValue = sDefaultValue
    '            If Not IsNothing(aCodesValidValues) Then
    '                With .ValidValues
    '                    For i = LBound(aCodesValidValues) To UBound(aCodesValidValues)
    '                        If i <> LBound(aCodesValidValues) Then .Add()
    '                        .Value = aCodesValidValues(i)
    '                        .Description = aNamesValidValues(i)
    '                    Next
    '                End With
    '            End If
    '            If iLinkedSystemObject <> SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone Then
    '                .LinkedSystemObject = iLinkedSystemObject
    '            End If
    '            If .Add <> 0 Then
    '                Throw New Exception(mComp.GetLastErrorDescription)
    '            End If

    '        End With
    '        '
    '        clsUtil.LiberarObjetoCOM(oCamp)
    '        '
    '    End If
    '    '
    'End Sub

    Protected Sub AddUserTable(ByVal sName As String, ByVal sDescription As String, ByVal iType As SAPbobsCOM.BoUTBTableType)
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

    Protected Function ExisteCampo(ByVal sTableName As String, ByVal sFieldName As String) As Boolean
        '
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        '
        oTmpRecordset = mComp.GetBusinessObject(BoRecordset)
        If Mid(sTableName, 1, 1) <> "@" And Len(sTableName) > 4 Then
            sTableName = "@" & sTableName
        End If
        '
        oTmpRecordset.DoQuery("Select Count(*) From CUFD Where TableId = '" & sTableName & "' And AliasID = '" & sFieldName & "'")
        '
        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            ExisteCampo = True
        Else
            ExisteCampo = False
        End If
        '
        clsUtil.LiberarObjetoCOM(oTmpRecordset)
        '
    End Function

    Protected Function ExisteClave(ByVal sTableName As String, ByVal sKeyName As String) As Boolean
        '
        Dim ls As String
        Dim oRecordset As SAPbobsCOM.Recordset
        '
        If Left(sTableName, 1) <> "@" And Len(sTableName) > 4 Then sTableName = "@" & sTableName
        '
        ls = "SELECT TOP 1 KeyId "
        ls = ls & " FROM OUKD "
        ls = ls & " WHERE TableName = '" & sTableName & "' "
        ls = ls & " AND KeyName = '" & sKeyName & "' "
        clsUtil.LiberarObjetoCOM(oRecordset)
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)
        If oRecordset.EoF Then
            ExisteClave = False
        Else
            ExisteClave = True
        End If
        '
        clsUtil.LiberarObjetoCOM(oRecordset)
        '
    End Function

    Protected Function IDCampo(ByVal sTableName As String, ByVal sFieldName As String) As Integer
        '
        Dim ls As String
        Dim oRecordset As SAPbobsCOM.Recordset
        '
        If Mid(sTableName, 1, 1) <> "@" And Len(sTableName) > 4 Then
            sTableName = "@" & sTableName
        End If
        '
        ls = "SELECT TOP 1 FieldId "
        ls = ls & " FROM CUFD "
        ls = ls & " WHERE TableId = '" & sTableName & "' "
        ls = ls & " AND AliasId = '" & sFieldName & "' "
        clsUtil.LiberarObjetoCOM(oRecordset)
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)
        If oRecordset.EoF Then
            IDCampo = -1
        Else
            IDCampo = oRecordset.Fields.Item("FieldId").Value
        End If
        '
        clsUtil.LiberarObjetoCOM(oRecordset)
        '
    End Function

    Protected Sub AddUserValidValue(ByVal sTable As String, ByVal sField As String, ByVal sValue As String, ByVal sDescription As String)
        '
        Dim ls As String
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim iField As Integer
        Dim iIndex As Integer
        '
        If Left(sTable, 1) <> "@" And Len(sTable) > 4 Then sTable = "@" & sTable
        '
        iField = Me.IDCampo(sTable, sField)
        If iField < 0 Then Exit Sub
        '
        ls = "SELECT TOP 1 IndexId "
        ls = ls & " FROM UFD1 "
        ls = ls & " WHERE TableId = '" & sTable & "' "
        ls = ls & " AND FieldId = " & iField.ToString
        ls = ls & " AND FldValue = '" & sValue.Replace("'", "''").Trim & "' "
        clsUtil.LiberarObjetoCOM(oRecordset)
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)
        If oRecordset.EoF Then
            ls = "SELECT MAX(IndexId) AS Valor "
            ls = ls & " FROM UFD1 "
            ls = ls & " WHERE TableId = '" & sTable & "' "
            ls = ls & " AND FieldId = " & iField.ToString
            clsUtil.LiberarObjetoCOM(oRecordset)
            oRecordset = mComp.GetBusinessObject(BoRecordset)
            oRecordset.DoQuery(ls)
            If oRecordset.EoF Then
                iIndex = 0
            Else
                iIndex = oRecordset.Fields.Item("Valor").Value + 1
            End If
            ls = "INSERT INTO UFD1 (TableId, FieldId, IndexId, FldValue, Descr) VALUES ("
            ls = ls & " '" & sTable & "', "
            ls = ls & iField.ToString & ", "
            ls = ls & iIndex.ToString & ", "
            ls = ls & " '" & sValue.Replace("'", "''").Trim & "', "
            ls = ls & " '" & sDescription.Replace("'", "''").Trim & "') "
        Else
            iIndex = oRecordset.Fields.Item("IndexId").Value
            ls = "UPDATE UFD1 "
            ls = ls & " SET Descr = '" & sDescription.Replace("'", "''").Trim & "' "
            ls = ls & " WHERE TableId = '" & sTable & "' "
            ls = ls & " AND FieldId = " & iField.ToString
            ls = ls & " AND IndexId = " & iIndex.ToString
            ls = ls & " AND FldValue = '" & sValue.Replace("'", "''").Trim & "' "
        End If
        clsUtil.LiberarObjetoCOM(oRecordset)
        oRecordset = mComp.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery(ls)
        '
        clsUtil.LiberarObjetoCOM(oRecordset)
        '
    End Sub


End Class