Public Class clsCampos

    Inherits clsCreaCampos

    Sub New(ByRef oComp As SAPbobsCOM.Company)
        MyBase.New(oComp)
    End Sub

    Public Sub CreaCampos()

        Dim cTabla As String

        '----------------
        ' Docuware
        '----------------

        cTabla = "OCRD"
        Me.AddUserField(cTabla, "SEIDraft", "Doc. Borrador DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "S", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEICtaDW", "Cuenta contable DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIICDW", "IC referencia DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIRecDW", "Factura recurrente DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEIImpDW", "Importe límite DW", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        cTabla = "OOAT"
        Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIESTDW", "Estado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "0", Split("0|1|2|3|4|5", "|"), Split("Nuevo|Pendiente firmar|Firmado|Rechazado|Caducado|Error", "|"))
        Me.AddUserField(cTabla, "SEIMOTDW", "Motivo DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        cTabla = "ODLN"
        Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIIMPDW", "Importe DW", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIREVDW", "Revisar importe DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEITRADW", "Tratado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEIESTDW", "Estado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "0", Split("0|1|2|3|4|5", "|"), Split("Nuevo|Pendiente firmar|Firmado|Rechazado|Caducado|Error", "|"))
        Me.AddUserField(cTabla, "SEIMOTDW", "Motivo DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        cTabla = "ORCT"
        Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIIMPDW", "Importe DW", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIREVDW", "Revisar importe DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEITRADW", "Tratado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEIESTDW", "Estado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "0", Split("0|1|2|3|4|5", "|"), Split("Nuevo|Pendiente firmar|Firmado|Rechazado|Caducado|Error", "|"))
        Me.AddUserField(cTabla, "SEIMOTDW", "Motivo DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        cTabla = "OJDT"
        Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIIMPDW", "Importe DW", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIREVDW", "Revisar importe DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEITRADW", "Tratado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))

        'cTabla = "OPRJ"
        'Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        'Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        cTabla = "OWTR"
        Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEITRADW", "Tratado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEIESTDW", "Estado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "0", Split("0|1|2|3|4|5", "|"), Split("Nuevo|Pendiente firmar|Firmado|Rechazado|Caducado|Error", "|"))
        Me.AddUserField(cTabla, "SEIMOTDW", "Motivo DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        cTabla = "OIGN"
        Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        Me.AddUserField(cTabla, "SEITRADW", "Tratado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N", "|"), Split("Sí|No", "|"))
        Me.AddUserField(cTabla, "SEIESTDW", "Estado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "0", Split("0|1|2|3|4|5", "|"), Split("Nuevo|Pendiente firmar|Firmado|Rechazado|Caducado|Error", "|"))
        Me.AddUserField(cTabla, "SEIMOTDW", "Motivo DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        'cTabla = "OOPR"
        'Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        'Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        'cTabla = "OSCL"
        'Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        'Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        'cTabla = "OCLG"
        'Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        'Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        'cTabla = "OWOR"
        'Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        'Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        'cTabla = "OIPF"
        'Me.AddUserField(cTabla, "SEIIDDW", "Id. DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
        'Me.AddUserField(cTabla, "SEIURLDW", "URL DW", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")

        'cTabla = "ATC1"
        'Me.AddUserField(cTabla, "SEIPROCDW", "Procesado DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", Split("S|N|E", "|"), Split("Sí|No|Error", "|"))
		'Me.AddUserField(cTabla, "SEIERRORDW", "Error DW", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", SAPbobsCOM.BoYesNoEnum.tNO, "")
    
	End Sub

End Class
