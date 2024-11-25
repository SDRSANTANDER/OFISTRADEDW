Option Strict On
Imports System

Public Class CAMPOS

    Private mCompany As SAPbobsCOM.Company = Nothing

    Private Function Login() As Boolean


        Try



            'CONECTAR A LA COMPANY

            '''''''''''''''
            'LOCAL
            '''''''''''''''
            mCompany = New SAPbobsCOM.Company
            mCompany.UserName = "manager"
            mCompany.Password = "Seidor@19"
            mCompany.DbUserName = "sa"
            mCompany.DbPassword = "Seidor2019"
            mCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019
            mCompany.UseTrusted = False
            mCompany.CompanyDB = "SEI_DOCUWARE"
            mCompany.Server = "MP2CSZVC"
            mCompany.LicenseServer = "MP2CSZVC:30000"

            '''''''''''''''
            'CLIENTE
            '''''''''''''''
            mCompany = New SAPbobsCOM.Company
            mCompany.UserName = "manager"
            mCompany.Password = "seidor"
            mCompany.DbUserName = "sa"
            mCompany.DbPassword = "Seidor2017"
            mCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019
            mCompany.UseTrusted = False

            'Productivo: SBOISOTUBI_ES / SBOSPB_ES / SBOGOODPOLISH_ES / SBOLAPUENTE_ES / SBOTUROLENSE_ES / SBOGESTORA_ES
            'Test: TEST_SBOISOTUBI_ES / SBOSPB_TEST_20220118 / SBOGOODPOLISH_TEST_20220118 / TEST_SBOLAPUENTE_ES
            mCompany.CompanyDB = "SBOGESTORA_ES"

            mCompany.Server = "SRVSAP"
            mCompany.LicenseServer = "localhost:30000"
            mCompany.SLDServer = "SRVSAP:40000"

            If mCompany.Connect <> 0 Then
                Throw New Exception(mCompany.GetLastErrorCode & " " & mCompany.GetLastErrorDescription)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        Finally

        End Try

        Return True

    End Function

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        If Login() Then

            If MessageBox.Show("Conectado a SAP en " & mCompany.CompanyName & " (" & mCompany.CompanyDB & "). Continuar?", "Atención", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                Dim oCrea As New clsCampos(mCompany)
                oCrea.CreaCampos()

                mCompany.Disconnect()
                MessageBox.Show("FIN DE PROCESO")

            End If

        End If

    End Sub

End Class