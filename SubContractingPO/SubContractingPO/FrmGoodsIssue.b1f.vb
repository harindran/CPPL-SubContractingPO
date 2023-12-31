﻿Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("720", "SubContractingPO/FrmGoodsIssue.b1f")>
    Friend Class FrmGoodsIssue
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter

        End Sub

        Private Sub OnCustomInitialize()
            Try
                'objform = objaddon.objapplication.Forms.GetForm("720", 1)

            Catch ex As Exception

            End Try
        End Sub
        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                'If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                'objform = objaddon.objapplication.Forms.GetForm("720", 1)
                objform.Items.Item("2310000079").Enabled = False
                objform.Items.Item("2310000078").Enabled = False
                Dim oUDFForm As SAPbouiCOM.Form
                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                oUDFForm.Items.Item("U_SubConNo").Enabled = False
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("720", pVal.FormTypeCount)

            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
