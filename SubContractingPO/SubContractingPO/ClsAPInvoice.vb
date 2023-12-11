Imports System.IO
Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    Public Class ClsAPInvoice
        Public Const Formtype = "141"
        Dim objAPform As SAPbouiCOM.Form
        Dim strsql As String
        Dim objrecset, objrsupdate As SAPbobsCOM.Recordset

        Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
            objAPform = objaddon.objapplication.Forms.Item(FormUID)
            If pVal.BeforeAction Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                End Select
            Else
                Try
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            If pVal.ActionSuccess Then
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    End Select
                Catch ex As Exception
                End Try
            End If
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objAPform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If BusinessObjectInfo.ActionSuccess Then
                                Dim oUDFForm As SAPbouiCOM.Form
                                oUDFForm = objaddon.objapplication.Forms.Item(objAPform.UDFFormUID)
                                Dim SubInvEntry As String = oUDFForm.Items.Item("U_SubInv").Specific.String
                                If SubInvEntry = "" Then Exit Sub
                                strsql = "Select T1.* from [@SUBINV] T0 join [@SUBINV1] T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry='" & SubInvEntry & "'"
                                objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objrsupdate = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objrecset.DoQuery(strsql)
                                If objrecset.RecordCount = 0 Then Exit Sub
                                Dim DocEntry As String = objAPform.DataSources.DBDataSources.Item("OPCH").GetValue("DocEntry", 0)
                                For i As Integer = 0 To objrecset.RecordCount - 1
                                    strsql = "Update OIGN Set U_SubInv='" & SubInvEntry & "',U_APInv='" & DocEntry & "' where DocEntry='" & Trim(objrecset.Fields.Item("U_GRNNum").Value.ToString) & "' and U_SubConNo='" & Trim(objrecset.Fields.Item("U_SDocEntry").Value.ToString) & "' and U_SubInv is null"
                                    objrsupdate.DoQuery(strsql)
                                    strsql = "Update [@MIPL_POR2] Set U_SubInv='" & SubInvEntry & "',U_APInv='" & DocEntry & "' where DocEntry='" & Trim(objrecset.Fields.Item("U_SDocEntry").Value.ToString) & "' and  U_GRNo='" & Trim(objrecset.Fields.Item("U_GRNNum").Value.ToString) & "' and U_SubInv is null"
                                    objrsupdate.DoQuery(strsql)
                                    strsql = "Update POR1 Set U_SubInv='" & SubInvEntry & "',U_APInv='" & DocEntry & "' where DocEntry='" & Trim(objrecset.Fields.Item("U_PoEntry").Value.ToString) & "' and LineNum='" & Trim(objrecset.Fields.Item("U_PoLine").Value.ToString) & "' and U_SubInv is null"
                                    objrsupdate.DoQuery(strsql)
                                    strsql = "Update [@SUBINV]  set U_TranEntry='" & DocEntry & "' where DocEntry='" & SubInvEntry & "' and U_TranEntry is not null"
                                    objrsupdate.DoQuery(strsql)
                                Next
                                objrsupdate = Nothing
                                objrecset = Nothing
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

    End Class
End Namespace

