Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("940", "SubContractingPO/FrmInventoryTransfer.b1f")>
    Friend Class FrmInventoryTransfer
        Inherits SystemFormBase
        Private WithEvents objSubConform As SAPbouiCOM.Form
        Private WithEvents objform As SAPbouiCOM.Form
        Public WithEvents objmatrix As SAPbouiCOM.Matrix
        Dim TypeCount As Integer
        Dim objRS As SAPbobsCOM.Recordset
        Dim StrQuery, LineId As String
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.StaticText0 = CType(Me.GetItem("SubConNo").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("U_SubConNo").Specific, SAPbouiCOM.EditText)
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.LinkedButton0 = CType(Me.GetItem("lnksubpo").Specific, SAPbouiCOM.LinkedButton)
            Me.ComboBox0 = CType(Me.GetItem("1250000074").Specific, SAPbouiCOM.ComboBox)
            Me.EditText1 = CType(Me.GetItem("3").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("23").Specific, SAPbouiCOM.Matrix)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DeactivateAfter, AddressOf Me.Form_DeactivateAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter

        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

        Private Sub OnCustomInitialize()
            Try
                'objform = objaddon.objapplication.Forms.GetForm("940", 1)
                'objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 1)

            Catch ex As Exception
            End Try

        End Sub

        Private WithEvents EditText0 As SAPbouiCOM.EditText

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
            Dim Status As String = ""
            Dim Docentry As String
            Dim objSubmatrix As SAPbouiCOM.Matrix
            If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True And pVal.BeforeAction = False Then
                    If TranList = "Y" Then
                        Docentry = objform.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0)
                    Else
                        Docentry = objform.DataSources.DBDataSources.Item("OWTR").GetValue("DocNum", 0)
                    End If
                    objmatrix = objform.Items.Item("23").Specific
                    If objform.Items.Item("U_SubConNo").Specific.String = "-1" Then
                        objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        For i As Integer = 1 To Matrix0.VisualRowCount
                            If Matrix0.Columns.Item("1").Cells.Item(i).Specific.String <> "" And Matrix0.Columns.Item("U_S_DocEntry").Cells.Item(i).Specific.String <> "" Then
                                If objaddon.HANA Then
                                    Status = objaddon.objglobalmethods.getSingleValue("select 1 from ""@MIPL_OPOR"" where ""DocEntry""='" & Matrix0.Columns.Item("U_S_DocEntry").Cells.Item(i).Specific.String & "' and ""U_InvTrNo"" like '%" & Docentry & "%'")
                                Else
                                    Status = objaddon.objglobalmethods.getSingleValue("select 1 from [@MIPL_OPOR] where DocEntry='" & Matrix0.Columns.Item("U_S_DocEntry").Cells.Item(i).Specific.String & "' and U_InvTrNo like '%" & Docentry & "%'")
                                End If
                                If Status = "1" Then Continue For
                                If objaddon.HANA Then
                                    Status = "Update ""@MIPL_OPOR"" set ""U_InvTrNo""=Case when ""U_InvTrNo"" is null then '" & Docentry & "' else ""U_InvTrNo""  +','+ '" & Docentry & "'  end where ""DocEntry""='" & Matrix0.Columns.Item("U_S_DocEntry").Cells.Item(i).Specific.String & "'"
                                Else
                                    Status = "Update [@MIPL_OPOR] set U_InvTrNo=Case when U_InvTrNo is null then '" & Docentry & "' else U_InvTrNo  +','+ '" & Docentry & "'  end where DocEntry='" & Matrix0.Columns.Item("U_S_DocEntry").Cells.Item(i).Specific.String & "'"
                                End If
                                objRS.DoQuery(Status)
                            End If

                        Next

                    Else
                        objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", SubConTypeCount)
                        objSubmatrix = objSubConform.Items.Item("MtxinputN").Specific
                        If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                            If objSubConform.Items.Item("TxtInvTr").Specific.String = "" Then
                                objSubConform.Items.Item("TxtInvTr").Specific.String = Docentry
                            Else
                                objSubConform.Items.Item("TxtInvTr").Specific.String += ", " & Docentry
                            End If
                            objSubmatrix.Columns.Item("SubWhse").Cells.Item(1).Click()
                            objSubmatrix.Columns.Item("PlanQty").Editable = False
                            objSubmatrix.Columns.Item("revbase").Editable = False
                        End If
                        SubConTypeCount = 0
                    End If

                    'If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    '    objSubConform.Items.Item("1").Click()
                    'End If
                    ' objform.Items.Item("2").Click()
                End If
                GC.Collect()
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                'StaticText0.Item.Top = objform.Items.Item("37").Top + 15 '502
                'StaticText0.Item.Left = objform.Items.Item("37").Left  '150
                'EditText0.Item.Top = objform.Items.Item("36").Top + 15 '19
                'EditText0.Item.Left = objform.Items.Item("36").Left '19
                objform = objaddon.objapplication.Forms.GetForm("940", pVal.FormTypeCount)
                StaticText0.Item.Top = objform.Items.Item("1470000099").Top + 16 '502
                StaticText0.Item.Left = objform.Items.Item("1470000099").Left  '150
                LinkedButton0.Item.Top = objform.Items.Item("1470000102").Top + 16
                LinkedButton0.Item.Left = objform.Items.Item("1470000102").Left
                EditText0.Item.Top = objform.Items.Item("1470000101").Top + 16 '19
                EditText0.Item.Left = objform.Items.Item("1470000101").Left '19
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DeactivateAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                ' If Not objaddon.objapplication.Forms.ActiveForm.TypeEx = "42" And Not objaddon.objapplication.Forms.ActiveForm.TypeEx = "21" Then  'Batch
                If objaddon.objapplication.Forms.ActiveForm.TypeEx = "SUBCTPO" Then
                    objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", SubConTypeCount)
                    If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        objSubConform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                End If

                'End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                'objform = objaddon.objapplication.Forms.GetForm("940", 1)
                Dim oUDFForm As SAPbouiCOM.Form
                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                'oUDFForm.Items.Item("U_SubConNo").Enabled = False
                'oUDFForm.Items.Item("U_ScrapNum").Enabled = False
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objmatrix = objform.Items.Item("23").Specific
                    Dim SubConEntry As String = ""
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If objform.Items.Item("U_SubConNo").Specific.String = "-1" Then SubConEntry = Matrix0.Columns.Item("U_S_DocEntry").Cells.Item(i).Specific.String Else SubConEntry = objform.Items.Item("U_SubConNo").Specific.String
                        If Matrix0.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                            If objaddon.HANA Then
                                StrQuery = "Select (A.""PlanQty"" - ifnull(A.""Transfer Qty""+A.""Return Qty"",0)) ""Qty"",* from (Select T0.""DocEntry"",T0.""DocNum"",T1.""U_ItemCode"" ""Item Code"",T1.""U_PlanQty"" ""PlanQty""" '+" & CDbl(Matrix0.Columns.Item("10").Cells.Item(i).Specific.String) & "
                                StrQuery += vbCrLf + "(Select ifnull(Sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_SubConNo""=T1.""DocEntry"" and B.""ItemCode""=T1.""U_Itemcode"") as ""Transfer Qty"","
                                StrQuery += vbCrLf + "(Select ifnull(Sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_ScrapNum""=T1.""DocEntry"" and B.""ItemCode""=T1.""U_Itemcode"") as ""Return Qty"","
                                StrQuery += vbCrLf + "(Select case when sum(B.""Quantity"")>0 then T1.""U_PlanQty""-ifnull(Sum(B.""Quantity""),0) else T1.""U_PlanQty"" end from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_SubConNo""=T1.""DocEntry"" and B.""ItemCode""=T1.""U_Itemcode"") as ""Pending Qty"""
                                StrQuery += vbCrLf + "from ""@MIPL_OPOR"" T0 join ""@MIPL_POR1"" T1 on T0.""DocEntry""=T1.""DocEntry"") A where A.""Pending Qty"">0 and A.""DocEntry""='" & SubConEntry & "' and A.""Item Code""='" & Matrix0.Columns.Item("1").Cells.Item(i).Specific.String & "'"

                            Else
                                StrQuery = "Select (A.[PlanQty] - isnull(A.[Transfer Qty]+A.[Return Qty],0)) Qty,* from (Select T0.DocEntry,T0.DocNum,T1.U_ItemCode [Item Code],T1.U_PlanQty [PlanQty]," '+" & CDbl(Matrix0.Columns.Item("10").Cells.Item(i).Specific.String) & "
                                StrQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T1.DocEntry and B.ItemCode=T1.U_Itemcode) as [Transfer Qty],"
                                StrQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T1.DocEntry and B.ItemCode=T1.U_Itemcode) as [Return Qty],"
                                StrQuery += vbCrLf + "(Select case when sum(B.Quantity)>0 then T1.U_PlanQty-isnull(Sum(B.Quantity),0) else T1.U_PlanQty end from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T1.DocEntry and B.ItemCode=T1.U_Itemcode) as [Pending Qty]"
                                StrQuery += vbCrLf + "from [@MIPL_OPOR] T0 join [@MIPL_POR1] T1 on T0.DocEntry=T1.DocEntry) A where A.[Pending Qty]>0 and A.DocEntry='" & SubConEntry & "' and A.[Item Code]='" & Matrix0.Columns.Item("1").Cells.Item(i).Specific.String & "'"
                            End If
                            objRS.DoQuery(StrQuery)
                            If objRS.RecordCount > 0 Then
                                If CDbl(objRS.Fields.Item("Qty").Value) <= 0.0 Then
                                    objaddon.objapplication.StatusBar.SetText("Excess Qty found for the Sub-Con Entry. Please check...Sub-Con Entry: " & SubConEntry & "Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                            If objmatrix.Columns.Item("U_PlanQty").Cells.Item(i).Specific.String = "" Then
                                objaddon.objapplication.StatusBar.SetText("Manually adding item not allowed for the Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                            If CDbl(objmatrix.Columns.Item("U_PlanQty").Cells.Item(i).Specific.String) < CDbl(objmatrix.Columns.Item("10").Cells.Item(i).Specific.String) Then
                                objaddon.objapplication.StatusBar.SetText("Plan Quantity mismatching from Sub-Con Entry Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next

                    If objform.Items.Item("U_SubConNo").Specific.String = "-1" Then
                    Else
                        'Dim CheckInvStat As Integer = 0
                        'If objaddon.HANA Then
                        '    StrQuery = "Select Case when Sum(T2.""Quantity"") >0 then T0.""U_PlanQty""-Sum(T2.""Quantity"") else T0.""U_PlanQty"" end "
                        '    StrQuery += vbCrLf + "+ ifnull((select Top 1 sum(T1.""U_Qty"") from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""U_Type""='2' and T1.""U_Status""='C' and T0.""DocEntry""='" & objform.Items.Item("U_SubConNo").Specific.String & "' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' "
                        '    StrQuery += vbCrLf + "group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc),0) as ""PendQty"",T0.""U_Itemcode"""
                        '    StrQuery += vbCrLf + " from ""@MIPL_POR1"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" left join OWTR T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join WTR1 T2  on T3.""DocEntry""=T2.""DocEntry"" And T0.""U_Itemcode""=T2.""ItemCode"" And T0.""LineId""= T2.""U_LineID"""
                        '    StrQuery += vbCrLf + "where T1.""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " group by T0.""U_PlanQty"",T0.""U_Itemcode"" order by ""PendQty"""
                        'Else
                        '    StrQuery = "Select Case when Sum(T2.Quantity) >0 then T0.U_PlanQty-Sum(T2.Quantity) else T0.U_PlanQty end "
                        '    StrQuery += vbCrLf + "+ isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and T0.DocEntry='" & objform.Items.Item("U_SubConNo").Specific.String & "' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' "
                        '    StrQuery += vbCrLf + "group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0) as PendQty,T0.U_Itemcode"
                        '    StrQuery += vbCrLf + "from [@MIPL_POR1] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry left join OWTR T3 on T3.U_SubConNo=T1.DocEntry left join WTR1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode and T0.LineId= T2.U_LineID"
                        '    StrQuery += vbCrLf + "where T1.DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " group by T0.U_PlanQty,T0.U_Itemcode order by PendQty"
                        'End If
                        'objRS.DoQuery(StrQuery)
                        'For Result As Integer = 0 To objRS.RecordCount - 1
                        '    If CDbl(objRS.Fields.Item("PendQty").Value) <= 0.0 Then
                        '        CheckInvStat += 1
                        '    End If
                        '    objRS.MoveNext()
                        'Next
                        'If CheckInvStat > 0 Then
                        '    objaddon.objapplication.StatusBar.SetText("Duplicate Inventory Transfer found for the Sub-Con Entry.please check...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False : Exit Sub
                        'End If

                        Dim Rowcount As Integer = objmatrix.VisualRowCount - 1
                        If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                            If objaddon.HANA Then
                                StrQuery = "select ""U_Itemcode"",""LineId"" from ""@MIPL_POR1"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & ""
                            Else
                                StrQuery = "select U_Itemcode,LineId from [@MIPL_POR1] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & ""
                            End If
                            objRS.DoQuery(StrQuery)
                        End If
                        If objRS.RecordCount > 0 And Rowcount > 0 Then
                            If Rowcount <> objRS.RecordCount Then
                                objaddon.objapplication.StatusBar.SetText("Item Count is mismatching with the Input Tab of Sub Contracting.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                        'For i As Integer = 1 To objmatrix.VisualRowCount
                        '    If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                        '        If objmatrix.Columns.Item("U_LineID").Cells.Item(i).Specific.String = "" Then
                        '            LineId = "0"
                        '        Else
                        '            LineId = objmatrix.Columns.Item("U_LineID").Cells.Item(i).Specific.String
                        '        End If
                        '        If objaddon.HANA Then
                        '            objRS.DoQuery("select ""U_Itemcode"",""U_WhsCode"",""U_SubWhse"" from ""@MIPL_POR1"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""LineId"" ='" & LineId & "'")
                        '        Else
                        '            objRS.DoQuery("select U_Itemcode,U_WhsCode,U_SubWhse from [@MIPL_POR1] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and LineId ='" & LineId & "'")
                        '        End If
                        '        If objRS.RecordCount > 0 Then
                        '            If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_Itemcode").Value.ToString Or objmatrix.Columns.Item("1470001039").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_WhsCode").Value.ToString Or objmatrix.Columns.Item("5").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_SubWhse").Value.ToString Then
                        '                objaddon.objapplication.StatusBar.SetText("Line Item Details is not matching.Please re-generate from Sub Contracting screen.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '                BubbleEvent = False : Exit Sub
                        '            End If
                        '        End If
                        '        If CDbl(objmatrix.Columns.Item("U_PlanQty").Cells.Item(i).Specific.String) <> CDbl(objmatrix.Columns.Item("10").Cells.Item(i).Specific.String) Then
                        '            objaddon.objapplication.SetStatusBarMessage("Quantity mismatching from Sub-Con Input Tab Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        '            BubbleEvent = False : Exit Sub
                        '        End If
                        '    End If
                        'Next

                    End If
                    objRS = Nothing
                    GC.Collect()
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton

        Private Sub LinkedButton0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles LinkedButton0.PressedAfter
            Try
                If objform.Items.Item("U_SubConNo").Specific.String = "-1" Then Exit Sub
                Link_Value = objform.Items.Item("U_SubConNo").Specific.String 'DocEntry
                Link_objtype = "SUBPO"
                Dim activeform As New SubContractingPO
                activeform.Show()
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

        Private Sub Matrix0_KeyDownBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.KeyDownBefore
            Try
                If pVal.ColUID = "U_S_DocEntry" Or pVal.ColUID = "U_PlanQty" Or pVal.ColUID = "U_LineID" Or pVal.ColUID = "U_Process" Then
                    BubbleEvent = False
                End If
            Catch ex As Exception

            End Try

        End Sub

    End Class
End Namespace
