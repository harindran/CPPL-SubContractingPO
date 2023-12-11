Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("721", "SubContractingPO/FrmGoodsReceipt.b1f")>
    Friend Class FrmGoodsReceipt
        Inherits SystemFormBase
        Private WithEvents objSubConform As SAPbouiCOM.Form
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents objUDFform As SAPbouiCOM.Form
        Dim objRs As SAPbobsCOM.Recordset

        Public Sub New()
        End Sub

#Region "Fields"

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

#End Region

        Public Overrides Sub OnInitializeComponent()
            Me.StaticText0 = CType(Me.GetItem("SubConNo").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("U_SubConNo").Specific, SAPbouiCOM.EditText)
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.LinkedButton0 = CType(Me.GetItem("lnksubpo").Specific, SAPbouiCOM.LinkedButton)
            Me.Matrix0 = CType(Me.GetItem("13").Specific, SAPbouiCOM.Matrix)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler DeactivateAfter, AddressOf Me.Form_DeactivateAfter
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter

        End Sub

        Private Sub OnCustomInitialize()
            Try
                'objform = objaddon.objapplication.Forms.GetForm("721", 1)
                'objform = objaddon.objapplication.Forms.ActiveForm
                'objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 0)

            Catch ex As Exception

            End Try
        End Sub

        Private Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                'Dim oUDFForm As SAPbouiCOM.Form
                'oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                'oUDFForm.Items.Item("U_SubConNo").Enabled = False
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                Dim StrQuery As String
                Dim objmatrix As SAPbouiCOM.Matrix
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objmatrix = objform.Items.Item("13").Specific
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    Dim FolderOutput, FolderScrap As SAPbouiCOM.Folder
                    objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", SubConTypeCount)
                    FolderOutput = objSubConform.Items.Item("flroutput").Specific
                    FolderScrap = objSubConform.Items.Item("flrscrap").Specific
                    'objmatrix = objform.Items.Item("13").Specific
                    'RemoveLastrow(objmatrix, "1")
                    Dim Rowcount As Integer = objmatrix.VisualRowCount - 1
                    If FolderOutput.Selected = True Then
                        If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                            If objaddon.HANA Then
                                StrQuery = "select ""U_Itemcode"" from ""@MIPL_POR2"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""U_Status"" ='O'"
                            Else
                                StrQuery = "select U_Itemcode from [@MIPL_POR2] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and U_Status ='O'"
                            End If
                            objRS.DoQuery(StrQuery)
                        End If
                        If objRS.RecordCount > 0 And Rowcount > 0 Then
                            If Rowcount <> objRS.RecordCount Then
                                objaddon.objapplication.StatusBar.SetText("Item Count is mismatching with the Output Tab of Sub Contracting.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                        For i As Integer = 1 To objmatrix.VisualRowCount
                            If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                                If objaddon.HANA Then
                                    objRS.DoQuery("select ""U_Itemcode"",""U_WhsCode"",""U_Qty"",""U_Price"" from ""@MIPL_POR2"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""U_Status"" ='O'")
                                Else
                                    objRS.DoQuery("select U_Itemcode,U_WhsCode,U_Qty,U_Price from [@MIPL_POR2] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and U_Status ='O'")
                                End If
                                If objRS.RecordCount > 0 Then
                                    'Dim Value As String = objmatrix.Columns.Item("10").Cells.Item(i).Specific.String
                                    'Value = Value.Replace("INR ", "")
                                    If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_Itemcode").Value.ToString Or objmatrix.Columns.Item("15").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_WhsCode").Value.ToString Or CDbl(objmatrix.Columns.Item("9").Cells.Item(i).Specific.String) <> CDbl(objRS.Fields.Item("U_Qty").Value) Then
                                        objaddon.objapplication.StatusBar.SetText("Line Item is not matching-Output Tab.Please re-generate from Sub Contracting screen.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                            End If
                        Next
                        For j = 1 To objmatrix.VisualRowCount
                            If objmatrix.Columns.Item("1").Cells.Item(j).Specific.String <> "" Then
                                If CDbl(objmatrix.Columns.Item("U_PlanQty").Cells.Item(j).Specific.String) <> CDbl(objmatrix.Columns.Item("9").Cells.Item(j).Specific.String) Then
                                    objaddon.objapplication.SetStatusBarMessage("Quantity mismatching from Sub-Con Output Tab Line: " & j, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                        Next
                    ElseIf FolderScrap.Selected = True Then
                        If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                            If objaddon.HANA Then
                                StrQuery = "select ""U_Itemcode"" from ""@MIPL_POR3"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""U_Status"" ='O'"
                            Else
                                StrQuery = "select U_Itemcode from [@MIPL_POR3] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and U_Status ='O'"
                            End If
                            objRS.DoQuery(StrQuery)
                        End If
                        If objRS.RecordCount > 0 And Rowcount > 0 Then
                            If Rowcount <> objRS.RecordCount Then
                                objaddon.objapplication.StatusBar.SetText("Item Count is mismatching with the Scrap Tab of Sub Contracting.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                        For i As Integer = 1 To objmatrix.VisualRowCount
                            If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                                If objaddon.HANA Then
                                    objRS.DoQuery("select ""U_Itemcode"",""U_WhsCode"",""U_Qty"",""U_Price"" from ""@MIPL_POR3"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""U_Status"" ='O'")
                                Else
                                    objRS.DoQuery("select U_Itemcode,U_WhsCode,U_Qty,U_Price from [@MIPL_POR3] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and U_Status ='O'")
                                End If
                                If objRS.RecordCount > 0 Then
                                    'Dim Value As String = objmatrix.Columns.Item("10").Cells.Item(i).Specific.String
                                    'Value = Value.Replace("INR ", "")
                                    If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_Itemcode").Value.ToString Or objmatrix.Columns.Item("15").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_WhsCode").Value.ToString Or CDbl(objmatrix.Columns.Item("9").Cells.Item(i).Specific.String) <> CDbl(objRS.Fields.Item("U_Qty").Value.ToString) Then
                                        objaddon.objapplication.StatusBar.SetText("Line Item is not matching -Scrap Tab.Please re-generate from Sub Contracting screen.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    For i As Integer = 1 To objmatrix.VisualRowCount
                        If objmatrix.Columns.Item("U_TabType").Cells.Item(i).Specific.String = "" Then
                            If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                                'If CDbl(objmatrix.Columns.Item("U_PlanQty").Cells.Item(i).Specific.String) <> CDbl(objmatrix.Columns.Item("9").Cells.Item(i).Specific.String) Then
                                '    objaddon.objapplication.SetStatusBarMessage("Quantity mismatching from Sub-Con Output Tab Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                '    BubbleEvent = False : Exit Sub
                                'End If
                                Dim Value As String = objmatrix.Columns.Item("10").Cells.Item(i).Specific.String
                                Value = Value.Replace("INR ", "")
                                If IIf(objmatrix.Columns.Item("U_Price").Cells.Item(i).Specific.String = "", 0, CDbl(objmatrix.Columns.Item("U_Price").Cells.Item(i).Specific.String)) <> IIf(Value = "", 0, Value) Then
                                    objaddon.objapplication.SetStatusBarMessage("UnitPrice should not to change...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                        ElseIf objmatrix.Columns.Item("U_TabType").Cells.Item(i).Specific.String = "Scrap" Then
                            If CDbl(objmatrix.Columns.Item("U_PlanQty").Cells.Item(i).Specific.String) <> CDbl(objmatrix.Columns.Item("9").Cells.Item(i).Specific.String) Then
                                objaddon.objapplication.SetStatusBarMessage("Quantity mismatching from Sub-Con Scrap Tab Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next

                End If

                GC.Collect()
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                Dim Status As String = "", TabName As String = "", StrQuery1 As String = "", StrQuery As String = "", GREntry As String = ""
                Dim objmatrix, objmatrixScrap, objmatrixOutput, objMatCosting As SAPbouiCOM.Matrix
                Dim objrs, objrecset As SAPbobsCOM.Recordset
                Dim FolderOutput, FolderScrap As SAPbouiCOM.Folder
                Try
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then Exit Sub
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True And pVal.BeforeAction = False Then
                        objmatrix = objform.Items.Item("13").Specific
                        objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", SubConTypeCount)
                        objmatrixOutput = objSubConform.Items.Item("mtxoutput").Specific
                        objmatrixScrap = objSubConform.Items.Item("mtxscrap").Specific
                        objMatCosting = objSubConform.Items.Item("MtxCosting").Specific
                        FolderOutput = objSubConform.Items.Item("flroutput").Specific
                        FolderScrap = objSubConform.Items.Item("flrscrap").Specific
                        objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                            If TranList = "Y" Then
                                GREntry = objform.DataSources.DBDataSources.Item("OIGN").GetValue("DocEntry", 0)
                            Else
                                GREntry = objform.DataSources.DBDataSources.Item("OIGN").GetValue("DocNum", 0)
                            End If
                            If objSubConform.Items.Item("txtGRNo").Specific.String = "" Then
                                objSubConform.Items.Item("txtGRNo").Specific.String = GREntry
                            Else
                                objSubConform.Items.Item("txtGRNo").Specific.String += ", " & GREntry
                            End If
                        End If
                        GREntry = objform.DataSources.DBDataSources.Item("OIGN").GetValue("DocEntry", 0)
                        Dim DocDate As Date = Date.ParseExact(objform.Items.Item("9").Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If FolderOutput.Selected = True Then
                            'Dim odbdsDetails As SAPbouiCOM.DBDataSource
                            'odbdsDetails = objSubConform.DataSources.DBDataSources.Item("@MIPL_POR2")
                            'objmatrixOutput.FlushToDataSource()
                            'For j As Integer = 0 To odbdsDetails.Size - 1
                            '    If odbdsDetails.GetValue("U_Itemcode", j) <> "" And odbdsDetails.GetValue("U_GRNo", j) = "" And odbdsDetails.GetValue("U_Status", j) = "O" Then
                            '        odbdsDetails.SetValue("U_GRNo", j, GREntry)
                            '        odbdsDetails.SetValue("U_Status", j, "C")
                            '        If objaddon.HANA Then
                            '            odbdsDetails.SetValue("U_InStock", j, objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  OITW where ""ItemCode""='" & odbdsDetails.GetValue("U_Itemcode", j) & "' and ""WhsCode""='" & odbdsDetails.GetValue("U_WhsCode", j) & "'"))
                            '        Else
                            '            odbdsDetails.SetValue("U_InStock", j, objaddon.objglobalmethods.getSingleValue("select OnHand from  OITW where ItemCode='" & odbdsDetails.GetValue("U_Itemcode", j) & "' and WhsCode='" & odbdsDetails.GetValue("U_WhsCode", j) & "'"))
                            '        End If
                            '        objmatrixOutput.CommonSetting.SetCellEditable(j + 1, 2, False) 'ItemCode
                            '        objmatrixOutput.CommonSetting.SetCellEditable(j + 1, 5, False)  'Qty
                            '        objmatrixOutput.CommonSetting.SetCellEditable(j + 1, 14, False) 'Whse
                            '        objmatrixOutput.CommonSetting.SetCellEditable(j + 1, 9, False)
                            '        objmatrixOutput.CommonSetting.SetCellEditable(j + 1, 10, False)
                            '    End If
                            'Next
                            'objmatrixOutput.LoadFromDataSource()

                            For j = 1 To objmatrixOutput.VisualRowCount
                                If objmatrixOutput.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And objmatrixOutput.Columns.Item("GRNo").Cells.Item(j).Specific.String = "" And objmatrixOutput.Columns.Item("Colstat").Cells.Item(j).Specific.String = "O" Then
                                    objmatrixOutput.Columns.Item("GRNo").Cells.Item(j).Specific.String = GREntry
                                    objmatrixOutput.Columns.Item("Colstat").Cells.Item(j).Specific.String = "C"
                                    If objaddon.HANA Then
                                        objmatrixOutput.Columns.Item("Instock").Cells.Item(j).Specific.String = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  OITW where ""ItemCode""='" & objmatrixOutput.Columns.Item("Code").Cells.Item(j).Specific.String & "' and ""WhsCode""='" & objmatrixOutput.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
                                    Else
                                        objmatrixOutput.Columns.Item("Instock").Cells.Item(j).Specific.String = objaddon.objglobalmethods.getSingleValue("select OnHand from  OITW where ItemCode='" & objmatrixOutput.Columns.Item("Code").Cells.Item(j).Specific.String & "' and WhsCode='" & objmatrixOutput.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
                                    End If
                                    'objSubConMatrix.CommonSetting.SetRowEditable(j, False)
                                    objmatrixOutput.CommonSetting.SetCellEditable(j, 2, False) 'ItemCode
                                    objmatrixOutput.CommonSetting.SetCellEditable(j, 5, False)  'Qty
                                    objmatrixOutput.CommonSetting.SetCellEditable(j, 14, False) 'Whse
                                    objmatrixOutput.CommonSetting.SetCellEditable(j, 9, False)
                                    objmatrixOutput.CommonSetting.SetCellEditable(j, 10, False)
                                End If
                            Next

                            If objaddon.HANA Then
                                StrQuery = objaddon.objglobalmethods.getSingleValue("Select ""U_Process"" from OCRD where CardCode=(Select ""U_CardCode"" from ""@MIPL_OPOR"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & ")")
                            Else
                                StrQuery = objaddon.objglobalmethods.getSingleValue("Select U_Process from OCRD where CardCode=(Select U_CardCode from [@MIPL_OPOR] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & ")")
                            End If
                            If (StrQuery.ToUpper) <> "OTHERS" Then Exit Sub
                            If Create_GoodsReceipt_PO(GREntry) Then

                            End If
                            ''StrQuery1 = objaddon.objglobalmethods.getSingleValue("select To_Varchar(""DocDate"",'dd/MM/yy') ""RefDate"" from OIGE where ""DocEntry""='" & objSubConMatrix.Columns.Item("GINo").Cells.Item(objSubConMatrix.VisualRowCount).Specific.String & "' ")
                            ''Dim RefDate As Date = Date.ParseExact(StrQuery1, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            ''StrQuery = "update T0 set T0.""U_Status""='C' from OJDT T0 join ""@MIPL_OPOR"" T1 on T1.""DocEntry""=T0.""U_SubConNo"" join ""@MIPL_POR5"" T2 on T2.""DocEntry""=T1.""DocEntry"""
                            ''StrQuery += vbCrLf + " where T0.""U_Status""='O' and T0.""U_SubConNo""='" & objSubConform.Items.Item("txtentry").Specific.String & "' and T0.""RefDate""='" & RefDate.ToString("yyyyMMdd") & "' and T2.""U_Status""='O'"
                            ''objrecset.DoQuery(StrQuery)
                            ''For j = 1 To objMatCosting.VisualRowCount
                            ''    If objMatCosting.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And objMatCosting.Columns.Item("Stat").Cells.Item(j).Specific.String = "O" Then
                            ''        objMatCosting.Columns.Item("Stat").Cells.Item(j).Specific.String = "C"
                            ''        objMatCosting.CommonSetting.SetRowEditable(j, False)
                            ''    End If
                            ''Next
                        ElseIf FolderScrap.Selected = True Then
                            For j = 1 To objmatrixScrap.VisualRowCount
                                If objmatrixScrap.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And objmatrixScrap.Columns.Item("Colstat").Cells.Item(j).Specific.String = "O" And objmatrixScrap.Columns.Item("GRNo").Cells.Item(j).Specific.String = "" Then
                                    objmatrixScrap.Columns.Item("GRNo").Cells.Item(j).Specific.String = GREntry
                                    objmatrixScrap.Columns.Item("Colstat").Cells.Item(j).Specific.String = "C"
                                    If objaddon.HANA Then
                                        objmatrixScrap.Columns.Item("Instock").Cells.Item(j).Specific.String = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  OITW where ""ItemCode""='" & objmatrixScrap.Columns.Item("Code").Cells.Item(j).Specific.String & "' and ""WhsCode""='" & objmatrixScrap.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
                                    Else
                                        objmatrixScrap.Columns.Item("Instock").Cells.Item(j).Specific.String = objaddon.objglobalmethods.getSingleValue("select OnHand from  OITW where ItemCode='" & objmatrixScrap.Columns.Item("Code").Cells.Item(j).Specific.String & "' and WhsCode='" & objmatrixScrap.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
                                    End If
                                    ' objmatrixScrap.CommonSetting.SetRowEditable(j, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 2, False) ' Type
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 3, False)   'ItemCode
                                    'objmatrixScrap.CommonSetting.SetCellEditable(j, 5, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 6, False)   'Qty
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 9, False)   'Whse
                                    'objmatrixScrap.CommonSetting.SetCellEditable(j, 16, False)
                                    'objmatrixScrap.CommonSetting.SetCellEditable(j, 17, False)
                                End If
                            Next
                        End If
                        SubConTypeCount = 0
                        If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            objSubConform.Items.Item("1").Click()
                            Exit Sub
                        End If
                        'objSubConform.Refresh()
                        'objSubConform.Update()

                    End If
                    objrs = Nothing
                    objrecset = Nothing
                    GC.Collect()
                Catch ex As Exception
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                'objform = objaddon.objapplication.Forms.GetForm("721", 1)
                objform.Items.Item("2310000079").Enabled = False
                objform.Items.Item("2310000078").Enabled = False
                Dim oUDFForm As SAPbouiCOM.Form
                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                oUDFForm.Items.Item("U_SubConNo").Enabled = False
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DeactivateAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                'If Not objaddon.objapplication.Forms.ActiveForm.TypeEx = "41" And Not objaddon.objapplication.Forms.ActiveForm.TypeEx = "21" Then  'Batch
                If objaddon.objapplication.Forms.ActiveForm.TypeEx = "SUBCTPO" Then
                    objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 1)
                    If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        objSubConform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub LinkedButton0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles LinkedButton0.PressedAfter
            Try
                Link_Value = objform.Items.Item("U_SubConNo").Specific.String 'DocEntry
                Link_objtype = "SUBPO"
                Dim activeform As New SubContractingPO
                activeform.Show()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("721", pVal.FormTypeCount)
                objform.Items.Item("2310000079").Enabled = False
                objform.Items.Item("2310000078").Enabled = False
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_KeyDownBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.KeyDownBefore
            Try
                If pVal.ColUID = "U_PlanQty" Or pVal.ColUID = "U_LineID" Or pVal.ColUID = "U_CardCode" Or pVal.ColUID = "U_Price" Then
                    BubbleEvent = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Function Create_GoodsReceipt_PO(ByVal GRNEntry As String)
            Try
                'If EditText5.Value <> "" Then Return True
                Dim objedit As SAPbouiCOM.EditText
                Dim objGRPO As SAPbobsCOM.Documents
                Dim DocEntry, strQuery As String
                objGRPO = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                Dim Lineflag As Boolean = False
                Dim Row As Integer = 1
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objaddon.objapplication.StatusBar.SetText("Creating Goods Receipt PO. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objedit = objform.Items.Item("9").Specific
                Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)

                If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()

                objGRPO.DocDate = DocDate
                objGRPO.JournalMemo = "Auto-Gen-> " & Now.ToString
                objGRPO.UserFields.Fields.Item("U_SubGR").Value = GRNEntry

                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("1").Cells.Item(i).Specific.String = "" Then Continue For
                    strQuery = "Select T0.DocEntry [Sub Entry],T0.U_SItemCode,T0.U_SItemDesc,T0.U_CardCode,T1.*"
                    strQuery += vbCrLf + "from [@MIPL_OPOR] T0 left join "
                    strQuery += vbCrLf + "(Select A.DocNum [PO Num],A.DocEntry [PO Entry],B.ItemCode,B.U_NonItemCode,B.LineNum,B.Quantity,B.Price,B.TaxCode,B.WhsCode,B.AcctCode"
                    strQuery += vbCrLf + "from OPOR A join POR1 B on A.DocEntry=B.DocEntry) T1 on T0.U_PurOrdrNo=T1.[PO Entry] and T0.U_SItemCode=T1.U_NonItemCode"
                    strQuery += vbCrLf + "where T0.DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " "
                    objRs.DoQuery(strQuery)
                    If objRs.RecordCount = 0 Then Exit Function
                    If objGRPO.CardCode = "" Then objGRPO.CardCode = Trim(objRs.Fields.Item("U_CardCode").Value.ToString)
                    objGRPO.Lines.ItemCode = Trim(objRs.Fields.Item("ItemCode").Value.ToString)
                    objGRPO.Lines.Quantity = CDbl(Matrix0.Columns.Item("9").Cells.Item(i).Specific.String) ' CDbl(Matrix0.Columns.Item("grnqty").Cells.Item(i).Specific.String) ' CDbl(objRs.Fields.Item("GRN Qty").Value.ToString) 
                    objGRPO.Lines.AccountCode = Trim(objRs.Fields.Item("AcctCode").Value.ToString)
                    objGRPO.Lines.TaxCode = Trim(objRs.Fields.Item("TaxCode").Value.ToString)
                    objGRPO.Lines.BaseType = 22
                    objGRPO.Lines.BaseEntry = CInt(objRs.Fields.Item("PO Entry").Value.ToString) ' CInt(objRs.Fields.Item("PO Entry").Value.ToString)
                    objGRPO.Lines.BaseLine = CInt(objRs.Fields.Item("LineNum").Value.ToString)
                    objGRPO.Lines.UnitPrice = Trim(objRs.Fields.Item("Price").Value.ToString)
                    objGRPO.Lines.WarehouseCode = Trim(objRs.Fields.Item("WhsCode").Value.ToString)
                    objGRPO.Lines.UserFields.Fields.Item("U_NonItemCode").Value = Trim(objRs.Fields.Item("U_SItemCode").Value.ToString)
                    objGRPO.Lines.UserFields.Fields.Item("U_NonItemDesc").Value = Trim(objRs.Fields.Item("U_SItemDesc").Value.ToString)
                    objGRPO.Lines.UserFields.Fields.Item("U_APInv").Value = ""
                    objGRPO.Lines.UserFields.Fields.Item("U_SubInv").Value = ""
                    objGRPO.Lines.Add()
                Next
                If objGRPO.Add() <> 0 Then
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("GRPO: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    objaddon.objapplication.MessageBox("GRPO: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode,, "OK")
                    Return False
                Else
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    DocEntry = objaddon.objcompany.GetNewObjectKey()
                    'EditText5.Value = DocEntry

                    objaddon.objapplication.StatusBar.SetText("GRPO Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Return True
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objGRPO)
                GC.Collect()
            Catch ex As Exception

            End Try
        End Function


    End Class
End Namespace
