Imports SAPbouiCOM
Namespace SubContractingPO

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "SUBCTPO"
                        SubContractingPO_MenuEvent(pVal, BubbleEvent)
                    Case "SUBBOM"
                        SubContractingBOM_MenuEvent(pVal, BubbleEvent)
                    Case "SUBINV"
                        SubContractingInvoice_MenuEvent(pVal, BubbleEvent)
                        'Case "SUBGEN"
                        '    GeneralSettings_MenuEvent(pVal, BubbleEvent)
                    Case "65211"
                        ProductionOrder_MenuEvent(pVal, BubbleEvent)
                    Case "940", "721", "720", "-940", "-721", "-720", "141", "-141"
                        Default_Sample_MenuEvent(pVal, BubbleEvent)
                        If objaddon.objapplication.Forms.ActiveForm.TypeEx = "940" And pVal.MenuUID = "SUBE" Then
                            If Not objaddon.FormExist("GetList") Then
                                Dim activeform As New FrmGetList
                                activeform.Show()
                                OEForm = objform
                                activeform.Load_SubContracting(objform.Items.Item("3").Specific.String)
                            End If
                        End If
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then

                Else
                    Dim oUDFForm As SAPbouiCOM.Form
                    Try
                        oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                    Catch ex As Exception
                        oUDFForm = objform
                    End Try

                    Select Case pval.MenuUID
                        Case "1281" 'Find
                            oUDFForm.Items.Item("U_SubConNo").Enabled = True
                            oUDFForm.Items.Item("U_ScrapNum").Enabled = True
                            oUDFForm.Items.Item("U_SubInv").Enabled = True
                        Case "1287"
                            If oUDFForm.Items.Item("U_SubConNo").Enabled = False Then
                                oUDFForm.Items.Item("U_SubConNo").Enabled = True
                            End If
                            oUDFForm.Items.Item("U_SubConNo").Specific.String = ""
                            oUDFForm.Items.Item("U_SubInv").Specific.String = ""
                        Case Else
                            oUDFForm.Items.Item("U_SubConNo").Enabled = False
                            objform.Items.Item("2310000079").Enabled = False
                            objform.Items.Item("2310000078").Enabled = False
                    End Select
                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "SubContractingPO"

        Private Sub SubContractingInvoice_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1284", "1286"
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                    End Select
                Else
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@SUBINV")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("tvencod").Enabled = True
                            objform.Items.Item("tvennam").Enabled = True
                            objform.Items.Item("tdocnum").Enabled = True
                            objform.Items.Item("tposdate").Enabled = True
                            objform.Items.Item("tinvnum").Enabled = True
                            objform.Items.Item("mtxdata").Enabled = False
                        Case "1282"
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource, "SCINV")
                            objform.Items.Item("tposdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")

                    End Select
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub SubContractingPO_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0, Matrix2, Matrix4, Matrix3, Matrix1 As SAPbouiCOM.Matrix
            Dim FolderInput, FolderOutput, FolderScrap, FolderRelDoc, FolderCosting As SAPbouiCOM.Folder
            Dim FolderID As String = ""
            Dim OnHand As String
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("MtxinputN").Specific
                Matrix2 = objform.Items.Item("mtxreldoc").Specific
                Matrix4 = objform.Items.Item("MtxCosting").Specific
                Matrix3 = objform.Items.Item("mtxoutput").Specific
                Matrix1 = objform.Items.Item("mtxscrap").Specific
                FolderInput = objform.Items.Item("flrinput").Specific
                FolderOutput = objform.Items.Item("flroutput").Specific
                FolderScrap = objform.Items.Item("flrscrap").Specific
                FolderRelDoc = objform.Items.Item("flrreldoc").Specific
                FolderCosting = objform.Items.Item("flrcosting").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        'Case "1283", "1284" 'Remove & Cancel

                        'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        'BubbleEvent = False

                        Case "1284", "1286"  'Cancel & Close
                            If pval.MenuUID = "1284" Then 'Cancel
                                If objaddon.objapplication.MessageBox("Cancelling a document is irreversible. Document status will be changed to ""Canceled"". Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                            Else 'Close
                                If objaddon.objapplication.MessageBox("Closing a document is irreversible. Document status will be changed to ""Closed"". Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub

                                If objaddon.objglobalmethods.Manual_StatusClose(objform) = False Then
                                    objaddon.objapplication.StatusBar.SetText("You Cannot Close the document the transaction is not fully Completed..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                            Dim objRS As SAPbobsCOM.Recordset
                            objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            If objaddon.HANA Then
                                objRS.DoQuery("Update ""@MIPL_OPOR"" set ""U_Clstat""='Y' where ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'")
                            Else
                                objRS.DoQuery("Update [@MIPL_OPOR] set U_Clstat='Y' where DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'")
                            End If
                            objRS = Nothing
                        Case "1293"  'Delete Row
                        Case "VPST" 'View Production Order Process Wise Status
                            If Not objaddon.FormExist("TRANVIEW") Then
                                Dim TranViewForm As New FrmViewTranData
                                TranViewForm.Show()
                                TranViewForm.objform.Title = "Production Order Process Wise Status"
                                TranViewForm.View_ProcessWise_SubCon_Entries(objform, objform.Items.Item("txtponum").Specific.String)
                            End If
                        Case "1288", "1289", "1290", "1291", "1287" '"1281", "1304",
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                objaddon.objapplication.StatusBar.SetText("Please update the Sub-Contracting Screen...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                            End If
                    End Select
                Else
                    If FolderInput.Selected = True Then
                        FolderID = "MtxinputN"
                    ElseIf FolderOutput.Selected = True Then
                        FolderID = "mtxoutput"
                    ElseIf FolderScrap.Selected = True Then
                        FolderID = "mtxscrap"
                    ElseIf FolderRelDoc.Selected = True Then
                        FolderID = "mtxreldoc"
                    ElseIf FolderCosting.Selected = True Then
                        FolderID = "MtxCosting"
                    End If
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@MIPL_OPOR")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("txtdocnum").Enabled = True
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtstat").Enabled = True
                            objform.Items.Item("txtcode").Enabled = True
                            objform.Items.Item("posdate").Enabled = True
                            objform.Items.Item("txtsitem").Enabled = True
                            'objform.Items.Item("txtbitem").Enabled = True
                            objform.Items.Item("docdate").Enabled = True
                            objform.Items.Item("deldate").Enabled = True
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtGINo").Enabled = True
                            objform.Items.Item("txtGRNo").Enabled = True
                            objform.Items.Item("TxtInvTr").Enabled = True
                            objform.Items.Item("tpodoc").Enabled = True
                            objform.Items.Item("tpoentry").Enabled = True
                            Matrix0.Item.Enabled = False
                            Matrix1.Item.Enabled = False
                            Matrix2.Item.Enabled = False
                            Matrix3.Item.Enabled = False
                            Matrix4.Item.Enabled = False
                            'objform.ActiveItem = "txtdocnum"
                            'objform.Items.Item("txtstat").Specific.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            'objform.Items.Item("txtstat").Specific.Select("O", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            objform.Items.Item("txtdocnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objform = objaddon.objapplication.Forms.GetForm("SUBCTPO", objform.TypeCount)
                            End If
                            Exit Sub
                        Case "1282" ' Add Mode
                            objform.Items.Item("btngendoc").Enabled = False
                            objform.Items.Item("btnload").Enabled = False
                            objform.Items.Item("BtnView").Enabled = False
                            objform.Items.Item("BtnInv").Enabled = False
                            objform.Items.Item("BtnInvO").Enabled = False
                            objform.Items.Item("AutoPO").Enabled = False
                            objform.Items.Item("BtnGIssue").Enabled = False
                            objform.Items.Item("btnOutput").Enabled = False
                            objform.Items.Item("BtnScrap").Enabled = False
                            objform.Items.Item("InvScrap").Enabled = False
                            objform.Items.Item("posdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("deldate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("docdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")

                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                            FolderScrap.Item.Click(BoCellClickType.ct_Regular)
                            FolderInput.Item.Click(BoCellClickType.ct_Regular)
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource, "SUBPO")
                            'objform.Items.Item("txtdocnum").Specific.string = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_OPOR")
                            'objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_OPOR")
                            For i As Integer = 1 To Matrix1.VisualRowCount
                                If Matrix1.Columns.Item("Code").Cells.Item(i).Specific.String = "" And Matrix1.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" Then
                                    Matrix1.Columns.Item("Code").Editable = True
                                    Matrix1.Columns.Item("Quant").Editable = True
                                    Matrix1.Columns.Item("Whse").Editable = True
                                End If
                            Next
                        Case "1288", "1289", "1290", "1291"

                        Case "1287"  'Duplicate
                            Matrix3.ClearRowData(Matrix3.VisualRowCount)
                            Matrix2.ClearRowData(Matrix2.VisualRowCount)
                            Matrix4.ClearRowData(Matrix4.VisualRowCount)
                            Matrix1.ClearRowData(Matrix1.VisualRowCount)
                            Matrix3.Clear()
                            Matrix4.Clear()
                            Matrix2.Clear()
                            Matrix1.Clear()
                            objform.Items.Item("txtGINo").Specific.String = ""
                            objform.Items.Item("txtGRNo").Specific.String = ""
                            objform.Items.Item("TxtInvTr").Specific.String = ""
                            objform.Items.Item("txtstat").Specific.Select("O", SAPbouiCOM.BoSearchKey.psk_ByValue) 'objform.Items.Item("txtstat").Specific.String = "O"
                            objform.Items.Item("txtentry").Specific.String = ""
                            objform.Items.Item("txtopqty").Specific.String = ""
                            objform.Items.Item("txtvobstk").Specific.String = ""
                            objform.Items.Item("txtvbal").Specific.String = ""
                            objform.Items.Item("deldate").Specific.String = ""
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource, "SUBPO")
                            objform.Items.Item("posdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("deldate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("docdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")
                            objform.Items.Item("btngendoc").Enabled = False
                            objform.Items.Item("btnload").Enabled = False
                            objform.Items.Item("BtnView").Enabled = False
                            objform.Items.Item("BtnInv").Enabled = False
                            objform.Items.Item("BtnGIssue").Enabled = False
                            FolderOutput.Item.Click(BoCellClickType.ct_Regular)
                            objform.Items.Item("btnOutput").Enabled = False
                            objform.Items.Item("BtnScrap").Enabled = False
                            objform.Items.Item("InvScrap").Enabled = False
                            objform.Items.Item("AutoPO").Enabled = False
                            objform.Items.Item("BtnInvO").Enabled = False
                            objform.Items.Item("txtponum").Specific.String = ""
                            objform.Items.Item("tpoentry").Specific.String = ""
                            objform.Items.Item("SubPoNum").Specific.String = ""
                            objform.Items.Item("tpodoc").Specific.String = ""
                            objform.Items.Item("POLine").Specific.String = ""
                            objform.Items.Item("toscrap").Specific.String = ""
                            For ii As Integer = 1 To Matrix0.VisualRowCount
                                If Matrix0.Columns.Item("distrule").Cells.Item(ii).Specific.String <> "" Then
                                    Matrix0.Columns.Item("distrule").Cells.Item(ii).Specific.String = ""
                                End If
                                Matrix0.Columns.Item("actplanqty").Cells.Item(ii).Specific.String = CDbl(objform.Items.Item("SQty").Specific.String) * CDbl(Matrix0.Columns.Item("Quant").Cells.Item(ii).Specific.String)
                                If objaddon.HANA Then
                                    OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(ii).Specific.String & "' and ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(ii).Specific.String & "'")
                                Else
                                    OnHand = objaddon.objglobalmethods.getSingleValue("select OnHand from  oitw where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(ii).Specific.String & "' and WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(ii).Specific.String & "'")
                                End If
                                Matrix0.Columns.Item("Instock").Cells.Item(ii).Specific.String = OnHand
                            Next
                            Matrix0.Columns.Item("PlanQty").Editable = True
                            Matrix0.Columns.Item("revbase").Editable = True
                            FolderInput.Item.Click(BoCellClickType.ct_Regular)
                        Case "1293"
                            Select Case FolderID
                                Case "MtxinputN"
                                    DeleteRow(Matrix0, "@MIPL_POR1")
                                Case "mtxoutput"
                                    DeleteRow(Matrix3, "@MIPL_POR2")
                                Case "mtxscrap"
                                    DeleteRow(Matrix1, "@MIPL_POR3")
                                Case "mtxreldoc"
                                    DeleteRow(Matrix2, "@MIPL_POR4")
                                Case "MtxCosting"
                                    DeleteRow(Matrix4, "@MIPL_POR5")
                            End Select
                        Case "1292"
                            Select Case FolderID
                                Case "MtxinputN"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                                Case "mtxoutput"
                                    'objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                                    FolderOutput.Item.Click()
                                    'Matrix3.CommonSetting.SetCellEditable(Matrix3.VisualRowCount, 2, True)
                                    'Matrix3.CommonSetting.SetCellEditable(Matrix3.VisualRowCount, 4, True)
                                    'Matrix3.CommonSetting.SetCellEditable(Matrix3.VisualRowCount, 10, True)
                                    'objaddon.objglobalmethods.GetSubItemQty(objform, Matrix3)
                                    'objaddon.objglobalmethods.ProcessCost_Output(objform, Matrix3)
                                Case "mtxscrap"
                                    'objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                                    FolderScrap.Item.Click()
                                Case "mtxreldoc"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                                Case "MtxCosting"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                            End Select
                        Case "1304" 'Refresh

                            Select Case FolderID
                                Case "MtxinputN"
                                    For i As Integer = 1 To Matrix0.VisualRowCount
                                        If objaddon.HANA Then
                                            OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")
                                        Else
                                            OnHand = objaddon.objglobalmethods.getSingleValue("select OnHand from  oitw where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")
                                        End If
                                        Matrix0.Columns.Item("Instock").Cells.Item(i).Specific.String = OnHand
                                    Next
                            End Select
                            'If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            'If objform.Mode = BoFormMode.fm_UPDATE_MODE Then
                            '    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'End If
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Finally
            End Try
        End Sub
#End Region

        Private Sub SubContractingBOM_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0, Matrix1 As SAPbouiCOM.Matrix
            Dim odbdsDetails As SAPbouiCOM.DBDataSource
            Dim FolderItem, FolderProcess As SAPbouiCOM.Folder
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_BOM1")
                Matrix0 = objform.Items.Item("mtxBOM").Specific
                Matrix1 = objform.Items.Item("mtxproc").Specific
                FolderItem = objform.Items.Item("fldritem").Specific
                FolderProcess = objform.Items.Item("fldrproc").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                            If objaddon.objapplication.MessageBox("Removal of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                        Case "1293"
                            If FolderItem.Selected Then
                                If Matrix0.VisualRowCount = 1 Then BubbleEvent = False
                            Else
                                If Matrix1.VisualRowCount = 1 Then BubbleEvent = False
                            End If

                        Case "1292"

                    End Select
                Else
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@MIPL_OBOM")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode                           
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtcode").Enabled = True
                            objform.Items.Item("mtxBOM").Enabled = False
                            objform.Items.Item("txtentry").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Case "1282" ' Add Mode
                            'If objaddon.HANA Then
                            '    objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 ""DocEntry"" from ""@MIPL_OBOM""")
                            'Else
                            '    objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 DocEntry from [@MIPL_OBOM]")
                            'End If
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "ItemNo", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "procc", "#")
                            If ResourceEnable <> "Y" Then
                                Dim cmbtype As SAPbouiCOM.Column = Matrix0.Columns.Item("Type")
                                cmbtype.ValidValues.Remove("290", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        Case "1288", "1289", "1290", "1291"
                            'objform.Items.Item("btngendoc").Enabled = True
                            objaddon.objapplication.Menus.Item("1300").Activate()
                        Case "1293"
                            If FolderItem.Selected Then
                                DeleteRow(Matrix0, "@MIPL_BOM1")
                                'For i As Integer = Matrix0.VisualRowCount To 1 Step -1
                                '    Matrix0.Columns.Item("#").Cells.Item(i).Specific.String = i
                                'Next
                            Else
                                DeleteRow(Matrix1, "@MIPL_BOM2")
                                'For i As Integer = Matrix1.VisualRowCount To 1 Step -1
                                '    Matrix1.Columns.Item("#").Cells.Item(i).Specific.String = i
                                'Next
                            End If
                            If objform.Mode = BoFormMode.fm_OK_MODE Then objform.Mode = BoFormMode.fm_UPDATE_MODE
                            objform.Update()
                            objform.Refresh()
                        Case "1292"
                            Try
                                If FolderItem.Selected Then
                                    If Matrix0.VisualRowCount > 0 Then
                                        If odbdsDetails.GetValue("U_Itemcode", Matrix0.VisualRowCount - 1) = "" Then Exit Sub
                                        objform.Freeze(True)
                                        odbdsDetails.InsertRecord(odbdsDetails.Size)
                                        odbdsDetails.SetValue("LineId", Matrix0.VisualRowCount, Matrix0.VisualRowCount + 1)
                                        Matrix0.LoadFromDataSource()
                                        objform.Freeze(False)
                                    End If
                                Else
                                    odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_BOM2")
                                    If Matrix1.VisualRowCount > 0 Then
                                        If odbdsDetails.GetValue("U_Proccode", Matrix1.VisualRowCount - 1) = "" Then Exit Sub
                                        objform.Freeze(True)
                                        odbdsDetails.InsertRecord(odbdsDetails.Size)
                                        odbdsDetails.SetValue("LineId", Matrix1.VisualRowCount, Matrix1.VisualRowCount + 1)
                                        Matrix0.LoadFromDataSource()
                                        objform.Freeze(False)
                                    End If
                                End If
                            Catch ex As Exception
                            End Try
                        Case "1287"  'Duplicate
                            objform.Items.Item("txtprodno").Specific.String = ""
                            objform.Items.Item("txtcode").Specific.String = ""
                            objform.Items.Item("txtname").Specific.String = ""
                            objform.Items.Item("txtentry").Specific.String = ""
                            'Matrix0.FlushToDataSource()
                            'odbdsDetails.RemoveRecord(odbdsDetails.Size - 1)
                            'Matrix0.LoadFromDataSource()
                            'odbdsDetails.Clear()
                            'Matrix0.Clear()
                            'Matrix0.FlushToDataSource()
                            'For i As Integer = 1 To Matrix0.VisualRowCount
                            '    Matrix0.GetLineData(i)
                            '    odbdsDetails.Offset = i - 1
                            '    odbdsDetails.SetValue("LineId", odbdsDetails.Offset, i)
                            '    Matrix0.SetLineData(i)
                            '    Matrix0.FlushToDataSource()
                            'Next
                            ''odbdsDetails.RemoveRecord(odbdsDetails.Size - 1)
                            'Matrix0.LoadFromDataSource()
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub GeneralSettings_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"

                        Case "1292"

                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode                           

                        Case "1282" ' Add Mode
                        Case "1288", "1289", "1290", "1291"
                        Case "1293"

                        Case "1292"

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub ProductionOrder_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm

                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode 

                        Case "1282" ' Add Mode
                        Case "1288", "1289", "1290", "1291", "1304"
                            If SubPOScreen <> "" Then
                                objaddon.objglobalmethods.SubConButtonEnabling(objform)
                            End If
                        Case "1293"
                        Case "1292"
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub
     
    End Class
End Namespace