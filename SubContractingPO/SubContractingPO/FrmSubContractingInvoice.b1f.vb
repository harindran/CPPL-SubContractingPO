Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("SUBINV", "SubContractingPO/FrmSubContractingInvoice.b1f")>
    Friend Class FrmSubContractingInvoice
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsHeader As SAPbouiCOM.DBDataSource
        Dim objRs As SAPbobsCOM.Recordset
        Dim FormCount As Integer = 0
        Dim strQuery As String

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.EditText0 = CType(Me.GetItem("tvencod").Specific, SAPbouiCOM.EditText)
            Me.StaticText0 = CType(Me.GetItem("lvencode").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("tvennam").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("mtxdata").Specific, SAPbouiCOM.Matrix)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Button2 = CType(Me.GetItem("btngendoc").Specific, SAPbouiCOM.Button)
            Me.Button3 = CType(Me.GetItem("btnload").Specific, SAPbouiCOM.Button)
            Me.StaticText2 = CType(Me.GetItem("ldocnum").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("tdocnum").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("lposdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("tposdate").Specific, SAPbouiCOM.EditText)
            Me.ComboBox0 = CType(Me.GetItem("Series").Specific, SAPbouiCOM.ComboBox)
            Me.LinkedButton0 = CType(Me.GetItem("lkbp").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText4 = CType(Me.GetItem("linvnum").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("tinvnum").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton1 = CType(Me.GetItem("lkapinv").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText5 = CType(Me.GetItem("lblremark").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("txtremark").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter

        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Button3 As SAPbouiCOM.Button
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton1 As SAPbouiCOM.LinkedButton

#End Region

        Private Sub OnCustomInitialize()
            Try
                odbdsHeader = objform.DataSources.DBDataSources.Item("@SUBINV")
                odbdsDetails = objform.DataSources.DBDataSources.Item("@SUBINV1")
                objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "SCINV")

                objform.Items.Item("tposdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "Series", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tvencod", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tposdate", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btngendoc", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btnload", True, False, False)
                Matrix0.Columns.Item("poline").Visible = False
            Catch ex As Exception

            End Try
        End Sub

        Private Sub EditText0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText0.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_BP")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    'strQuery = "select distinct ""U_CardCode"",""U_CardName"" from ""@MIPL_OPOR"" where ""U_CardCode"" <>''"
                    strQuery = "Select distinct ""CardCode"" from OCRD where ""CardCode"" in (select distinct T0.""U_CardCode"" from ""@MIPL_OPOR"" T0 where T0.""U_CardCode"" <>'' )"
                Else
                    'strQuery = "select distinct U_CardCode,U_CardName from [@MIPL_OPOR] where U_CardCode <>''"
                    strQuery = "Select distinct CardCode from OCRD where CardCode in (select distinct T0.U_CardCode from [@MIPL_OPOR] T0 where T0.U_CardCode <>'' )"
                End If
                objRs.DoQuery(strQuery)
                If objRs.RecordCount > 0 Then
                    oCond = oConds.Add()
                    oCond.Alias = "CardType"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "S"
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    For Val As Integer = 0 To objRs.RecordCount - 1
                        If Val = 0 Then
                            oCond = oConds.Add()
                            oCond.Alias = "CardCode"
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(objRs.Fields.Item(0).Value)
                        Else
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                            oCond = oConds.Add()
                            oCond.Alias = "CardCode"
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(objRs.Fields.Item(0).Value)
                        End If
                        objRs.MoveNext()
                    Next
                Else
                    oCond = oConds.Add()
                    oCond.Alias = "CardCode"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = ""
                End If
                oCFL.SetConditions(oConds)
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText0_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        odbdsHeader.SetValue("U_BPCode", 0, pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        odbdsHeader.SetValue("U_BPName", 0, pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    'objform.Update()
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("SUBINV", pVal.FormTypeCount)
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button3_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try
                If objaddon.HANA Then
                    ProductQuery = "Select * from "
                    ProductQuery += vbCrLf + "(Select T0.""DocNum"" ""Sub Con Doc#"",T0.""DocEntry"" ""Sub Con DocEntry"",TO_VARCHAR(T0.""U_DocDate"",'yyyyMMdd') ""Sub Con Doc Date"", T0.""U_PurEnt"" ""PO#"","
                    ProductQuery += vbCrLf + "(Select ""DocDate"" from OPOR where ""DocEntry""=T0.""U_PurOrdrNo"") ""PO Date"",T0.""U_SItemCode"" ""Sub Item Code"",T0.""U_SItemDesc"" ""Sub Item Description"","
                    ProductQuery += vbCrLf + "T0.""U_SQty"" ""Plan Qty"",T2.""DocNum"" ""GRN#"",T2.""DocEntry"" ""GRN Entry"", TO_VARCHAR(T2.""DocDate"",'yyyyMMdd') ""GRN Date"","
                    ProductQuery += vbCrLf + "(Select B.""Quantity"" from IGN1 B where B.""DocEntry""=T2.""DocEntry"" and B.""DocEntry""=T1.""U_GRNo"" and B.""ItemCode""=T1.""U_Itemcode"") ""GRN Qty"""
                    ProductQuery += vbCrLf + "from ""@MIPL_OPOR"" T0 join ""@MIPL_POR2"" T1 on T0.""DocEntry""=T1.""DocEntry"""
                    ProductQuery += vbCrLf + "left join (Select ""DocNum"",""DocEntry"",""DocDate"" ,""U_SubConNo"",""U_SubInv"" from OIGN ) T2 on T2.""U_SubConNo""=T0.""DocEntry"" and T2.""DocEntry""=T1.""U_GRNo"""
                    ProductQuery += vbCrLf + "where T0.""U_CardCode""='" & EditText0.Value & "' and T0.""U_DocDate"">='20221201' and  T2.""U_SubInv"" is null) T3"
                    ProductQuery += vbCrLf + "where T3.""GRN Qty"">0"
                Else
                    'ProductQuery = "Select T0.DocNum [Sub Con Doc#],T0.DocEntry [Sub Con DocEntry],T0.U_DocDate [Sub Con Doc Date], T0.U_PurEnt [PO#],"
                    'ProductQuery += vbCrLf + "(Select DocDate from OPOR where DocEntry=T0.U_PurOrdrNo) [PO Date],T0.U_SItemCode [Sub Item Code],"
                    'ProductQuery += vbCrLf + "T0.U_SItemDesc [Sub Item Description],T0.U_SQty [Plan Qty],(Select DocNum from OIGN where DocEntry=T1.U_GRNo) [GRN#],"
                    'ProductQuery += vbCrLf + "(Select DocDate from OIGN where DocEntry=T1.U_GRNo) [GRN Date],T1.U_Qty [GRN Qty]"
                    'ProductQuery += vbCrLf + "from [@MIPL_OPOR] T0 join [@MIPL_POR2] T1 on T0.DocEntry=T1.DocEntry where T0.U_CardCode='" & EditText0.Value & "'"

                    ProductQuery = "Select * from "
                    ProductQuery += vbCrLf + "(Select T0.DocNum [Sub Con Doc#],T0.DocEntry [Sub Con DocEntry],Format(T0.U_DocDate,'yyyyMMdd') [Sub Con Doc Date], T0.U_PurEnt [PO#],T0.U_PurOrdrNo [PO Entry],"
                    ProductQuery += vbCrLf + "(Select LineNum from POR1 where DocEntry=T0.U_PurOrdrNo and U_NonItemCode=T0.U_SItemCode) [PO Line],"
                    ProductQuery += vbCrLf + "(Select DocEntry from OPDN where U_SubGR=T2.DocEntry) [GRPO Entry],(Select Format(DocDate,'yyyyMMdd') from OPDN where U_SubGR=T2.DocEntry) [GRPO Date],"
                    ProductQuery += vbCrLf + "(Select Format(DocDate,'yyyyMMdd') from OPOR where DocEntry=T0.U_PurOrdrNo) [PO Date],T0.U_SItemCode [Sub Item Code],T0.U_SItemDesc [Sub Item Description],"
                    ProductQuery += vbCrLf + "T0.U_SQty [Plan Qty],T2.DocNum [GRN#],T2.DocEntry [GRN Entry], Format(T2.DocDate,'yyyyMMdd') [GRN Date],"
                    ProductQuery += vbCrLf + "(Select B.Quantity from IGN1 B where B.DocEntry=T2.DocEntry and B.DocEntry=T1.U_GRNo and B.ItemCode=T1.U_Itemcode) [GRN Qty]"
                    ProductQuery += vbCrLf + "from [@MIPL_OPOR] T0 join [@MIPL_POR2] T1 on T0.DocEntry=T1.DocEntry"
                    ProductQuery += vbCrLf + "left join (Select DocNum,DocEntry,DocDate ,U_SubConNo,U_SubInv from OIGN ) T2 on T2.U_SubConNo=T0.DocEntry and T2.DocEntry=T1.U_GRNo"
                    ProductQuery += vbCrLf + "where T0.U_CardCode='" & EditText0.Value & "' and T0.U_DocDate>='20221201' and  T2.U_SubInv is null) T3"
                    ProductQuery += vbCrLf + "where T3.[GRN Qty]>0 and T3.[GRPO Entry] is not null and T3.[GRN Entry] not in (Select distinct Y.U_GRNNum from [@SUBINV] X join [@SUBINV1] Y on X.DocEntry=Y.DocEntry where X.U_TranEntry is not null)"
                End If
                FrmMultiSel = objaddon.objapplication.Forms.ActiveForm
                If Not objaddon.FormExist("MULSEL") Then
                    FieldName = "subinvoice"
                    Dim Multiselect As New FrmMultiSelect
                    Multiselect.Show()
                    Multiselect.UIAPIRawForm.Left = objform.Left + 100
                    Multiselect.UIAPIRawForm.Top = objform.Top + 100
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If EditText0.Value = "" Then
                    objaddon.objapplication.StatusBar.SetText("Vendor Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
                If Matrix0.VisualRowCount = 0 Then
                    objaddon.objapplication.StatusBar.SetText("Row is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Function Create_AP_Invoice() As Boolean
            Try
                If EditText5.Value <> "" Then Return True
                Dim objAPInvoice As SAPbobsCOM.Documents
                'Dim objedit As SAPbouiCOM.EditText
                Dim DocEntry As String
                objAPInvoice = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                Dim Lineflag As Boolean = False
                Dim Row As Integer = 1
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objaddon.objapplication.StatusBar.SetText(" Creating A/P Invoice Draft. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'QCDocNum = objform.BusinessObject.GetNextSerialNumber(objform.Items.Item("14").Specific.Selected.value)
                'objedit = objform.Items.Item("tposdate").Specific
                Dim DocDate As Date = Date.ParseExact(EditText4.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) 'yyyyMMdd EditText4
                Try
                    If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                    objAPInvoice.CardCode = EditText0.Value
                    objAPInvoice.DocDate = DocDate
                    'objAPInvoice.Series = 186
                    objAPInvoice.JournalMemo = "Auto-Gen-> " & Now.ToString
                    'objAPInvoice.Comments = "QCA DocNum-> " & CStr(QCDocNum)
                    objAPInvoice.UserFields.Fields.Item("U_SubInv").Value = odbdsHeader.GetValue("DocEntry", 0)
                    objAPInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                    objAPInvoice.GSTTransactionType = SAPbobsCOM.GSTTransactionTypeEnum.gsttrantyp_GSTTaxInvoice
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If Matrix0.Columns.Item("scentry").Cells.Item(i).Specific.String = "" Then Continue For
                        strQuery = objaddon.objglobalmethods.getSingleValue("Select 1 as Status from ODRF T0 join DRF1 T1 on T0.DocEntry=T1.DocEntry where T1.BaseEntry=" & Trim(Matrix0.Columns.Item("grpoentry").Cells.Item(i).Specific.String) & " and T0.ObjType='18'")
                        If strQuery <> "" Then Continue For
                        strQuery = "Select T0.DocEntry [Sub Entry],T0.U_SItemCode,T0.U_SItemDesc,T1.*"
                        strQuery += vbCrLf + "from [@MIPL_OPOR] T0 left join "
                        strQuery += vbCrLf + "(Select A.DocNum [GRPO Num],A.DocEntry [GRPO Entry],B.ItemCode,B.U_NonItemCode,B.LineNum,B.Quantity,B.Price,B.TaxCode,B.WhsCode,B.AcctCode,B.BaseEntry,B.LineStatus"
                        strQuery += vbCrLf + "from OPDN A join PDN1 B on A.DocEntry=B.DocEntry where B.BaseType='22') T1 on T0.U_PurOrdrNo=T1.BaseEntry  and T0.U_SItemCode=T1.U_NonItemCode"
                        strQuery += vbCrLf + "where T1.LineStatus='O' and T0.DocEntry=" & Matrix0.Columns.Item("scentry").Cells.Item(i).Specific.String & " and T1.[GRPO Entry]='" & Matrix0.Columns.Item("grpoentry").Cells.Item(i).Specific.String & "' "
                        objRs.DoQuery(strQuery)
                        If objRs.RecordCount = 0 Then Exit Function
                        objAPInvoice.Lines.ItemCode = Trim(objRs.Fields.Item("ItemCode").Value.ToString)
                        objAPInvoice.Lines.Quantity = CDbl(Matrix0.Columns.Item("grnqty").Cells.Item(i).Specific.String) ' CDbl(objRs.Fields.Item("GRN Qty").Value.ToString) 
                        objAPInvoice.Lines.AccountCode = Trim(objRs.Fields.Item("AcctCode").Value.ToString)
                        objAPInvoice.Lines.TaxCode = Trim(objRs.Fields.Item("TaxCode").Value.ToString)
                        objAPInvoice.Lines.BaseType = 20 ' 22
                        objAPInvoice.Lines.BaseEntry = CInt(objRs.Fields.Item("GRPO Entry").Value.ToString) ' CInt(objRs.Fields.Item("PO Entry").Value.ToString)
                        objAPInvoice.Lines.BaseLine = CInt(objRs.Fields.Item("LineNum").Value.ToString)
                        objAPInvoice.Lines.UnitPrice = Trim(objRs.Fields.Item("Price").Value.ToString)
                        objAPInvoice.Lines.WarehouseCode = Trim(objRs.Fields.Item("WhsCode").Value.ToString)
                        objAPInvoice.Lines.UserFields.Fields.Item("U_NonItemCode").Value = Trim(objRs.Fields.Item("U_SItemCode").Value.ToString)
                        objAPInvoice.Lines.UserFields.Fields.Item("U_NonItemDesc").Value = Trim(objRs.Fields.Item("U_SItemDesc").Value.ToString)
                        objAPInvoice.Lines.UserFields.Fields.Item("U_GRNEntry").Value = Trim(Matrix0.Columns.Item("grnnum").Cells.Item(i).Specific.String)
                        objAPInvoice.Lines.Add()
                    Next
                    If objAPInvoice.Add() <> 0 Then
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objaddon.objapplication.SetStatusBarMessage("A/P Invoice: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        objaddon.objapplication.MessageBox("A/P Invoice: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode,, "OK")
                        Return False
                    Else
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objaddon.objcompany.GetNewObjectKey()
                        EditText5.Value = DocEntry
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                        objaddon.objapplication.StatusBar.SetText("A/P Invoice Draft Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Return True
                    End If
                Catch ex As Exception
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.StatusBar.SetText("Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try

                System.Runtime.InteropServices.Marshal.ReleaseComObject(objAPInvoice)
                GC.Collect()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Function

        Private Sub LinkedButton1_PressedBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles LinkedButton1.PressedBefore
            Try
                If EditText5.Value = "" Then Exit Sub
                Dim TEntry As String = objaddon.objglobalmethods.getSingleValue("Select * from ODRF where ObjType='18' and isnull(DocStatus,'')='O' and DocEntry='" & EditText5.Value & "'")
                LinkedButton1.LinkedObject = "-1"
                If TEntry <> "" Then
                    LinkedButton1.LinkedObject = "112"
                    StaticText4.Caption = "AP Draft Entry"
                Else
                    LinkedButton1.LinkedObject = "18"
                    StaticText4.Caption = "AP Doc Entry"
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button3_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button3.ClickBefore
            Try
                If Button3.Item.Enabled = False Then BubbleEvent = False : Exit Sub
                If EditText0.Value = "" Then
                    objaddon.objapplication.StatusBar.SetText("Vendor Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                If Create_AP_Invoice() Then
                    objaddon.objapplication.StatusBar.SetText("Invoice Draft Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button2_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button2.ClickBefore
            Try
                If Button2.Item.Enabled = False Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then BubbleEvent = False : Exit Sub
                If EditText5.Value <> "" Then
                    strQuery = objaddon.objglobalmethods.getSingleValue("Select DocStatus from ODRF where ObjType='18' and DocEntry =" & EditText5.Value & "")
                    If strQuery = "C" Then
                        objaddon.objapplication.StatusBar.SetText("Already Created Invoice Document...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    ElseIf strQuery = "O" Then
                        objaddon.objapplication.StatusBar.SetText("Already Created Invoice Draft...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If

                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                Matrix0.AutoResizeColumns()
                If EditText5.Value <> "" Then Button2.Item.Enabled = False
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_LinkPressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LinkPressedAfter
            Try
                If pVal.ColUID = "scentry" Then
                    Link_Value = Matrix0.Columns.Item("scentry").Cells.Item(pVal.Row).Specific.String 'DocEntry
                    Link_objtype = "SUBPO"
                    Dim activeform As New SubContractingPO
                    activeform.Show()
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If pVal.ActionSuccess = True And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "SCINV")
                    objform.Items.Item("tposdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                    objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
    End Class
End Namespace
