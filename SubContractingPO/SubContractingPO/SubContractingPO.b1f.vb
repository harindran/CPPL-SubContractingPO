Option Strict Off
Option Explicit On

Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Namespace SubContractingPO
    <FormAttribute("SUBCTPO", "SubContractingPO/SubContractingPO.b1f")>
    Friend Class SubContractingPO
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsHeader As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsDetails2 As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsDetails3 As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsDetails4 As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsDetails5 As SAPbouiCOM.DBDataSource
        Dim objRs As SAPbobsCOM.Recordset
        'Private Shared FormCount As Integer = 0
        Dim FormCount As Integer = 0
        Dim strQuery As String
        Public Shared CostCenter As String = ""
        Private WithEvents objcombo As SAPbouiCOM.ComboBox
        Public Sub New()

        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lblvendor").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("txtcode").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lblname").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("txtname").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lblctper").Specific, SAPbouiCOM.StaticText)
            Me.StaticText3 = CType(Me.GetItem("lblvrefno").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("txtvrefno").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("lbldocnum").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("txtdocnum").Specific, SAPbouiCOM.EditText)
            Me.StaticText5 = CType(Me.GetItem("lblstat").Specific, SAPbouiCOM.StaticText)
            Me.StaticText6 = CType(Me.GetItem("lblposdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("posdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText7 = CType(Me.GetItem("lbldeldate").Specific, SAPbouiCOM.StaticText)
            Me.EditText7 = CType(Me.GetItem("deldate").Specific, SAPbouiCOM.EditText)
            Me.StaticText8 = CType(Me.GetItem("lblsitem").Specific, SAPbouiCOM.StaticText)
            Me.EditText8 = CType(Me.GetItem("txtsitem").Specific, SAPbouiCOM.EditText)
            Me.StaticText9 = CType(Me.GetItem("lblvenbal").Specific, SAPbouiCOM.StaticText)
            Me.EditText9 = CType(Me.GetItem("txtvbal").Specific, SAPbouiCOM.EditText)
            Me.StaticText10 = CType(Me.GetItem("lbldocdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText10 = CType(Me.GetItem("docdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText11 = CType(Me.GetItem("lblponum").Specific, SAPbouiCOM.StaticText)
            Me.EditText11 = CType(Me.GetItem("txtponum").Specific, SAPbouiCOM.EditText)
            Me.StaticText12 = CType(Me.GetItem("lblvobstk").Specific, SAPbouiCOM.StaticText)
            Me.EditText12 = CType(Me.GetItem("txtvobstk").Specific, SAPbouiCOM.EditText)
            Me.StaticText13 = CType(Me.GetItem("lblitemqty").Specific, SAPbouiCOM.StaticText)
            Me.EditText13 = CType(Me.GetItem("txtopqty").Specific, SAPbouiCOM.EditText)
            Me.StaticText14 = CType(Me.GetItem("lblremark").Specific, SAPbouiCOM.StaticText)
            Me.EditText14 = CType(Me.GetItem("txtremark").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton0 = CType(Me.GetItem("lnkcode").Specific, SAPbouiCOM.LinkedButton)
            Me.LinkedButton1 = CType(Me.GetItem("lnkicode").Specific, SAPbouiCOM.LinkedButton)
            Me.ButtonCombo1 = CType(Me.GetItem("btngendoc").Specific, SAPbouiCOM.ButtonCombo)
            Me.LinkedButton2 = CType(Me.GetItem("lnkponum").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText22 = CType(Me.GetItem("lblInvTr").Specific, SAPbouiCOM.StaticText)
            Me.EditText21 = CType(Me.GetItem("TxtInvTr").Specific, SAPbouiCOM.EditText)
            Me.StaticText23 = CType(Me.GetItem("lblginum").Specific, SAPbouiCOM.StaticText)
            Me.EditText22 = CType(Me.GetItem("txtGINo").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("flrinput").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("flroutput").Specific, SAPbouiCOM.Folder)
            Me.Folder2 = CType(Me.GetItem("flrscrap").Specific, SAPbouiCOM.Folder)
            Me.Folder3 = CType(Me.GetItem("flrreldoc").Specific, SAPbouiCOM.Folder)
            Me.Matrix1 = CType(Me.GetItem("mtxreldoc").Specific, SAPbouiCOM.Matrix)
            Me.Matrix2 = CType(Me.GetItem("mtxscrap").Specific, SAPbouiCOM.Matrix)
            Me.Matrix3 = CType(Me.GetItem("mtxoutput").Specific, SAPbouiCOM.Matrix)
            Me.StaticText20 = CType(Me.GetItem("lblgrnum").Specific, SAPbouiCOM.StaticText)
            Me.EditText23 = CType(Me.GetItem("txtGRNo").Specific, SAPbouiCOM.EditText)
            Me.StaticText24 = CType(Me.GetItem("lpurnum").Specific, SAPbouiCOM.StaticText)
            Me.EditText24 = CType(Me.GetItem("SubPoNum").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton6 = CType(Me.GetItem("lnkpurnum").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText25 = CType(Me.GetItem("liqty").Specific, SAPbouiCOM.StaticText)
            Me.EditText25 = CType(Me.GetItem("SQty").Specific, SAPbouiCOM.EditText)
            Me.EditText26 = CType(Me.GetItem("txtentry").Specific, SAPbouiCOM.EditText)
            Me.Button3 = CType(Me.GetItem("btnload").Specific, SAPbouiCOM.Button)
            Me.Button4 = CType(Me.GetItem("BtnView").Specific, SAPbouiCOM.Button)
            Me.Button5 = CType(Me.GetItem("BtnInv").Specific, SAPbouiCOM.Button)
            Me.Button6 = CType(Me.GetItem("BtnGIssue").Specific, SAPbouiCOM.Button)
            Me.ButtonCombo0 = CType(Me.GetItem("btnOutput").Specific, SAPbouiCOM.ButtonCombo)
            Me.ButtonCombo2 = CType(Me.GetItem("BtnScrap").Specific, SAPbouiCOM.ButtonCombo)
            Me.StaticText15 = CType(Me.GetItem("lsitemd").Specific, SAPbouiCOM.StaticText)
            Me.EditText15 = CType(Me.GetItem("SItemDesc").Specific, SAPbouiCOM.EditText)
            Me.Folder4 = CType(Me.GetItem("flrcosting").Specific, SAPbouiCOM.Folder)
            Me.Matrix4 = CType(Me.GetItem("MtxCosting").Specific, SAPbouiCOM.Matrix)
            Me.EditText17 = CType(Me.GetItem("POLine").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("MtxinputN").Specific, SAPbouiCOM.Matrix)
            Me.Button7 = CType(Me.GetItem("AutoPO").Specific, SAPbouiCOM.Button)
            Me.Button8 = CType(Me.GetItem("BtnInvO").Specific, SAPbouiCOM.Button)
            Me.Button9 = CType(Me.GetItem("InvScrap").Specific, SAPbouiCOM.Button)
            Me.EditText27 = CType(Me.GetItem("tpoentry").Specific, SAPbouiCOM.EditText)
            Me.EditText18 = CType(Me.GetItem("tpodoc").Specific, SAPbouiCOM.EditText)
            Me.Button10 = CType(Me.GetItem("BtnJE").Specific, SAPbouiCOM.Button)
            Me.ComboBox0 = CType(Me.GetItem("Series").Specific, SAPbouiCOM.ComboBox)
            Me.CheckBox0 = CType(Me.GetItem("EditQty").Specific, SAPbouiCOM.CheckBox)
            Me.ComboBox1 = CType(Me.GetItem("txtctper").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText16 = CType(Me.GetItem("lInvUOM").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("tInvUom").Specific, SAPbouiCOM.EditText)
            Me.CheckBox1 = CType(Me.GetItem("clstat").Specific, SAPbouiCOM.CheckBox)
            Me.EditText19 = CType(Me.GetItem("txtbitem").Specific, SAPbouiCOM.EditText)
            Me.StaticText17 = CType(Me.GetItem("lblproc").Specific, SAPbouiCOM.StaticText)
            Me.Button11 = CType(Me.GetItem("btnbomref").Specific, SAPbouiCOM.Button)
            Me.EditText20 = CType(Me.GetItem("tprocseq").Specific, SAPbouiCOM.EditText)
            Me.EditText28 = CType(Me.GetItem("tproc").Specific, SAPbouiCOM.EditText)
            Me.StaticText18 = CType(Me.GetItem("lblowner").Specific, SAPbouiCOM.StaticText)
            Me.EditText29 = CType(Me.GetItem("townern").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton3 = CType(Me.GetItem("lnkowner").Specific, SAPbouiCOM.LinkedButton)
            Me.EditText30 = CType(Me.GetItem("townerc").Specific, SAPbouiCOM.EditText)
            Me.ComboBox2 = CType(Me.GetItem("txtstat").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText21 = CType(Me.GetItem("loscrap").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("toscrap").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler ActivateAfter, AddressOf Me.Form_ActivateAfter
            AddHandler CloseBefore, AddressOf Me.Form_CloseBefore
            AddHandler LayoutKeyBefore, AddressOf Me.Form_LayoutKeyBefore
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DataLoadBefore, AddressOf Me.Form_DataLoadBefore
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

        Private Sub OnCustomInitialize()
            Try
                'objform = objaddon.objapplication.Forms.GetForm("SUBCTPO", Me.FormCount)'SUBCTPO
                'objform = objaddon.objapplication.Forms.ActiveForm
                If Title <> "" Then
                    objform.Title = Title
                End If
                'objaddon.objglobalmethods.Update_UserFormSettings_UDF(objform, "-SUBCTPO", 1)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objform.Freeze(True)
                objaddon.objglobalmethods.setReport("PO Layout", Me.FormCount)
                'objaddon.objglobalmethods.setReport("DC Layout", Me.FormCount)
                Try
                    CreateDynamicUDF()
                Catch ex As Exception
                End Try
                If objaddon.HANA Then
                    CostCenter = objaddon.objglobalmethods.getSingleValue("select ""MDStyle"" from OADM")
                Else
                    CostCenter = objaddon.objglobalmethods.getSingleValue("select MDStyle from OADM")
                End If
                Matrix0.Columns.Item("cc1").Visible = False
                Matrix0.Columns.Item("cc2").Visible = False
                Matrix0.Columns.Item("cc3").Visible = False
                Matrix0.Columns.Item("cc4").Visible = False
                Matrix0.Columns.Item("cc5").Visible = False
                Matrix0.Columns.Item("tweight").Visible = False
                Matrix0.Columns.Item("cmblttype").Visible = False
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If CostCenter = "S" Then
                    Matrix0.Columns.Item("distrule").Visible = False
                    If objaddon.HANA Then
                        objRs.DoQuery("select 'cc'||""DimCode"" as ""Code"",* from ODIM where ""DimActive""='Y'")
                    Else
                        objRs.DoQuery("select CONCAT('cc',DimCode) as Code,* from ODIM where DimActive='Y'")
                    End If
                    If objRs.RecordCount > 0 Then
                        For i As Integer = 0 To objRs.RecordCount - 1
                            Matrix0.Columns.Item(objRs.Fields.Item("Code").Value.ToString).Visible = True
                            Matrix0.Columns.Item(objRs.Fields.Item("Code").Value.ToString).TitleObject.Caption = objRs.Fields.Item("DimDesc").Value.ToString
                            objRs.MoveNext()
                        Next
                    End If
                End If
                'EditText8.Item.Visible = False
                If objaddon.HANA Then
                    objRs.DoQuery("select * from ""@MIPL_GEN""")
                Else
                    objRs.DoQuery("select * from [@MIPL_GEN]")
                End If
                If objRs.RecordCount > 0 Then
                    SuperUser = objRs.Fields.Item("U_SUser").Value.ToString
                    POItem = objRs.Fields.Item("U_POItem").Value.ToString
                    GetDatePO = objRs.Fields.Item("U_DatePO").Value.ToString
                    Costing = objRs.Fields.Item("U_Costing").Value.ToString
                    AutoProduction = objRs.Fields.Item("U_AutoPO").Value.ToString
                    ReceiptAutoLoad = objRs.Fields.Item("U_RecLoad").Value.ToString
                    APAutoLoad = objRs.Fields.Item("U_APLoad").Value.ToString
                    AutoAssignItem = objRs.Fields.Item("U_AutoItem").Value.ToString
                    TypeInScrap = objRs.Fields.Item("U_Type").Value.ToString
                    SubPOScreen = objRs.Fields.Item("U_SubScreen").Value.ToString
                    ItemBOM = objRs.Fields.Item("U_ItemBOM").Value.ToString
                    VendorSPlPrice = objRs.Fields.Item("U_Price").Value.ToString
                    ToLocationInOutput = objRs.Fields.Item("U_ToWhseO").Value.ToString
                    ToWhseInvTransfer = objRs.Fields.Item("U_InvWhse").Value.ToString
                    TranList = objRs.Fields.Item("U_TranList").Value.ToString
                    ToWhseInInput = objRs.Fields.Item("U_ToWhse").Value.ToString
                    WPrice = objRs.Fields.Item("U_WPrice").Value.ToString
                    'PlanQtyUDFEditable = objRs.Fields.Item("U_Field1").Value.ToString
                    UDF0 = objRs.Fields.Item("U_UDF0").Value.ToString
                    UDF1 = objRs.Fields.Item("U_UDF1").Value.ToString
                    ValUDF0 = objRs.Fields.Item("U_Val0").Value.ToString
                    ValUDF1 = objRs.Fields.Item("U_Val1").Value.ToString
                    BPWhse = objRs.Fields.Item("U_BPWhse").Value.ToString
                    ProdOrderStat_Close = objRs.Fields.Item("U_StatPO").Value.ToString
                    BomWhse = objRs.Fields.Item("U_BomWhse").Value.ToString
                    ItemProcess = objRs.Fields.Item("U_Process").Value.ToString
                    BOMRefresh = objRs.Fields.Item("U_BomRef").Value.ToString
                    RowDelete = objRs.Fields.Item("U_RowDel").Value.ToString
                    ProcessValidation = objRs.Fields.Item("U_ProcVal").Value.ToString
                    RevPlanQty = objRs.Fields.Item("U_RevPlanQty").Value.ToString
                    GIFIFO = objRs.Fields.Item("U_GIFIFO").Value.ToString
                End If
                Button11.Item.Visible = False
                If SuperUser = objaddon.objcompany.UserName Then
                    If BOMRefresh = "Y" Then
                        Button11.Item.Visible = True
                    End If
                End If
                objform.EnableMenu("1283", False)
                objform.EnableMenu("1285", False) 'Restore
                objform.EnableMenu("1284", False)
                ManageAttributes()
                Matrix_Total()
                FieldEnabled_CommonSettings()
                If Link_Value <> "-1" And Link_objtype = "202" And NewLink = "-1" Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    EditText26.Item.Enabled = True   'EditText11
                    EditText26.Value = Link_Value
                    objform.ActiveItem = "docdate"
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    EditText26.Item.Enabled = False
                    Link_Value = "-1" : Link_objtype = "-1" : NewLink = "-1"
                    Exit Sub
                ElseIf Link_Value <> "-1" And Link_objtype = "SUBPO" And NewLink = "-1" Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    EditText26.Item.Enabled = True   'EditText11
                    EditText26.Value = Link_Value
                    objform.ActiveItem = "docdate"
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    EditText26.Item.Enabled = False
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    Link_Value = "-1" : Link_objtype = "-1" : NewLink = "-1"
                    Exit Sub
                End If
                'odbdsDetails = objform.DataSources.DBDataSources.Item(CType(1, Object))

                odbdsHeader = objform.DataSources.DBDataSources.Item("@MIPL_OPOR")
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_POR1")
                odbdsDetails2 = objform.DataSources.DBDataSources.Item("@MIPL_POR2")
                odbdsDetails3 = objform.DataSources.DBDataSources.Item("@MIPL_POR3")
                odbdsDetails4 = objform.DataSources.DBDataSources.Item("@MIPL_POR4")
                odbdsDetails5 = objform.DataSources.DBDataSources.Item("@MIPL_POR5")
                objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "SUBPO")
                'EditText4.Value = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_OPOR")
                'EditText26.Value = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_OPOR")
                objform.Items.Item("posdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                objform.Items.Item("deldate").Specific.string = Now.Date.ToString("dd/MM/yy")
                objform.Items.Item("docdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")
                If objaddon.HANA Then
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select ""MltpBrnchs"" from OADM")
                Else
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select MltpBrnchs from OADM")
                End If
                CheckBox0.Item.Height = CheckBox0.Item.Height + 3
                CheckBox0.Item.Width = CheckBox0.Item.Width + 5
                CheckBox1.Item.Visible = False ' CheckBox1.Item.Height = CheckBox1.Item.Height + 3
                CheckBox0.Item.Enabled = False
                ''FieldDisableInMatrix()
                objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                'objaddon.objapplication.Menus.Item("1300").Activate()
                objform.Settings.Enabled = True
                Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'objform.EnableMenu("779", True)
                objform.Freeze(False)

                objRs = Nothing


            Catch ex As Exception
                objform.Freeze(False)
            Finally
                objform.Freeze(False)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

#Region "FormItems"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
        Private WithEvents StaticText8 As SAPbouiCOM.StaticText
        Private WithEvents EditText8 As SAPbouiCOM.EditText
        Private WithEvents StaticText9 As SAPbouiCOM.StaticText
        Private WithEvents EditText9 As SAPbouiCOM.EditText
        Private WithEvents StaticText10 As SAPbouiCOM.StaticText
        Private WithEvents EditText10 As SAPbouiCOM.EditText
        Private WithEvents StaticText11 As SAPbouiCOM.StaticText
        Private WithEvents EditText11 As SAPbouiCOM.EditText
        Private WithEvents StaticText12 As SAPbouiCOM.StaticText
        Private WithEvents EditText12 As SAPbouiCOM.EditText
        Private WithEvents StaticText13 As SAPbouiCOM.StaticText
        Private WithEvents EditText13 As SAPbouiCOM.EditText
        Private WithEvents StaticText14 As SAPbouiCOM.StaticText
        Private WithEvents EditText14 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton
        Private WithEvents LinkedButton1 As SAPbouiCOM.LinkedButton
        Private WithEvents ButtonCombo1 As SAPbouiCOM.ButtonCombo
        Private WithEvents LinkedButton2 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText22 As SAPbouiCOM.StaticText
        Private WithEvents EditText21 As SAPbouiCOM.EditText
        Private WithEvents StaticText23 As SAPbouiCOM.StaticText
        Private WithEvents EditText22 As SAPbouiCOM.EditText
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents Folder2 As SAPbouiCOM.Folder
        Private WithEvents Folder3 As SAPbouiCOM.Folder
        Private WithEvents Matrix1 As SAPbouiCOM.Matrix
        Private WithEvents Matrix2 As SAPbouiCOM.Matrix
        Private WithEvents Matrix3 As SAPbouiCOM.Matrix
        Private WithEvents StaticText20 As SAPbouiCOM.StaticText
        Private WithEvents EditText23 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton5 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText24 As SAPbouiCOM.StaticText
        Private WithEvents EditText24 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton6 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText25 As SAPbouiCOM.StaticText
        Private WithEvents EditText25 As SAPbouiCOM.EditText
        Private WithEvents EditText26 As SAPbouiCOM.EditText
        Private WithEvents Button3 As SAPbouiCOM.Button
        Private WithEvents Button4 As SAPbouiCOM.Button
        Private WithEvents Button5 As SAPbouiCOM.Button
        Private WithEvents Button6 As SAPbouiCOM.Button
        Private WithEvents ButtonCombo0 As SAPbouiCOM.ButtonCombo
        Private WithEvents ButtonCombo2 As SAPbouiCOM.ButtonCombo
        Private WithEvents StaticText15 As SAPbouiCOM.StaticText
        Private WithEvents EditText15 As SAPbouiCOM.EditText
        Private WithEvents Folder4 As SAPbouiCOM.Folder
        Private WithEvents Matrix4 As SAPbouiCOM.Matrix
        Private WithEvents EditText16 As SAPbouiCOM.EditText
        Private WithEvents EditText17 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents Button7 As SAPbouiCOM.Button
        Private WithEvents Button8 As SAPbouiCOM.Button
        Private WithEvents Button9 As SAPbouiCOM.Button
        Private WithEvents EditText27 As SAPbouiCOM.EditText
        Private WithEvents EditText18 As SAPbouiCOM.EditText
        Private WithEvents Button10 As SAPbouiCOM.Button
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents ComboBox1 As SAPbouiCOM.ComboBox
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox
        Private WithEvents StaticText16 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText17 As SAPbouiCOM.StaticText
        Private WithEvents Button11 As SAPbouiCOM.Button
        Private WithEvents CheckBox1 As SAPbouiCOM.CheckBox
        Private WithEvents EditText19 As SAPbouiCOM.EditText
        Private WithEvents EditText20 As SAPbouiCOM.EditText
        Private WithEvents EditText28 As SAPbouiCOM.EditText
        Private WithEvents StaticText18 As SAPbouiCOM.StaticText
        Private WithEvents EditText29 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton3 As SAPbouiCOM.LinkedButton
        Private WithEvents EditText30 As SAPbouiCOM.EditText
#End Region

#Region "Header"

        Private Sub EditText0_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.LostFocusAfter
            Try
                Dim objrs As SAPbobsCOM.Recordset
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If EditText0.Value = "" Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If EditText0.Value <> "" Then    'Getting Vendor Balance
                        If objaddon.HANA Then
                            'EditText9.Value = objaddon.objglobalmethods.getSingleValue("select ""Balance"" from OCRD where ""CardCode""='" & EditText0.Value & "'")
                            odbdsHeader.SetValue("U_VBalnce", 0, objaddon.objglobalmethods.getSingleValue("select ""Balance"" from OCRD where ""CardCode""='" & EditText0.Value & "'"))
                        Else
                            'EditText9.Value = objaddon.objglobalmethods.getSingleValue("select Balance from OCRD where CardCode='" & EditText0.Value & "'")
                            odbdsHeader.SetValue("U_VBalnce", 0, objaddon.objglobalmethods.getSingleValue("select Balance from OCRD where CardCode='" & EditText0.Value & "'"))
                        End If
                    End If
                    If EditText0.Value = "" Then
                        If EditText1.Value <> "" Then
                            EditText1.Value = ""
                        End If
                    End If
                    Dim VNamesize As Size = TextRenderer.MeasureText(EditText1.Value, New Font("Arial", 12.0F))
                    If VNamesize.Width <= 135 Then
                        EditText1.Item.Width = 135
                    Else
                        EditText1.Item.Width = VNamesize.Width
                    End If
                    If objaddon.HANA Then
                        objrs.DoQuery("select ""Name"" from  OCPR where ""CardCode""='" & EditText0.Value & "'")
                        'EditText2.Value = objaddon.objglobalmethods.getSingleValue("select ""Name"" from  OCPR where ""CardCode""='" & EditText0.Value & "'")
                    Else
                        objrs.DoQuery("select Name from  OCPR where CardCode='" & EditText0.Value & "'")
                        'EditText2.Value = objaddon.objglobalmethods.getSingleValue("select Name from  OCPR where CardCode='" & EditText0.Value & "'")
                    End If
                    If ComboBox1.ValidValues.Count > 0 Then
                        For i As Integer = ComboBox1.ValidValues.Count - 1 To 0 Step -1
                            ComboBox1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                    If objrs.RecordCount > 0 Then
                        For Rec As Integer = 0 To objrs.RecordCount - 1
                            ComboBox1.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
                            objrs.MoveNext()
                        Next
                    End If
                    ComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                    objrs = Nothing
                End If

                'If objaddon.HANA Then
                '    objrs.DoQuery("select ""Name"" from  OCPR where ""CardCode""='" & EditText0.Value & "'")
                '    'EditText2.Value = objaddon.objglobalmethods.getSingleValue("select ""Name"" from  OCPR where ""CardCode""='" & EditText0.Value & "'")
                'Else
                '    objrs.DoQuery("select Name from  OCPR where CardCode='" & EditText0.Value & "'")
                '    'EditText2.Value = objaddon.objglobalmethods.getSingleValue("select Name from  OCPR where CardCode='" & EditText0.Value & "'")
                'End If
                ''ComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                'If ComboBox1.ValidValues.Count > 0 Then
                '    For i As Integer = ComboBox1.ValidValues.Count - 1 To 0 Step -1
                '        ComboBox1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                '    Next
                'End If
                'If objrs.RecordCount > 0 Then
                '    For Rec As Integer = 0 To objrs.RecordCount - 1
                '        ComboBox1.ValidValues.Add(Rec, objrs.Fields.Item(0).Value.ToString)
                '        objrs.MoveNext()
                '    Next
                'End If
                objrs = Nothing
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText8_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText8.LostFocusAfter
            'Try
            '    If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then Exit Sub
            '    Dim objrs As SAPbobsCOM.Recordset
            '    Dim strsql As String = "", GetValue As String = "", WhsCode As String = "", InvUOM As String = ""
            '    Dim i As Integer = 0
            '    objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    If objaddon.HANA Then
            '        EditText15.Value = objaddon.objglobalmethods.getSingleValue("SELECT  ""ItemName"" FROM OITM where ""ItemCode""='" & EditText8.Value & "'")
            '    Else
            '        EditText15.Value = objaddon.objglobalmethods.getSingleValue("SELECT  ItemName FROM OITM where ItemCode='" & EditText8.Value & "'")
            '    End If
            '    Dim ItemDescsize As Size = TextRenderer.MeasureText(EditText15.Value, New Font("Arial", 12.0F))
            '    If ItemDescsize.Width <= 135 Then
            '        EditText15.Item.Width = 135
            '    Else
            '        EditText15.Item.Width = ItemDescsize.Width
            '    End If
            '    If ItemBOM = "Y" Then  'Loading SUB-BOM
            '        If EditText0.Value = "" Then objaddon.objapplication.StatusBar.SetText("Please Select Vendor..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Sub
            '        If Matrix0.VisualRowCount = 0 Then objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
            '        If Matrix0.Columns.Item("Item1").Cells.Item(1).Specific.String <> "" Then
            '            If EditText8.Value = Matrix0.Columns.Item("Item1").Cells.Item(1).Specific.String Then
            '                Exit Sub
            '            End If
            '        End If
            '        If EditText8.Value <> "" Then
            '            If objaddon.HANA Then
            '                EditText2.Value = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & EditText8.Value & "'")
            '            Else
            '                EditText2.Value = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & EditText8.Value & "'")
            '            End If
            '            If objaddon.HANA Then
            '                strsql = "select T1.""U_Itemcode"",T1.""U_ItemDesc"",T1.""U_Qty"",T1.""U_Whse"",T1.""U_Unitprice"",T1.""U_Total"" "
            '                strsql += vbCrLf + "from ""@MIPL_OBOM"" T0 join ""@MIPL_BOM1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & EditText8.Value & "'"
            '            Else
            '                strsql = "select T1.U_Itemcode,T1.U_ItemDesc,T1.U_Qty,T1.U_Whse,T1.U_Unitprice,T1.U_Total "
            '                strsql += vbCrLf + "from [@MIPL_OBOM] T0 join [@MIPL_BOM1] T1 on T0.Code=T1.Code where T0.Code='" & EditText8.Value & "'"
            '            End If
            '            objrs.DoQuery(strsql)
            '            If AutoProduction = "Y" Then
            '                If objaddon.HANA Then
            '                    WhsCode = objaddon.objglobalmethods.getSingleValue("select ""U_WAREHOUSE"" from OCRD where ""CardCode""='" & EditText0.Value & "'")
            '                Else
            '                    WhsCode = objaddon.objglobalmethods.getSingleValue("select U_WAREHOUSE from OCRD where CardCode='" & EditText0.Value & "'")
            '                End If
            '                'ElseIf ToWhseInInput = "Y" Then
            '                '    If objaddon.HANA Then
            '                '        WhsCode = objaddon.objglobalmethods.getSingleValue("Select ""DflWhs"" from OBPL where ""BPLId""=(select  distinct T0.""BPLid"" from OWHS T0 left join OLCT T1 on T0.""Location""=T1.""Code"" where T1.""Code""=(Select ""U_LCode"" from ""@MIPL_GEN"") and ifnull(T0.""Inactive"",'') ='N'and T0.""BPLid"" is not null)")
            '                '    Else
            '                '        WhsCode = objaddon.objglobalmethods.getSingleValue("Select DflWhs from OBPL where BPLId=(select  distinct T0.BPLid from OWHS T0 left join OLCT T1 on T0.Location=T1.Code where T1.Code=(Select U_LCode from [@MIPL_GEN]) and isnull(T0.Inactive,'') ='N' and T0.BPLid is not null)")
            '                '    End If
            '            End If
            '            objform.Freeze(True)
            '            odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_POR1")
            '            Matrix0.Clear()
            '            odbdsDetails.Clear()
            '            If objrs.RecordCount > 0 Then
            '                objaddon.objapplication.StatusBar.SetText("Loading from SUB-BOM Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '                While Not objrs.EoF
            '                    Matrix0.AddRow()
            '                    'odbdsDetails.Clear()
            '                    Matrix0.GetLineData(Matrix0.VisualRowCount)
            '                    odbdsDetails.SetValue("LineId", 0, i + 1)
            '                    odbdsDetails.SetValue("U_Item1", 0, EditText8.Value.ToString)
            '                    odbdsDetails.SetValue("U_Itemcode", 0, objrs.Fields.Item("U_Itemcode").Value.ToString)
            '                    odbdsDetails.SetValue("U_ItemDesc", 0, objrs.Fields.Item("U_ItemDesc").Value.ToString)
            '                    odbdsDetails.SetValue("U_Qty", 0, objrs.Fields.Item("U_Qty").Value.ToString)
            '                    odbdsDetails.SetValue("U_PlanQty", 0, CDbl(objrs.Fields.Item("U_Qty").Value.ToString) * 1)
            '                    Try
            '                        If objaddon.HANA Then
            '                            InvUOM = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "'")
            '                        Else
            '                            InvUOM = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "'")
            '                        End If
            '                    Catch ex As Exception
            '                    End Try
            '                    odbdsDetails.SetValue("U_InvUom", 0, InvUOM)
            '                    Dim Price As String = ""
            '                    If VendorSPlPrice = "Y" Then
            '                        If objaddon.HANA Then
            '                            Price = objaddon.objglobalmethods.getSingleValue("Select ""AvgPrice"" from OITW where ""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and ""WhsCode""='" & objrs.Fields.Item("U_Whse").Value.ToString & "';")
            '                        Else
            '                            Price = objaddon.objglobalmethods.getSingleValue("Select AvgPrice from OITW where ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and WhsCode='" & objrs.Fields.Item("U_Whse").Value.ToString & "'")
            '                        End If
            '                        If Price <> "0" Then
            '                            odbdsDetails.SetValue("U_Price", 0, Price)
            '                        Else
            '                            odbdsDetails.SetValue("U_Price", 0, objrs.Fields.Item("U_Unitprice").Value.ToString)
            '                        End If
            '                    Else
            '                        odbdsDetails.SetValue("U_Price", 0, objrs.Fields.Item("U_Unitprice").Value.ToString)
            '                    End If
            '                    odbdsDetails.SetValue("U_WhsCode", 0, objrs.Fields.Item("U_Whse").Value.ToString)
            '                    odbdsDetails.SetValue("U_LineTot", 0, CDbl(odbdsDetails.GetValue("U_Qty", 0)) * CDbl(odbdsDetails.GetValue("U_Price", 0)))
            '                    odbdsDetails.SetValue("U_SubWhse", 0, WhsCode)
            '                    Matrix0.SetLineData(Matrix0.VisualRowCount)
            '                    objrs.MoveNext()
            '                    i += 1
            '                End While
            '                objaddon.objapplication.StatusBar.SetText("Loaded from SUB-BOM Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '            Else
            '                objaddon.objapplication.StatusBar.SetText("No records found from SUB_BOM...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '            End If

            '        Else
            '            Exit Sub
            '        End If
            '    End If
            '    For j As Integer = 1 To Matrix0.VisualRowCount
            '        If Matrix0.Columns.Item("Code").Cells.Item(j).Specific.String <> "" Then
            '            If objaddon.HANA Then
            '                GetValue = objaddon.objglobalmethods.getSingleValue("select T1.""ChapterID"" from OITM T0 join OCHP T1 on T0.""ChapterID""=T1.""AbsEntry"" where T0.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(j).Specific.String & "'")
            '            Else
            '                GetValue = objaddon.objglobalmethods.getSingleValue("select T1.ChapterID from OITM T0 join OCHP T1 on T0.ChapterID=T1.AbsEntry where T0.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(j).Specific.String & "'")
            '            End If
            '            If GetValue <> "" Then
            '                Matrix0.Columns.Item("HSN").Cells.Item(j).Specific.String = GetValue
            '            Else
            '                Matrix0.Columns.Item("HSN").Cells.Item(j).Specific.String = 0
            '            End If
            '            If Matrix0.Columns.Item("Whse").Cells.Item(j).Specific.String <> "" Then
            '                If objaddon.HANA Then
            '                    GetValue = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  OITW where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(j).Specific.String & "' and ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
            '                Else
            '                    GetValue = objaddon.objglobalmethods.getSingleValue("select OnHand from  OITW where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(j).Specific.String & "' and WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
            '                End If
            '                Matrix0.Columns.Item("Instock").Cells.Item(j).Specific.String = GetValue
            '            Else
            '                Matrix0.Columns.Item("Instock").Cells.Item(j).Specific.String = 0
            '            End If
            '            If Matrix0.Columns.Item("Code").Cells.Item(j).Specific.String <> "" Then
            '                Matrix0.Columns.Item("Item1").Cells.Item(j).Specific.String = EditText8.Value
            '            End If
            '        End If
            '    Next
            '    objrs = Nothing
            '    objaddon.objapplication.Menus.Item("1300").Activate()
            '    'objform.ActiveItem = "SQty"
            '    objform.Freeze(False)
            'Catch ex As Exception
            '    objform.Freeze(False)
            'End Try
        End Sub

        Private Sub EditText25_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText25.LostFocusAfter
            Try  'Calculate Planned Qty in Input Tab
                'CalcPlanQty()

            Catch ex As Exception
            End Try

        End Sub

        Private Sub LinkedButton1_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles LinkedButton1.PressedAfter
            Try
                If EditText8.Value = "" Then Exit Sub
                If ItemBOM = "Y" Then 'Getting Sub-BOM Screen
                    EditText19.Item.LinkTo = "SUBBOM"
                    Link_Value = EditText19.Value : Link_objtype = "SUBBOM"
                    'Dim activeform As New SubContract
                    'activeform.Show()
                    If Not objaddon.FormExist("BOM") Then
                        Dim activeform As New SubContract
                        activeform.Show()
                        'Else
                        '    Try
                        '        Dim cflForm As SAPbouiCOM.Form
                        '        If objaddon.objapplication.Forms.Count > 0 Then
                        '            For frm As Integer = 0 To objaddon.objapplication.Forms.Count - 1
                        '                If objaddon.objapplication.Forms.Item(frm).UniqueID = "SUBBOM" Then
                        '                    cflForm = objaddon.objapplication.Forms.Item("SUBBOM")
                        '                    cflForm.Close()
                        '                    Exit For
                        '                End If
                        '            Next
                        '            GoTo Label1
                        '        End If
                        '    Catch ex As Exception
                        '    End Try
                    End If
                Else                   ' Getting Item Master Screen
                    EditText19.Item.LinkTo = "4"
                    LinkedButton1.LinkedObjectType = "4"
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Button0_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If pVal.ActionSuccess = True And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "SUBPO")
                    'EditText4.Value = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_OPOR")
                    'EditText26.Value = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_OPOR")
                    objform.Items.Item("posdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                    objform.Items.Item("deldate").Specific.string = Now.Date.ToString("dd/MM/yy")
                    objform.Items.Item("docdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                    objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")
                End If
            Catch ex As Exception
            End Try


        End Sub

        Private Sub EditText0_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText0.ChooseFromListBefore
            'CardCode CFL Condition
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_V")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "S"
                ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub EditText1_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText1.ChooseFromListBefore
            'CardName CFL Condition
            'If pVal.ActionSuccess = True Then Exit Sub
            'Try
            '    Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_VN")
            '    Dim oConds As SAPbouiCOM.Conditions
            '    Dim oCond As SAPbouiCOM.Condition
            '    Dim oEmptyConds As New SAPbouiCOM.Conditions
            '    oCFL.SetConditions(oEmptyConds)
            '    oConds = oCFL.GetConditions()

            '    oCond = oConds.Add()
            '    oCond.Alias = "CardType"
            '    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCond.CondVal = "S"
            '    ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        
            '    oCFL.SetConditions(oConds)
            'Catch ex As Exception
            '    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'End Try

        End Sub

        Private Sub EditText1_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.ChooseFromListAfter
            'CardName CFL Condition After
            'Try
            '    If pVal.ActionSuccess = False Then Exit Sub
            '    pCFL = pVal
            '    If Not pCFL.SelectedObjects Is Nothing Then
            '        Try
            '            EditText0.Value = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
            '        Catch ex As Exception
            '        End Try
            '        Try
            '            EditText1.Value = pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value
            '        Catch ex As Exception
            '        End Try
            '    End If
            '    If objaddon.HANA Then
            '        EditText2.Value = objaddon.objglobalmethods.getSingleValue("select ""Name"" from  OCPR where ""CardCode""='" & EditText0.Value & "'")
            '    Else
            '        EditText2.Value = objaddon.objglobalmethods.getSingleValue("select Name from  OCPR where CardCode='" & EditText0.Value & "'")
            '    End If

            'Catch ex As Exception
            'Finally
            '    GC.Collect()
            '    GC.WaitForPendingFinalizers()
            'End Try

        End Sub

        Private Sub EditText0_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.ChooseFromListAfter
            Try
                'Dim objrs As SAPbobsCOM.Recordset
                'objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        'EditText0.Value = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                        odbdsHeader.SetValue("U_CardCode", 0, pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        'EditText1.Value = pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value
                        odbdsHeader.SetValue("U_CardName", 0, pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    'objform.Update()
                End If
                'Dim Fieldsize As Size = TextRenderer.MeasureText(EditText1.Value, New Font("Arial", 12.0F))
                'If Fieldsize.Width <= 135 Then
                '    EditText1.Item.Width = 135
                'Else
                '    EditText1.Item.Width = Fieldsize.Width
                'End If
                'If objaddon.HANA Then
                '    objrs.DoQuery("select ""Name"" from  OCPR where ""CardCode""='" & EditText0.Value & "'")
                '    'EditText2.Value = objaddon.objglobalmethods.getSingleValue("select ""Name"" from  OCPR where ""CardCode""='" & EditText0.Value & "'")
                'Else
                '    objrs.DoQuery("select Name from  OCPR where CardCode='" & EditText0.Value & "'")
                '    'EditText2.Value = objaddon.objglobalmethods.getSingleValue("select Name from  OCPR where CardCode='" & EditText0.Value & "'")
                'End If
                'If ComboBox1.ValidValues.Count > 0 Then
                '    For i As Integer = ComboBox1.ValidValues.Count - 1 To 0 Step -1
                '        ComboBox1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                '    Next
                'End If
                'If objrs.RecordCount > 0 Then
                '    For Rec As Integer = 0 To objrs.RecordCount - 1
                '        ComboBox1.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(0).Value.ToString)
                '        objrs.MoveNext()
                '    Next
                'End If
                'ComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

                'objrs = Nothing
            Catch ex As Exception
                'Finally
                'GC.Collect()
                'GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub EditText8_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText8.ChooseFromListBefore
            'Sub-Con ItemCode CFL Condition
            If pVal.ActionSuccess = True Then Exit Sub
            'Try
            '    Dim oEdit As SAPbouiCOM.EditText
            '    oEdit = objform.Items.Item("txtsitem").Specific
            '    Dim oConds As SAPbouiCOM.Conditions
            '    Dim oCond As SAPbouiCOM.Condition
            '    Dim oEmptyConds As New SAPbouiCOM.Conditions
            '    Dim oCFL As SAPbouiCOM.ChooseFromList

            '    If ItemBOM <> "Y" Then   'From SUB_BOM
            '        oCFL = objform.ChooseFromLists.Item("CFL_ItemM")
            '        oCFL.SetConditions(oEmptyConds)
            '        oConds = oCFL.GetConditions()
            '        oEdit.ChooseFromListUID = "CFL_ItemM"
            '        oEdit.ChooseFromListAlias = "ItemCode"
            '        oCond = oConds.Add()
            '        oCond.Alias = "U_SubCont"
            '        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCond.CondVal = "Y"
            '        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            '        oCond = oConds.Add()
            '        oCond.Alias = "validFor"
            '        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCond.CondVal = "Y"
            '    Else                    'From Item Master
            '        oCFL = objform.ChooseFromLists.Item("CFL_Item")
            '        oCFL.SetConditions(oEmptyConds)
            '        oConds = oCFL.GetConditions()
            '        oEdit.ChooseFromListUID = "CFL_Item"
            '        oEdit.ChooseFromListAlias = "Code"
            '    End If
            '    oCFL.SetConditions(oConds)
            'Catch ex As Exception
            '    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'End Try
        End Sub

        Private Sub EditText8_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText8.ChooseFromListAfter
            'Sub-Con ItemCode CFL Condition After
            '            Try
            '                If pVal.ActionSuccess = False Then Exit Sub
            '                pCFL = pVal
            '                If Not pCFL.SelectedObjects Is Nothing Then
            '                    If ItemBOM <> "Y" Then
            '                        Try
            '                            EditText8.Value = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
            '                        Catch ex As Exception
            '                        End Try
            '                    Else
            '                        Try
            '                            EditText19.Value = pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value
            '                        Catch ex As Exception
            '                        End Try
            '                        Try
            '                            EditText8.Value = pCFL.SelectedObjects.Columns.Item("U_ItemCode").Cells.Item(0).Value
            '                        Catch ex As Exception
            '                        End Try
            '                    End If
            '                End If
            '                Dim objrs As SAPbobsCOM.Recordset
            '                objrs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            '                If objaddon.HANA Then
            '                    objrs.DoQuery("select T2.""U_Type"",(Select distinct T1.""Descr"" from CUFD T0 join UFD1 T1 on T0.""TableID""=T1.""TableID"" " &
            '" where T0.""TableID""='@MIPL_OBOM' and T1.""FldValue""=T2.""U_Type"") as ""Desc""  from ""@MIPL_OBOM"" T2 where T2.""U_ItemCode""='" & EditText8.Value & "'")
            '                Else
            '                    objrs.DoQuery("select T2.U_Type,(Select distinct T1.Descr from CUFD T0 join UFD1 T1 on T0.TableID=T1.TableID " &
            '" where T0.TableID='@MIPL_OBOM' and T1.FldValue=T2.U_Type) as Desc  from [@MIPL_OBOM] T2 where T2.U_ItemCode='" & EditText8.Value & "'")
            '                End If
            '                ComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            '                If ComboBox2.ValidValues.Count > 0 Then
            '                    For i As Integer = ComboBox2.ValidValues.Count - 1 To 0 Step -1
            '                        ComboBox2.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            '                    Next
            '                End If
            '                If objrs.RecordCount > 0 Then
            '                    For Rec As Integer = 0 To objrs.RecordCount - 1
            '                        ComboBox2.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
            '                        objrs.MoveNext()
            '                    Next
            '                    ComboBox2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            '                End If
            '                objrs = Nothing
            '                Dim Fieldsize As Size = TextRenderer.MeasureText(EditText15.Value, New Font("Arial", 12.0F))
            '                If Fieldsize.Width <= 135 Then
            '                    EditText15.Item.Width = 135
            '                Else
            '                    EditText15.Item.Width = Fieldsize.Width
            '                End If
            '                EditText25.Value = "1"

            '            Catch ex As Exception
            '            Finally
            '                GC.Collect()
            '                GC.WaitForPendingFinalizers()
            '            End Try
        End Sub

        Private Sub EditText11_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText11.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                EditText27.Value = ""
                If POItem <> "Y" Then  'Production Order CFL Condition
                    EditText11.ChooseFromListUID = "CFL_PO"
                    EditText11.ChooseFromListAlias = "DocEntry"
                    Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_PO")
                    Dim oConds As SAPbouiCOM.Conditions
                    Dim oCond As SAPbouiCOM.Condition
                    Dim oEmptyConds As New SAPbouiCOM.Conditions
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()

                    'oCond = oConds.Add
                    'oCond.Alias = "Status"
                    'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    'oCond.CondVal = "P"
                    'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCond = oConds.Add
                    oCond.Alias = "Status"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "R"
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCond = oConds.Add
                    oCond.Alias = "Status"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                    oCond.CondVal = "C"
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCond = oConds.Add
                    oCond.Alias = "Status"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                    oCond.CondVal = "L"
                    oCFL.SetConditions(oConds)
                End If
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub EditText11_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText11.ChooseFromListAfter
            'Production Order CFL Condition After
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If ItemProcess = "Y" Then
                    If EditText19.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Sub-Item is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Exit Sub
                    End If
                    If EditText20.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Sub-Item Process is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Exit Sub
                    End If
                End If
                If Not pCFL.SelectedObjects Is Nothing Then
                    If ItemProcess = "Y" Then
                        If objaddon.HANA Then
                            strQuery = objaddon.objglobalmethods.getSingleValue("Select Count(*) from ""@MIPL_OPOR"" T0 where T0.""U_PONum""='" & pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value & "' and T0.""U_BOMCode""='" & objform.Items.Item("txtbitem").Specific.string & "' and T0.""U_Process""='" & objform.Items.Item("tproc").Specific.string & "'")
                        Else
                            strQuery = objaddon.objglobalmethods.getSingleValue("Select Count(*) from [@MIPL_OPOR] T0 where T0.U_PONum='" & pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value & "' and T0.U_BOMCode='" & objform.Items.Item("txtbitem").Specific.string & "' and T0.U_Process='" & objform.Items.Item("tproc").Specific.string & "'")
                        End If
                        If CDbl(strQuery) > 0 Then
                            If objaddon.objapplication.MessageBox("Already created process in Sub Contracting against Production order. Do you want to Continue?", 2, "Yes", "No") <> 1 Then objform.Items.Item("txtponum").Specific.string = "" : objform.Items.Item("tpoentry").Specific.string = "" : Exit Sub 'BubbleEvent = False :
                            'objaddon.objapplication.MessageBox("Already created process in Sub Contracting against Production order", , "OK")
                        End If
                    End If
                    Try
                        'EditText11.Value = pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value
                        odbdsHeader.SetValue("U_PONum", 0, pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        'EditText27.Value = pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value
                        odbdsHeader.SetValue("U_PODoc", 0, pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                End If

            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub EditText24_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText24.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_PurOdr")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                If POItem = "Y" Then   'Purchase Order CFL Condition
                    If EditText0.Value = "" Then
                        oCond = oConds.Add()
                        oCond.Alias = "DocStatus"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "O"
                    Else
                        oCond = oConds.Add()
                        oCond.Alias = "CardCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = EditText0.Value
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        oCond = oConds.Add()
                        oCond.Alias = "DocStatus"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "O"
                    End If
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCond = oConds.Add()
                    oCond.Alias = "DocType"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "S"
                    If EditText24.Value <> "" Then
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        oCond = oConds.Add()
                        oCond.Alias = "DocEntry"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = EditText24.Value
                    End If
                Else
                    oCond = oConds.Add()
                    oCond.Alias = "DocStatus"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "O"
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCond = oConds.Add()
                    oCond.Alias = "CardCode"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = EditText0.Value
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCond = oConds.Add()
                    oCond.Alias = "DocType"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "I"
                End If
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub EditText24_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText24.ChooseFromListAfter
            'Purchase Order CFL Condition After
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        'EditText24.Value = pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value
                        odbdsHeader.SetValue("U_PurOrdrNo", 0, pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        'EditText18.Value = pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value
                        odbdsHeader.SetValue("U_PurEnt", 0, pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                End If

            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
            Try
                If pVal.InnerEvent = True Then BubbleEvent = False : Exit Sub
                'objform.Refresh()
                'objform.Update()
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If ItemProcess = "Y" Then
                        If EditText28.Value = "" Then
                            objaddon.objapplication.SetStatusBarMessage("Sub-Item Process is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        If objaddon.HANA Then
                            strQuery = objaddon.objglobalmethods.getSingleValue("Select Count(*) from ""@MIPL_OPOR"" T0 where T0.""U_PONum""='" & objform.Items.Item("txtponum").Specific.string & "' and T0.""U_BOMCode""='" & objform.Items.Item("txtbitem").Specific.string & "' and T0.""U_Process""='" & objform.Items.Item("tproc").Specific.string & "'")
                        Else
                            strQuery = objaddon.objglobalmethods.getSingleValue("Select Count(*) from [@MIPL_OPOR] T0 where T0.U_PONum='" & objform.Items.Item("txtponum").Specific.string & "' and T0.U_BOMCode='" & objform.Items.Item("txtbitem").Specific.string & "' and T0.U_Process='" & objform.Items.Item("tproc").Specific.string & "'")
                        End If
                        If CDbl(strQuery) > 0 Then
                            'If objaddon.objapplication.MessageBox("Already created process in Sub Contracting against Production order. Do you want to Continue?", 2, "Yes", "No") <> 1 Then Exit Sub 'BubbleEvent = False :
                            'objaddon.objapplication.MessageBox("Already created process in Sub Contracting against Production order", , "OK")
                        End If
                        'If EditText11.Value = "" Then
                        '    objaddon.objapplication.SetStatusBarMessage("Since Item Process is enabled. Please update Production Order...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '    BubbleEvent = False : Exit Sub
                        'End If
                    End If
                    If AutoProduction = "Y" Then
                        Dim Stat As String = ""
                        If objaddon.HANA Then
                            Stat = objaddon.objglobalmethods.getSingleValue("select  case when count(*)>1 then 1 else 0 end as ""Status"" from ""@MIPL_OPOR"" where ""U_PurOrdrNo""='" & EditText24.Value & "' and ""U_PurEnt""='" & EditText18.Value & "' and ""U_POLine""='" & EditText17.Value & "' ")
                        Else
                            Stat = objaddon.objglobalmethods.getSingleValue("select  case when count(*)>1 then 1 else 0 end as Status from [@MIPL_OPOR] where U_PurOrdrNo='" & EditText24.Value & "' and U_PurEnt='" & EditText18.Value & "' and U_POLine='" & EditText17.Value & "'")
                        End If
                        If Stat = "1" Then
                            objaddon.objapplication.SetStatusBarMessage("Duplicate PO found.Please Change...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            EditText17.Item.Click()
                            BubbleEvent = False : Exit Sub
                        End If
                    Else
                        'If EditText24.Value = "" Then
                        '    objaddon.objapplication.SetStatusBarMessage("Purchase Order is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        '    'BubbleEvent = False : Exit Sub
                        'End If
                    End If
                    If EditText0.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Vendor Code is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    If EditText7.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Delivery Date is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If

                    Try
                        Dim DocDate As Date = Date.ParseExact(EditText6.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        Dim DelDate As Date = Date.ParseExact(EditText7.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If DateTime.Compare(DocDate, DelDate) > 0 Then
                            objaddon.objapplication.StatusBar.SetText("In ""Due Date"" field, enter date that is equal to or later than posting date.  Field: Delivery Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            EditText10.Item.Click()
                            BubbleEvent = False : Exit Sub
                        End If
                    Catch ex As Exception
                    End Try
                    If EditText8.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Sub Contracting Item Code is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    'If AutoProduction <> "Y" Then
                    '    If EditText11.Value = "" Then
                    '        objaddon.objapplication.SetStatusBarMessage("Production Order Number is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    '        BubbleEvent = False : Exit Sub
                    '    End If
                    'End If
                    If EditText25.Value = "" Or EditText25.Value = 0 Then
                        objaddon.objapplication.SetStatusBarMessage("Sub Item Quantity is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    If Matrix0.Columns.Item("Code").Cells.Item(1).Specific.string = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Input Tab-> Line Data is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If Matrix0.Columns.Item("Code").Cells.Item(i).Specific.string <> "" Then
                            If Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.string = "" Or Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.string = "" Then
                                objaddon.objapplication.SetStatusBarMessage("One of the column value is missing in Input Tab.Please update", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                            If CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(i).Specific.String) = 0 Or CDbl(Matrix0.Columns.Item("bomqty").Cells.Item(i).Specific.String) = 0 Or CDbl(Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String) = 0 Or CDbl(Matrix0.Columns.Item("actplanqty").Cells.Item(i).Specific.String) = 0 Then
                                objaddon.objapplication.SetStatusBarMessage("Qty Field should not be zero. in Input Tab-Line " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                            If Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.string = Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.string Then
                                objaddon.objapplication.SetStatusBarMessage("Receipt warehouse cannot be identical to the release warehouse. in Input Tab-Line " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next
                    If ReceiptAutoLoad <> "Y" Then
                        Dim Quantity As Double
                        Dim Itemcount As Integer = 0
                        If Matrix3.Columns.Item("Code").Cells.Item(1).Specific.string = "" Then
                            objaddon.objapplication.SetStatusBarMessage("Output Tab-> Line Data is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        For i As Integer = 1 To Matrix3.VisualRowCount
                            If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string <> "" Then
                                If Matrix3.Columns.Item("Whse").Cells.Item(i).Specific.string = "" Then
                                    objaddon.objapplication.SetStatusBarMessage("One of the column value is missing in Output Tab.Please update", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                            If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String Then
                                Quantity = Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String
                                'objaddon.objapplication.SetStatusBarMessage("Quantity mismatch For Sub Item...Output Tab Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                'BubbleEvent = False : Exit Sub
                            End If
                        Next
                        For i As Integer = 1 To Matrix3.VisualRowCount
                            'If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String And Quantity > CDbl(objform.Items.Item("SQty").Specific.String) Then
                            '    objaddon.objapplication.SetStatusBarMessage("Output Tab-> Quantity doesnot exceed of Sub Item quantity  ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            '    BubbleEvent = False : Exit Sub
                            'End If
                            If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String Then
                                Itemcount += 1
                            End If
                        Next
                        If Itemcount = 0 Then
                            objaddon.objapplication.SetStatusBarMessage("Output Tab-> Sub Item missing. Please update.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        If Matrix0.VisualRowCount <= 0 And Matrix3.VisualRowCount <= 0 Then
                            objaddon.objapplication.SetStatusBarMessage("Please fill data in Input or Output tab Row level", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                    RemoveLastrow(Matrix0, "Code")
                    RemoveLastrow(Matrix3, "Code")
                    RemoveLastrow(Matrix2, "Code")
                    RemoveLastrow(Matrix1, "DocNum")
                    RemoveLastrow(Matrix4, "Code")
                    'CalcPlanQty()
                    'Dim SubItemCount As Integer = 0
                    'For i As Integer = 1 To Matrix2.VisualRowCount
                    '    objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                    '    If objcombo.Selected.Value = "1" Then
                    '        If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String Then
                    '            SubItemCount += 1
                    '        End If
                    '    End If
                    'Next
                    'If SubItemCount > 0 Then
                    '    objaddon.objapplication.SetStatusBarMessage("Scrap Tab-> Sub Item Found Please Change...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    '    BubbleEvent = False : Exit Sub
                    'End If
                End If

            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                Dim Status As String = ""
                'Dim objRS As SAPbobsCOM.Recordset
                '                Dim OpenQty() As Double
                'objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'If Not ReceiptAutoLoad = "Y" Then
                'Dim OpenQty() As Double
                If objform.Items.Item("txtentry").Specific.String = "" Then Exit Sub
                'OpenQty = GetOpen_VendorQty(objform.Items.Item("txtentry").Specific.String)
                'odbdsHeader.SetValue("U_OpenQty", 0, StatusClosing_OpenQty("N"))
                odbdsHeader.SetValue("U_VOBStock", 0, GetVendor_OpenReceiptQty(objform.Items.Item("txtentry").Specific.String))

                'End If

                'If objaddon.HANA Then
                '    'Status = objaddon.objglobalmethods.getSingleValue("Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join ""@MIPL_OPOR"" T2 on T0.""U_SubConNo""=T2.""DocEntry"" and T2.""U_SItemCode""=  T1.""ItemCode""  " & _
                '    '                                                                                "    where T1.""ItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.""U_SQty"" having T2.""U_SQty""- sum(T1.""Quantity"")<=0 ")
                '    Status = objaddon.objglobalmethods.getSingleValue("Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join ""@MIPL_OPOR"" T2 on T0.""U_SubConNo""=T2.""DocEntry"" and T2.""U_SItemCode""=  T1.""ItemCode"" " & _
                '                                                      "where T1.""ItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.""U_SQty"" having " & _
                '                                                      "T2.""U_SQty""- ifnull((sum(T1.""Quantity"")+ " & _
                '                                                      "(Select ifnull(sum(T0.""U_Qty""),0) from ""@MIPL_POR3"" T0 where T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and ifnull(T0.""U_InvNo"",'') <>'' and ifnull(T0.""U_GRNo"",'') ='' )),0) <=0")
                'Else
                '    'Status = objaddon.objglobalmethods.getSingleValue("Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.DocEntry=T1.DocEntry left join @MIPL_OPOR T2 on T0.U_SubConNo=T2.DocEntry and T2.U_SItemCode=  T1.ItemCode  " & _
                '    '                                                                                "    where T1.ItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.U_SQty having T2.U_SQty- sum(T1.Quantity)<=0 ")
                '    Status = objaddon.objglobalmethods.getSingleValue("Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.DocEntry=T1.DocEntry left join [@MIPL_OPOR] T2 on T0.U_SubConNo=T2.DocEntry and T2.U_SItemCode=  T1.ItemCode " & _
                '                                                    "where T1.ItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.U_SQty having " & _
                '                                                    "T2.U_SQty- isnull((sum(T1.Quantity)+ " & _
                '                                                    "(Select isnull(sum(T0.U_Qty),0) from [@MIPL_POR3] T0 where T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and isnull(T0.U_InvNo,'') <>'' and isnull(T0.U_GRNo,'') ='' )),0) <=0")
                'End If
                If StatusClosing_OpenQty() = "0" Then
                    'If ComboBox2.Selected.Value = "O" Then
                    '    ComboBox2.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '    'EditText5.Value = "C"
                    '    'Matrix2.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    '    Matrix3.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    'End If

                    'If objaddon.HANA Then
                    '    Status = "Update ""@MIPL_OPOR"" set ""Status""='C' where ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'"
                    'Else
                    '    Status = "Update [@MIPL_OPOR] set Status='C' where DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'"
                    'End If
                    'objRS.DoQuery(Status)
                    ' objRS = Nothing
                Else
                    'If EditText5.Value = "C" Then
                    '    EditText5.Value = "O"
                    'End If
                End If

            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button4_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button4.ClickAfter
            'Getting Goods Receipt Posted data
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If AutoProduction = "Y" Then
                        Dim TranViewForm As New FrmViewTranData
                        TranViewForm.Show()
                        TranViewForm.objform.Title = "Goods Receipt"
                        TranViewForm.ViewTransactionData("OIGN", "IGN1", "59", "N", EditText11.Value)
                    Else
                        'CreateMySimpleForm("GoodsReceiptStatus", "Goods Receipt List", "OIGN", "IGN1", "59", "N")
                        CreateMySimpleForm("ViewD", "Goods Receipt List", "OIGN", "IGN1", "59", "N")
                    End If
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button5_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button5.ClickAfter
            'Getting Inventory Transfer Posted data
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If AutoProduction = "Y" Then
                        Dim TranViewForm As New FrmViewTranData
                        TranViewForm.Show()
                        TranViewForm.objform.Title = "Inventory Transfer"
                        TranViewForm.ViewTransactionData("OWTR", "WTR1", "67", "Y", EditText26.Value)
                    Else
                        'CreateMySimpleForm("InvTransferStatus", "Inventory Transfer List", "OWTR", "WTR1", "67", "Y")
                        CreateMySimpleForm("ViewD", "Inventory Transfer List", "OWTR", "WTR1", "67", "Y")
                    End If
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button6_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button6.ClickAfter
            'Getting Goods Issue Posted data
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If AutoProduction = "Y" Then
                        Dim TranViewForm As New FrmViewTranData
                        TranViewForm.Show()
                        TranViewForm.objform.Title = "Goods Issue"
                        TranViewForm.ViewTransactionData("OIGE", "IGE1", "60", "N", EditText11.Value)
                    Else
                        'CreateMySimpleForm("GoodsIssueStatus", "Goods Issue List", "OIGE", "IGE1", "60", "N")
                        CreateMySimpleForm("ViewD", "Goods Issue List", "OIGE", "IGE1", "60", "N")
                    End If
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim Status As String = "", StrQuery As String = "", GetEntry As String = ""
                'Dim OpenQty() As Double
                Try
                    'Dim cc As String = objform.Items.Item("txtcode").Specific.String
                    'If cc = "CPPL-SD" Or cc = "CPPL-MD" Then
                    '    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    '    Exit Sub
                    'End If
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        'Matrix0.Columns.Item("distrule").Editable = True
                        'CheckBox0.Item.Enabled = True
                        Matrix0.Columns.Item("PlanQty").Editable = False
                        If EditText21.Value <> "" Then
                            Matrix0.Columns.Item("PlanQty").Editable = False
                            Matrix0.Columns.Item("revbase").Editable = False
                        Else
                            Matrix0.Columns.Item("PlanQty").Editable = True
                            Matrix0.Columns.Item("revbase").Editable = True
                        End If
                        If ComboBox2.Selected.Value = "C" Then
                            Matrix0.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            Matrix3.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        Else
                            If AutoAssignItem = "Y" Then
                                Matrix0.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                Matrix3.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            End If
                        End If
                        For iRow As Integer = 1 To Matrix0.VisualRowCount
                            If Matrix0.Columns.Item("Code").Cells.Item(iRow).Specific.String <> "" Then
                                If objaddon.HANA Then
                                    StrQuery = "Select  * from (Select T0.""DocEntry"",T0.""DocNum"",T1.""U_ItemCode"" ""Item Code"","
                                    StrQuery += vbCrLf + "(Select ifnull(Sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where (A.""U_SubConNo""=T1.""DocEntry"" or B.""U_S_DocEntry""=T1.""DocEntry"") and B.""ItemCode""=T1.""U_Itemcode"") as ""Transfer Qty"","
                                    StrQuery += vbCrLf + "(Select ifnull(Sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_ScrapNum""=T1.""DocEntry"" and B.""ItemCode""=T1.""U_Itemcode"") as ""Return Qty"","
                                    StrQuery += vbCrLf + "(Select case when sum(B.""Quantity"")>0 then T1.""U_PlanQty""-ifnull(Sum(B.""Quantity""),0) else T1.""U_PlanQty"" end from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where (A.""U_SubConNo""=T1.""DocEntry"" or B.""U_S_DocEntry""=T1.""DocEntry"") and B.""ItemCode""=T1.""U_Itemcode"") as ""Pending Qty"""
                                    StrQuery += vbCrLf + "from ""@MIPL_OPOR"" T0 join ""@MIPL_POR1"" T1 on T0.""DocEntry""=T1.""DocEntry"") A where A.""Pending Qty"">0 and A.""DocEntry""='" & EditText26.Value & "' and A.""Item Code""='" & Matrix0.Columns.Item("Code").Cells.Item(iRow).Specific.String & "'"

                                Else
                                    StrQuery = "Select  * from (Select T0.DocEntry,T0.DocNum,T1.U_ItemCode [Item Code],"
                                    StrQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where (A.U_SubConNo=T1.DocEntry or B.U_S_DocEntry=T1.DocEntry) and B.ItemCode=T1.U_Itemcode) as [Transfer Qty],"
                                    StrQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T1.DocEntry and B.ItemCode=T1.U_Itemcode) as [Return Qty],"
                                    StrQuery += vbCrLf + "(Select case when sum(B.Quantity)>0 then T1.U_PlanQty-isnull(Sum(B.Quantity),0) else T1.U_PlanQty end from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where (A.U_SubConNo=T1.DocEntry or B.U_S_DocEntry=T1.DocEntry) and B.ItemCode=T1.U_Itemcode) as [Pending Qty]"
                                    StrQuery += vbCrLf + "from [@MIPL_OPOR] T0 join [@MIPL_POR1] T1 on T0.DocEntry=T1.DocEntry) A where A.[Pending Qty]>0 and A.DocEntry='" & EditText26.Value & "' and A.[Item Code]='" & Matrix0.Columns.Item("Code").Cells.Item(iRow).Specific.String & "'"
                                End If
                                objRs.DoQuery(StrQuery)
                                If objRs.RecordCount > 0 Then
                                    If Matrix0.Columns.Item("ProcQty").Cells.Item(iRow).Specific.String <> objRs.Fields.Item("Transfer Qty").Value Then Matrix0.Columns.Item("ProcQty").Cells.Item(iRow).Specific.String = objRs.Fields.Item("Transfer Qty").Value : objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    If Matrix0.Columns.Item("OpenQty").Cells.Item(iRow).Specific.String <> objRs.Fields.Item("Pending Qty").Value Then Matrix0.Columns.Item("OpenQty").Cells.Item(iRow).Specific.String = objRs.Fields.Item("Pending Qty").Value
                                    If Matrix0.Columns.Item("RetQty").Cells.Item(iRow).Specific.String <> objRs.Fields.Item("Return Qty").Value Then Matrix0.Columns.Item("RetQty").Cells.Item(iRow).Specific.String = objRs.Fields.Item("Return Qty").Value
                                End If
                            End If
                        Next
                        'If EditText5.Value = "C" Then
                        '    objform.Items.Item("clstat").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        '    objform.Items.Item("EditQty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        '    'CheckBox1.Checked = True
                        '    'CheckBox1.Item.Enabled = False
                        '    'CheckBox0.Item.Enabled = False
                        'Else
                        '    objform.Items.Item("clstat").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_Default)
                        '    '    CheckBox1.Item.Enabled = True
                        '    '    CheckBox1.Checked = False
                        'End If

                        'For ii As Integer = 1 To Matrix0.VisualRowCount
                        '    If Val(Matrix0.Columns.Item("actplanqty").Cells.Item(ii).Specific.String) = 0 Then
                        '        Matrix0.Columns.Item("actplanqty").Cells.Item(ii).Specific.String = CDbl(objform.Items.Item("SQty").Specific.String) * CDbl(Matrix0.Columns.Item("Quant").Cells.Item(ii).Specific.String)
                        '    End If
                        '    'If RevPlanQty = "Y" Then
                        '    '    objcombo = Matrix0.Columns.Item("revbase").Cells.Item(ii).Specific
                        '    '    If objcombo.Selected Is Nothing Then Continue For
                        '    '    If objcombo.Selected.Value = "Y" Then
                        '    '        Matrix0.CommonSetting.SetCellEditable(ii, 8, True)
                        '    '    Else
                        '    '        Matrix0.CommonSetting.SetCellEditable(ii, 8, False)
                        '    '    End If
                        '    'End If
                        'Next
                        'If Not TypeInScrap = "Y" Then
                        '    Matrix2.Columns.Item("Type").Visible = False
                        'End If
                        'CheckBox0.Checked = False
                        Dim Fieldsize As Size = TextRenderer.MeasureText(EditText15.Value, New Font("Arial", 12.0F))
                        If Fieldsize.Width <= 135 Then
                            EditText15.Item.Width = 135
                        Else
                            EditText15.Item.Width = Fieldsize.Width
                        End If
                        Fieldsize = TextRenderer.MeasureText(EditText1.Value, New Font("Arial", 12.0F))
                        If Fieldsize.Width <= 135 Then
                            EditText1.Item.Width = 135
                        Else
                            EditText1.Item.Width = Fieldsize.Width
                        End If
                        Fieldsize = TextRenderer.MeasureText(EditText8.Value, New Font("Arial", 12.0F))
                        If Fieldsize.Width <= 135 Then
                            EditText8.Item.Width = 135
                        Else
                            EditText8.Item.Width = Fieldsize.Width
                        End If
                        If ReceiptAutoLoad = "Y" Then
                            EditText11.Item.Enabled = False
                            If EditText11.Value <> "" Then
                                If objaddon.HANA Then
                                    GetEntry = "select Distinct T0.""DocEntry"" from OIGN T0 join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""BaseType""=202 And T1.""BaseEntry""='" & objform.Items.Item("txtponum").Specific.String & "'"
                                Else
                                    GetEntry = "select Distinct T0.DocEntry from OIGN T0 join IGN1 T1 on T0.DocEntry=T1.DocEntry where T1.BaseType=202 and T1.BaseEntry='" & objform.Items.Item("txtponum").Specific.String & "'"
                                End If
                                objRs.DoQuery(GetEntry)
                                If objRs.RecordCount > 0 Then
                                    For i As Integer = 0 To objRs.RecordCount - 1
                                        If i = 0 Then
                                            GetEntry = objRs.Fields.Item("DocEntry").Value.ToString
                                        Else
                                            GetEntry += ", " + objRs.Fields.Item("DocEntry").Value.ToString
                                        End If
                                        objRs.MoveNext()
                                    Next
                                    EditText23.Value = GetEntry
                                End If
                                If objaddon.HANA Then
                                    GetEntry = "select Distinct T0.""DocEntry"" from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""BaseType""=202 and T1.""BaseEntry""='" & objform.Items.Item("txtponum").Specific.String & "'"
                                Else
                                    GetEntry = "select Distinct T0.DocEntry from OIGE T0 join IGE1 T1 on T0.DocEntry=T1.DocEntry where T1.BaseType=202 and T1.BaseEntry='" & objform.Items.Item("txtponum").Specific.String & "'"
                                End If
                                objRs.DoQuery(GetEntry)
                                If objRs.RecordCount > 0 Then
                                    For i As Integer = 0 To objRs.RecordCount - 1
                                        If i = 0 Then
                                            GetEntry = objRs.Fields.Item("DocEntry").Value.ToString
                                        Else
                                            GetEntry += ", " + objRs.Fields.Item("DocEntry").Value.ToString
                                        End If
                                        objRs.MoveNext()
                                    Next
                                    EditText22.Value = GetEntry
                                End If
                            End If
                            If objaddon.HANA Then
                                GetEntry = objaddon.objglobalmethods.getSingleValue("select ifnull(""PlannedQty"",0)-ifnull((""CmpltQty""+""RjctQty""),0) ""OpenQty"" from OWOR where ""U_SubPONum""='" & objform.Items.Item("txtentry").Specific.String & "'")
                            Else
                                GetEntry = objaddon.objglobalmethods.getSingleValue("select isnull(PlannedQty,0)-isnull((CmpltQty+RjctQty),0) OpenQty from OWOR where U_SubPONum='" & objform.Items.Item("txtentry").Specific.String & "'")
                            End If

                            If GetEntry <> "" Then
                                EditText13.Value = GetEntry
                                If CDbl(GetEntry) <= 0 Then
                                    Status = "1"
                                End If
                            End If
                        Else
                            'OpenQty = GetOpen_VendorQty(objform.Items.Item("txtentry").Specific.String)
                            StatusClosing_OpenQty()
                            'Dim OQty As String = StatusClosing_OpenQty("N")
                            'EditText13.Value = OQty ' CDbl(OpenQty(0)) 'CDbl(CDbl(objform.Items.Item("SQty").Specific.String) - OpenQty(0))
                            EditText12.Value = GetVendor_OpenReceiptQty(objform.Items.Item("txtentry").Specific.String)
                            'If objaddon.HANA Then
                            '    Status = objaddon.objglobalmethods.getSingleValue("Select Distinct 1 from ""@MIPL_POR2"" T1 left join  ""@MIPL_OPOR"" T0 on T0.""DocEntry""=T1.""DocEntry""  left join OIGN T2 on T0.""DocEntry""=T2.""U_SubConNo"" left join IGN1 T3 on T2.""DocEntry""=T3.""DocEntry"" and  T1.""U_Itemcode""=T3.""ItemCode""  and T0.""U_SItemCode"" =T3.""ItemCode"" " & _
                            '                                             "    where T0.""U_SItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and T0.""Status""='O' and T0.""DocNum""='" & objform.Items.Item("txtdocnum").Specific.String & "' group by T0.""U_SQty"" having T0.""U_SQty""- sum(T3.""Quantity"")<=0;  ")
                            'Else
                            '    Status = objaddon.objglobalmethods.getSingleValue("Select Distinct 1 from [@MIPL_POR2] T1 left join  [@MIPL_OPOR] T0 on T0.DocEntry=T1.DocEntry  left join OIGN T2 on T0.DocEntry=T2.U_SubConNo left join IGN1 T3 on T2.DocEntry=T3.DocEntry and  T1.U_Itemcode=T3.ItemCode  and T0.U_SItemCode =T3.ItemCode " & _
                            '                                                 "    where T0.U_SItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and T0.Status='O' and T0.DocNum='" & objform.Items.Item("txtdocnum").Specific.String & "' group by T0.U_SQty having T0.U_SQty- sum(T3.Quantity)<=0  ")
                            'End If
                            'Status = OQty
                            'If objaddon.HANA Then
                            '    'Status = objaddon.objglobalmethods.getSingleValue("Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join ""@MIPL_OPOR"" T2 on T0.""U_SubConNo""=T2.""DocEntry"" and T2.""U_SItemCode""=  T1.""ItemCode""  " & _
                            '    '                                                                                "    where T1.""ItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.""U_SQty"" having T2.""U_SQty""- sum(T1.""Quantity"")<=0 ")
                            '    Status = objaddon.objglobalmethods.getSingleValue("Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join ""@MIPL_OPOR"" T2 on T0.""U_SubConNo""=T2.""DocEntry"" and T2.""U_SItemCode""=  T1.""ItemCode"" " & _
                            '                                          "where T1.""ItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.""U_SQty"" having " & _
                            '                                          "T2.""U_SQty""- ifnull((sum(T1.""Quantity"")+ " & _
                            '                                          "(Select ifnull(sum(T0.""U_Qty""),0) from ""@MIPL_POR3"" T0 where T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and ifnull(T0.""U_InvNo"",'') <>'' and ifnull(T0.""U_GRNo"",'') ='' )),0) <=0")

                            'Else
                            '    'Status = objaddon.objglobalmethods.getSingleValue("Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.DocEntry=T1.DocEntry left join @MIPL_OPOR T2 on T0.U_SubConNo=T2.DocEntry and T2.U_SItemCode=  T1.ItemCode  " & _
                            '    '                                                                                "    where T1.ItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.U_SQty having T2.U_SQty- sum(T1.Quantity)<=0 ")
                            '    Status = objaddon.objglobalmethods.getSingleValue("Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.DocEntry=T1.DocEntry left join [@MIPL_OPOR] T2 on T0.U_SubConNo=T2.DocEntry and T2.U_SItemCode=  T1.ItemCode " & _
                            '                                        "where T1.ItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.U_SQty having " & _
                            '                                        "T2.U_SQty- isnull((sum(T1.Quantity)+ " & _
                            '                                        "(Select isnull(sum(T0.U_Qty),0) from [@MIPL_POR3] T0 where T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and isnull(T0.U_InvNo,'') <>'' and isnull(T0.U_GRNo,'') ='' )),0) <=0")
                            'End If
                            SetTransEntryToSubPO(objform.Items.Item("txtentry").Specific.String, "OIGN", "txtGRNo")
                            SetTransEntryToSubPO(objform.Items.Item("txtentry").Specific.String, "OIGE", "txtGINo")
                        End If
                        SetTransEntryToSubPO(objform.Items.Item("txtentry").Specific.String, "OWTR", "TxtInvTr")
                        'If Status = "0" Then
                        '    'If ComboBox2.Selected.Value = "O" Then
                        '    '    ComboBox2.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        '    '    'EditText5.Value = "C"
                        '    'End If
                        '    'If objaddon.HANA Then
                        '    '    StrQuery = "Update ""@MIPL_OPOR"" set ""Status""='C' where ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'"
                        '    'Else
                        '    '    StrQuery = "Update [@MIPL_OPOR] set Status='C' where DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'"
                        '    'End If
                        '    'objRs.DoQuery(StrQuery)
                        'Else
                        '    'If EditText5.Value = "C" Then
                        '    '    EditText5.Value = "O"
                        '    'End If
                        'End If
                        'If ProdOrderStat_Close = "Y" Then
                        '    If EditText11.Value <> "" Then
                        '        Try
                        '            Dim updatequery As String = ""
                        '            If objaddon.HANA Then
                        '                Status = objaddon.objglobalmethods.getSingleValue("SELECT 1 AS ""Status"" FROM OWOR T4 WHERE T4.""ItemCode"" = '" & EditText8.Value & "' AND T4.""DocEntry"" = '" & EditText11.Value & "' GROUP BY T4.""ItemCode"", T4.""DocEntry"", T4.""PlannedQty"" " &
                        '                    " HAVING IFNULL((T4.""PlannedQty"") - (SELECT IFNULL(CASE WHEN SUM(T1.""Quantity"") > 0 THEN SUM(T1.""Quantity"") ELSE 0 END, 0)   " &
                        '                    " FROM IGN1 T1 INNER JOIN OIGN T0 ON T1.""DocEntry"" = T0.""DocEntry"" INNER JOIN ""@MIPL_OPOR"" T3 ON T3.""DocEntry"" = T0.""U_SubConNo"" AND T3.""U_SItemCode"" = T1.""ItemCode"" AND T3.""U_SItemCode"" = T4.""ItemCode"" AND T3.""U_PONum"" = T4.""DocEntry""), 0) <= 0")
                        '                If Status = "1" Then
                        '                    updatequery = "update OWOR set ""Status""='L' where ""DocEntry""='" & EditText11.Value & "' and ""ItemCode"" ='" & EditText8.Value & "' and ""Status""='R' "
                        '                End If
                        '            Else
                        '                Status = objaddon.objglobalmethods.getSingleValue("Select 1 as Status from OWOR T4 where T4.ItemCode='" & EditText8.Value & "' and T4.DocEntry='" & EditText11.Value & "' group by T4.ItemCode,T4.DocEntry,T4.PlannedQty" &
                        '                    " having isnull((T4.PlannedQty)-(Select isnull(case when sum(T1.Quantity)>0 then Sum(T1.Quantity) else 0 end,0)  " &
                        '                    " from IGN1 T1 join OIGN T0 on T1.DocEntry=T0.DocEntry join [@MIPL_OPOR] T3 on T3.DocEntry=T0.U_SubConNo and T3.U_SItemCode=T1.ItemCode and T3.U_SItemCode= T4.ItemCode and T3.U_PONum=T4.DocEntry),0)<=0")
                        '                If Status = "1" Then
                        '                    updatequery = "update OWOR set Status='L' where DocEntry='" & EditText11.Value & "' and ItemCode ='" & EditText8.Value & "' and Status='R'"
                        '                End If
                        '            End If
                        '            If updatequery <> "" Then
                        '                objRs.DoQuery(updatequery)
                        '            End If
                        '        Catch ex As Exception
                        '        End Try
                        '    End If
                        'End If
                        objRs = Nothing
                    End If

                    If ReceiptAutoLoad = "Y" Then
                        Folder1.Item.Click()
                        ButtonCombo0.Item.Visible = False
                        LoadProdOrderReceiptToOutputScreen()
                    End If
                    If APAutoLoad = "Y" Then
                        Folder3.Item.Click()
                        LoadAPInvoiceToRelatedDocTab()
                    End If
                    If AutoProduction <> "Y" Then
                        objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If objaddon.HANA Then
                            objRs.DoQuery("select ""DocEntry"" from OIGN where ""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "' and ""Ref2""='Output' order by ""DocEntry""")
                        Else
                            objRs.DoQuery("select DocEntry from OIGN where U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "' and Ref2='Output' order by DocEntry")
                        End If

                        If objRs.RecordCount > 0 Then
                            For i As Integer = 0 To objRs.RecordCount - 1
                                If Matrix3.Columns.Item("GINo").Cells.Item(i + 1).Specific.String <> "" Then
                                    Matrix3.Columns.Item("GRNo").Cells.Item(i + 1).Specific.String = objRs.Fields.Item("DocEntry").Value
                                End If
                                objRs.MoveNext()
                            Next
                        End If
                    Else
                        objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If objaddon.HANA Then
                            StrQuery = "Select Case when Sum(T2.""Quantity"") >0 then T0.""U_PlanQty""-Sum(T2.""Quantity"") else T0.""U_PlanQty"" end as ""OpenQty"","
                            StrQuery += vbCrLf + " Sum(T2.""Quantity"") ""ProcessedQty"" from ""@MIPL_POR1"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" "
                            StrQuery += vbCrLf + " left join owtr T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join wtr1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T0.""U_Itemcode""=T2.""ItemCode"""
                            StrQuery += vbCrLf + " where T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & "  group by T0.""U_PlanQty"",T0.""U_Itemcode"",T0.""U_WhsCode"",T0.""U_SubWhse"""
                        Else
                            StrQuery = "Select Case when Sum(T2.Quantity) >0 then T0.U_PlanQty-Sum(T2.Quantity) else T0.U_PlanQty end as OpenQty,"
                            StrQuery += vbCrLf + " Sum(T2.Quantity) ProcessedQty from [@MIPL_POR1] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry "
                            StrQuery += vbCrLf + " left join owtr T3 on T3.U_SubConNo=T1.DocEntry left join wtr1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode"
                            StrQuery += vbCrLf + " where T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & "  group by T0.U_PlanQty,T0.U_Itemcode,T0.U_WhsCode,T0.U_SubWhse"
                        End If
                        objRs.DoQuery(StrQuery)
                        If objRs.RecordCount > 0 Then
                            For i As Integer = 0 To objRs.RecordCount - 1
                                If objaddon.HANA Then
                                    StrQuery = objaddon.objglobalmethods.getSingleValue("select 1 from OITM where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i + 1).Specific.String & "'")
                                Else
                                    StrQuery = objaddon.objglobalmethods.getSingleValue("select 1 from OITM where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i + 1).Specific.String & "'")
                                End If
                                If StrQuery = "1" Then
                                    Matrix0.Columns.Item("ProcQty").Cells.Item(i + 1).Specific.String = objRs.Fields.Item("ProcessedQty").Value
                                    Matrix0.Columns.Item("OpenQty").Cells.Item(i + 1).Specific.String = objRs.Fields.Item("OpenQty").Value
                                    'objform.Items.Item("txtopqty").Specific.String = objRs.Fields.Item("OpenQty").Value
                                Else
                                    Matrix0.Columns.Item("ProcQty").Cells.Item(i + 1).Specific.String = "0"
                                    Matrix0.Columns.Item("OpenQty").Cells.Item(i + 1).Specific.String = "0"
                                End If
                                objRs.MoveNext()
                            Next
                        End If
                    End If
                    Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    'If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    '    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'End If

                    'objaddon.objapplication.Menus.Item("1300").Activate()
                Catch ex As Exception
                End Try
                objRs = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText24_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText24.LostFocusAfter
            Try
                Dim objRs As SAPbobsCOM.Recordset
                Dim StrQuery As String
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If POItem = "Y" Then
                    If EditText24.Value = "" Then
                        objform.ActiveItem = "SubPoNum"
                        'objaddon.objapplication.StatusBar.SetText("Please select Sub-Contract PO Entry & ItemCode...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    Else
                        'DocNum = objaddon.objglobalmethods.getSingleValue("Select  ""DocNum"" from OPOR where ""DocEntry""='" & EditText24.Value & "'")
                        'EditText18.Value = DocNum
                    End If
                End If
                If GetDatePO = "Y" Then
                    If objaddon.HANA Then
                        StrQuery = "select To_Varchar(""DocDate"",'dd/MM/yy') as ""DocDate"",To_Varchar(""DocDueDate"",'dd/MM/yy') as ""DocDueDate"",To_Varchar(""TaxDate"",'dd/MM/yy') as ""TaxDate"" from OPOR where ""DocEntry""='" & EditText24.Value & "'" '" & EditText24.Value & "
                    Else
                        StrQuery = "select Format(DocDate,'dd/MM/yy') as DocDate,Format(DocDueDate,'dd/MM/yy') as DocDueDate,Format(TaxDate,'dd/MM/yy') as TaxDate from OPOR where DocEntry='" & EditText24.Value & "'" '" & EditText24.Value & "
                    End If
                    objRs.DoQuery(StrQuery)
                    If objRs.RecordCount > 0 Then
                        objform.Items.Item("posdate").Specific.string = objRs.Fields.Item("DocDate").Value
                        objform.Items.Item("deldate").Specific.string = objRs.Fields.Item("DocDueDate").Value
                        objform.Items.Item("docdate").Specific.string = objRs.Fields.Item("TaxDate").Value
                    End If
                End If
                objRs = Nothing
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText17_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText17.LostFocusAfter
            Try
                Dim Qty As String = ""
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If EditText17.Value <> "" Then
                    If objaddon.HANA Then
                        Qty = objaddon.objglobalmethods.getSingleValue("Select ""Quantity"" from POR1 where ""DocEntry""='" & EditText24.Value & "' and ""LineNum""='" & EditText17.Value & "' and ""LineStatus""='O'")
                    Else
                        Qty = objaddon.objglobalmethods.getSingleValue("Select Quantity from POR1 where DocEntry='" & EditText24.Value & "' and LineNum='" & EditText17.Value & "' and LineStatus='O'")
                    End If
                    If Qty <> "" Then
                        'EditText25.Value = Qty
                        odbdsHeader.SetValue("U_SQty", 0, Qty)
                    End If
                End If

            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText1_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.LostFocusAfter
            Try
                If EditText1.Value = "" Then
                    If EditText0.Value <> "" Then
                        EditText0.Value = ""
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Form_CloseBefore(pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean)
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    'If objaddon.objapplication.MessageBox("Do you want to Save the Data?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                    objaddon.objapplication.StatusBar.SetText("Please update the Sub-Contracting Screen...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub

                    'objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular) : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText25_ValidateAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText25.ValidateAfter
            Try
                If pVal.InnerEvent = True Then Exit Sub
                CalcPlanQty()
                If Matrix3.VisualRowCount > 0 Then
                    Matrix3.SetCellWithoutValidation(Matrix3.VisualRowCount, "Quant", CDbl(EditText25.Value))
                End If
                'Matrix_Total()
                'Matrix0.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("PlanQty").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'objform.Update()
                'objform.Refresh()
                'Matrix0.CommonSetting.EnableArrowKey = True
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText7_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText7.LostFocusAfter
            'Delivery Date Checking not less than the posting date
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    Dim DocDate As Date = Date.ParseExact(EditText6.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    Dim DelDate As Date = Date.ParseExact(EditText7.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    If DateTime.Compare(DocDate, DelDate) > 0 Then
                        objaddon.objapplication.StatusBar.SetText("In ""Due Date"" field, enter date that is equal to or later than posting date.  Field: Delivery Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        EditText10.Item.Click()
                    End If
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub CheckBox0_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox0.PressedAfter
            Try
                If ComboBox2.Selected.Value = "C" Then Exit Sub
                Dim StrSql As String
                Dim objrs As SAPbobsCOM.Recordset
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If CheckBox0.Checked = True Then
                        'CheckBox0.Checked = False
                        Dim CheckInvStat As Integer = 0
                        If objaddon.HANA Then
                            StrSql = "Select Case when Sum(T2.""Quantity"") >0 then T0.""U_PlanQty""-Sum(T2.""Quantity"") else T0.""U_PlanQty"" end as ""InvQty"",T0.""U_Itemcode"""
                            'StrSql += vbCrLf + "- ifnull((select Top 1 sum(T1.""U_Qty"") from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""U_Type""='2' and T1.""U_Status""='C' and T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' "
                            'StrSql += vbCrLf + "group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc),0) "
                            StrSql += vbCrLf + " from ""@MIPL_POR1"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" left join OWTR T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join WTR1 T2  on T3.""DocEntry""=T2.""DocEntry"" And T0.""U_Itemcode""=T2.""ItemCode"" And T0.""LineId""= T2.""U_LineID"""
                            StrSql += vbCrLf + "where T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " group by T0.""U_PlanQty"",T0.""U_Itemcode"" order by ""InvQty"""
                        Else
                            StrSql = "Select Case when Sum(T2.Quantity) >0 then T0.U_PlanQty-Sum(T2.Quantity) else T0.U_PlanQty end as InvQty,T0.U_Itemcode"
                            'StrSql += vbCrLf + "- isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' "
                            'StrSql += vbCrLf + "group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0) "
                            StrSql += vbCrLf + "from [@MIPL_POR1] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry left join OWTR T3 on T3.U_SubConNo=T1.DocEntry left join WTR1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode and T0.LineId= T2.U_LineID"
                            StrSql += vbCrLf + "where T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " group by T0.U_PlanQty,T0.U_Itemcode order by InvQty"
                        End If
                        objrs.DoQuery(StrSql)
                        For Result As Integer = 0 To objrs.RecordCount - 1
                            If CDbl(objrs.Fields.Item("InvQty").Value) > 0.0 Then
                                CheckInvStat += 1
                            End If
                            objrs.MoveNext()
                        Next
                        If CheckInvStat > 0 Then
                            EditText25.Item.Enabled = True
                            EditText0.Item.Enabled = True
                            EditText7.Item.Enabled = True
                            EditText11.Item.Enabled = True
                            EditText19.Item.Enabled = True
                            EditText24.Item.Enabled = True
                            EditText28.Item.Enabled = True
                            CheckBox0.Item.Enabled = True
                        Else
                            CheckBox0.Item.Enabled = False
                            EditText25.Item.Enabled = False
                            EditText0.Item.Enabled = False
                            EditText7.Item.Enabled = False
                            EditText11.Item.Enabled = False
                            EditText19.Item.Enabled = False
                            EditText24.Item.Enabled = False
                            EditText28.Item.Enabled = False
                            objaddon.objapplication.StatusBar.SetText("Qty not to be editable due to inventory posted...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                        'objform.Refresh()
                        'objform.Update()
                    End If
                Else
                    CheckBox0.Checked = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ActivateAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                'If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            Catch ex As Exception
            End Try

        End Sub

        Private Sub CheckBox1_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox1.PressedAfter
            Try
                'If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And pVal.ActionSuccess Then
                '    If EditText5.Value = "O" Then
                '        If CheckBox1.Checked = True Then
                '            If objaddon.objapplication.MessageBox("Do you want to Close the Status?", 2, "Yes", "No") <> 1 Then CheckBox1.Checked = False : Exit Sub
                '            If SuperUser = objaddon.objcompany.UserName Then
                '                Dim objRS As SAPbobsCOM.Recordset
                '                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '                If objaddon.HANA Then
                '                    objRS.DoQuery("Update ""@MIPL_OPOR"" set ""Status""='C',""U_Clstat""='Y' where ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'")
                '                Else
                '                    objRS.DoQuery("Update [@MIPL_OPOR] set Status='C',U_Clstat='Y' where DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'")
                '                End If
                '                objRS = Nothing
                '                'Try
                '                '    Dim oGeneralService As SAPbobsCOM.GeneralService
                '                '    'Dim oGeneralData As SAPbobsCOM.GeneralData
                '                '    Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                '                '    oGeneralService = objaddon.objcompany.GetCompanyService.GetGeneralService("SUBPO")
                '                '    'oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                '                '    oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                '                '    oGeneralParams.SetProperty("DocEntry", "6938") 'objform.Items.Item("txtentry").Specific.String
                '                '    'oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                '                '    oGeneralService.Close(oGeneralParams)
                '                'Catch ex As Exception

                '                'End Try
                '                'EditText5.Value = "C"
                '                'CheckBox1.Item.Enabled = False
                '                objform.Items.Item("clstat").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                '                objform.Refresh()
                '                objform.Update()

                '                'objRS = Nothing
                '            Else
                '                objform.Items.Item("clstat").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_Default)
                '                'CheckBox1.Item.Enabled = True
                '                'CheckBox1.Checked = False
                '                objaddon.objapplication.StatusBar.SetText("You are not authorized to close the status due to SuperUser is not assigned..." & objaddon.objcompany.UserName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '            End If

                '        End If
                '    Else
                '        objform.Items.Item("clstat").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                '        'CheckBox1.Item.Enabled = False
                '    End If
                'End If

            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText25_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText25.KeyDownAfter
            Try
                'Matrix_Total()
                If EditText19.Value = "" Then Exit Sub
                If Val(EditText25.Value) <= 0 Then
                    objaddon.objapplication.StatusBar.SetText("In ""Qty"" (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    EditText25.Item.Click()
                    EditText25.Value = "1"
                Else
                    Exit Sub
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText19_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText19.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                EditText8.Value = ""
                Dim oEdit As SAPbouiCOM.EditText
                oEdit = objform.Items.Item("txtbitem").Specific
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                Dim oCFL As SAPbouiCOM.ChooseFromList

                If ItemBOM <> "Y" Then   'From Item Master
                    oCFL = objform.ChooseFromLists.Item("CFL_ItemM")
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()
                    oEdit.ChooseFromListUID = "CFL_ItemM"
                    oEdit.ChooseFromListAlias = "ItemCode"
                    oCond = oConds.Add()
                    oCond.Alias = "U_SubCont"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "Y"
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCond = oConds.Add()
                    oCond.Alias = "validFor"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "Y"
                Else                    ' From SUB_BOM
                    oCFL = objform.ChooseFromLists.Item("CFL_Item")
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()
                    oEdit.ChooseFromListUID = "CFL_Item"
                    oEdit.ChooseFromListAlias = "Code"
                    oCond = oConds.Add()
                    oCond.Alias = "U_Active"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "Y"
                End If
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub EditText19_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText19.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    If ItemBOM <> "Y" Then
                        Try
                            'EditText8.Value = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                            odbdsHeader.SetValue("U_SItemCode", 0, pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value)
                        Catch ex As Exception
                        End Try
                        Try
                            'EditText19.Value = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                            odbdsHeader.SetValue("U_BOMCode", 0, pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value)
                        Catch ex As Exception
                        End Try
                    Else
                        Try
                            'EditText19.Value = pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value
                            odbdsHeader.SetValue("U_BOMCode", 0, pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value)
                        Catch ex As Exception
                        End Try
                        Try
                            'EditText8.Value = pCFL.SelectedObjects.Columns.Item("U_ItemCode").Cells.Item(0).Value
                            odbdsHeader.SetValue("U_SItemCode", 0, pCFL.SelectedObjects.Columns.Item("U_ItemCode").Cells.Item(0).Value)
                        Catch ex As Exception
                        End Try
                    End If
                    Dim Getval As String = ""
                    If objaddon.HANA Then
                        Getval = objaddon.objglobalmethods.getSingleValue("Select 1 as ""Status"" from ""@MIPL_BOM2"" where ""U_Proccode""<>'' and ""Code""='" & EditText19.Value & "'")
                    Else
                        Getval = objaddon.objglobalmethods.getSingleValue("Select 1 as Status from [@MIPL_BOM2] where U_Proccode<>'' and Code='" & EditText19.Value & "'")
                    End If
                    If Getval = "" Then
                        If objaddon.HANA Then
                            Getval = objaddon.objglobalmethods.getSingleValue("Select ""U_Type"" as ""Process"" from ""@MIPL_OBOM"" where ""U_Type""<>'-' and ""Code""='" & EditText19.Value & "'")
                        Else
                            Getval = objaddon.objglobalmethods.getSingleValue("Select U_Type as Process from [@MIPL_OBOM] where U_Type<>'-' and Code='" & EditText19.Value & "'")
                        End If
                        'If Getval <> "" Then EditText28.Value = Getval : EditText25.Item.Click() : EditText20.Value = Getval : EditText28.Item.Enabled = False ' objform.ActiveItem = "SQty" :
                        If Getval <> "" Then odbdsHeader.SetValue("U_Process", 0, Getval) : EditText25.Item.Click() : EditText20.Value = Getval : EditText28.Item.Enabled = False ' objform.ActiveItem = "SQty" :
                    End If
                    'Dim Fieldsize As Size = TextRenderer.MeasureText(EditText15.Value, New Font("Arial", 12.0F))
                    'If Fieldsize.Width <= 135 Then
                    '    EditText15.Item.Width = 135
                    'Else
                    '    EditText15.Item.Width = Fieldsize.Width
                    'End If
                    EditText25.Value = "1"
                    EditText19.Item.Click()
                End If
                '                Dim objrs As SAPbobsCOM.Recordset
                '                objrs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                '                If objaddon.HANA Then
                '                    objrs.DoQuery("select T2.""U_Type"",(Select distinct T1.""Descr"" from CUFD T0 join UFD1 T1 on T0.""TableID""=T1.""TableID"" " &
                '" where T0.""TableID""='@MIPL_OBOM' and T1.""FldValue""=T2.""U_Type"") as ""Desc""  from ""@MIPL_OBOM"" T2 where T2.""U_ItemCode""='" & EditText8.Value & "'")
                '                Else
                '                    objrs.DoQuery("select T2.U_Type,(Select distinct T1.Descr from CUFD T0 join UFD1 T1 on T0.TableID=T1.TableID " &
                '" where T0.TableID='@MIPL_OBOM' and T1.FldValue=T2.U_Type) as Desc  from [@MIPL_OBOM] T2 where T2.U_ItemCode='" & EditText8.Value & "'")
                '                End If
                '                ComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                '                If ComboBox2.ValidValues.Count > 0 Then
                '                    For i As Integer = ComboBox2.ValidValues.Count - 1 To 0 Step -1
                '                        ComboBox2.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                '                    Next
                '                End If
                '                If objrs.RecordCount > 0 Then
                '                    For Rec As Integer = 0 To objrs.RecordCount - 1
                '                        ComboBox2.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
                '                        objrs.MoveNext()
                '                    Next
                '                    ComboBox2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                '                End If
                '                objrs = Nothing
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub EditText19_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText19.LostFocusAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If EditText19.Value = "" Then Exit Sub
                Dim objrs As SAPbobsCOM.Recordset
                Dim strsql As String = "", GetValue As String = "", WhsCode As String = "", InvUOM As String = "", HSNCode As String
                Dim i As Integer = 0
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If ItemBOM = "Y" Then  'Loading SUB-BOM
                    If objaddon.HANA Then
                        'EditText15.Value = objaddon.objglobalmethods.getSingleValue("SELECT  ""ItemName"" FROM OITM where ""ItemCode""='" & EditText8.Value & "'")
                        odbdsHeader.SetValue("U_SItemDesc", 0, objaddon.objglobalmethods.getSingleValue("SELECT  ""ItemName"" FROM OITM where ""ItemCode""='" & EditText8.Value & "'"))
                    Else
                        'EditText15.Value = objaddon.objglobalmethods.getSingleValue("SELECT  ItemName FROM OITM where ItemCode='" & EditText8.Value & "'")
                        odbdsHeader.SetValue("U_SItemDesc", 0, objaddon.objglobalmethods.getSingleValue("SELECT  ItemName FROM OITM where ItemCode='" & EditText8.Value & "'"))
                    End If
                    Dim ItemDescsize As Size = TextRenderer.MeasureText(EditText15.Value, New Font("Arial", 12.0F))
                    If ItemDescsize.Width <= 135 Then
                        EditText15.Item.Width = 135
                    Else
                        EditText15.Item.Width = ItemDescsize.Width
                    End If
                    Dim VNamesize As Size = TextRenderer.MeasureText(EditText8.Value, New Font("Arial", 12.0F))
                    If VNamesize.Width <= 135 Then
                        EditText8.Item.Width = 135
                    Else
                        EditText8.Item.Width = VNamesize.Width
                    End If
                    If EditText8.Value = "" Then
                        EditText8.Value = EditText19.Value
                    End If
                    If EditText0.Value = "" Then objaddon.objapplication.StatusBar.SetText("Please Select Vendor..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Sub
                    If Matrix0.VisualRowCount = 0 Then objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                    If Matrix0.Columns.Item("Item1").Cells.Item(1).Specific.String <> "" Then
                        If EditText19.Value = Matrix0.Columns.Item("Item1").Cells.Item(1).Specific.String Then
                            Exit Sub
                        End If
                    End If
                    If EditText19.Value <> "" Then
                        If objaddon.HANA Then
                            EditText2.Value = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & EditText8.Value & "'")
                        Else
                            EditText2.Value = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & EditText8.Value & "'")
                        End If
                        If objaddon.HANA Then
                            strsql = "select T1.""U_Itemcode"",T1.""U_ItemDesc"",T1.""U_Qty"",T1.""U_Whse"",T1.""U_Unitprice"",T1.""U_Total"" "
                            strsql += vbCrLf + "from ""@MIPL_OBOM"" T0 join ""@MIPL_BOM1"" T1 on T0.""Code""=T1.""Code"" where ifnull(T1.""U_SCType"",'')='I' and T0.""Code""='" & EditText19.Value & "'"
                        Else
                            strsql = "select T1.U_Itemcode,T1.U_ItemDesc,T1.U_Qty,T1.U_Whse,T1.U_Unitprice,T1.U_Total "
                            strsql += vbCrLf + "from [@MIPL_OBOM] T0 join [@MIPL_BOM1] T1 on T0.Code=T1.Code where isnull(T1.U_SCType,'')='I' and T0.Code='" & EditText19.Value & "'"
                        End If
                        objrs.DoQuery(strsql)
                        'objaddon.objglobalmethods.WriteErrorLog("Query: " + strsql)
                        If BPWhse = "Y" Then
                            If objaddon.HANA Then
                                WhsCode = objaddon.objglobalmethods.getSingleValue("select ""U_WAREHOUSE"" from OCRD where ""CardCode""='" & EditText0.Value & "'")
                            Else
                                WhsCode = objaddon.objglobalmethods.getSingleValue("select U_WAREHOUSE from OCRD where CardCode='" & EditText0.Value & "'")
                            End If
                            'objaddon.objglobalmethods.WriteErrorLog("WhsCode: " + WhsCode)
                            'ElseIf ToWhseInInput = "Y" Then
                            '    If objaddon.HANA Then
                            '        WhsCode = objaddon.objglobalmethods.getSingleValue("Select ""DflWhs"" from OBPL where ""BPLId""=(select  distinct T0.""BPLid"" from OWHS T0 left join OLCT T1 on T0.""Location""=T1.""Code"" where T1.""Code""=(Select ""U_LCode"" from ""@MIPL_GEN"") and ifnull(T0.""Inactive"",'') ='N'and T0.""BPLid"" is not null)")
                            '    Else
                            '        WhsCode = objaddon.objglobalmethods.getSingleValue("Select DflWhs from OBPL where BPLId=(select  distinct T0.BPLid from OWHS T0 left join OLCT T1 on T0.Location=T1.Code where T1.Code=(Select U_LCode from [@MIPL_GEN]) and isnull(T0.Inactive,'') ='N' and T0.BPLid is not null)")
                            '    End If
                        End If
                        odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_POR1")
                        Matrix0.Clear()
                        odbdsDetails.Clear()
                        If objrs.RecordCount > 0 Then
                            'objaddon.objglobalmethods.WriteErrorLog("Reccount: " + CStr(objrs.RecordCount.ToString))
                            objform.Freeze(True)
                            objaddon.objapplication.StatusBar.SetText("Loading from SUB-BOM Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            While Not objrs.EoF
                                If CDbl(objrs.Fields.Item("U_Qty").Value) > 0 Then
                                    Matrix0.AddRow()
                                    'odbdsDetails.Clear()
                                    Matrix0.GetLineData(Matrix0.VisualRowCount)
                                    odbdsDetails.SetValue("LineId", 0, i + 1)
                                    odbdsDetails.SetValue("U_Item1", 0, EditText19.Value.ToString)
                                    odbdsDetails.SetValue("U_Itemcode", 0, objrs.Fields.Item("U_Itemcode").Value.ToString)
                                    odbdsDetails.SetValue("U_ItemDesc", 0, objrs.Fields.Item("U_ItemDesc").Value.ToString)
                                    odbdsDetails.SetValue("U_BOMQty", 0, objrs.Fields.Item("U_Qty").Value.ToString)
                                    odbdsDetails.SetValue("U_Qty", 0, objrs.Fields.Item("U_Qty").Value.ToString)
                                    Dim Sqty As Double = IIf(CDbl(objform.Items.Item("SQty").Specific.String) = 0, 1, CDbl(objform.Items.Item("SQty").Specific.String))
                                    'Dim Recqty As Double = IIf(CDbl(objform.Items.Item("SQty").Specific.String) = 0, 1, CDbl(objform.Items.Item("SQty").Specific.String))
                                    odbdsDetails.SetValue("U_PlanQty", 0, CDbl(objrs.Fields.Item("U_Qty").Value) * Sqty)
                                    odbdsDetails.SetValue("U_ActPlanQty", 0, CDbl(objrs.Fields.Item("U_Qty").Value) * Sqty)
                                    Try
                                        If objaddon.HANA Then
                                            InvUOM = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "'")
                                        Else
                                            InvUOM = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "'")
                                        End If
                                        If objaddon.HANA Then
                                            HSNCode = objaddon.objglobalmethods.getSingleValue("select T1.""ChapterID"" from OITM T0 left join OCHP T1 on T0.""ChapterID""=T1.""AbsEntry"" where T0.""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "'")
                                        Else
                                            HSNCode = objaddon.objglobalmethods.getSingleValue("select T1.ChapterID from OITM T0 left join OCHP T1 on T0.ChapterID=T1.AbsEntry where T0.ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "'")
                                        End If

                                        odbdsDetails.SetValue("U_InvUom", 0, InvUOM)
                                        odbdsDetails.SetValue("U_HSNCode", 0, HSNCode)
                                        'objaddon.objglobalmethods.WriteErrorLog("InvUOM: " + InvUOM)
                                        'objaddon.objglobalmethods.WriteErrorLog("HSNCode: " + HSNCode)
                                    Catch ex As Exception
                                    End Try

                                    Dim Price As String = "", OnHand As String
                                    'If VendorSPlPrice = "Y" Then
                                    If objaddon.HANA Then
                                        Price = objaddon.objglobalmethods.getSingleValue("Select ""AvgPrice"" from OITW where ""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and ""WhsCode""='" & objrs.Fields.Item("U_Whse").Value.ToString & "';")
                                    Else
                                        Price = objaddon.objglobalmethods.getSingleValue("Select AvgPrice from OITW where ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and WhsCode='" & objrs.Fields.Item("U_Whse").Value.ToString & "'")
                                    End If
                                    If Price <> "0" Then
                                        odbdsDetails.SetValue("U_Price", 0, Price)
                                    Else
                                        odbdsDetails.SetValue("U_Price", 0, objrs.Fields.Item("U_Unitprice").Value.ToString)
                                    End If
                                    'Else
                                    '    odbdsDetails.SetValue("U_Price", 0, objrs.Fields.Item("U_Unitprice").Value.ToString)
                                    'End If
                                    odbdsDetails.SetValue("U_WhsCode", 0, objrs.Fields.Item("U_Whse").Value.ToString)
                                    odbdsDetails.SetValue("U_LineTot", 0, CDbl(odbdsDetails.GetValue("U_Qty", 0)) * CDbl(odbdsDetails.GetValue("U_Price", 0)))
                                    If WhsCode <> "" Then
                                        If objaddon.HANA Then
                                            GetValue = objaddon.objglobalmethods.getSingleValue("Select 1 from OWHS where  ""WhsCode""='" & WhsCode & "';")
                                        Else
                                            GetValue = objaddon.objglobalmethods.getSingleValue("Select 1 from OWHS where WhsCode='" & WhsCode & "'")
                                        End If
                                        If GetValue = "" Then objaddon.objapplication.StatusBar.SetText("Invalid data : Sub-Whse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub : Matrix0.Clear()
                                    End If
                                    odbdsDetails.SetValue("U_SubWhse", 0, WhsCode)
                                    If objaddon.HANA Then
                                        OnHand = objaddon.objglobalmethods.getSingleValue("select ifnull(""OnHand"",0) from  oitw where ""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and ""WhsCode""='" & objrs.Fields.Item("U_Whse").Value.ToString & "'")
                                    Else
                                        OnHand = objaddon.objglobalmethods.getSingleValue("select Isnull(OnHand,0) from  oitw where ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and WhsCode='" & objrs.Fields.Item("U_Whse").Value.ToString & "'")
                                    End If
                                    odbdsDetails.SetValue("U_InStock", 0, OnHand)
                                    'objaddon.objglobalmethods.WriteErrorLog("OnHand: " + OnHand)
                                    If WPrice = "Y" Then
                                        odbdsDetails.SetValue("U_LTType", 0, "1")
                                        'objaddon.objglobalmethods.WriteErrorLog("U_LTType: " + "1")
                                    End If
                                    If RevPlanQty = "Y" Then
                                        odbdsDetails.SetValue("U_RevBase", 0, "N")
                                    End If
                                    Matrix0.SetLineData(Matrix0.VisualRowCount)
                                    i += 1
                                End If
                                objrs.MoveNext()
                            End While
                            objform.Freeze(False)
                            Matrix0.Columns.Item("PlanQty").Editable = False
                            If Matrix0.Columns.Item("PlanQty").Editable = False Or Matrix0.Columns.Item("revbase").Editable = False Then
                                'Matrix0.Columns.Item("PlanQty").Editable = True
                                Matrix0.Columns.Item("revbase").Editable = True
                            End If

                            objrs = Nothing
                            objaddon.objapplication.StatusBar.SetText("Loaded from SUB-BOM Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Matrix0.AutoResizeColumns()
                        Else
                            objaddon.objapplication.StatusBar.SetText("No records found from SUB_BOM...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                    Else
                        Exit Sub
                    End If
                End If
                'objaddon.objapplication.Menus.Item("1300").Activate()
                'objform.ActiveItem = "SQty"
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Form_LayoutKeyBefore(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean)
            Try
                'frmActCarriedOut = oGFun.oApplication.Forms.Item(eventInfo.FormUID)
                eventInfo.LayoutKey = objform.Items.Item("txtentry").Specific.string
            Catch ex As Exception
            End Try




        End Sub

        Private Sub Button11_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button11.ClickBefore
            Try
                If Button11.Item.Enabled = False Then BubbleEvent = False : Exit Sub
                If EditText21.Value = "" Then 'Check inventory transaction not posted 
                    If objaddon.objapplication.MessageBox("Do you want to refresh the Input line Items?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                Else
                    objaddon.objapplication.StatusBar.SetText("BOM refresh will not be happen due to Inventory transactions posted...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button11_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button11.ClickAfter
            Try
                If EditText21.Value = "" Then 'Check inventory transaction not posted 
                    Dim objrs As SAPbobsCOM.Recordset
                    Dim strsql As String = "", GetValue As String = "", WhsCode As String = "", InvUOM As String = ""
                    Dim i As Integer = 0
                    objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If EditText19.Value <> "" Then
                        If objaddon.HANA Then
                            EditText2.Value = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & EditText8.Value & "'")
                        Else
                            EditText2.Value = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & EditText8.Value & "'")
                        End If
                        If objaddon.HANA Then
                            strsql = "select T1.""U_Itemcode"",T1.""U_ItemDesc"",T1.""U_Qty"",T1.""U_Whse"",T1.""U_Unitprice"",T1.""U_Total"" "
                            strsql += vbCrLf + "from ""@MIPL_OBOM"" T0 join ""@MIPL_BOM1"" T1 on T0.""Code""=T1.""Code"" where T1.""U_SCType""='I' and T0.""Code""='" & EditText19.Value & "'"
                        Else
                            strsql = "select T1.U_Itemcode,T1.U_ItemDesc,T1.U_Qty,T1.U_Whse,T1.U_Unitprice,T1.U_Total "
                            strsql += vbCrLf + "from [@MIPL_OBOM] T0 join [@MIPL_BOM1] T1 on T0.Code=T1.Code where T1.U_SCType='I' and T0.Code='" & EditText19.Value & "'"
                        End If
                        objrs.DoQuery(strsql)
                        If BPWhse = "Y" Then
                            If objaddon.HANA Then
                                WhsCode = objaddon.objglobalmethods.getSingleValue("select ""U_WAREHOUSE"" from OCRD where ""CardCode""='" & EditText0.Value & "'")
                            Else
                                WhsCode = objaddon.objglobalmethods.getSingleValue("select U_WAREHOUSE from OCRD where CardCode='" & EditText0.Value & "'")
                            End If
                        End If
                        objform.Freeze(True)
                        odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_POR1")
                        Matrix0.Clear()
                        odbdsDetails.Clear()
                        If objrs.RecordCount > 0 Then
                            objaddon.objapplication.StatusBar.SetText("Re-Loading from SUB-BOM Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            While Not objrs.EoF
                                If CDbl(objrs.Fields.Item("U_Qty").Value) > 0 Then
                                    Matrix0.AddRow()
                                    Matrix0.GetLineData(Matrix0.VisualRowCount)
                                    odbdsDetails.SetValue("LineId", 0, i + 1)
                                    odbdsDetails.SetValue("U_Item1", 0, EditText19.Value.ToString)
                                    odbdsDetails.SetValue("U_Itemcode", 0, objrs.Fields.Item("U_Itemcode").Value.ToString)
                                    odbdsDetails.SetValue("U_ItemDesc", 0, objrs.Fields.Item("U_ItemDesc").Value.ToString)
                                    odbdsDetails.SetValue("U_Qty", 0, objrs.Fields.Item("U_Qty").Value.ToString)
                                    odbdsDetails.SetValue("U_BOMQty", 0, objrs.Fields.Item("U_Qty").Value.ToString)
                                    odbdsDetails.SetValue("U_PlanQty", 0, CDbl(objrs.Fields.Item("U_Qty").Value.ToString) * CDbl(objform.Items.Item("SQty").Specific.String))
                                    odbdsDetails.SetValue("U_ActPlanQty", 0, CDbl(objrs.Fields.Item("U_Qty").Value.ToString) * CDbl(objform.Items.Item("SQty").Specific.String))
                                    Try
                                        If objaddon.HANA Then
                                            InvUOM = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "'")
                                        Else
                                            InvUOM = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "'")
                                        End If
                                    Catch ex As Exception
                                    End Try
                                    odbdsDetails.SetValue("U_InvUom", 0, InvUOM)
                                    Dim Price As String = "", OnHand As String
                                    If VendorSPlPrice = "Y" Then
                                        If objaddon.HANA Then
                                            Price = objaddon.objglobalmethods.getSingleValue("Select ""AvgPrice"" from OITW where ""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and ""WhsCode""='" & objrs.Fields.Item("U_Whse").Value.ToString & "';")
                                        Else
                                            Price = objaddon.objglobalmethods.getSingleValue("Select AvgPrice from OITW where ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and WhsCode='" & objrs.Fields.Item("U_Whse").Value.ToString & "'")
                                        End If
                                        If Price <> "0" Then
                                            odbdsDetails.SetValue("U_Price", 0, Price)
                                        Else
                                            odbdsDetails.SetValue("U_Price", 0, objrs.Fields.Item("U_Unitprice").Value.ToString)
                                        End If
                                    Else
                                        odbdsDetails.SetValue("U_Price", 0, objrs.Fields.Item("U_Unitprice").Value.ToString)
                                    End If
                                    odbdsDetails.SetValue("U_WhsCode", 0, objrs.Fields.Item("U_Whse").Value.ToString)
                                    'odbdsDetails.SetValue("U_LineTot", 0, CDbl(odbdsDetails.GetValue("U_Qty", 0)) * CDbl(odbdsDetails.GetValue("U_Price", 0)))
                                    odbdsDetails.SetValue("U_LineTot", 0, ((CDbl(objrs.Fields.Item("U_Qty").Value.ToString) * CDbl(objform.Items.Item("SQty").Specific.String)) * CDbl(objrs.Fields.Item("U_Unitprice").Value.ToString)))
                                    odbdsDetails.SetValue("U_SubWhse", 0, WhsCode)
                                    If objaddon.HANA Then
                                        OnHand = objaddon.objglobalmethods.getSingleValue("select ifnull(""OnHand"",0) from  oitw where ""ItemCode""='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and ""WhsCode""='" & objrs.Fields.Item("U_Whse").Value.ToString & "'")
                                    Else
                                        OnHand = objaddon.objglobalmethods.getSingleValue("select Isnull(OnHand,0) from  oitw where ItemCode='" & objrs.Fields.Item("U_Itemcode").Value.ToString & "' and WhsCode='" & objrs.Fields.Item("U_Whse").Value.ToString & "'")
                                    End If
                                    odbdsDetails.SetValue("U_InStock", 0, OnHand)
                                    If WPrice = "Y" Then
                                        odbdsDetails.SetValue("U_LTType", 0, "1")
                                    End If
                                    If RevPlanQty = "Y" Then
                                        odbdsDetails.SetValue("U_RevBase", 0, "N")
                                    End If
                                    Matrix0.SetLineData(Matrix0.VisualRowCount)
                                    i += 1
                                End If
                                objrs.MoveNext()
                            End While
                            objform.Freeze(False)
                            If Matrix0.Columns.Item("PlanQty").Editable = False Or Matrix0.Columns.Item("revbase").Editable = False Then
                                Matrix0.Columns.Item("PlanQty").Editable = True
                                Matrix0.Columns.Item("revbase").Editable = True
                            End If
                            Matrix0.AutoResizeColumns()
                            'objform.Update()
                            'objform.Refresh()
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            'If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            '    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'End If
                            objaddon.objapplication.StatusBar.SetText("Re-Loaded from SUB-BOM Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Else
                            objaddon.objapplication.StatusBar.SetText("No records found from SUB_BOM...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                    Else
                        Exit Sub
                    End If
                End If

            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub EditText6_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText6.LostFocusAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    EditText10.Value = EditText6.Value
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub ComboBox0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                odbdsHeader.SetValue("DocNum", 0, objaddon.objglobalmethods.GetDocNum("SUBPO", CInt(ComboBox0.Selected.Value)))
            Catch ex As Exception
            End Try
        End Sub

        Private Sub EditText29_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText29.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                EditText30.Value = ""
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_OwnC")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "Active"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub EditText29_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText29.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        odbdsHeader.SetValue("U_OwnerC", 0, pCFL.SelectedObjects.Columns.Item("empID").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        odbdsHeader.SetValue("U_OwnerN", 0, pCFL.SelectedObjects.Columns.Item("lastName").Cells.Item(0).Value + "," + pCFL.SelectedObjects.Columns.Item("firstName").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    'objform.Update()
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub EditText28_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText28.KeyDownAfter
            Try
                If EditText19.Value = "" Then Exit Sub
                If ItemProcess = "Y" Then
                    If pVal.CharPressed = 13 Then
                        If objaddon.HANA Then
                            ProductQuery = "select Distinct 'N' ""Select"", T1.""U_Sequence"" ""Sequence"",T1.""U_Procname"" ""Process"",Case when T1.""U_Priority""='1' then 'Optional' else 'Mandatory' end as ""Priority"" "
                            ProductQuery += vbCrLf + "from ""@MIPL_OBOM"" T0 join ""@MIPL_BOM2"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & EditText19.Value & "' and T1.""U_Proccode""<>'' order by T1.""U_Sequence"" "
                        Else
                            ProductQuery = "select Distinct 'N' 'Select',T1.U_Sequence as Sequence,T1.U_Procname as Process,Case when T1.U_Priority='1' then 'Optional' else 'Mandatory' end as Priority"
                            ProductQuery += vbCrLf + "from [@MIPL_OBOM] T0 join [@MIPL_BOM2] T1 on T0.Code=T1.Code where T0.Code='" & EditText19.Value & "' and T1.U_Proccode<>'' order by T1.U_Sequence"
                        End If

                        FrmMultiSel = objaddon.objapplication.Forms.ActiveForm
                        If Not objaddon.FormExist("MULSEL") Then
                            FieldName = "tprocseq"
                            Dim Multiselect As New FrmMultiSelect
                            Multiselect.Show()
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText28_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText28.ValidateBefore
            Try
                If pVal.InnerEvent = True Then Exit Sub
                'If pVal.ItemChanged = False Then Exit Sub
                'If ItemProcess = "Y" Then
                '    If EditText28.Value = "" Then
                '        objaddon.objapplication.StatusBar.SetText("Item Process is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '        'BubbleEvent = False
                '    End If
                'End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_DoubleClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.DoubleClickBefore
            Try
                BubbleEvent = False
            Catch ex As Exception

            End Try

        End Sub


#End Region

#Region "Input Tab"

        Private Sub Matrix0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ClickAfter
            Try
                If pVal.Row = 0 Then
                    Matrix0.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                Else
                    Matrix0.SelectRow(pVal.Row, True, False)
                End If
                'Matrix_Total()
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try
                If pVal.Row = 0 Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                'Dim OnHand As Double
                'If pVal.ItemChanged = False Then Exit Sub
                Select Case pVal.ColUID
                    Case "Code"
                        If Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String <> "" Then
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                        End If
                    Case "Quant"
                        If CDbl(Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) > 0 Then
                            Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String = CDbl(objform.Items.Item("SQty").Specific.String) * CDbl(Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String)
                        End If
                    Case "Price"
                        'If WPrice = "N" Then
                        If CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String) > 0 Then
                            Matrix0.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                        End If
                        'End If
                        objform.Update()
                    Case "Whse"
                        'If Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String <> "" Then
                        '    If objaddon.HANA Then
                        '        OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String & "'")
                        '    Else
                        '        OnHand = objaddon.objglobalmethods.getSingleValue("select OnHand from  oitw where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String & "'")
                        '    End If
                        '    Matrix0.Columns.Item("Instock").Cells.Item(pVal.Row).Specific.String = OnHand
                        'Else
                        '    Matrix0.Columns.Item("Instock").Cells.Item(pVal.Row).Specific.String = 0
                        'End If
                    Case "distrule"
                        If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String <> "" Then
                            Dim code As String
                            code = Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String
                            Dim CostCodes() As String = code.Split(";")
                            If CostCodes.Count > 0 Then
                                If CostCodes.ElementAtOrDefault(0) <> "" Then Matrix0.Columns.Item("cc1").Cells.Item(pVal.Row).Specific.String = CostCodes(0)
                                If CostCodes.ElementAtOrDefault(1) <> "" Then Matrix0.Columns.Item("cc2").Cells.Item(pVal.Row).Specific.String = CostCodes(1)
                                If CostCodes.ElementAtOrDefault(2) <> "" Then Matrix0.Columns.Item("cc3").Cells.Item(pVal.Row).Specific.String = CostCodes(2)
                                If CostCodes.ElementAtOrDefault(3) <> "" Then Matrix0.Columns.Item("cc4").Cells.Item(pVal.Row).Specific.String = CostCodes(3)
                                If CostCodes.ElementAtOrDefault(4) <> "" Then Matrix0.Columns.Item("cc5").Cells.Item(pVal.Row).Specific.String = CostCodes(4)
                            End If
                        End If
                        'If CostCenter = "U" Then
                        '    If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String <> "" Then Exit Sub
                        '    Dim oform As New FrmDistRule
                        '    oform.Show()
                        'End If

                        'Case "tweight"
                        'If WPrice = "Y" Then
                        '    Dim objcombo As SAPbouiCOM.ComboBox
                        '    objcombo = Matrix0.Columns.Item("cmblttype").Cells.Item(pVal.Row).Specific
                        'Select Case pVal.ColUID
                        '    Case "tweight"
                        '        If objcombo.Selected Is Nothing Then Exit Sub
                        '        If objcombo.Selected.Value = "2" Then
                        '            If CDbl(Matrix0.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) > 0 Then
                        '                Matrix0.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                        '            Else
                        '                Exit Sub
                        '            End If
                        '        Else
                        '            Matrix0.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                        '        End If
                        'End Select
                        'End If
                        'objaddon.objapplication.Menus.Item("1300").Activate()
                        Matrix0.AutoResizeColumns()
                End Select
                'Matrix_Total()
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
            End Try
        End Sub

        Private Sub ButtonCombo1_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles ButtonCombo1.ClickBefore
            Try
                If ComboBox2.Selected.Value = "C" Then
                    objaddon.objapplication.StatusBar.SetText("Document status is closed.You cannot perform this action!!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ButtonCombo1.Item.Enabled = False
                    BubbleEvent = False : Exit Sub
                End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objaddon.objapplication.StatusBar.SetText("Please update the Sub-Contracting Screen...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            Catch ex As Exception

            End Try

        End Sub

        Private Sub ButtonCombo1_ComboSelectBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles ButtonCombo1.ComboSelectBefore
            Try
                'CalcPlanQty()
            Catch ex As Exception
            End Try
        End Sub

        Private Sub ButtonCombo1_ComboSelectAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ButtonCombo1.ComboSelectAfter
            If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
            Dim objrs, objrecset As SAPbobsCOM.Recordset
            Dim StrSql As String = "", ItemValid As String = "", StrQuery As String = ""
            Dim objMatrix As SAPbouiCOM.Matrix
            Dim objInvTrform As SAPbouiCOM.Form
            objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            'If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            Try
                If ButtonCombo1.Selected.Value = "Inventory Transfer" Then
                    'Dim CheckInvStat As Integer = 0
                    'StrSql = "Call ""MIPL_GetPendQty_Inv_GI_GR"" ('OWTR','WTR1'," & objform.Items.Item("txtentry").Specific.String & ")"
                    'If objaddon.HANA Then
                    '    StrSql = "Select Case when Sum(T2.""Quantity"") >0 then T0.""U_PlanQty""-Sum(T2.""Quantity"") else T0.""U_PlanQty"" end as ""PendQty"",T0.""U_Itemcode"""
                    '    StrSql += vbCrLf + " from ""@MIPL_POR1"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" left join OWTR T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join WTR1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T0.""U_Itemcode""=T2.""ItemCode"" and T0.""LineId""= T2.""U_LineID"""
                    '    StrSql += vbCrLf + "where T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " group by T0.""U_PlanQty"",T0.""U_Itemcode"" order by ""PendQty"""
                    'Else
                    '    StrSql = "Select Case when Sum(T2.Quantity) >0 then T0.U_PlanQty-Sum(T2.Quantity) else T0.U_PlanQty end as PendQty,T0.U_Itemcode"
                    '    StrSql += vbCrLf + "from [@MIPL_POR1] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry left join OWTR T3 on T3.U_SubConNo=T1.DocEntry left join WTR1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode and T0.LineId= T2.U_LineID"
                    '    StrSql += vbCrLf + "where T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " group by T0.U_PlanQty,T0.U_Itemcode order by PendQty"
                    'End If
                    'If objaddon.HANA Then
                    '    StrSql = "Select Case when Sum(T2.""Quantity"") >0 then T0.""U_PlanQty""-Sum(T2.""Quantity"") else T0.""U_PlanQty"" end "
                    '    StrSql += vbCrLf + "+ ifnull((select Top 1 sum(T1.""U_Qty"") from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""U_Type""='2' and T1.""U_Status""='C' and T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' "
                    '    StrSql += vbCrLf + "group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc),0) as ""PendQty"",T0.""U_Itemcode"""
                    '    StrSql += vbCrLf + " from ""@MIPL_POR1"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" left join OWTR T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join WTR1 T2  on T3.""DocEntry""=T2.""DocEntry"" And T0.""U_Itemcode""=T2.""ItemCode"" And T0.""LineId""= T2.""U_LineID"""
                    '    StrSql += vbCrLf + "where T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " group by T0.""U_PlanQty"",T0.""U_Itemcode"" order by ""PendQty"""
                    'Else
                    '    StrSql = "Select Case when Sum(T2.Quantity) >0 then T0.U_PlanQty-Sum(T2.Quantity) else T0.U_PlanQty end "
                    '    StrSql += vbCrLf + "+ isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' "
                    '    StrSql += vbCrLf + "group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0) as PendQty,T0.U_Itemcode"
                    '    StrSql += vbCrLf + "from [@MIPL_POR1] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry left join OWTR T3 on T3.U_SubConNo=T1.DocEntry left join WTR1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode and T0.LineId= T2.U_LineID"
                    '    StrSql += vbCrLf + "where T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " group by T0.U_PlanQty,T0.U_Itemcode order by PendQty"
                    'End If
                    If objaddon.HANA Then
                        StrSql = "Select (A.""PendQty"" + A.""Return Qty"") ""Qty"",* from (Select T0.""DocEntry"",T0.""DocNum"",T1.""U_ItemCode"",T1.""U_PlanQty"" ""PlanQty"",T1.""U_WhsCode"",T1.""U_SubWhse"",T1.""LineId""," '+" & CDbl(Matrix0.Columns.Item("10").Cells.Item(i).Specific.String) & "
                        StrSql += vbCrLf + "(Select ifnull(Sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where (A.""U_SubConNo""=T1.""DocEntry"" or B.""U_S_DocEntry""=T1.""DocEntry"") and B.""ItemCode""=T1.""U_Itemcode"") as ""Transfer Qty"","
                        StrSql += vbCrLf + "(Select ifnull(Sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_ScrapNum""=T1.""DocEntry"" and B.""ItemCode""=T1.""U_Itemcode"") as ""Return Qty"","
                        StrSql += vbCrLf + "(Select case when sum(B.""Quantity"")>0 then T1.""U_PlanQty""-ifnull(Sum(B.""Quantity""),0) else T1.""U_PlanQty"" end from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where (A.""U_SubConNo""=T1.""DocEntry"" or B.""U_S_DocEntry""=T1.""DocEntry"") and B.""ItemCode""=T1.""U_Itemcode"") as ""PendQty"""
                        StrSql += vbCrLf + "from ""@MIPL_OPOR"" T0 join ""@MIPL_POR1"" T1 on T0.""DocEntry""=T1.""DocEntry"") A where  A.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' order by (A.""PendQty"" + A.""Return Qty"") desc "
                    Else
                        StrSql = "Select (A.[PendQty] + A.[Return Qty]) Qty,* from (Select T0.DocEntry,T0.DocNum,T1.U_ItemCode,T1.U_PlanQty [PlanQty],T1.U_WhsCode,T1.U_SubWhse,T1.LineId," '+" & CDbl(Matrix0.Columns.Item("10").Cells.Item(i).Specific.String) & "
                        StrSql += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where (A.U_SubConNo=T1.DocEntry or B.U_S_DocEntry=T1.DocEntry) and B.ItemCode=T1.U_Itemcode) as [Transfer Qty],"
                        StrSql += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T1.DocEntry and B.ItemCode=T1.U_Itemcode) as [Return Qty],"
                        StrSql += vbCrLf + "(Select case when sum(B.Quantity)>0 then T1.U_PlanQty-isnull(Sum(B.Quantity),0) else T1.U_PlanQty end from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where (A.U_SubConNo=T1.DocEntry or B.U_S_DocEntry=T1.DocEntry) and B.ItemCode=T1.U_Itemcode) as [PendQty]"
                        StrSql += vbCrLf + "from [@MIPL_OPOR] T0 join [@MIPL_POR1] T1 on T0.DocEntry=T1.DocEntry) A where  A.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' order by (A.[PendQty] + A.[Return Qty]) desc  "
                    End If
                    objrs.DoQuery(StrSql)
                    If CDbl(objrs.Fields.Item("Qty").Value) <= 0.0 Then
                        objaddon.objapplication.StatusBar.SetText("Inventory Transferred for the Planned Quantity...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ButtonCombo1.Caption = "Generate Document"
                        Exit Sub
                    End If
                    'For Result As Integer = 0 To objrs.RecordCount - 1
                    '    If CDbl(objrs.Fields.Item("Qty").Value) <= 0.0 Then
                    '        CheckInvStat += 1
                    '    End If
                    '    objrs.MoveNext()
                    'Next
                    'If CheckInvStat > 0 Then
                    '    objaddon.objapplication.StatusBar.SetText("Inventory Transferred for the Planned Quantity...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    ButtonCombo1.Caption = "Generate Document"
                    '    Exit Sub
                    'End If

                    ButtonCombo1.Caption = "Generate Document"
                    objaddon.objapplication.StatusBar.SetText("Looking for Inventory Transfer screen Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objaddon.objapplication.Menus.Item("3080").Activate()
                    objInvTrform = objaddon.objapplication.Forms.ActiveForm
                    objInvTrform = objaddon.objapplication.Forms.Item(objInvTrform.UniqueID)
                    objInvTrform.Visible = True

                    If objInvTrform.IsSystem = False Then
                        objaddon.objapplication.StatusBar.SetText("Problem with the Inventory Transfer screen Please check...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If
                    objInvTrform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                    objInvTrform.State = SAPbouiCOM.BoFormStateEnum.fs_Restore
                    objMatrix = objInvTrform.Items.Item("23").Specific

                    Try
                        objaddon.objapplication.StatusBar.SetText("Data Loading to Inventory Transfer Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        SubConTypeCount = pVal.FormTypeCount
                        objform.Freeze(True)
                        objInvTrform.Freeze(True)
                        objInvTrform.Items.Item("3").Specific.String = odbdsHeader.GetValue("U_CardCode", 0)
                        objInvTrform.Items.Item("22").Specific.String = "From Sub-Con Add-on" & " -> " & Now.ToString
                        objInvTrform.Items.Item("18").Specific.String = Matrix0.Columns.Item("Whse").Cells.Item(1).Specific.String
                        objInvTrform.Items.Item("1470000101").Specific.String = Matrix0.Columns.Item("SubWhse").Cells.Item(1).Specific.String
                        objInvTrform.Items.Item("U_SubConNo").Specific.String = objform.Items.Item("txtentry").Specific.String 'SubPOHeader.GetValue("DocEntry", 0) 'CStr(objform.Items.Item("txtdocnum").Specific.String)
                        'Dim oUDFForm As SAPbouiCOM.Form
                        'oUDFForm = objaddon.objapplication.Forms.Item(objInvTrform.UDFFormUID)
                        'oUDFForm.Items.Item("U_SubConNo").Enabled = False
                        If BranchEnabled = "Y" Then
                            Dim cmbSeries As SAPbouiCOM.ComboBox
                            cmbSeries = objInvTrform.Items.Item("40").Specific
                            Dim series As String
                            If objaddon.HANA Then
                                series = objaddon.objglobalmethods.getSingleValue("select Top 1 ""Series"" From NNM1 where ""ObjectCode""='67' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                              " and ""BPLId""=(select Top 1 ""BPLid"" from OWHS where ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(1).Specific.String & "') order by ""CreateDate"" desc")
                            Else
                                series = objaddon.objglobalmethods.getSingleValue("select Top 1 Series From NNM1 where ObjectCode='67' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                                                                                                  " and BPLId=(select Top 1 BPLid from OWHS where WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(1).Specific.String & "') ") 'order by CreateDate desc")
                            End If

                            If series <> "" Then
                                cmbSeries.Select(series, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        End If
                        'If objaddon.HANA Then
                        '    StrQuery = "select Case when Sum(T2.""Quantity"") >0 then T0.""U_PlanQty""-Sum(T2.""Quantity"") else T0.""U_PlanQty"" end as ""PendQty"",T0.""U_Itemcode"",T0.""U_WhsCode"",T0.""U_SubWhse"",T0.""LineId"""
                        '    StrQuery += vbCrLf + " from ""@MIPL_POR1"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" "
                        '    StrQuery += vbCrLf + " left join owtr T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join wtr1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T0.""U_Itemcode""=T2.""ItemCode"" and T0.""LineId""= T2.""U_LineID"""
                        '    StrQuery += vbCrLf + " where T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " group by T0.""U_PlanQty"",T0.""U_Itemcode"",T0.""U_WhsCode"",T0.""U_SubWhse"",T0.""LineId"""
                        'Else
                        '    StrQuery = "select Case when Sum(T2.Quantity) >0 then T0.U_PlanQty-Sum(T2.Quantity) else T0.U_PlanQty end as PendQty,T0.U_Itemcode,T0.U_WhsCode,T0.U_SubWhse,T0.LineId"
                        '    StrQuery += vbCrLf + " from [@MIPL_POR1] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry "
                        '    StrQuery += vbCrLf + " left join owtr T3 on T3.U_SubConNo=T1.DocEntry left join wtr1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode and T0.LineId= T2.U_LineID"
                        '    StrQuery += vbCrLf + " where T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " group by T0.U_PlanQty,T0.U_Itemcode,T0.U_WhsCode,T0.U_SubWhse,T0.LineId"
                        'End If

                        objrecset.DoQuery(StrSql)
                        If objrecset.RecordCount > 0 Then
                            'If objMatrix.Columns.Item("U_PlanQty").Editable = False Or objMatrix.Columns.Item("U_LineID").Editable = False Then
                            '    objMatrix.Columns.Item("U_PlanQty").Editable = True
                            '    objMatrix.Columns.Item("U_LineID").Editable = True
                            'End If
                            If objMatrix.Columns.Item("U_PlanQty").Editable = False Then objMatrix.Columns.Item("U_PlanQty").Editable = True
                            If objMatrix.Columns.Item("U_LineID").Editable = False Then objMatrix.Columns.Item("U_LineID").Editable = True
                            For i As Integer = 0 To objrecset.RecordCount - 1
                                If objrecset.Fields.Item("PendQty").Value <> 0 And objrecset.Fields.Item("PendQty").Value.ToString <> String.Empty Then
                                    If objaddon.HANA Then
                                        ItemValid = objaddon.objglobalmethods.getSingleValue(" select 1 from OITM where ""ItemCode""='" & objrecset.Fields.Item("U_Itemcode").Value & "'")
                                    Else
                                        ItemValid = objaddon.objglobalmethods.getSingleValue(" select 1 from OITM where ItemCode='" & objrecset.Fields.Item("U_Itemcode").Value & "'")
                                    End If
                                    If ItemValid = "1" Then
                                        objMatrix.Columns.Item("1").Cells.Item(i + 1).Specific.String = objrecset.Fields.Item("U_Itemcode").Value 'Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String
                                        objMatrix.Columns.Item("10").Cells.Item(i + 1).Specific.String = objrecset.Fields.Item("PendQty").Value 'MatrixID.Columns.Item("Quant").Cells.Item(i).Specific.String
                                        objMatrix.Columns.Item("U_PlanQty").Cells.Item(i + 1).Specific.String = objrecset.Fields.Item("PendQty").Value
                                        objMatrix.Columns.Item("U_LineID").Cells.Item(i + 1).Specific.String = objrecset.Fields.Item("LineId").Value
                                        objMatrix.Columns.Item("1470001039").Cells.Item(i + 1).Specific.String = objrecset.Fields.Item("U_WhsCode").Value 'Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.String
                                        objMatrix.Columns.Item("5").Cells.Item(i + 1).Specific.String = objrecset.Fields.Item("U_SubWhse").Value ' Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String
                                        If ItemProcess = "Y" Then
                                            If Trim(EditText20.Value) <> "" Then
                                                objMatrix.Columns.Item("U_Process").Cells.Item(i + 1).Specific.String = Trim(EditText20.Value)
                                            End If
                                        End If
                                    End If
                                End If
                                objrecset.MoveNext()
                            Next
                            objMatrix.Columns.Item("1").Cells.Item(1).Click()
                        End If
                        If objMatrix.VisualRowCount = 1 Then
                            If objMatrix.Columns.Item("1").Cells.Item(1).Specific.String = "" Then
                                objaddon.objapplication.StatusBar.SetText("No More Quantity to Post the Inventory Transfer ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objInvTrform.Items.Item("3").Specific.String = ""
                                objInvTrform.Items.Item("22").Specific.String = ""
                                objInvTrform.Items.Item("U_SubConNo").Specific.String = ""
                            End If
                        End If
                        'objMatrix.Columns.Item("U_PlanQty").Editable = False
                        'objMatrix.Columns.Item("U_LineID").Editable = False
                        objrecset = Nothing
                        objInvTrform.Freeze(False)
                        objform.Freeze(False)
                        objaddon.objapplication.StatusBar.SetText("Data Loaded to Inventory Transfer Screen ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Catch ex As Exception
                        objInvTrform.Freeze(False)
                        objInvTrform.Close()
                        objform.Freeze(False)
                        objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
            'End If

        End Sub

        Private Sub Folder0_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder0.PressedAfter
            Try
                objform.Freeze(True)
                'If ComboBox2.Selected.Value = "C" Then
                '    Matrix0.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                'Else
                '    If AutoAssignItem = "Y" Then
                '        Matrix0.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                '    End If
                'End If
                'Matrix_Total()
                'objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                'objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                'Matrix0.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("PlanQty").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual
                'objaddon.objapplication.Menus.Item("1300").Activate()
                objform.Settings.MatrixUID = "MtxinputN"
                Matrix0.AutoResizeColumns()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Matrix0_LinkPressedBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix0.LinkPressedBefore
            Try
                Select Case pVal.ColUID
                    Case "Code"
                        Dim ColItem As SAPbouiCOM.Column = Matrix0.Columns.Item("Code")
                        Dim objlink As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
                        Dim Status As String = ""
                        If objaddon.HANA Then
                            Status = objaddon.objglobalmethods.getSingleValue("select 1 from OITM where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                        Else
                            Status = objaddon.objglobalmethods.getSingleValue("select 1 from OITM where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                        End If
                        If Status = "1" Then
                            objlink.LinkedObjectType = "4"
                            objlink.Item.LinkTo = "Code"
                        Else
                            objlink.LinkedObjectType = "290"
                            objlink.Item.LinkTo = "Code"
                        End If
                    Case "distrule"
                        If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String <> "" Then
                            Link_Value = Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String
                            Dim oform As New FrmDistRule
                            oform.Show()
                        End If
                End Select

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix0_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix0.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                Dim LocCode As String
                If pVal.ColUID = "Code" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_I")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "InvntItem"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "Y"
                        ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "SubWhse" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_Whse")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()
                        If ToWhseInInput = "Y" Then
                            If objaddon.HANA Then
                                LocCode = objaddon.objglobalmethods.getSingleValue("select ""U_LCode"" from ""@MIPL_GEN""")
                            Else
                                LocCode = objaddon.objglobalmethods.getSingleValue("select U_LCode from [@MIPL_GEN]")
                            End If
                            Dim SplitLoc() As String = Split(LocCode, ",")
                            oCond = oConds.Add()
                            oCond.Alias = "Inactive"
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                            oCond.CondVal = "Y"
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                            For Val As Integer = 0 To SplitLoc.Length - 1
                                If Val = 0 Then
                                    oCond = oConds.Add()
                                    oCond.Alias = "Location"
                                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                    oCond.CondVal = SplitLoc(Val)
                                    'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                                    'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                                    'oCond = oConds.Add()
                                    'oCond.Alias = "WhsCode"
                                    'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                                    'oCond.CondVal = Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String
                                Else
                                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                    oCond = oConds.Add()
                                    oCond.Alias = "Location"
                                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                    oCond.CondVal = SplitLoc(Val)
                                    'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                                    'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                                    'oCond = oConds.Add()
                                    'oCond.Alias = "WhsCode"
                                    'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                                    'oCond.CondVal = Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String
                                End If
                            Next
                            'oCond = oConds.Add()
                            'oCond.Alias = "Location"
                            'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            'oCond.CondVal = LocCode
                        End If
                        ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc1" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_24")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "1"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc2" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_25")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "2"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc3" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_26")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "3"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc4" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_27")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "4"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc5" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_28")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "5"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                    'ElseIf pVal.ColUID = "distrule" Then
                    '    If CostCenter = "U" Then
                    '        If Not objaddon.FormExist("DistRule") Then
                    '            Dim oform As New FrmDistRule
                    '            oform.Show()
                    '        End If
                    '    End If
                    '    BubbleEvent = False
                End If

            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub Matrix0_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ChooseFromListAfter
            Try
                Dim HSNCode As String = ""
                Dim OnHand, SQty, Qty As Double
                If pVal.ColUID = "Code" And pVal.ActionSuccess = True Then
                    Try
                        'If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                                Matrix0.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                                Matrix0.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                                Matrix0.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                                Matrix0.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                            End Try
                        End If
                        SQty = CDbl(objform.Items.Item("SQty").Specific.String)
                        Qty = CDbl(Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String)
                        If Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String <> "" Then
                            Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = "1"
                            Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String = SQty * Qty
                        End If
                        If objaddon.HANA Then
                            HSNCode = objaddon.objglobalmethods.getSingleValue("select ""ChapterID"" from OCHP where ""AbsEntry""='" & pCFL.SelectedObjects.Columns.Item("ChapterID").Cells.Item(0).Value & "'")
                        Else
                            HSNCode = objaddon.objglobalmethods.getSingleValue("select ChapterID from OCHP where AbsEntry='" & pCFL.SelectedObjects.Columns.Item("ChapterID").Cells.Item(0).Value & "'")
                        End If

                        If HSNCode <> "" Then
                            Matrix0.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = HSNCode
                        Else
                            Matrix0.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = 0
                        End If
                        'objform.Update()
                        'Try
                        '    If objaddon.HANA Then
                        '        Matrix0.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                        '    Else
                        '        Matrix0.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                        '    End If
                        'Catch ex As Exception
                        'End Try
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "Whse" And pVal.ActionSuccess = True Then
                    Try
                        'If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                            End Try
                        End If
                        If Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String <> "" Then
                            If objaddon.HANA Then
                                OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String & "'")
                            Else
                                OnHand = objaddon.objglobalmethods.getSingleValue("select OnHand from  oitw where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String & "'")
                            End If
                            Matrix0.Columns.Item("Instock").Cells.Item(pVal.Row).Specific.String = OnHand
                        Else
                            Matrix0.Columns.Item("Instock").Cells.Item(pVal.Row).Specific.String = 0
                        End If
                        'objform.Update()
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "SubWhse" And pVal.ActionSuccess = True Then
                    Try
                        'If pVal.ActionSuccess = False Then Exit Sub
                        odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_POR1")
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("SubWhse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                                'Matrix0.FlushToDataSource()
                                'odbdsDetails.SetValue("U_SubWhse", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value)
                                'Matrix0.LoadFromDataSource()
                                'Matrix0.Columns.Item("SubWhse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                            Catch ex As Exception
                                'Matrix0.Columns.Item("SubWhse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                            End Try
                        End If
                        'objform.Update()
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc1" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc1").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc1").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc2" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc2").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc2").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc3" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc3").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc3").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc4" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc4").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc4").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc5" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc4").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc4").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                Matrix0.AutoResizeColumns()
                'objaddon.objapplication.Menus.Item("1300").Activate()
                GC.Collect()
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Button7_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button7.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                AutoProductionOrder()
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button7_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button7.ClickBefore
            Try
                If ComboBox2.Selected.Value = "C" Then
                    objaddon.objapplication.StatusBar.SetText("Document status is closed.You cannot perform this action!!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Button7.Item.Enabled = False
                    Exit Sub
                End If
                If Matrix0.VisualRowCount = 1 Then
                    If Matrix0.Columns.Item("Code").Cells.Item(1).Specific.String = "" Then
                        If ItemBOM = "Y" Then
                            objaddon.objapplication.StatusBar.SetText("Please select a Sub Item-BOM... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                        Else
                            objaddon.objapplication.StatusBar.SetText("Please select a line level data in Input Tab... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                        End If
                    End If
                End If

                For Input As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("Code").Cells.Item(Input).Specific.String <> "" Then
                        If Matrix0.Columns.Item("Whse").Cells.Item(Input).Specific.String = "" Then
                            objaddon.objapplication.StatusBar.SetText("Please update the sub-Con warehouse...in Line : " & Input, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Matrix0_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ValidateAfter
            Try
                'Dim sumdec As String
                If pVal.Row = 0 Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                objform.Freeze(True)
                Matrix0.FlushToDataSource()
                Select Case pVal.ColUID
                    Case "PlanQty", "revbase"
                        If pVal.ItemChanged = False Then Exit Sub
                        objcombo = Matrix0.Columns.Item("revbase").Cells.Item(pVal.Row).Specific
                        If objcombo.Selected Is Nothing Then Exit Sub
                        If RevPlanQty = "Y" Then
                            If objcombo.Selected.Value = "Y" Then
                                'If objaddon.HANA Then
                                '    sumdec = objaddon.objglobalmethods.getSingleValue("Select ""QtyDec"" from OADM")
                                'Else
                                '    sumdec = objaddon.objglobalmethods.getSingleValue("Select QtyDec from OADM")
                                'End If
                                odbdsDetails.SetValue("U_Qty", pVal.Row - 1, CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) / CDbl(EditText25.Value))

                                'Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) / CDbl(EditText25.Value) ' Math.Round(CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) / CDbl(EditText25.Value), CInt(sumdec))
                            Else
                                odbdsDetails.SetValue("U_Qty", pVal.Row - 1, CDbl(Matrix0.Columns.Item("bomqty").Cells.Item(pVal.Row).Specific.String))
                                'Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("bomqty").Cells.Item(pVal.Row).Specific.String)
                            End If
                            Matrix0.LoadFromDataSource()
                            objform.Update()
                            objform.Refresh()
                        End If
                    Case "tweight", "cmblttype"
                        If pVal.ItemChanged = False Then Exit Sub
                        objcombo = Matrix0.Columns.Item("cmblttype").Cells.Item(pVal.Row).Specific
                        If objcombo.Selected Is Nothing Then Exit Sub
                        If WPrice = "Y" Then
                            If objcombo.Selected.Value = "2" Then
                                If Val(Matrix0.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) > 0 Then
                                    odbdsDetails.SetValue("U_LineTot", pVal.Row - 1, CDbl(Matrix0.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String))
                                    'Matrix0.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                                Else
                                    Exit Sub
                                End If
                            Else
                                odbdsDetails.SetValue("U_LineTot", pVal.Row - 1, CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String))
                                'Matrix0.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                            End If
                            Matrix0.LoadFromDataSource()
                            objform.Update()
                            objform.Refresh()
                        End If
                        'Case "PlanQty"
                        '    If PlanQtyUDFEditable = "Y" Then
                        '        If Val(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) > 0 Then
                        '            If objaddon.HANA Then
                        '                sumdec = objaddon.objglobalmethods.getSingleValue("Select ""SumDec"" from OADM")
                        '            Else
                        '                sumdec = objaddon.objglobalmethods.getSingleValue("Select SumDec from OADM")
                        '            End If
                        '            Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = Math.Round(CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) / CDbl(EditText25.Value), CInt(sumdec))
                        '        Else
                        '            Exit Sub
                        '        End If
                        '        objform.Update()
                        '        objform.Refresh()
                        '    End If

                End Select

            Catch ex As Exception
            Finally
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Matrix0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ComboSelectAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub

                Select Case pVal.ColUID
                    Case "revbase"
                        If pVal.ItemChanged = False Then Exit Sub
                        objcombo = Matrix0.Columns.Item("revbase").Cells.Item(pVal.Row).Specific
                        If objcombo.Selected Is Nothing Then Exit Sub
                        If RevPlanQty = "Y" Then
                            If objcombo.Selected.Value = "Y" Then
                                'Matrix0.CommonSetting.SetCellEditable(pVal.Row, 8, True) ' Plan Qty
                                Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) / CDbl(EditText25.Value) ' Math.Round(CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) / CDbl(EditText25.Value), CInt(sumdec))
                            Else
                                'Matrix0.CommonSetting.SetCellEditable(pVal.Row, 8, False) ' Plan Qty
                                Matrix0.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("bomqty").Cells.Item(pVal.Row).Specific.String)
                            End If
                            objform.Update()
                            objform.Refresh()
                        End If
                    Case "cmblttype"
                        If WPrice = "Y" Then
                            objcombo = Matrix0.Columns.Item("cmblttype").Cells.Item(pVal.Row).Specific
                            If objcombo.Selected Is Nothing Then Exit Sub
                            If pVal.ItemChanged = False Then Exit Sub
                            If CDbl(Matrix0.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) = 0 Then Exit Sub
                            If objcombo.Selected.Value = "2" Then
                                If CDbl(Matrix0.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) > 0 Then
                                    Matrix0.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                                Else
                                    Exit Sub
                                End If
                            Else
                                Matrix0.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                            End If
                            objform.Update()
                            objform.Refresh()
                        End If
                End Select

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.KeyDownAfter
            Try
                Dim ColID As Integer = Matrix0.GetCellFocus().ColumnIndex
                If pVal.CharPressed = 38 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then  'up
                    Matrix0.SetCellFocus(pVal.Row - 1, ColID)
                    Matrix0.SelectRow(pVal.Row - 1, True, False)
                ElseIf pVal.CharPressed = 40 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'down
                    Matrix0.SetCellFocus(pVal.Row + 1, ColID)
                    Matrix0.SelectRow(pVal.Row + 1, True, False)
                End If
                Select Case pVal.ColUID
                    Case "distrule"
                        If pVal.CharPressed = 13 Then
                            If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String = "" Then
                                If CostCenter = "U" Then
                                    OEForm = objaddon.objapplication.Forms.ActiveForm
                                    Dim oform As New FrmDistRule
                                    oform.Show()
                                End If
                            End If
                        End If

                End Select
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button8_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button8.ClickBefore
            Try
                If ComboBox2.Selected.Value = "C" Then
                    objaddon.objapplication.StatusBar.SetText("Document status is closed.You cannot perform this action!!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Button8.Item.Enabled = False
                    Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

#End Region

#Region "Output Tab"

        Private Sub Matrix3_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix3.ClickAfter
            Try
                If pVal.Row = 0 Then
                    Matrix3.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                Else
                    Matrix3.SelectRow(pVal.Row, True, False)
                End If
                Try
                    If Matrix3.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.String = "" Then
                        Matrix3.CommonSetting.SetCellEditable(pVal.Row, 20, True)
                    ElseIf Matrix3.Columns.Item("RefNo").Cells.Item(pVal.Row).Specific.String = "" Then
                        Matrix3.CommonSetting.SetCellEditable(pVal.Row, 21, True)
                    End If
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix3_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix3.LostFocusAfter
            'Dim PendQty As String
            Try
                If pVal.Row <= 0 Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                Dim Quantity As Double, Price As Double
                objform.Freeze(True)
                Matrix3.FlushToDataSource()
                Quantity = CDbl(odbdsDetails2.GetValue("U_Qty", pVal.Row - 1)) 'Val(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String)
                Price = odbdsDetails2.GetValue("U_Price", pVal.Row - 1) 'Val(Matrix3.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)

                Select Case pVal.ColUID
                    Case "Code"
                        If odbdsDetails2.GetValue("U_Itemcode", pVal.Row - 1) <> "" Then 'Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String
                            If objaddon.HANA Then
                                odbdsDetails2.SetValue("U_Itemcode", pVal.Row - 1, objaddon.objglobalmethods.getSingleValue("SELECT  ""ItemName"" FROM OITM where ""ItemCode""='" & odbdsDetails2.GetValue("U_Itemcode", pVal.Row - 1) & "'"))
                                'Matrix3.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("SELECT  ""ItemName"" FROM OITM where ""ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                            Else
                                odbdsDetails2.SetValue("U_Itemcode", pVal.Row - 1, objaddon.objglobalmethods.getSingleValue("SELECT  ItemName FROM OITM where ItemCode='" & odbdsDetails2.GetValue("U_Itemcode", pVal.Row - 1) & "'"))
                                'Matrix3.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("SELECT  ItemName FROM OITM where ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                            End If
                            If objaddon.HANA Then
                                odbdsDetails2.SetValue("U_HSNCode", pVal.Row - 1, objaddon.objglobalmethods.getSingleValue("select T1.""ChapterID"" from OITM T0 join OCHP T1 on T0.""ChapterID""=T1.""AbsEntry"" where T0.""ItemCode""='" & odbdsDetails2.GetValue("U_Itemcode", pVal.Row - 1) & "'"))
                                'Matrix3.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("select T1.""ChapterID"" from OITM T0 join OCHP T1 on T0.""ChapterID""=T1.""AbsEntry"" where T0.""ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                            Else
                                odbdsDetails2.SetValue("U_HSNCode", pVal.Row - 1, objaddon.objglobalmethods.getSingleValue("select T1.ChapterID from OITM T0 join OCHP T1 on T0.ChapterID=T1.AbsEntry where T0.ItemCode='" & odbdsDetails2.GetValue("U_Itemcode", pVal.Row - 1) & "'"))
                                'Matrix3.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("select T1.ChapterID from OITM T0 join OCHP T1 on T0.ChapterID=T1.AbsEntry where T0.ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                            End If
                        End If
                    Case "Quant"
                        If POItem = "Y" Then
                            odbdsDetails2.SetValue("U_Itemcode", pVal.Row - 1, EditText17.Value)
                            'Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String = EditText17.Value
                            If objaddon.HANA Then
                                odbdsDetails2.SetValue("U_ItemDesc", pVal.Row - 1, objaddon.objglobalmethods.getSingleValue("select ""ItemName"" from OITM where ""ItemCode""='" & EditText17.Value & "'"))
                                'Matrix3.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("select ""ItemName"" from OITM where ""ItemCode""='" & EditText17.Value & "'")
                            Else
                                odbdsDetails2.SetValue("U_ItemDesc", pVal.Row - 1, objaddon.objglobalmethods.getSingleValue("select ItemName from OITM where ItemCode='" & EditText17.Value & "'"))
                                'Matrix3.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("select ItemName from OITM where ItemCode='" & EditText17.Value & "'")
                            End If
                        End If

                        If odbdsDetails2.GetValue("U_Qty", pVal.Row - 1) > 0 Then 'Val(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) 
                            odbdsDetails2.SetValue("U_LineTot", pVal.Row - 1, Quantity * Price)
                            'Matrix3.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = Quantity * Price
                        End If
                        'objform.Freeze(False)
                    Case "Price"
                        'objform.Freeze(True)
                        If odbdsDetails2.GetValue("U_Price", pVal.Row - 1) <> "" Then 'Matrix3.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String
                            odbdsDetails2.SetValue("U_LineTot", pVal.Row - 1, Quantity * Price)
                            'Matrix3.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = Quantity * Price
                        End If
                        'objform.Freeze(False)
                End Select
                Matrix3.LoadFromDataSource()
                objform.Update()
                'objaddon.objapplication.Menus.Item("1300").Activate()
                'End If
            Catch ex As Exception
                'objform.Freeze(False)
            Finally
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub ButtonCombo0_ComboSelectAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ButtonCombo0.ComboSelectAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If ComboBox2.Selected.Value = "C" Then
                    objaddon.objapplication.StatusBar.SetText("Document status is closed.You cannot perform this action!!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ButtonCombo0.Item.Enabled = False
                    Exit Sub
                End If
                'If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
                'If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                If ButtonCombo0.Selected.Value = "Issue & Receipt" Then
                    If ReceiptAutoLoad <> "Y" Then
                        ButtonCombo0.Caption = "Generate Document"
                        If Matrix3.Columns.Item("Code").Cells.Item(Matrix3.VisualRowCount).Specific.string <> "" Then
                            If objaddon.objapplication.MessageBox("Do you want to Post the Goods Issue & Receipt ?", 2, "Yes", "No") <> 1 Then Exit Sub
                            PostingGoodsIssue()
                        End If
                        'If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    End If
                End If
                'End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Folder1_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder1.PressedAfter
            Try
                If AutoProduction = "Y" Then
                Else
                    Button8.Item.Visible = False
                End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                'If ComboBox2.Selected.Value = "C" Then
                '    Matrix3.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                'Else
                '    If AutoAssignItem = "Y" Then
                '        Matrix3.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                '        'objform.Items.Item("mtxoutput").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                '    End If
                'End If
                objform.Freeze(True)
                If Matrix3.VisualRowCount = 0 Then
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                End If
                'Matrix_Total()
                If AutoAssignItem = "Y" Then
                    If Matrix3.VisualRowCount = 1 Then Matrix3.Columns.Item("Code").Cells.Item(1).Specific.String = EditText8.Value
                    'Matrix3.Columns.Item("Quant").Cells.Item(Matrix3.VisualRowCount).Specific.String = EditText25.Value
                    If objaddon.objglobalmethods.AutoAssign_SubItem(Folder1, Matrix3) Then
                        If Not objform.Items.Item("txtstat").Specific.Selected.Value = "C" Then 'objform.Items.Item("txtstat").Specific.String
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "GRNo", "#")
                            objaddon.objglobalmethods.GetSubItemQty(objform, Matrix3)
                            'If VendorSPlPrice = "2" Then
                            '    CalcGRPrice()
                            'End If
                        End If
                    End If
                Else
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                    Matrix3.Columns.Item("Date").Cells.Item(Matrix3.VisualRowCount).Specific.String = Now.Date.ToString("dd/MM/yy")
                End If
                FieldDisableInMatrix()
                Status_Closing_RowDisable()
                'Dim txtDate As SAPbouiCOM.EditText
                'txtDate = Matrix3.Columns.Item("Date").Cells.Item(Matrix3.VisualRowCount).Specific
                'txtDate.Value = Now.Date.ToString("yyyyMMdd")

                'Matrix3.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix3.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix3.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix3.Columns.Item("ProCost").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix3.Columns.Item("TProCost").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objform.Settings.MatrixUID = "mtxoutput"
                Matrix3.AutoResizeColumns()
                'objaddon.objapplication.Menus.Item("1300").Activate()
                objform.Freeze(False)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub Matrix3_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix3.ChooseFromListAfter
            If pVal.ColUID = "Code" And pVal.ActionSuccess = True Then
                Try
                    Dim HSNCode As String = ""
                    Matrix3.FlushToDataSource()
                    If pVal.ActionSuccess = False Then Exit Sub
                    pCFL = pVal
                    If Not pCFL.SelectedObjects Is Nothing Then
                        'If POItem <> "Y" Then
                        Try
                            odbdsDetails2.SetValue("U_Itemcode", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value)
                            odbdsDetails2.SetValue("U_ItemDesc", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value)
                            odbdsDetails2.SetValue("U_InvUom", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value)


                            'Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                            'Matrix3.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                            'Matrix3.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                        Catch ex As Exception
                            'Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                            'Matrix3.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                            'Matrix3.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                        End Try

                        odbdsDetails2.SetValue("U_Date", pVal.Row - 1, Now.Date.ToString("yyyyMMdd"))
                        odbdsDetails2.SetValue("U_Status", pVal.Row - 1, "O")

                        'Matrix3.Columns.Item("Date").Cells.Item(pVal.Row).Specific.String = Now.Date.ToString("dd/MM/yy")
                        'Matrix3.Columns.Item("Colstat").Cells.Item(pVal.Row).Specific.String = "O"
                        If AutoAssignItem = "Y" Then
                            objaddon.objglobalmethods.GetSubItemQty(objform, Matrix3)
                            'If VendorSPlPrice = "2" Then
                            '    CalcGRPrice()
                            'End If
                        Else
                            odbdsDetails2.SetValue("U_Qty", pVal.Row - 1, CDbl(1))
                            'Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = "1"
                        End If
                        'If Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String <> "" Then
                        '    'If ToLocationInOutput = "Y" Then
                        '    '    If objaddon.HANA Then
                        '    '        Matrix3.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("Select ""DflWhs"" from OBPL where ""BPLId""=(select  distinct T0.""BPLid"" from OWHS T0 left join OLCT T1 on T0.""Location""=T1.""Code"" where T1.""Code""=(Select ""U_LCodeO"" from ""@MIPL_GEN"") and ifnull(T0.""Inactive"",'') ='N' and T0.""BPLid"" is not null)")
                        '    '    Else
                        '    '        Matrix3.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("Select DflWhs from OBPL where BPLId=(select  distinct T0.BPLid from OWHS T0 left join OLCT T1 on T0.Location=T1.Code where T1.Code=(Select U_LCodeO from [@MIPL_GEN]) and isnull(T0.Inactive,'') ='N' and T0.BPLid is not null)")
                        '    '    End If
                        '    'End If
                        'End If
                        'If objaddon.HANA Then
                        '    HSNCode = objaddon.objglobalmethods.getSingleValue("select ""ChapterID"" from OCHP where ""AbsEntry""='" & pCFL.SelectedObjects.Columns.Item("ChapterID").Cells.Item(0).Value & "'")
                        'Else
                        '    HSNCode = objaddon.objglobalmethods.getSingleValue("select ChapterID from OCHP where AbsEntry='" & pCFL.SelectedObjects.Columns.Item("ChapterID").Cells.Item(0).Value & "'")
                        'End If
                        'If HSNCode <> "" Then
                        '    Matrix3.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = HSNCode
                        'Else
                        '    Matrix3.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = 0
                        'End If
                    End If
                    odbdsDetails2.SetValue("U_TabType", pVal.Row - 1, "Output")
                    'Matrix3.Columns.Item("TabType").Cells.Item(pVal.Row).Specific.String = "Output"
                    'objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                    Matrix3.LoadFromDataSource()
                Catch ex As Exception
                End Try
            ElseIf pVal.ColUID = "Whse" And pVal.ActionSuccess = True Then
                Dim OnHand As Double
                Try
                    Matrix3.FlushToDataSource()
                    If pVal.ActionSuccess = False Then Exit Sub
                    pCFL = pVal
                    If Not pCFL.SelectedObjects Is Nothing Then
                        Try
                            odbdsDetails2.SetValue("U_WhsCode", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value)
                            'Matrix3.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                        Catch ex As Exception
                            'Matrix3.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                        End Try
                    End If
                    If odbdsDetails2.GetValue("U_WhsCode", pVal.Row - 1) <> "" Then 'Matrix3.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String <> ""
                        If objaddon.HANA Then
                            OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and ""WhsCode""='" & Matrix3.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String & "'")
                        Else
                            OnHand = objaddon.objglobalmethods.getSingleValue("select OnHand from  oitw where ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and WhsCode='" & Matrix3.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String & "'")
                        End If
                        odbdsDetails2.SetValue("U_InStock", pVal.Row - 1, OnHand)
                        'Matrix3.Columns.Item("Instock").Cells.Item(pVal.Row).Specific.String = OnHand
                    Else
                        odbdsDetails2.SetValue("U_InStock", pVal.Row - 1, 0)
                        'Matrix3.Columns.Item("Instock").Cells.Item(pVal.Row).Specific.String = 0
                    End If
                    Matrix3.LoadFromDataSource()
                Catch ex As Exception
                End Try

            End If
        End Sub

        Private Sub ButtonCombo0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles ButtonCombo0.ClickBefore
            Try
                Dim StrSql As String
                Dim ProcessedQty As Double
                'Dim OpenQty() As Double
                Dim objrecset As SAPbobsCOM.Recordset
                If ComboBox2.Selected.Value = "C" Then
                    objaddon.objapplication.StatusBar.SetText("Document status is closed.You cannot perform this action!!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ButtonCombo0.Item.Enabled = False
                    Exit Sub
                End If
                objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    'If objaddon.HANA Then
                    '    StrSql = "select T0.""DocEntry"", from OWTR T0 left join ""@MIPL_OPOR"" T1 on T0.""U_SubConNo""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "';"
                    '    'StrSql = "select T0.""DocEntry"",T1.""ItemCode"",T1.""Quantity"",T2.""U_Qty"",(Select Sum(""U_Qty"") from ""@MIPL_POR2"" where ""DocEntry""=T2.""DocEntry"" and ""U_Qty"">0) ""OpenRecQty"" from OWTR T0 left join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join ""@MIPL_POR1"" T2 on T2.""U_Itemcode""=T1.""ItemCode"" and T2.""DocEntry""=T0.""U_SubConNo"" "
                    '    'StrSql += vbCrLf + "where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "'"
                    '    'GetVal = objaddon.objglobalmethods.getSingleValue("Select Sum(""U_Qty"") from ""@MIPL_POR2"" where ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'")
                    'Else
                    '    StrSql = "select T0.DocEntry from OWTR T0 left join [@MIPL_OPOR] T1 on T0.U_SubConNo=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "'"
                    '    'StrSql = "select T0.DocEntry,T1.ItemCode,T1.Quantity,T2.U_Qty,(Select Sum(U_Qty) from [@MIPL_POR2] where DocEntry=T2.DocEntry and U_Qty>0) OpenRecQty from OWTR T0 left join WTR1 T1 on T0.DocEntry=T1.DocEntry left join [@MIPL_POR1] T2 on T2.U_Itemcode=T1.ItemCode and T2.DocEntry=T0.U_SubConNo "
                    '    'StrSql += vbCrLf + "where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "'"
                    '    'GetVal = objaddon.objglobalmethods.getSingleValue("Select Sum(U_Qty) from [@MIPL_POR2] where DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'")
                    'End If
                    If objaddon.HANA Then
                        StrSql = "select T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"",T1.""Quantity"" from OWTR T0 left join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join  ""@MIPL_OPOR"" T2 on (T2.""DocEntry""=T0.""U_SubConNo"" or T2.""DocEntry""=T1.""U_S_DocEntry"") where T2.""DocEntry""='" & EditText26.Value & "' "
                    Else
                        StrSql = "select T0.DocEntry,T0.DocNum,T0.DocDate,T1.Quantity from OWTR T0 left join WTR1 T1 on T0.DocEntry=T1.DocEntry left join  [@MIPL_OPOR] T2 on (T2.DocEntry=T0.U_SubConNo or T2.DocEntry=T1.U_S_DocEntry) where T2.DocEntry='" & EditText26.Value & "' "
                    End If
                    objrecset.DoQuery(StrSql)
                    If objrecset.RecordCount = 0 Then
                        objaddon.objapplication.StatusBar.SetText("Inventory Transfer not created for the Planned Quantity...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    Else
                        For i As Integer = 1 To Matrix0.VisualRowCount
                            If Matrix0.Columns.Item("Code").Cells.Item(i).Specific.string <> "" Then
                                If objaddon.HANA Then
                                    StrSql = "select T2.""DocEntry"",T1.""ItemCode"",sum(T1.""Quantity"") ""Quantity"",T2.""U_Qty"",(Select Sum(""U_Qty"") from ""@MIPL_POR2"" where ""DocEntry""=T2.""DocEntry"" and ""U_Qty"">0) ""OpenRecQty"" from OWTR T0 left join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join ""@MIPL_POR1"" T2 on T2.""U_Itemcode""=T1.""ItemCode"" and (T2.""DocEntry""=T0.""U_SubConNo"" or T1.""U_S_DocEntry""=T2.""DocEntry"") "
                                    StrSql += vbCrLf + "where T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T1.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.string & "' group by T2.""DocEntry"",T1.""ItemCode"",T2.""U_Qty"""

                                Else
                                    StrSql = "select T2.DocEntry,T1.ItemCode,Sum(T1.Quantity) Quantity,T2.U_Qty,(Select Sum(U_Qty) from [@MIPL_POR2] where DocEntry=T2.DocEntry and U_Qty>0) OpenRecQty from OWTR T0 left join WTR1 T1 on T0.DocEntry=T1.DocEntry left join [@MIPL_POR1] T2 on T2.U_Itemcode=T1.ItemCode and (T2.DocEntry=T0.U_SubConNo or T1.U_S_DocEntry=T2.DocEntry) "
                                    StrSql += vbCrLf + "where T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T1.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.string & "' group by T2.DocEntry,T1.ItemCode,T2.U_Qty"

                                End If
                                objrecset.DoQuery(StrSql)
                                If objrecset.RecordCount > 0 Then
                                    If Math.Round((objrecset.Fields.Item("OpenRecQty").Value * objrecset.Fields.Item("U_Qty").Value), 6) > Math.Round(objrecset.Fields.Item("Quantity").Value, 6) Then 'Math.Round((OpenRecQty * Qty), 6)
                                        objaddon.objapplication.SetStatusBarMessage("Issue Qty should not be greater than the transferred Qty - Check the Input Item : " & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.string, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                Else
                                    objaddon.objapplication.SetStatusBarMessage("Inventory not transferred - Check the Input Item : " & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.string, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                        Next

                    End If
                    For i As Integer = 1 To Matrix3.VisualRowCount
                        If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string <> "" Then
                            If Matrix3.Columns.Item("Whse").Cells.Item(i).Specific.string = "" Or Val(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.string) = 0 Or Matrix3.Columns.Item("Price").Cells.Item(i).Specific.string = "" Then
                                objaddon.objapplication.SetStatusBarMessage("One of the column value is missing.Please update", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next
                    'OpenQty = GetOpen_VendorQty(objform.Items.Item("txtentry").Specific.String)
                    'If OpenQty(0) <= 0 Then
                    '    objaddon.objapplication.StatusBar.SetText("No more pending transactions...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    ButtonCombo1.Caption = "Generate Document"
                    '    ButtonCombo0.Item.Enabled = False
                    '    BubbleEvent = False : Exit Sub
                    'End If
                    If ProcessValidation = "Y" Then
                        If ItemProcess = "Y" Then
                            ProcessedQty = GetProcessQty()
                            If Matrix3.Columns.Item("Quant").Cells.Item(Matrix3.VisualRowCount).Specific.string > ProcessedQty Then
                                objaddon.objapplication.StatusBar.SetText("Seems Line Quantity exceeds. " & ProcessedQty & " qty will be allowed in this process.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    End If

                    If Folder1.Selected = True Then
                        For i As Integer = 1 To Matrix3.VisualRowCount
                            If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string <> "" Then
                                If UDF0 = "Y" And (ValUDF0 = "Y" Or ValUDF0 = "2") Then
                                    If Matrix3.Columns.Item("DocDate").Cells.Item(i).Specific.string = "" Then
                                        objaddon.objapplication.SetStatusBarMessage("Output Tab-> Document Date missing. Please update on line : " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                                If UDF1 = "Y" And (ValUDF1 = "Y" Or ValUDF1 = "2") Then
                                    If Matrix3.Columns.Item("RefNo").Cells.Item(i).Specific.string = "" Then
                                        objaddon.objapplication.SetStatusBarMessage("Output Tab-> Reference No missing. Please update on line : " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If

                    'ValOpenQty = OpenQty(0)
                    'For i As Integer = 1 To Matrix3.VisualRowCount
                    '    If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String) <> 0.0 And Matrix3.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" Then
                    '        ValOpenQty = CDbl(OpenQty(0) + CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)) 'CDbl(objform.Items.Item("SQty").Specific.String) - (OpenQty(0) + CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String))
                    '        If ValOpenQty < 0 Then
                    '            'RemoveLastrow(Matrix3, "GRNo")
                    '            objaddon.objapplication.StatusBar.SetText("Line Quantity should not exceed Sub Item Quantity...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            'BubbleEvent = False : Exit Sub
                    '        End If
                    '    End If
                    'Next
                    'Dim CheckIGIStat As Boolean = False
                    ''StrSql = "Call ""MIPL_GetPendQty_Inv_GI_GR"" ('OIGE','IGE1'," & objform.Items.Item("txtentry").Specific.String & ")"
                    'If objaddon.HANA Then
                    '    StrSql = "Select Case when Sum(T2.""Quantity"") >0 then T0.""U_PlanQty""-Sum(T2.""Quantity"") else T0.""U_PlanQty"" end as ""PendQty"",T0.""U_Itemcode"""
                    '    StrSql += vbCrLf + " from ""@MIPL_POR1"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" left join OIGE T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGE1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T0.""U_Itemcode""=T2.""ItemCode"""
                    '    StrSql += vbCrLf + "where T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " group by T0.""U_PlanQty"",T0.""U_Itemcode"" order by ""PendQty"""
                    'Else
                    '    StrSql = "Select Case when Sum(T2.Quantity) >0 then T0.U_PlanQty-Sum(T2.Quantity) else T0.U_PlanQty end as PendQty,T0.U_Itemcode"
                    '    StrSql += vbCrLf + "from [@MIPL_POR1] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry left join OIGE T3 on T3.U_SubConNo=T1.DocEntry left join IGE1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode"
                    '    StrSql += vbCrLf + "where T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " group by T0.U_PlanQty,T0.U_Itemcode order by PendQty"
                    'End If
                    'objrs.DoQuery(StrSql)
                    'For Result As Integer = 0 To objrs.RecordCount - 1
                    '    If CDbl(objrs.Fields.Item("PendQty").Value) <= 0.0 Then
                    '        CheckIGIStat = True
                    '    Else
                    '        CheckIGIStat = False
                    '    End If
                    '    objrs.MoveNext()
                    'Next
                    'Dim ChkGRStat As String
                    ''ChkGRStat = objaddon.objglobalmethods.getSingleValue("Call ""MIPL_ValidateGRQty"" ('@MIPL_POR2'," & objform.Items.Item("txtentry").Specific.String & ",'" & EditText8.Value & "')")
                    'If objaddon.HANA Then

                    '    ChkGRStat = objaddon.objglobalmethods.getSingleValue("Select  Distinct T1.""U_SQty""-T4.""Qty"" as ""PendQty""from  ""@MIPL_OPOR"" T1 inner join ""@MIPL_POR2"" T0 on T0.""DocEntry""=T1.""DocEntry"" and T0.""U_Itemcode""=T1.""U_SItemCode"" inner join " &
                    '                                                      "(select T2.""U_SubConNo"",sum(T3.""Quantity"") as ""Qty"",T3.""ItemCode"" from  OIGN T2 join IGN1 T3 on T2.""DocEntry""=T3.""DocEntry"" group by T2.""U_SubConNo"",T3.""ItemCode"" ) as T4 " &
                    '                                                     "  on T4.""U_SubConNo""=T1.""DocEntry"" and T0.""U_Itemcode""=T4.""ItemCode"" where T1.""U_SItemCode""='" & EditText8.Value & "'  and T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " ")

                    'Else
                    '    ChkGRStat = objaddon.objglobalmethods.getSingleValue("Select  Distinct T1.U_SQty-T4.Qty as PendQty from  [@MIPL_OPOR] T1 inner join [@MIPL_POR2] T0 on T0.DocEntry=T1.DocEntry and T0.U_Itemcode=T1.U_SItemCode inner join " &
                    '                                                        "(select T2.U_SubConNo,sum(T3.Quantity) as Qty,T3.ItemCode from  OIGN T2 join IGN1 T3 on T2.DocEntry=T3.DocEntry group by T2.U_SubConNo,T3.ItemCode ) as T4 " &
                    '                                                       "  on T4.U_SubConNo=T1.DocEntry and T0.U_Itemcode=T4.ItemCode where T1.U_SItemCode='" & EditText8.Value & "'  and T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " ")
                    'End If
                    'If ChkGRStat = "0" And CheckIGIStat = True Then
                    '    objaddon.objapplication.StatusBar.SetText("Goods Issue & Goods Receipt Created for the Planned Quantity...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    ButtonCombo0.Item.Enabled = False
                    '    'RemoveLastrow(Matrix3, "Code")
                    '    'For i As Integer = 1 To Matrix3.VisualRowCount
                    '    '    If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix3.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" Then
                    '    '        'RemoveLastrow(Matrix3, "Code")
                    '    '    End If
                    '    'Next
                    '    BubbleEvent = False : Exit Sub
                    'End If
                    'If CheckIGIStat = True Then
                    '    objaddon.objapplication.StatusBar.SetText("Goods Issue Created for the Planned Quantity...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    ButtonCombo1.Caption = "Generate Document"
                    '    'BubbleEvent = False : Exit Sub
                    'End If
                    RemoveLastrow(Matrix3, "Code")
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objaddon.objapplication.StatusBar.SetText("Please update the Sub-Contracting Screen...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                    'If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    '    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button8_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button8.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                Auto_InventoryTransfer(Matrix3, "U_OutNum")
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix3_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix3.ChooseFromListBefore
            Try  'Output
                If pVal.ActionSuccess = True Then Exit Sub
                If pVal.ColUID = "Code" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_IOut")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()
                        Dim Qty As Double = 0
                        Dim Row As Integer = 0
                        oCond = oConds.Add()
                        oCond.Alias = "InvntItem"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "Y"
                        If AutoAssignItem = "Y" Then
                            For i As Integer = 1 To Matrix3.VisualRowCount
                                If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String Then
                                    Qty += CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.string)
                                    Row = i
                                End If
                            Next
                            If Qty >= CDbl(objform.Items.Item("SQty").Specific.String) And pVal.Row <> Row Then
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                                oCond = oConds.Add()
                                oCond.Alias = "ItemCode"
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                                oCond.CondVal = objform.Items.Item("txtsitem").Specific.String
                            End If
                        End If
                        oCFL.SetConditions(oConds)

                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "Whse" Then
                    Try
                        Dim LocCode As String = ""
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_W")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()
                        If ToLocationInOutput = "Y" Then
                            If objaddon.HANA Then
                                LocCode = objaddon.objglobalmethods.getSingleValue("select ""U_LCodeO"" from ""@MIPL_GEN""")
                            Else
                                LocCode = objaddon.objglobalmethods.getSingleValue("select U_LCodeO from [@MIPL_GEN]")
                            End If
                            Dim SplitLoc() As String = Split(LocCode, ",")
                            For Val As Integer = 0 To SplitLoc.Length - 1
                                If Val = 0 Then
                                    oCond = oConds.Add()
                                    oCond.Alias = "Location"
                                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                    oCond.CondVal = SplitLoc(Val)
                                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                                    oCond = oConds.Add()
                                    oCond.Alias = "Inactive"
                                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                                    oCond.CondVal = "Y"
                                Else
                                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                    oCond = oConds.Add()
                                    oCond.Alias = "Location"
                                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                    oCond.CondVal = SplitLoc(Val)
                                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                                    oCond = oConds.Add()
                                    oCond.Alias = "Inactive"
                                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                                    oCond.CondVal = "Y"
                                End If
                            Next
                            'oCond = oConds.Add()
                            'oCond.Alias = "Location"
                            'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            'oCond.CondVal = LocCode
                        End If
                        ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try

                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub Matrix3_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix3.KeyDownAfter
            Try
                If pVal.ActionSuccess = True Then
                    Select Case pVal.ColUID
                        Case "Quant"
                            If CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) < 0 Then
                                objaddon.objapplication.StatusBar.SetText("In ""Qty"" (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Click()
                                Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = "1"
                            End If
                    End Select
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix3_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix3.ValidateBefore
            Try
                If pVal.InnerEvent = True Then Exit Sub
                If pVal.ItemChanged = False Then Exit Sub
                Dim Qty As Double = 0, InvQty As Double = 0, ExcessQty As Double = 0
                Dim objRs As SAPbobsCOM.Recordset
                'If pVal.ActionSuccess = False Then Exit Sub
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If AutoAssignItem = "Y" Then
                    For i As Integer = 1 To Matrix3.VisualRowCount
                        If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String Then
                            Qty += CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.string)
                        End If
                    Next
                    If objform.Items.Item("txtentry").Specific.String <> "" Then
                        'If objaddon.HANA Then
                        '    objRs.DoQuery("select Top 1 T1.""U_Itemcode"",sum(T1.""U_Qty"") ""Qty"" from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                        '      "where T1.""U_Type""='2' and T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T1.""U_Status""='C' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' " &
                        '      " group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc")
                        'Else
                        '    objRs.DoQuery("select Top 1 T1.U_Itemcode,sum(T1.U_Qty) Qty from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry " &
                        '      "where T1.U_Type='2' and T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T1.U_Status='C' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' " &
                        '      " group by T1.U_Itemcode order by sum(T1.U_Qty) desc")
                        'End If
                        If objaddon.HANA Then
                            'strQuery = "SELECT IFNULL(CASE WHEN SUM(T2.""Quantity"") > 0 THEN (T0.""U_PlanQty"") - (SUM(T2.""Quantity"") * T0.""U_BOMQty"") ELSE T0.""U_PlanQty"" END"
                            'strQuery += vbCrLf + "- IFNULL((SELECT TOP 1 SUM(B.""U_Qty"") FROM ""@MIPL_OPOR"" A INNER JOIN ""@MIPL_POR3"" B ON A.""DocEntry"" = B.""DocEntry"" INNER JOIN ""@MIPL_POR1"" C ON B.""DocEntry"" = C.""DocEntry"""
                            'strQuery += vbCrLf + "WHERE B.""U_Type"" = '2' AND B.""U_Status"" = 'C' AND A.""DocEntry"" = T0.""DocEntry"" AND IFNULL(B.""U_InvNo"", '') <> '' AND IFNULL(B.""U_GRNo"", '') = '' GROUP BY B.""U_Itemcode"", C.""U_BOMQty"""
                            'strQuery += vbCrLf + "ORDER BY SUM(B.""U_Qty"") DESC), 0), 0) / T0.""U_BOMQty"" AS ""OpenQty"" FROM ""@MIPL_OPOR"" T1 LEFT OUTER JOIN ""@MIPL_POR1"" T0 ON T1.""DocEntry"" = T0.""DocEntry"""
                            'strQuery += vbCrLf + "LEFT OUTER JOIN OIGN T3 ON T3.""U_SubConNo"" = T1.""DocEntry"" LEFT OUTER JOIN IGN1 T2 ON T3.""DocEntry"" = T2.""DocEntry"" AND T1.""U_SItemCode"" = T2.""ItemCode"""
                            'strQuery += vbCrLf + "WHERE T1.""DocEntry"" = '" & objform.Items.Item("txtentry").Specific.String & "' GROUP BY T0.""DocEntry"", T0.""U_PlanQty"", T0.""U_BOMQty"""

                            'strQuery = "Select MIN(((Case when A.""ReceiptQty"">0 then A.""U_PlanQty""-(A.""ReceiptQty""*A.""U_Qty"") Else A.""U_PlanQty"" End)-A.""ReturnQty"")/A.""U_Qty"") as ""OpenQty"",A.* "
                            'strQuery += vbCrLf + "from (Select T1.""DocEntry"",T0.""U_PlanQty"",T0.""U_Itemcode"",T0.""U_Qty"","
                            'strQuery += vbCrLf + "(Select ifnull((select top 1 sum(B.""U_Qty"") from ""@MIPL_OPOR"" A join ""@MIPL_POR3"" B on A.""DocEntry""=B.""DocEntry"" where B.""U_Type""='2' and B.""U_Status""='C' and A.""DocEntry""=T1.""DocEntry"" and B.""U_Itemcode""=T0.""U_Itemcode"""
                            'strQuery += vbCrLf + "and ifnull(B.""U_InvNo"",'') <>'' and ifnull(B.""U_GRNo"",'') ='' group by B.""U_Itemcode"" order by sum(B.""U_Qty"") ),0)) as ""ReturnQty"","
                            'strQuery += vbCrLf + "(Select Sum(T2.""Quantity"") from OIGN T3 join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" where T3.""U_SubConNo""=T1.""DocEntry"" and T1.""U_SItemCode""=T2.""ItemCode"") as ""ReceiptQty"""
                            'strQuery += vbCrLf + "from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"") A where A.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by A.""DocEntry"",A.""U_PlanQty"",A.""U_Itemcode"",A.""U_Qty"",A.""ReturnQty"",A.""ReceiptQty"""
                            'strQuery += vbCrLf + "order by MIN(((Case when A.""ReceiptQty"">0 then A.""U_PlanQty""-(A.""ReceiptQty""*A.""U_Qty"") Else A.""U_PlanQty"" End)-A.""ReturnQty"")/A.""U_Qty"")"
                            'objRs.DoQuery(strQuery)

                            strQuery = "Select MIN(((Case when A.""ReceiptQty"">0 then A.""U_PlanQty""-(A.""ReceiptQty""*A.""U_Qty"") Else A.""U_PlanQty"" End)-A.""ReturnQty"")/A.""U_Qty"") as ""OpenQty"",A.* "
                            strQuery += vbCrLf + "from (Select T1.""DocEntry"",T0.""U_PlanQty"",T0.""U_Itemcode"",T0.""U_Qty"","
                            strQuery += vbCrLf + "(Select ifnull((select top 1 sum(B.""U_Qty"") from ""@MIPL_OPOR"" A join ""@MIPL_POR3"" B on A.""DocEntry""=B.""DocEntry"" where B.""U_Type""='2' and B.""U_Status""='C' and A.""DocEntry""=T1.""DocEntry"" and B.""U_Itemcode""=T0.""U_Itemcode"""
                            strQuery += vbCrLf + "and ifnull(B.""U_InvNo"",'') <>'' and ifnull(B.""U_GRNo"",'') ='' group by B.""U_Itemcode"" order by sum(B.""U_Qty"") ),0)) as ""ReturnQty"","
                            strQuery += vbCrLf + "(Select Sum(T2.""Quantity"") from OIGN T3 join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" where T3.""U_SubConNo""=T1.""DocEntry"" and T1.""U_SItemCode""=T2.""ItemCode"") as ""ReceiptQty"""
                            strQuery += vbCrLf + "from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"") A where A.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by A.""DocEntry"",A.""U_PlanQty"",A.""U_Itemcode"",A.""U_Qty"",A.""ReturnQty"",A.""ReceiptQty"""
                            strQuery += vbCrLf + "order by MIN(((Case when A.""ReceiptQty"">0 then A.""U_PlanQty""-(A.""ReceiptQty""*A.""U_Qty"") Else A.""U_PlanQty"" End)-A.""ReturnQty"")/A.""U_Qty"")"
                            objRs.DoQuery(strQuery)
                        Else
                            'strQuery = "Select isnull(case when Sum(T2.Quantity)> 0 Then (T0.U_PlanQty)-(Sum(T2.Quantity)*T0.U_BOMQty) Else T0.U_PlanQty End"
                            'strQuery += vbCrLf + "- isnull((select Top 1 sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B on A.DocEntry=B.DocEntry join [@MIPL_POR1] C on B.DocEntry=C.DocEntry where B.U_Type='2' and B.U_Status='C' and A.DocEntry=T0.DocEntry"
                            'strQuery += vbCrLf + "and isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' group by B.U_Itemcode,C.U_BOMQty order by sum(B.U_Qty) desc),0),0)/T0.U_BOMQty  as OpenQty"
                            'strQuery += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry And T1.U_SItemCode=T2.ItemCode"
                            'strQuery += vbCrLf + "where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T0.DocEntry,T0.U_PlanQty,T0.U_BOMQty"

                            'strQuery = "Select ((Case when A.[Receipt Qty]>0 then A.[Header Qty]-A.[Receipt Qty] Else A.[Header Qty] End)-A.[Return Qty]) As OpenQty,A.* "
                            'strQuery += vbCrLf + "from (Select T1.DocEntry,T0.U_PlanQty [Line Qty],T0.U_Itemcode,T0.U_Qty [Base Qty],T1.U_SQty [Header Qty] ,T1.U_InvUom,T1.U_SItemCode,"
                            'strQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A inner join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T0.DocEntry and B.ItemCode=T0.U_Itemcode) as [Return Qty],"
                            'strQuery += vbCrLf + "(Select isnull(Sum(T2.Quantity),0) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T1.U_SItemCode=T2.ItemCode) as [Receipt Qty]"
                            'strQuery += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry) A where A.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'"
                            'strQuery += vbCrLf + "group by A.DocEntry,A.[Line Qty],A.U_Itemcode,A.[Base Qty],A.[Return Qty],A.[Header Qty], A.[Receipt Qty],A.U_SItemCode,A.U_InvUom"
                            'objRs.DoQuery(strQuery)

                            strQuery = "Select case when isnull((select Top 1 U_Roundoff from [@MI_ITMRND] where U_InvUOM=(SELECT InvntryUom FROM OITM where ItemCode=A.U_SItemCode and InvntryUom <>'') order by Code desc),'') ='Y' then "
                            strQuery += vbCrLf + "cast(MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty) as decimal(18,0)) Else MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty) End As OpenQty,A.* "
                            strQuery += vbCrLf + "from (Select T1.DocEntry,T0.U_PlanQty,T0.U_Itemcode,T0.U_Qty,T1.U_InvUom,T1.U_SItemCode,"
                            strQuery += vbCrLf + "(Select isnull((Select top 1 sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B On A.DocEntry= B.DocEntry where B.U_Type='2' and B.U_Status='C' and A.DocEntry=T1.DocEntry and B.U_Itemcode=T0.U_Itemcode"
                            strQuery += vbCrLf + "and isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' group by B.U_Itemcode order by sum(B.U_Qty) ),0)) as ReturnQty,"
                            strQuery += vbCrLf + "(Select Sum(T2.Quantity) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T1.U_SItemCode=T2.ItemCode) as ReceiptQty"
                            strQuery += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry) A where A.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by A.DocEntry,A.U_PlanQty,A.U_Itemcode,A.U_Qty,A.ReturnQty,A.ReceiptQty,A.U_SItemCode,A.U_InvUom"
                            strQuery += vbCrLf + "order by MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty)"
                            objRs.DoQuery(strQuery)
                        End If
                        'InvQty = CDbl(objRs.Fields.Item("Qty").Value.ToString)
                        Qty = CDbl(objRs.Fields.Item(0).Value.ToString)
                    Else
                        Qty = CDbl(objform.Items.Item("SQty").Specific.String)
                    End If

                    If ProcessValidation = "Y" Then
                        If ItemProcess = "Y" Then
                            If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                Dim SQuery As String, Getseq As String
                                If objform.Items.Item("txtponum").Specific.String <> "" And Trim(objform.Items.Item("tprocseq").Specific.String) <> "" Then
                                    If objaddon.HANA Then
                                        Getseq = objaddon.objglobalmethods.getSingleValue("Select Top 1 T0.""U_Sequence"" as ""Seq"" from ""@MIPL_BOM2"" T0 where T0.""Code""='" & Trim(objform.Items.Item("txtbitem").Specific.String) & "' and T0.""U_Priority""='2' and T0.""U_Sequence""<'" & Trim(objform.Items.Item("tprocseq").Specific.String) & "' order by T0.""U_Sequence"" desc")

                                        SQuery = "Select Case when A.""BOMPrior""='2' and A.""ReceiptQty"" >0 then  A.""ReceiptQty"" Else A.""SubItemQty""-A.""ReceiptQty"" End as ""Qty"",* from(Select (Select ifnull(Sum(T2.""Quantity""),0) from OIGN T3 left join IGN1 T2  On T3.""DocEntry""= T2.""DocEntry"""
                                        SQuery += vbCrLf + "And T1.""U_SItemCode""=T2.""ItemCode"" And T3.""U_SubConNo""=T1.""DocEntry"" "
                                        If Getseq = "" Then
                                            SQuery += vbCrLf + "where T2.""U_Process"" ='" & Trim(objform.Items.Item("tprocseq").Specific.String) & "')"
                                        Else
                                            SQuery += vbCrLf + "where T2.""U_Process"" ='" & Getseq & "')"
                                        End If
                                        SQuery += vbCrLf + "as ""ReceiptQty"",T1.""U_SQty"" as ""SubItemQty"",(Select ""U_Priority"" from ""@MIPL_BOM2"" where ""Code""=T1.""U_BOMCode"""
                                        If Getseq = "" Then
                                            SQuery += vbCrLf + "and ""U_Sequence"" ='" & Getseq & "')"
                                        Else
                                            SQuery += vbCrLf + "and ""U_Sequence"" ='" & Trim(objform.Items.Item("tprocseq").Specific.String) & "')"
                                        End If
                                        SQuery += vbCrLf + " as ""BOMPrior"" from ""@MIPL_OPOR"" T1 where T1.""U_PONum""='" & objform.Items.Item("txtponum").Specific.String & "' "
                                        If Getseq = "" Then
                                            SQuery += vbCrLf + "And T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'"
                                        Else
                                            SQuery += vbCrLf + "And T1.""DocEntry""=(Select ""DocEntry"" from ""@MIPL_OPOR"" where ""U_PONum""='" & objform.Items.Item("txtponum").Specific.String & "' and ""U_Procseq""='" & Getseq & "')"
                                        End If
                                        SQuery += vbCrLf + ") A"
                                    Else
                                        Getseq = objaddon.objglobalmethods.getSingleValue("Select Top 1 T0.U_Sequence as Seq from [@MIPL_BOM2] T0 where T0.Code='" & Trim(objform.Items.Item("txtbitem").Specific.String) & "' and T0.U_Priority='2' and T0.U_Sequence<'" & Trim(objform.Items.Item("tprocseq").Specific.String) & "' order by T0.U_Sequence desc")

                                        SQuery = "Select Case when A.BOMPrior='2' and A.ReceiptQty >0 then  A.ReceiptQty Else A.SubItemQty-A.ReceiptQty End as Qty,* from(Select (Select isnull(Sum(T2.Quantity),0) from OIGN T3 left join IGN1 T2  On T3.DocEntry= T2.DocEntry"
                                        SQuery += vbCrLf + "And T1.U_SItemCode=T2.ItemCode And T3.U_SubConNo=T1.DocEntry "
                                        If Getseq = "" Then
                                            SQuery += vbCrLf + "where T2.U_Process ='" & Trim(objform.Items.Item("tprocseq").Specific.String) & "')"
                                        Else
                                            SQuery += vbCrLf + "where T2.U_Process ='" & Getseq & "')"
                                        End If
                                        SQuery += vbCrLf + "as ReceiptQty,T1.U_SQty as SubItemQty,(Select U_Priority from [@MIPL_BOM2] where Code=T1.U_BOMCode"
                                        If Getseq = "" Then
                                            SQuery += vbCrLf + "and U_Sequence ='" & Getseq & "')"
                                        Else
                                            SQuery += vbCrLf + "and U_Sequence ='" & Getseq & "')"
                                            'SQuery += vbCrLf + "and U_Sequence ='" & Trim(objform.Items.Item("tprocseq").Specific.String) & "')"
                                        End If
                                        SQuery += vbCrLf + " as BOMPrior from [@MIPL_OPOR] T1 where T1.U_PONum='" & objform.Items.Item("txtponum").Specific.String & "' "
                                        If Getseq = "" Then
                                            SQuery += vbCrLf + "And T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'"
                                        Else
                                            SQuery += vbCrLf + "And T1.DocEntry=(Select DocEntry from [@MIPL_OPOR] where U_PONum='" & objform.Items.Item("txtponum").Specific.String & "' and U_Procseq='" & Getseq & "')"
                                        End If
                                        SQuery += vbCrLf + ") A"
                                        'Select Case when A.BOMPrior='2' and A.ReceiptQty >0 then  A.ReceiptQty Else A.SubItemQty-A.ReceiptQty End as Qty,* 
                                        'from(Select (Select Sum(T2.Quantity) from OIGN T3 left join IGN1 T2  On T3.DocEntry= T2.DocEntry
                                        'And T1.U_SItemCode=T2.ItemCode And T3.U_SubConNo=T1.DocEntry where T2.U_Process ='2') as ReceiptQty,T1.U_SQty as SubItemQty,
                                        '(Select U_Priority from [@MIPL_BOM2] where Code=T1.U_BOMCode And U_Sequence='2') as BOMPrior,
                                        '(Select U_Priority from [@MIPL_BOM2] where Code=T1.U_BOMCode And U_Sequence=(Select T2.U_Process from OIGN T3 left join IGN1 T2  on T3.DocEntry=T2.DocEntry  
                                        'And T1.U_SItemCode=T2.ItemCode And T3.U_SubConNo=T1.DocEntry where T2.U_Process ='2')) as _Prior
                                        'from [@MIPL_OPOR] T1 where T1.U_PONum='239' And T1.DocEntry=37) A


                                        'Getseq = objaddon.objglobalmethods.getSingleValue("Select T0.U_Sequence as Seq from [@MIPL_BOM2] T0 where T0.Code='" & Trim(objform.Items.Item("txtbitem").Specific.String) & "' and T0.U_Priority='2' and T0.U_Sequence<'" & Trim(objform.Items.Item("tprocseq").Specific.String) & "'")

                                        'SQuery = "Select Case when Sum(T2.Quantity) >0 then (Sum(T2.Quantity)+ "
                                        'SQuery += vbCrLf + "isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' "
                                        'SQuery += vbCrLf + " group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0)) else "
                                        'SQuery += vbCrLf + "isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' "
                                        'SQuery += vbCrLf + " group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0) end as OpenQty "
                                        'SQuery += vbCrLf + " from [@MIPL_OPOR] T1 left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry and T1.U_SItemCode=T2.ItemCode  "
                                        'SQuery += vbCrLf + " where T1.U_PONum='" & objform.Items.Item("txtponum").Specific.String & "'"
                                        'If Getseq <> "" Then SQuery += vbCrLf + " And T1.U_ProcSeq='" & Getseq & "'"
                                        'SQuery += vbCrLf + "  Group by T1.U_SQty "

                                    End If
                                    objRs.DoQuery(SQuery)
                                    If objRs.RecordCount > 0 Then
                                        If CDbl(objRs.Fields.Item(0).Value.ToString) < CDbl(Qty) Then
                                            If Getseq = "" Then
                                                ExcessQty = CDbl(CDbl(Qty + InvQty) - CDbl(objform.Items.Item("SQty").Specific.String))
                                                Matrix3.SetCellWithoutValidation(Matrix3.VisualRowCount, "Quant", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty)
                                            Else
                                                ExcessQty = CDbl(CDbl(Qty) - CDbl(objRs.Fields.Item(0).Value.ToString))
                                                Matrix3.SetCellWithoutValidation(Matrix3.VisualRowCount, "Quant", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty)
                                            End If
                                            'ExcessQty = CDbl(CDbl(Qty) - CDbl(objRs.Fields.Item(0).Value.ToString)) 'CDbl(objform.Items.Item("SQty").Specific.String))
                                            'Dim cc As Double = CDbl(objRs.Fields.Item("Qty").Value.ToString)
                                            'Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.string = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty ' CDbl(objRs.Fields.Item(0).Value.ToString)
                                            'Matrix3.SetCellWithoutValidation(Matrix3.VisualRowCount, "Quant", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty) 'ExcessQty 'Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String
                                        End If

                                    End If
                                End If
                                'If CDbl((objform.Items.Item("SQty").Specific.String) - CDbl(CDbl(Qty) + InvQty)) < 0 Then
                                '    ExcessQty = CDbl(CDbl(Qty + InvQty) - CDbl(objform.Items.Item("SQty").Specific.String))
                                '    'Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty
                                '    Matrix3.SetCellWithoutValidation(Matrix3.VisualRowCount, "Quant", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty)
                                'End If
                            End If
                        Else
                            Dim vv As Double = CDbl(objform.Items.Item("SQty").Specific.String)
                            If CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) > (Qty + InvQty) Then 'CDbl(objform.Items.Item("SQty").Specific.String) - (Qty + InvQty) < 0
                                ExcessQty = CDbl(CDbl(Qty + InvQty) - CDbl(objform.Items.Item("SQty").Specific.String))
                                'Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty
                                Matrix3.SetCellWithoutValidation(Matrix3.VisualRowCount, "Quant", (Qty + InvQty)) 'CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty
                            End If
                        End If
                    Else
                        'If CDbl((objform.Items.Item("SQty").Specific.String) - CDbl(CDbl(Qty) + InvQty)) < 0 Then
                        '    ExcessQty = CDbl(CDbl(Qty + InvQty) - CDbl(objform.Items.Item("SQty").Specific.String))
                        '    'Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty
                        '    Matrix3.SetCellWithoutValidation(Matrix3.VisualRowCount, "Quant", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty)
                        'End If
                        If CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) > Qty Then
                            ExcessQty = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - Qty
                            'Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty
                            Matrix3.SetCellWithoutValidation(Matrix3.VisualRowCount, "Quant", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ExcessQty)
                        End If
                    End If
                End If

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix3_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix3.ComboSelectAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                If WPrice = "Y" Then
                    If Matrix3.Columns.Item("Colstat").Cells.Item(pVal.Row).Specific.String = "C" Then Exit Sub
                    objcombo = Matrix3.Columns.Item("cmblttype").Cells.Item(pVal.Row).Specific
                    Select Case pVal.ColUID
                        Case "cmblttype"
                            If objcombo.Selected Is Nothing Then Exit Sub
                            If pVal.ItemChanged = False Then Exit Sub
                            If CDbl(Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) = 0 Then Exit Sub
                            If objcombo.Selected.Value = "2" Then
                                If Val(Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) > 0 Then
                                    'Matrix3.SetCellWithoutValidation(pVal.Row, "tweight", CDbl(Matrix3.Columns.Item("actwt").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String))
                                    Matrix3.SetCellWithoutValidation(pVal.Row, "TProCost", CDbl(Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String))

                                    'Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("actwt").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String)
                                    'Matrix3.Columns.Item("TProCost").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String)
                                Else
                                    Exit Sub
                                End If
                            Else
                                Matrix3.SetCellWithoutValidation(pVal.Row, "tweight", CDbl(Matrix3.Columns.Item("actwt").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String))
                                'Matrix3.Columns.Item("TProCost").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String)
                                Matrix3.SetCellWithoutValidation(pVal.Row, "TProCost", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String))
                            End If
                            objform.Update()
                            objform.Refresh()
                        Case Else
                            Exit Sub
                    End Select

                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Matrix3_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix3.ValidateAfter
            Try
                If pVal.Row = 0 Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                Select Case pVal.ColUID
                    Case "tweight", "cmblttype", "Quant"
                        If pVal.ItemChanged = False Then Exit Sub
                        objcombo = Matrix3.Columns.Item("cmblttype").Cells.Item(pVal.Row).Specific
                        If objcombo.Selected Is Nothing Then Exit Sub
                        If WPrice = "Y" Then
                            If Matrix3.Columns.Item("Colstat").Cells.Item(pVal.Row).Specific.String = "C" Then Exit Sub
                            If objcombo.Selected.Value = "2" Then
                                If Val(Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) > 0 Then
                                    'Matrix3.SetCellWithoutValidation(pVal.Row, "tweight", CDbl(Matrix3.Columns.Item("actwt").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String))
                                    Matrix3.SetCellWithoutValidation(pVal.Row, "TProCost", CDbl(Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String))

                                    'Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("actwt").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String)
                                    'Matrix3.Columns.Item("TProCost").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("tweight").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String)
                                Else
                                    Exit Sub
                                End If
                            Else
                                Matrix3.SetCellWithoutValidation(pVal.Row, "tweight", CDbl(Matrix3.Columns.Item("actwt").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String))
                                Matrix3.SetCellWithoutValidation(pVal.Row, "TProCost", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String))
                                ' Matrix3.Columns.Item("TProCost").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String)
                            End If
                        Else
                            Matrix3.SetCellWithoutValidation(pVal.Row, "TProCost", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("ProCost").Cells.Item(pVal.Row).Specific.String))

                        End If
                        'Case "Price"

                End Select
                Matrix3.SetCellWithoutValidation(pVal.Row, "Linetot", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix3.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String))
                objform.Update()
                objform.Refresh()
            Catch ex As Exception

            End Try

        End Sub


#End Region

#Region "Scrap Tab"

        Private Sub Matrix2_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix2.ClickAfter
            Try
                If pVal.Row = 0 Then
                    Matrix2.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                Else
                    Matrix2.SelectRow(pVal.Row, True, False)
                End If
                Try
                    If Matrix2.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.String = "" Then
                        Matrix2.CommonSetting.SetCellEditable(pVal.Row, 16, True)
                    ElseIf Matrix2.Columns.Item("RefNo").Cells.Item(pVal.Row).Specific.String = "" Then
                        Matrix2.CommonSetting.SetCellEditable(pVal.Row, 17, True)
                    End If
                Catch ex As Exception
                End Try

            Catch ex As Exception
            End Try
        End Sub

        Private Sub Matrix2_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix2.LostFocusAfter
            Try
                If pVal.Row = 0 Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                Dim Quantity As Double, Price As Double
                'If pVal.ActionSuccess = True Then
                'If Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String <> "" Then
                '    Quantity = CDbl(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String)
                'Else
                '    Quantity = 0
                'End If
                'If Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String <> "" Then
                '    Price = CDbl(Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                'Else
                '    Price = 0
                'End If
                Quantity = IIf(Val(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) > 0, CDbl(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String), 0)
                Price = IIf(Val(Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String) > 0, CDbl(Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String), 0)
                Select Case pVal.ColUID
                    Case "Code"
                        If Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String <> "" Then
                            If objaddon.HANA Then
                                Matrix2.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("select T1.""ChapterID"" from OITM T0 join OCHP T1 on T0.""ChapterID""=T1.""AbsEntry"" where T0.""ItemCode""='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                            Else
                                Matrix2.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("select T1.ChapterID from OITM T0 join OCHP T1 on T0.ChapterID=T1.AbsEntry where T0.ItemCode='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'")
                            End If
                            'objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                            ''If objaddon.HANA Then
                            ''    unitprice = objaddon.objglobalmethods.getSingleValue("select ""Price"" from  itm1 where ""ItemCode""='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and ""PriceList""=1")
                            ''Else
                            ''    unitprice = objaddon.objglobalmethods.getSingleValue("select Price from  itm1 where ItemCode='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and PriceList=1")
                            ''End If

                            ''If unitprice > 0 Then
                            ''    Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String = unitprice
                            ''Else
                            ''    Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String = 0
                            ''End If
                        End If
                    Case "Quant"
                        If Val(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) <> 0 Then
                            Matrix2.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = Quantity * Price
                        End If

                    Case "Price"
                        If Val(Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String) <> 0 Then
                            Matrix2.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = Quantity * Price
                        End If
                End Select
                objcombo = Matrix2.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                If objcombo.Selected.Description = "Scrap" Then
                    Matrix2.CommonSetting.SetCellEditable(pVal.Row, 9, True)
                ElseIf objcombo.Selected.Description = "Return" Then
                    ' Matrix2.CommonSetting.SetCellEditable(pVal.Row, 9, False)
                End If
                'objaddon.objapplication.Menus.Item("1300").Activate()
                'End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub ButtonCombo2_ComboSelectAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ButtonCombo2.ComboSelectAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If ButtonCombo2.Selected.Value = "Goods Receipt" Then
                        ButtonCombo0.Caption = "Generate Document"
                        If Matrix2.Columns.Item("Code").Cells.Item(Matrix2.VisualRowCount).Specific.string <> "" Then
                            CreateGoodsReceipt("Scrap")
                        Else
                            objaddon.objapplication.SetStatusBarMessage("Please update the Row level data...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            Exit Sub
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Folder2_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder2.PressedAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If ComboBox2.Selected.Value = "C" Then
                    'Button9.Item.Enabled = False
                    'Matrix2.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                Else
                    'If AutoAssignItem = "Y" Then
                    '    Matrix2.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_Default)
                    'End If
                End If
                objform.Freeze(True)
                'Matrix_Total()
                FieldDisableInMatrix()
                Status_Closing_RowDisable()
                'Matrix2.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix2.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix2.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                If Matrix2.VisualRowCount = 0 Then
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                End If
                If AutoAssignItem = "Y" Then
                    If objaddon.objglobalmethods.AutoAssign_SubItem(Folder2, Matrix2) Then
                        objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                    End If
                Else
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#") 'Scrap
                End If
                If AutoProduction = "Y" Then
                    EditText11.Item.Enabled = False
                    'ButtonCombo2.Item.Visible = False
                Else
                    'Button9.Item.Visible = False
                    ButtonCombo2.Item.Visible = True
                End If
                'Matrix2.Columns.Item("Date").Cells.Item(Matrix2.VisualRowCount).Specific.String = Now.Date.ToString("dd/MM/yy")
                'objaddon.objapplication.Menus.Item("1300").Activate()
                objform.Settings.MatrixUID = "mtxscrap"
                Matrix2.AutoResizeColumns()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

        End Sub

        Private Sub Matrix2_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix2.ChooseFromListBefore
            'In Scrap Tab - Scrap Items Groups and Input Items only listing 
            Try
                Dim combotype As SAPbouiCOM.ComboBox
                If pVal.ActionSuccess = True Then Exit Sub
                'Dim Sub_Input_Item As Boolean = False  ' For checking input item is having a header item not
                If pVal.ColUID = "Code" Then
                    Try
                        Dim GroupCode As String
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_IScrap")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()
                        If objaddon.HANA Then
                            GroupCode = objaddon.objglobalmethods.getSingleValue("select ""U_ScrapCon"" from ""@MIPL_GEN""")
                        Else
                            GroupCode = objaddon.objglobalmethods.getSingleValue("select U_ScrapCon from [@MIPL_GEN]")
                        End If
                        combotype = Matrix2.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                        If combotype.Selected.Value = "1" Then   'Scrap Items
                            If GroupCode = "Y" Then
                                If objaddon.HANA Then
                                    GroupCode = objaddon.objglobalmethods.getSingleValue("select ""U_SGroup"" from ""@MIPL_GEN""")
                                Else
                                    GroupCode = objaddon.objglobalmethods.getSingleValue("select U_SGroup from [@MIPL_GEN]")
                                End If
                                oCond = oConds.Add()
                                oCond.Alias = "ItmsGrpCod"
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                oCond.CondVal = GroupCode
                            Else
                                Dim rsetCFL As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If objaddon.HANA Then
                                    strQuery = "select T1.""U_Itemcode"",T1.""U_Qty"" ,T1.""U_Whse"" from ""@MIPL_OBOM"" T0 join ""@MIPL_BOM1"" T1 on T0.""Code""=T1.""Code"" where T1.""U_SCType""='S' and T0.""Code""='" & EditText19.Value & "'"
                                Else
                                    strQuery = "select T1.U_Itemcode,T1.U_Qty ,T1.U_Whse from [@MIPL_OBOM] T0 join [@MIPL_BOM1] T1 on T0.Code=T1.Code where T1.U_SCType='S' and T0.Code='" & EditText19.Value & "'"
                                End If
                                rsetCFL.DoQuery(strQuery)
                                If rsetCFL.RecordCount > 0 Then
                                    For i As Integer = 1 To rsetCFL.RecordCount
                                        If i = rsetCFL.RecordCount Then
                                            oCond = oConds.Add()
                                            oCond.Alias = "ItemCode"
                                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                                        Else
                                            oCond = oConds.Add()
                                            oCond.Alias = "ItemCode"
                                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                        End If
                                        rsetCFL.MoveNext()
                                    Next
                                Else
                                    For i As Integer = 1 To Matrix0.VisualRowCount
                                        If Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String <> "" Then
                                            If i = 1 Then
                                                oCond = oConds.Add()
                                                oCond.Alias = "ItemCode"
                                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                                oCond.CondVal = Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String
                                            Else
                                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                                oCond = oConds.Add()
                                                oCond.Alias = "ItemCode"
                                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                                oCond.CondVal = Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        ElseIf combotype.Selected.Value = "2" Then   'Input Items for Return
                            Dim Status As String = "", InvStatus As String = ""
                            If AutoProduction = "Y" Then
                                If objaddon.HANA Then
                                    Status = "0" ' objaddon.objglobalmethods.getSingleValue("select ifnull(""PlannedQty"",0)-ifnull((""CmpltQty""+""RjctQty""),0) ""OpenQty"" from OWOR where ""U_SubPONum""='" & objform.Items.Item("txtentry").Specific.String & "'")
                                Else
                                    Status = "0" 'objaddon.objglobalmethods.getSingleValue("select isnull(PlannedQty,0)-isnull((CmpltQty+RjctQty),0) OpenQty from OWOR where U_SubPONum='" & objform.Items.Item("txtentry").Specific.String & "'")
                                End If
                            Else
                                'If objaddon.HANA Then
                                '    Status = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T3.""U_PlanQty"")-(Select sum(T1.""Quantity"") from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" and T0.""U_SubConNo""=T2.""DocEntry"" and T1.""ItemCode""=T3.""U_Itemcode"") <=0 Then 1 else 0 end as ""Status"" " &
                                '                                                                                        " from ""@MIPL_OPOR"" T2 join ""@MIPL_POR1"" T3 on T2.""DocEntry""=T3.""DocEntry"" where T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.""DocEntry"",T3.""U_Itemcode""")
                                'Else
                                '    Status = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T3.U_PlanQty)-(Select sum(T1.Quantity) from OIGE T0 join IGE1 T1 on T0.DocEntry=T1.DocEntry and T0.U_SubConNo=T2.DocEntry and T1.ItemCode=T3.U_Itemcode) <=0 Then 1 else 0 end as Status " &
                                '                                                                                        " from [@MIPL_OPOR] T2 join [@MIPL_POR1] T3 on T2.DocEntry=T3.DocEntry where T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.DocEntry,T3.U_Itemcode ")
                                'End If
                                If objaddon.HANA Then
                                    Status = objaddon.objglobalmethods.getSingleValue("Select top 1 case when A.""OpenReturnQty"">0 then 1 Else 0 End ""Status"" from (Select T2.""DocEntry"",(Select ifnull(sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_SubConNo""=T2.""DocEntry"" and B.""ItemCode""=T3.""U_Itemcode"")" &
                                                                                       " - ((Select ifnull(sum(B.""Quantity""),0) from OIGE A join IGE1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_SubConNo""=T2.""DocEntry"" and B.""ItemCode""=T3.""U_Itemcode"") + (Select ifnull(sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_ScrapNum""=T2.""DocEntry"" and B.""ItemCode""=T3.""U_Itemcode"")) as ""OpenReturnQty"" " &
                                                                                      " from ""@MIPL_OPOR"" T2 join ""@MIPL_POR1"" T3 on T2.""DocEntry""=T3.""DocEntry"" where T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.""DocEntry"",T3.""U_Itemcode"") A order by A.""OpenReturnQty"" desc")

                                Else
                                    Status = objaddon.objglobalmethods.getSingleValue("Select top 1 case when A.OpenReturnQty>0 then 1 Else 0 End Status from (Select T2.DocEntry,(Select isnull(sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T2.DocEntry and B.ItemCode=T3.U_Itemcode)" &
                                                                                       " - ((Select isnull(sum(B.Quantity),0) from OIGE A join IGE1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T2.DocEntry and B.ItemCode=T3.U_Itemcode) + (Select isnull(sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T2.DocEntry and B.ItemCode=T3.U_Itemcode)) as OpenReturnQty " &
                                                                                      " from [@MIPL_OPOR] T2 join [@MIPL_POR1] T3 on T2.DocEntry=T3.DocEntry where T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.DocEntry,T3.U_Itemcode) A order by A.OpenReturnQty desc")
                                End If
                            End If
                            'If objaddon.HANA Then
                            '    InvStatus = objaddon.objglobalmethods.getSingleValue("select 1 from OWTR T0 left join ""@MIPL_OPOR"" T1 on T0.""U_SubConNo""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "'")
                            'Else
                            '    InvStatus = objaddon.objglobalmethods.getSingleValue("select 1 from OWTR T0 left join [@MIPL_OPOR] T1 on T0.U_SubConNo=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "'")
                            'End If
                            If Status = "1" Then 'And EditText5.Value <> "C"  (Status <> "1" And InvStatus = "1")
                                If objaddon.HANA Then
                                    strQuery = "Select * from (Select T2.""DocEntry"",T3.""U_Itemcode"",(Select ifnull(sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_SubConNo""=T2.""DocEntry"" and B.""ItemCode""=T3.""U_Itemcode"")" &
                                             " - ((Select ifnull(sum(B.""Quantity""),0) from OIGE A join IGE1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_SubConNo""=T2.""DocEntry"" and B.""ItemCode""=T3.""U_Itemcode"")+ (Select ifnull(sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_ScrapNum""=T2.""DocEntry"" and B.""ItemCode""=T3.""U_Itemcode"")) as ""OpenReturnQty"" " &
                                    " from ""@MIPL_OPOR"" T2 join ""@MIPL_POR1"" T3 on T2.""DocEntry""=T3.""DocEntry"" where T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.""DocEntry"",T3.""U_Itemcode"") A where A.""OpenReturnQty""<>0 order by A.""OpenReturnQty"" desc"

                                Else
                                    strQuery = "Select * from (Select T2.DocEntry,T3.U_Itemcode,(Select isnull(sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T2.DocEntry and B.ItemCode=T3.U_Itemcode)" &
                                             " - ((Select isnull(sum(B.Quantity),0) from OIGE A join IGE1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T2.DocEntry and B.ItemCode=T3.U_Itemcode)+ (Select isnull(sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T2.DocEntry and B.ItemCode=T3.U_Itemcode)) as OpenReturnQty " &
                                         " from [@MIPL_OPOR] T2 join [@MIPL_POR1] T3 on T2.DocEntry=T3.DocEntry where T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.DocEntry,T3.U_Itemcode) A where A.OpenReturnQty<>0 order by A.OpenReturnQty desc"
                                End If
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRs.DoQuery(strQuery)
                                If objRs.RecordCount > 0 Then
                                    For i As Integer = 0 To objRs.RecordCount - 1
                                        If i = 0 Then
                                            oCond = oConds.Add()
                                            oCond.Alias = "ItemCode"
                                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                            oCond.CondVal = objRs.Fields.Item(1).Value.ToString
                                        Else
                                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                            oCond = oConds.Add()
                                            oCond.Alias = "ItemCode"
                                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                            oCond.CondVal = objRs.Fields.Item(1).Value.ToString
                                        End If
                                        objRs.MoveNext()
                                    Next
                                End If
                            Else
                                oCond = oConds.Add()
                                oCond.Alias = "ItemCode"
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                oCond.CondVal = ""
                            End If
                        End If


                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "Whse" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_WhScrap")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()
                        oCond = oConds.Add()
                        oCond.Alias = "Inactive"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                        oCond.CondVal = "Y"

                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try

                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub Matrix2_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix2.ChooseFromListAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Dim HSNCode As String = "", WhsCode As String = ""
                Dim OnHand, Qty, ScrapQty As Double
                Dim ScrapEnabled As String
                Dim combotype As SAPbouiCOM.ComboBox
                Dim Getval As String
                objform.Freeze(True)
                If pVal.ColUID = "Code" And pVal.ActionSuccess = True Then
                    combotype = Matrix2.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                    Matrix2.FlushToDataSource()
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                odbdsDetails3.SetValue("U_Itemcode", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value)
                                odbdsDetails3.SetValue("U_ItemDesc", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value)
                                odbdsDetails3.SetValue("U_InvUom", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value)
                                'Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                                'Matrix2.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                                'Matrix2.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                            odbdsDetails3.SetValue("U_Date", pVal.Row - 1, Now.Date.ToString("yyyyMMdd"))
                            odbdsDetails3.SetValue("U_Status", pVal.Row - 1, "O")

                            'Matrix2.Columns.Item("Date").Cells.Item(pVal.Row).Specific.String = Now.Date.ToString("dd/MM/yy")
                            'Matrix2.Columns.Item("Colstat").Cells.Item(pVal.Row).Specific.String = "O"
                            ''Try
                            ''    Matrix2.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                            ''Catch ex As Exception
                            ''    Matrix2.Columns.Item("Desc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                            ''End Try
                            ''Try
                            ''    Matrix2.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                            ''Catch ex As Exception
                            ''    Matrix2.Columns.Item("InvUom").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                            ''End Try
                            If objaddon.HANA Then
                                HSNCode = objaddon.objglobalmethods.getSingleValue("select ""ChapterID"" from OCHP where ""AbsEntry""='" & pCFL.SelectedObjects.Columns.Item("ChapterID").Cells.Item(0).Value & "'")
                            Else
                                HSNCode = objaddon.objglobalmethods.getSingleValue("select ChapterID from OCHP where AbsEntry='" & pCFL.SelectedObjects.Columns.Item("ChapterID").Cells.Item(0).Value & "'")
                            End If
                            odbdsDetails3.SetValue("U_HSNCode", pVal.Row - 1, HSNCode)
                            'If HSNCode <> "" Then
                            '    Matrix2.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = HSNCode
                            'Else
                            '    Matrix2.Columns.Item("HSN").Cells.Item(pVal.Row).Specific.String = 0
                            'End If
                        End If
                        If odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) <> "" Then 'Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String <> ""
                            If TypeInScrap = "Y" Then
                                If combotype.Selected.Description = "Return" Then
                                    If AutoProduction = "Y" Then
                                        Dim GetEntry As String
                                        If objaddon.HANA Then
                                            GetEntry = objaddon.objglobalmethods.getSingleValue("select ifnull(""PlannedQty"",0)-ifnull((""CmpltQty""+""RjctQty""),0) ""OpenQty"" from OWOR where ""U_SubPONum""='" & objform.Items.Item("txtentry").Specific.String & "'")
                                        Else
                                            GetEntry = objaddon.objglobalmethods.getSingleValue("select isnull(PlannedQty,0)-isnull((CmpltQty+RjctQty),0) OpenQty from OWOR where U_SubPONum='" & objform.Items.Item("txtentry").Specific.String & "'")
                                        End If
                                        odbdsDetails3.SetValue("U_Qty", pVal.Row - 1, GetEntry)
                                        'Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = GetEntry
                                    Else
                                        ' Dim OpenQty() As Double
                                        'OpenQty = GetOpen_VendorQty(objform.Items.Item("txtentry").Specific.String)
                                        'If EditText13.Value > 0 Then
                                        'If objaddon.HANA Then
                                        '    Getval = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.""Quantity"")> 0 Then Isnull((T0.""U_PlanQty"")-(Sum(T2.""Quantity"")*T0.""U_Qty""),0) Else T0.""U_PlanQty"" End " &
                                        '                                                      "- ifnull((select Top 1 sum(T1.""U_Qty"") from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""U_Type""='2' and T1.""U_Status""='C' and T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc),0) " &
                                        '                                                      "as ""OpenInvQty"" from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"" left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" And T1.""U_SItemCode""=T2.""ItemCode"" where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Itemcode""='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' group by T0.U_PlanQty,T0.U_Qty")
                                        'Else
                                        '    Getval = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.Quantity)> 0 Then Isnull((T0.U_PlanQty)-(Sum(T2.Quantity)*T0.U_Qty),0) Else T0.U_PlanQty End " &
                                        '                                                       "- isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0) " &
                                        '                                                      "as OpenInvQty from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry And T1.U_SItemCode=T2.ItemCode where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Itemcode='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'  group by T0.U_PlanQty,T0.U_Qty")
                                        'End If
                                        If objaddon.HANA Then
                                            Getval = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.""Quantity"")> 0 Then Ifnull((T0.""U_PlanQty"")-(Sum(T2.""Quantity"")*T0.""U_Qty""),0) Else T0.""U_PlanQty"" End - " &
                                                                                               "ifnull((select Top 1 sum(B.""U_Qty"") from ""@MIPL_OPOR"" A join ""@MIPL_POR3"" B on A.""DocEntry""=B.""DocEntry"" where B.""U_Type""='2' and B.""U_Status""='C' and A.""DocEntry""=T1.""DocEntry"" and B.""U_Itemcode""=T0.""U_Itemcode"" " &
                                                                                              "and ifnull(B.""U_InvNo"",'') <>'' and ifnull(B.""U_GRNo"",'') ='' group by B.""U_Itemcode"" order by sum(B.""U_Qty"") desc),0) as ""OpenInvQty"" " &
                                                                                              "from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"" left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" And T1.""U_SItemCode""=T2.""ItemCode"" " &
                                                                                              "where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Itemcode""='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' group by T1.""DocEntry"",T0.""U_Itemcode"",T0.""U_PlanQty"",T0.""U_Qty""") 'Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String

                                        Else
                                            Getval = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.Quantity)> 0 Then Isnull((T0.U_PlanQty)-(Sum(T2.Quantity)*T0.U_Qty),0) Else T0.U_PlanQty End - " &
                                                                                               "isnull((select Top 1 sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B on A.DocEntry=B.DocEntry where B.U_Type='2' and B.U_Status='C' and A.DocEntry=T1.DocEntry and B.U_Itemcode=T0.U_Itemcode " &
                                                                                              "and isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' group by B.U_Itemcode order by sum(B.U_Qty) desc),0) as OpenInvQty " &
                                                                                              "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry And T1.U_SItemCode=T2.ItemCode " &
                                                                                              "where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Itemcode='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' group by T1.DocEntry,T0.U_Itemcode,T0.U_PlanQty,T0.U_Qty")
                                        End If
                                        odbdsDetails3.SetValue("U_Qty", pVal.Row - 1, CDbl(Getval))
                                        'Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Getval) 'CDbl(OpenQty(0)) ' CDbl(CDbl(objform.Items.Item("SQty").Specific.String) - OpenQty(0)) 'EditText25.Value
                                        'End If
                                    End If
                                    odbdsDetails3.SetValue("U_TabType", pVal.Row - 1, "Return")
                                Else
                                    If objaddon.HANA Then
                                        Getval = objaddon.objglobalmethods.getSingleValue("Select (Ifnull(T1.""U_Qty"",0) * T2.""U_SQty"")-(Select ifnull(Sum(A.""Quantity""),0) from IGN1 A left join OIGN B On A.""DocEntry""=B.""DocEntry"" where B.""U_ScrapNum""=T2.""DocEntry"" " &
                                                 "and A.""ItemCode""=T1.""U_Itemcode"") ""Open Scrap Qty"" from ""@MIPL_OBOM"" T0 left join ""@MIPL_OPOR"" T2 On T2.""U_BOMCode""=T0.""Code"" join ""@MIPL_BOM1"" T1 on T0.""Code""=T1.""Code"" where T1.""U_SCType""='S'" &
                                                 "and T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T1.""U_Itemcode""='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' ")
                                    Else
                                        Getval = objaddon.objglobalmethods.getSingleValue("Select (Isnull(T1.U_Qty,0) * T2.U_SQty)-(Select isnull(Sum(A.Quantity),0) from IGN1 A left join OIGN B On A.DocEntry=B.DocEntry where B.U_ScrapNum=T2.DocEntry " &
                                            "and A.ItemCode=T1.U_Itemcode) [Open Scrap Qty] from [@MIPL_OBOM] T0 left join [@MIPL_OPOR] T2 On T2.U_BOMCode=T0.Code join [@MIPL_BOM1] T1 on T0.Code=T1.Code where T1.U_SCType='S'" &
                                            "and T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T1.U_Itemcode='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' ")
                                    End If
                                    If Getval <> "" Then odbdsDetails3.SetValue("U_Qty", pVal.Row - 1, CDbl(Getval)) Else odbdsDetails3.SetValue("U_Qty", pVal.Row - 1, CDbl(1))
                                    odbdsDetails3.SetValue("U_TabType", pVal.Row - 1, "Scrap")
                                    'Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = "1" ' EditText25.Value
                                End If
                            Else
                                odbdsDetails3.SetValue("U_Qty", pVal.Row - 1, CDbl(1))
                                'Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = "1" 'EditText25.Value
                            End If
                        End If

                        If AutoProduction = "Y" Then
                            If objaddon.HANA Then
                                WhsCode = objaddon.objglobalmethods.getSingleValue("select ""U_WAREHOUSE"" from OCRD where ""CardCode""='" & EditText0.Value & "'")
                            Else
                                WhsCode = objaddon.objglobalmethods.getSingleValue("select U_WAREHOUSE from OCRD where CardCode='" & EditText0.Value & "'")
                            End If
                            odbdsDetails3.SetValue("U_WhsCode", pVal.Row - 1, WhsCode)
                            'Matrix2.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = WhsCode
                        End If

                        If TypeInScrap = "Y" Then
                            objcombo = Matrix2.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                            If objcombo.Selected.Description = "Return" Then
                                If objaddon.HANA Then
                                    WhsCode = objaddon.objglobalmethods.getSingleValue("Select Top 1 ""U_WhsCode"" from ""@MIPL_POR1""  where ""U_Itemcode""='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' and ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' order by ""LineId"" desc")
                                Else
                                    WhsCode = objaddon.objglobalmethods.getSingleValue("Select Top 1 U_WhsCode from [@MIPL_POR1]  where U_Itemcode='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' and DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' order by LineId desc")
                                End If
                                odbdsDetails3.SetValue("U_WhsCode", pVal.Row - 1, WhsCode)
                                'Matrix2.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = WhsCode
                            End If
                        End If

                        If objaddon.HANA Then
                            ScrapEnabled = objaddon.objglobalmethods.getSingleValue("select ""U_ScrapCon"" from ""@MIPL_GEN""")
                        Else
                            ScrapEnabled = objaddon.objglobalmethods.getSingleValue("select U_ScrapCon from [@MIPL_GEN]")
                        End If
                        If ScrapEnabled = "N" Then
                            If combotype.Selected.Description <> "Return" Then
                                Dim rsetCFL As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For i As Integer = 1 To Matrix2.VisualRowCount
                                    combotype = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                                    If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value And combotype.Selected.Description = "Scrap" Then
                                        Qty = Math.Round(CDbl(Qty + CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string)), 6)
                                    End If
                                Next

                                If objaddon.HANA Then
                                    'strQuery = "select T1.""U_Itemcode"",T1.""U_Qty"" * " & CDbl(EditText25.Value) & ",T1.""U_Whse"",T1.""U_Unitprice"" from ""@MIPL_OBOM"" T0 join ""@MIPL_BOM1"" T1 on T0.""Code""=T1.""Code"" where T1.""U_SCType""='S' and T1.""U_Itemcode""='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "'"
                                Else
                                    'strQuery = "select T1.U_Itemcode,T1.U_Qty * " & CDbl(EditText25.Value) & " ,T1.U_Whse,T1.U_Unitprice from [@MIPL_OBOM] T0 join [@MIPL_BOM1] T1 on T0.Code=T1.Code where T1.U_SCType='S' and T1.U_Itemcode='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "'"
                                    strQuery = "Select T0.U_Itemcode,isnull(Sum(T0.U_Qty*T1.U_SQty),0)- (Select isnull(Sum(T2.Quantity),0) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_ScrapNum=T1.DocEntry and T2.ItemCode=T0.U_Itemcode) as ScrapQty,"
                                    strQuery += vbCrLf + "T0.U_Whse,T0.U_Unitprice,isnull(Sum(T0.U_Qty*T1.U_SQty),0) [Tot Scrap] from [@MIPL_BOM1] T0 join [@MIPL_OPOR] T1 on T0.Code=T1.U_BOMCode where T0.U_SCType='S' and T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " and T0.U_Itemcode='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' group by T1.DocEntry,T0.U_Itemcode,T0.U_Whse,T0.U_Unitprice"
                                End If
                                rsetCFL.DoQuery(strQuery)
                                If rsetCFL.RecordCount > 0 Then
                                    ScrapQty = rsetCFL.Fields.Item("Tot Scrap").Value
                                    odbdsDetails3.SetValue("U_Qty", pVal.Row - 1, Trim(rsetCFL.Fields.Item(1).Value))
                                    If (ScrapQty - Qty) <= 0 Then
                                        ScrapQty = ScrapQty - CDbl(Qty) 'CDbl(CDbl(Qty) - CDbl(ScrapQty))
                                        odbdsDetails3.SetValue("U_Qty", pVal.Row - 1, ScrapQty)
                                    End If
                                    odbdsDetails3.SetValue("U_WhsCode", pVal.Row - 1, Trim(rsetCFL.Fields.Item(2).Value))
                                    If Trim(rsetCFL.Fields.Item(3).Value) <> 0 Then
                                        odbdsDetails3.SetValue("U_Price", pVal.Row - 1, Trim(rsetCFL.Fields.Item(3).Value))
                                    Else
                                        If objaddon.HANA Then
                                            odbdsDetails3.SetValue("U_Price", pVal.Row - 1, objaddon.objglobalmethods.getSingleValue("Select ""AvgPrice"" from OITW where ""ItemCode""='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' and ""WhsCode""='" & odbdsDetails3.GetValue("U_WhsCode", pVal.Row - 1) & "';"))
                                        Else
                                            odbdsDetails3.SetValue("U_Price", pVal.Row - 1, objaddon.objglobalmethods.getSingleValue("Select AvgPrice from OITW where ItemCode='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' and WhsCode='" & odbdsDetails3.GetValue("U_WhsCode", pVal.Row - 1) & "'"))
                                            'Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("Select AvgPrice from OITW where ItemCode='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' and WhsCode='" & Matrix2.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String & "'")
                                        End If
                                    End If
                                    odbdsDetails3.SetValue("U_LineTot", pVal.Row - 1, odbdsDetails3.GetValue("U_Qty", pVal.Row - 1) * odbdsDetails3.GetValue("U_Price", pVal.Row - 1))
                                    'Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = Trim(rsetCFL.Fields.Item(1).Value)
                                    'Matrix2.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = Trim(rsetCFL.Fields.Item(2).Value)
                                End If
                            End If
                        End If

                        'Matrix2.Columns.Item("TabType").Cells.Item(pVal.Row).Specific.String = "Scrap"
                        ' objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                        Matrix2.LoadFromDataSource()
                    Catch ex As Exception
                    End Try

                ElseIf pVal.ColUID = "Whse" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix2.FlushToDataSource()
                                odbdsDetails3.SetValue("U_WhsCode", pVal.Row - 1, pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value)
                                Matrix2.LoadFromDataSource()
                                'Matrix2.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                            Catch ex As Exception
                                'Matrix2.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If odbdsDetails3.GetValue("U_WhsCode", pVal.Row - 1) = "" Then Exit Sub
                Matrix2.FlushToDataSource()

                odbdsDetails3.SetValue("U_LineTot", pVal.Row - 1, CDbl(Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String))
                'Matrix2.Columns.Item("Linetot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix2.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String) * CDbl(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String)
                If objaddon.HANA Then
                    OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' and ""WhsCode""='" & odbdsDetails3.GetValue("U_WhsCode", pVal.Row - 1) & "'")
                Else
                    OnHand = objaddon.objglobalmethods.getSingleValue("select OnHand from  oitw where ItemCode='" & odbdsDetails3.GetValue("U_Itemcode", pVal.Row - 1) & "' and WhsCode='" & odbdsDetails3.GetValue("U_WhsCode", pVal.Row - 1) & "'")
                End If
                odbdsDetails3.SetValue("U_InStock", pVal.Row - 1, OnHand)
                'Matrix2.Columns.Item("Instock").Cells.Item(pVal.Row).Specific.String = OnHand
                Matrix2.AutoResizeColumns()
                Matrix2.LoadFromDataSource()
            Catch ex As Exception
            Finally
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub ButtonCombo2_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles ButtonCombo2.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                'If EditText5.Value = "C" Then
                '    objaddon.objapplication.StatusBar.SetText("Document status is closed.You cannot perform this action!!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '    ButtonCombo2.Item.Enabled = False
                '    Exit Sub
                'End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                    If Not ReceiptAutoLoad = "Y" Then
                        Dim SubItemCount As Integer = 0
                        For i As Integer = 1 To Matrix2.VisualRowCount
                            'objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                            If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" Then
                                If Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string = "" Or Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string = "" Or Matrix2.Columns.Item("Price").Cells.Item(i).Specific.string = "" Then
                                    objaddon.objapplication.SetStatusBarMessage("One of the column value is missing.Please update", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    BubbleEvent = False : Exit Sub
                                End If
                                'If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String Then
                                '    SubItemCount += 1
                                'End If
                                If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string) = 0 Then
                                    objaddon.objapplication.SetStatusBarMessage("Scrap Qty must be greater than 0...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                        Next
                        'If SubItemCount > 0 Then
                        '    objaddon.objapplication.SetStatusBarMessage("Scrap Tab-> Sub Item Found Please Change...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '    BubbleEvent = False : Exit Sub
                        'End If
                    End If
                    If Folder2.Selected = True Then
                        For i As Integer = 1 To Matrix2.VisualRowCount
                            If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" Then
                                If UDF0 = "Y" And (ValUDF0 = "Y" Or ValUDF0 = "2") Then
                                    If Matrix2.Columns.Item("DocDate").Cells.Item(i).Specific.string = "" Then
                                        objaddon.objapplication.SetStatusBarMessage("Scrap Tab-> Document Date missing. Please update on line : " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                                If UDF1 = "Y" And (ValUDF1 = "Y" Or ValUDF1 = "2") Then
                                    If Matrix2.Columns.Item("RefNo").Cells.Item(i).Specific.string = "" Then
                                        objaddon.objapplication.SetStatusBarMessage("Scrap Tab-> Reference No missing. Please update on line : " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                            End If
                        Next

                    End If

                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button9_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button9.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                If TypeInScrap = "Y" Then
                    Creating_Return_Auto_InventoryTransfer()
                Else
                    Auto_InventoryTransfer(Matrix2, "U_ScrapNum")
                End If

            Catch ex As Exception
            End Try
        End Sub

        Private Sub Button9_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button9.ClickBefore
            Try
                If Button9.Item.Enabled = False Then BubbleEvent = False : Exit Sub
                'If EditText5.Value = "C" Then
                '    objaddon.objapplication.StatusBar.SetText("Document status is closed.You cannot perform this action!!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '    Button9.Item.Enabled = False
                '    BubbleEvent = False : Exit Sub
                'End If
                'If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                'If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'End If
                If AutoProduction = "Y" Then
                    For i As Integer = 1 To Matrix2.VisualRowCount
                        If Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.ToString <> "" Then
                            If CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.ToString) = 0 Or Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.ToString = "" Then
                                objaddon.objapplication.StatusBar.SetText("Please update the values in row level in Scrap/Return Tab...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next
                Else
                    'Dim Status As String
                    'Dim ValidateQty() As Double = GetOpen_VendorQty(objform.Items.Item("txtentry").Specific.String)
                    'If objaddon.HANA Then
                    '    Status = objaddon.objglobalmethods.getSingleValue("Select Distinct 1 from ""@MIPL_POR2"" T1 left join  ""@MIPL_OPOR"" T0 on T0.""DocEntry""=T1.""DocEntry""  left join OIGN T2 on T0.""DocEntry""=T2.""U_SubConNo"" left join IGN1 T3 on T2.""DocEntry""=T3.""DocEntry"" and  T1.""U_Itemcode""=T3.""ItemCode""  and T0.""U_SItemCode"" =T3.""ItemCode"" " &
                    '                                                                                    "    where T0.""U_SItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and  T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T0.""U_SQty"" having T0.""U_SQty""- sum(T3.""Quantity"")<=0;  ")
                    'Else
                    '    Status = objaddon.objglobalmethods.getSingleValue(" Select Distinct 1 from [@MIPL_POR2] T1 left join  [@MIPL_OPOR] T0 on T0.DocEntry=T1.DocEntry  left join OIGN T2 on T0.DocEntry=T2.U_SubConNo left join IGN1 T3 on T2.DocEntry=T3.DocEntry and  T1.U_Itemcode=T3.ItemCode  and T0.U_SItemCode =T3.ItemCode " &
                    '                                                                                                    "    where T0.U_SItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T0.U_SQty having T0.U_SQty- sum(T3.Quantity)<=0  ")
                    'End If
                    objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    Dim flag As Boolean = False
                    Dim Item As String = ""
                    For i As Integer = 1 To Matrix2.VisualRowCount
                        objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                        If objcombo.Selected.Description = "Return" Then
                            If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" Then
                                'strQuery = "Select  T0.U_Itemcode, Case When Sum(T2.Quantity) >0 Then T1.U_SQty-(Sum(T2.Quantity)+ "
                                'strQuery += vbCrLf + "isnull((select sum(B.Quantity)/T0.U_Qty from OWTR A join WTR1 B On A.DocEntry= B.DocEntry where A.U_ScrapNum=T1.DocEntry And B.ItemCode=T0.U_Itemcode),0)) Else T1.U_SQty-"
                                'strQuery += vbCrLf + "isnull((select sum(B.Quantity)/T0.U_Qty  from OWTR A join WTR1 B On A.DocEntry= B.DocEntry where A.U_ScrapNum=T1.DocEntry And B.ItemCode=T0.U_Itemcode),0) End As OpenQty "
                                'strQuery += vbCrLf + " from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 On T1.DocEntry=T0.DocEntry left join OIGN T3 On T3.U_SubConNo=T1.DocEntry left join IGN1 T2  On T3.DocEntry=T2.DocEntry And T1.U_SItemCode=T2.ItemCode "
                                'strQuery += vbCrLf + " where T1.DocEntry ='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Itemcode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String & "' group by T1.U_SQty,T1.DocEntry,T0.U_Itemcode,T0.U_Qty"
                                'objRs.DoQuery(strQuery)
                                If objaddon.HANA Then
                                    strQuery = "Select T0.""U_Itemcode"",case when Sum(T2.""Quantity"")> 0 Then Ifnull((T0.""U_PlanQty"")-(Sum(T2.""Quantity"")*T0.""U_Qty""),0) Else T0.""U_PlanQty"" End - "
                                    strQuery += vbCrLf + "ifnull((select Top 1 sum(B.""U_Qty"") from ""@MIPL_OPOR"" A join ""@MIPL_POR3"" B on A.""DocEntry""=B.""DocEntry"" where B.""U_Type""='2' and B.""U_Status""='C' and A.""DocEntry""=T1.""DocEntry"" and B.""U_Itemcode""=T0.""U_Itemcode"" "
                                    strQuery += vbCrLf + "and ifnull(B.""U_InvNo"",'') <>'' and ifnull(B.""U_GRNo"",'') ='' group by B.""U_Itemcode"" order by sum(B.""U_Qty"") desc),0) as ""OpenInvQty"" "
                                    strQuery += vbCrLf + "from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"" left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" And T1.""U_SItemCode""=T2.""ItemCode"" "
                                    strQuery += vbCrLf + "where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Itemcode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String & "' group by T1.""DocEntry"",T0.""U_Itemcode"",T0.""U_PlanQty"",T0.""U_Qty"""

                                Else
                                    strQuery = "Select T0.U_Itemcode,case when Sum(T2.Quantity)> 0 Then Isnull((T0.U_PlanQty)-(Sum(T2.Quantity)*T0.U_Qty),0) Else T0.U_PlanQty End -  "
                                    strQuery += vbCrLf + "isnull((select Top 1 sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B on A.DocEntry=B.DocEntry where B.U_Type='2' and B.U_Status='C' and A.DocEntry=T1.DocEntry and B.U_Itemcode=T0.U_Itemcode "
                                    strQuery += vbCrLf + "and isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' group by B.U_Itemcode order by sum(B.U_Qty) desc),0) as OpenInvQty "
                                    strQuery += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry And T1.U_SItemCode=T2.ItemCode "
                                    strQuery += vbCrLf + "where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Itemcode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String & "' group by T1.DocEntry,T0.U_Itemcode,T0.U_PlanQty,T0.U_Qty"

                                End If
                                objRs.DoQuery(strQuery)
                                If objRs.RecordCount > 0 Then
                                    If objRs.Fields.Item(1).Value = 0 Then
                                        flag = True
                                        If Item = "" Then Item = objRs.Fields.Item(0).Value.ToString Else Item += ",'" & objRs.Fields.Item(0).Value.ToString & "'"
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If flag Then
                        'objcombo = Matrix2.Columns.Item("Type").Cells.Item(Matrix2.VisualRowCount).Specific
                        'If objcombo.Selected.Description = "Return" Then
                        objaddon.objapplication.StatusBar.SetText("Return not available..." & Item, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Matrix2.DeleteRow(Matrix2.VisualRowCount)
                        BubbleEvent = False : Exit Sub
                        'End If
                    End If
                    'If Validate_Batch_Serial() = False Then
                    '    BubbleEvent = False : Exit Sub
                    'End If
                End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objaddon.objapplication.StatusBar.SetText("Please update the Sub-Contracting Screen...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub

            Catch ex As Exception
            End Try
        End Sub

        Private Sub Matrix2_ComboSelectAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix2.ComboSelectAfter
            Try
                Select Case pVal.ColUID
                    Case "Type"
                        If TypeInScrap = "Y" Then
                            Dim combotype As SAPbouiCOM.ComboBox
                            combotype = Matrix2.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                            If Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String <> "" Then
                                If combotype.Selected.Value = "2" Then
                                    Matrix2.ClearRowData(pVal.Row)
                                    combotype.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Else
                                    Matrix2.ClearRowData(pVal.Row)
                                End If

                            End If
                            'If combotype.Selected.Value = "2" Then
                            '    'Matrix2.CommonSetting.SetCellEditable(pVal.Row, 8, False)
                            'Else
                            '    'Matrix2.CommonSetting.SetCellEditable(pVal.Row, 8, True)
                            'End If
                        End If
                End Select


            Catch ex As Exception
            End Try

        End Sub

        Private Sub ButtonCombo2_ComboSelectBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles ButtonCombo2.ComboSelectBefore
            'Scrap
            Try
                If Not AutoProduction = "Y" Then
                    Dim RecCount As String
                    If objaddon.HANA Then
                        RecCount = objaddon.objglobalmethods.getSingleValue("select Count(T0.""DocEntry"") from OWTR T0 left join ""@MIPL_OPOR"" T1 on T0.""U_SubConNo""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "';")
                    Else
                        RecCount = objaddon.objglobalmethods.getSingleValue("select count(T0.DocEntry) from OWTR T0 left join [@MIPL_OPOR] T1 on T0.U_SubConNo=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "'")
                    End If
                    If RecCount = "0" Then
                        objaddon.objapplication.StatusBar.SetText("Scrap- Goods Receipt not allowed due to Inventory transfer not created...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False : Exit Sub
                    End If
                    'If objaddon.HANA Then
                    '    RecCount = objaddon.objglobalmethods.getSingleValue("select Count(T0.""DocEntry"") from OIGE T0 left join ""@MIPL_OPOR"" T1 on T0.""U_SubConNo""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "';")
                    'Else
                    '    RecCount = objaddon.objglobalmethods.getSingleValue("select count(T0.DocEntry) from OIGE T0 left join [@MIPL_OPOR] T1 on T0.U_SubConNo=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "'")
                    'End If
                    'If RecCount = "0" Then
                    '    objaddon.objapplication.StatusBar.SetText("Scrap- Goods Receipt not allowed due to Goods is not Issued...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    BubbleEvent = False : Exit Sub
                    'End If
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix2_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix2.KeyDownAfter
            Try
                If pVal.ActionSuccess = True Then
                    Select Case pVal.ColUID
                        Case "Quant"
                            If CDbl(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) < 0 Then
                                objaddon.objapplication.StatusBar.SetText("In ""Qty"" (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Click()
                                Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = "1"
                            End If
                    End Select
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix2_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix2.ValidateBefore
            Try
                If pVal.ItemChanged = False Then Exit Sub
                'Dim OpenQty() As Double
                Dim Qty, ScrapQty As Double
                'OpenQty = GetOpen_VendorQty(objform.Items.Item("txtentry").Specific.String)
                Dim combotype As SAPbouiCOM.ComboBox
                Dim ItemCode As String
                ItemCode = Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.string
                combotype = Matrix2.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                If combotype.Selected.Description = "Return" Then
                    For i As Integer = 1 To Matrix2.VisualRowCount
                        combotype = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                        If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string = ItemCode And combotype.Selected.Description = "Return" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" Then
                            Qty += CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string)
                        End If
                    Next
                    'If objaddon.HANA Then
                    '    strQuery = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.""Quantity"")> 0 Then Isnull((T0.""U_PlanQty"")-(Sum(T2.""Quantity"")*T0.""U_Qty""),0) Else T0.""U_PlanQty"" End as ""OpenInvQty"" from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"" left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T1.""U_SItemCode""=T2.""ItemCode"" where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Itemcode""='" & ItemCode & "' group by T0.U_PlanQty,T0.U_Qty")
                    'Else
                    '    strQuery = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.Quantity)> 0 Then Isnull((T0.U_PlanQty)-(Sum(T2.Quantity)*T0.U_Qty),0) Else T0.U_PlanQty End as OpenInvQty from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry and T1.U_SItemCode=T2.ItemCode where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Itemcode='" & ItemCode & "' group by T0.U_PlanQty,T0.U_Qty")
                    'End If
                    'If objaddon.HANA Then
                    '    strQuery = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.""Quantity"")> 0 Then Isnull((T0.""U_PlanQty"")-(Sum(T2.""Quantity"")*T0.""U_Qty""),0) Else T0.""U_PlanQty"" End " &
                    '                                                                              "- ifnull((select Top 1 sum(T1.""U_Qty"") from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""U_Type""='2' and T1.""U_Status""='C' and T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc),0) " &
                    '                                                                              "as ""OpenInvQty"" from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"" left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" And T1.""U_SItemCode""=T2.""ItemCode"" where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Itemcode""='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' group by T0.U_PlanQty,T0.U_Qty")
                    'Else
                    '    strQuery = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.Quantity)> 0 Then Isnull((T0.U_PlanQty)-(Sum(T2.Quantity)*T0.U_Qty),0) Else T0.U_PlanQty End " &
                    '                                                                               "- isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0) " &
                    '                                                                              "as OpenInvQty from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry And T1.U_SItemCode=T2.ItemCode where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Itemcode='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "'  group by T0.U_PlanQty,T0.U_Qty")

                    'End If
                    If objaddon.HANA Then
                        strQuery = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.""Quantity"")> 0 Then Ifnull((T0.""U_PlanQty"")-(Sum(T2.""Quantity"")*T0.""U_Qty""),0) Else T0.""U_PlanQty"" End - " &
                                                                                                   "ifnull((select Top 1 sum(B.""U_Qty"") from ""@MIPL_OPOR"" A join ""@MIPL_POR3"" B on A.""DocEntry""=B.""DocEntry"" where B.""U_Type""='2' and B.""U_Status""='C' and A.""DocEntry""=T1.""DocEntry"" and B.""U_Itemcode""=T0.""U_Itemcode"" " &
                                                                                                  "and ifnull(B.""U_InvNo"",'') <>'' and ifnull(B.""U_GRNo"",'') ='' group by B.""U_Itemcode"" order by sum(B.""U_Qty"") desc),0) as ""OpenInvQty"" " &
                                                                                                  "from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"" left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" And T1.""U_SItemCode""=T2.""ItemCode"" " &
                                                                                                  "where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Itemcode""='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' group by T1.""DocEntry"",T0.""U_Itemcode"",T0.""U_PlanQty"",T0.""U_Qty""")

                    Else
                        strQuery = objaddon.objglobalmethods.getSingleValue("Select case when Sum(T2.Quantity)> 0 Then Isnull((T0.U_PlanQty)-(Sum(T2.Quantity)*T0.U_Qty),0) Else T0.U_PlanQty End - " &
                                                                                                   "isnull((select Top 1 sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B on A.DocEntry=B.DocEntry where B.U_Type='2' and B.U_Status='C' and A.DocEntry=T1.DocEntry and B.U_Itemcode=T0.U_Itemcode " &
                                                                                                  "and isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' group by B.U_Itemcode order by sum(B.U_Qty) desc),0) as OpenInvQty " &
                                                                                                  "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry And T1.U_SItemCode=T2.ItemCode " &
                                                                                                  "where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Itemcode='" & Matrix2.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String & "' group by T1.DocEntry,T0.U_Itemcode,T0.U_PlanQty,T0.U_Qty")
                    End If
                    ScrapQty = CDbl(strQuery)
                    If (ScrapQty - Qty) < 0 Then
                        ScrapQty = CDbl(CDbl(Qty) - CDbl(ScrapQty))
                        Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ScrapQty 'CDbl(strQuery) '
                    End If

                    'If CDbl((OpenQty(0)) - CDbl(Qty)) < 0 Then
                    '    'ExcessQty = CDbl(CDbl(Qty) - CDbl(objform.Items.Item("SQty").Specific.String))
                    '    Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(OpenQty(0))
                    'End If
                Else
                    For i As Integer = 1 To Matrix2.VisualRowCount
                        combotype = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                        If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string = ItemCode And combotype.Selected.Description = "Scrap" Then
                            Qty += CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string)
                        End If
                    Next
                    If objaddon.HANA Then
                        strQuery = objaddon.objglobalmethods.getSingleValue("select Ifnull(T1.""U_Qty"",0) * " & CDbl(EditText25.Value) & " from ""@MIPL_OBOM"" T0 join ""@MIPL_BOM1"" T1 on T0.""Code""=T1.""Code"" where T1.""U_SCType""='S' and T1.""U_Itemcode""='" & ItemCode & "'")
                    Else
                        strQuery = objaddon.objglobalmethods.getSingleValue("select Isnull(T1.U_Qty,0) * " & CDbl(EditText25.Value) & "  from [@MIPL_OBOM] T0 join [@MIPL_BOM1] T1 on T0.Code=T1.Code where T1.U_SCType='S' and T1.U_Itemcode='" & ItemCode & "'")
                    End If
                    ScrapQty = CDbl(strQuery)
                    If (ScrapQty - Qty) < 0 Then
                        ScrapQty = CDbl(CDbl(Qty) - CDbl(ScrapQty))
                        Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix2.Columns.Item("Quant").Cells.Item(pVal.Row).Specific.String) - ScrapQty
                        'ElseIf (ScrapQty - Qty) = 0 Then
                        '    objform.Freeze(True)
                        '    Matrix2.ClearRowData(pVal.Row)
                        '    objform.Freeze(False)
                    End If
                End If

            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

#End Region

#Region "Related Document Tab"

        Private Sub Matrix1_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.ClickAfter
            Try
                If pVal.Row = 0 Then
                    Matrix1.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                Else
                    Matrix1.SelectRow(pVal.Row, True, False)
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Matrix1_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.LostFocusAfter
            'Dim DocDate As Date
            'Dim SDocDate As String
            ''Dim cmbtype As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("DocType").Cells.Item(pVal.Row).Specific
            ''Dim CmbName As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("Docname").Cells.Item(pVal.Row).Specific
            Try
                Select Case pVal.ColUID
                    Case "Code"
                    Case "DocNum"
                        'If CmbName.Selected.Value = "13" Then
                        '    If cmbtype.Selected.Value = "I" Then
                        '        If Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String <> "" Then
                        '            If objaddon.HANA Then
                        '                SDocDate = objaddon.objglobalmethods.getSingleValue("select TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"" from OINV T0 where T0.""DocEntry""='" & Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String & "'")
                        '            Else
                        '                SDocDate = objaddon.objglobalmethods.getSingleValue("select Format(T0.DocDate,'dd/MM/yy') DocDate from OINV T0 where T0.DocEntry='" & Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String & "'")
                        '            End If
                        '            DocDate = Date.ParseExact(SDocDate, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        '            Matrix1.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.String = DocDate.ToString("dd/MM/yy")
                        '        End If
                        '    ElseIf cmbtype.Selected.Value = "S" Then
                        '        If Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String <> "" Then
                        '            If objaddon.HANA Then
                        '                SDocDate = objaddon.objglobalmethods.getSingleValue("select TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"" from OINV T0 where T0.""DocEntry""='" & Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String & "'")
                        '            Else
                        '                SDocDate = objaddon.objglobalmethods.getSingleValue("select Format(T0.DocDate,'dd/MM/yy') DocDate from OINV T0 where T0.DocEntry='" & Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String & "'")
                        '            End If
                        '            DocDate = Date.ParseExact(SDocDate, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        '            Matrix1.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.String = DocDate.ToString("dd/MM/yy")
                        '        End If
                        '    End If
                        'ElseIf CmbName.Selected.Value = "18" Then
                        '    If cmbtype.Selected.Value = "I" Then
                        '        If Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String <> "" Then
                        '            If objaddon.HANA Then
                        '                SDocDate = objaddon.objglobalmethods.getSingleValue("select TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"" from OPCH T0 where T0.""DocEntry""='" & Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String & "'")
                        '            Else
                        '                SDocDate = objaddon.objglobalmethods.getSingleValue("select Format(T0.DocDate,'dd/MM/yy') DocDate from OPCH T0 where T0.DocEntry='" & Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String & "'")
                        '            End If
                        '            DocDate = Date.ParseExact(SDocDate, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        '            Matrix1.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.String = DocDate.ToString("dd/MM/yy")
                        '        End If
                        '    ElseIf cmbtype.Selected.Value = "S" Then
                        '        If Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String <> "" Then
                        '            If objaddon.HANA Then
                        '                SDocDate = objaddon.objglobalmethods.getSingleValue("select TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"" from OPCH T0 where T0.""DocEntry""='" & Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String & "'")
                        '            Else
                        '                SDocDate = objaddon.objglobalmethods.getSingleValue("select Format(T0.DocDate,'dd/MM/yy') DocDate from OPCH T0 where T0.DocEntry='" & Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String & "'")
                        '            End If
                        '            DocDate = Date.ParseExact(SDocDate, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        '            Matrix1.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.String = DocDate.ToString("dd/MM/yy")
                        '        End If
                        '    End If
                        'End If
                End Select
                'Matrix1.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix1.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Folder3_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder3.PressedAfter
            Try
                objform.Freeze(True)
                'Matrix_Total()
                objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "DocNum", "#")
                'Matrix1.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix1.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                If APAutoLoad = "Y" Then   'For enabled automatic AP Invoice loading in Related Doc tab
                    Button3.Item.Visible = False
                End If
                'objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                'objaddon.objapplication.Menus.Item("1300").Activate()
                objform.Settings.MatrixUID = "mtxreldoc"
                Matrix1.AutoResizeColumns()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Matrix1_LinkPressedBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix1.LinkPressedBefore
            Try
                Dim cmbdocname As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("Docname").Cells.Item(pVal.Row).Specific
                Dim ColItem As SAPbouiCOM.Column = Matrix1.Columns.Item("DocNum")
                Dim objlink As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
                If cmbdocname.Selected.Value = "13" Then
                    objlink.LinkedObjectType = "13"
                    objlink.Item.LinkTo = "DocEntry"
                ElseIf cmbdocname.Selected.Value = "18" Then
                    objlink.LinkedObjectType = "18"
                    objlink.Item.LinkTo = "DocEntry"
                End If

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button3_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.PressedAfter
            Try
                Dim objrs As SAPbobsCOM.Recordset
                Dim strsql As String = ""
                If APAutoLoad = "Y" Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_POR4")
                    Dim rowno As Integer = 0
                    Try
                        rowno = Matrix1.GetCellFocus().rowIndex
                    Catch ex As Exception
                        rowno = Matrix1.VisualRowCount
                    End Try
                    Dim cmbtype As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("DocType").Cells.Item(rowno).Specific
                    Dim CmbName As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("Docname").Cells.Item(rowno).Specific
                    If Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String <> "" Then
                        If CmbName.Selected.Value = "13" Then
                            If cmbtype.Selected.Value = "I" Then
                                If objaddon.HANA Then
                                    strsql = "select T0.""DocEntry"",TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"",T1.""WhsCode"",T1.""LineTotal"" "
                                    strsql += vbCrLf + " from OINV T0 join INV1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""='" & Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String & "' "
                                Else
                                    strsql = "select T0.DocEntry,Format(T0.DocDate,'dd/MM/yy') DocDate,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.LineTotal "
                                    strsql += vbCrLf + " from OINV T0 join INV1 T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry='" & Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String & "' "
                                End If
                                objrs.DoQuery(strsql)
                                objform.Freeze(True)
                                'Matrix1.Clear()
                                odbdsDetails.Clear()
                                Try
                                    If objrs.RecordCount > 0 Then
                                        'AttDate = objaddon.objglobalmethods.GetDateTimeValue(objrs.Fields.Item("DocDate").Value.ToString) 'Format(objrs.Fields.Item("DocDate").Value.ToString, "dd/MM/yy") 'Date.ParseExact(objrs.Fields.Item("DocDate").Value, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                        While Not objrs.EoF
                                            If Matrix1.Columns.Item("Code").Cells.Item(rowno).Specific.String <> "" Then
                                                objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                                                'Matrix1.AddRow()
                                            End If
                                            odbdsDetails.Clear()
                                            Matrix1.GetLineData(rowno)
                                            odbdsDetails.SetValue("LineId", 0, rowno) 'rowno + 1)
                                            odbdsDetails.SetValue("U_DocDet", 0, CmbName.Selected.Value)
                                            odbdsDetails.SetValue("U_DocType", 0, cmbtype.Selected.Value)
                                            odbdsDetails.SetValue("U_DocNum", 0, objrs.Fields.Item("DocEntry").Value.ToString)
                                            odbdsDetails.SetValue("U_DocDate", 0, objrs.Fields.Item("DocDate").Value.ToString)
                                            odbdsDetails.SetValue("U_Itemcode", 0, objrs.Fields.Item("ItemCode").Value.ToString)
                                            odbdsDetails.SetValue("U_ItemDesc", 0, objrs.Fields.Item("Dscription").Value.ToString)
                                            odbdsDetails.SetValue("U_Qty", 0, objrs.Fields.Item("Quantity").Value.ToString)
                                            odbdsDetails.SetValue("U_WhsCode", 0, objrs.Fields.Item("WhsCode").Value.ToString)
                                            odbdsDetails.SetValue("U_LineTot", 0, objrs.Fields.Item("LineTotal").Value.ToString)
                                            Matrix1.SetLineData(rowno)
                                            rowno += 1
                                            objrs.MoveNext()
                                        End While
                                    Else
                                        objaddon.objapplication.SetStatusBarMessage("No Records Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        Exit Sub
                                    End If
                                    objform.Freeze(False)
                                Catch ex As Exception
                                End Try
                            ElseIf cmbtype.Selected.Value = "S" Then
                                If objaddon.HANA Then
                                    strsql = "select T0.""DocEntry"",TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"",T1.""AcctCode"" ""ItemCode"", "
                                    strsql += vbCrLf + " (Select ""AcctName"" from OACT where ""AcctCode""= T1.""AcctCode"" )""Dscription"", T1.""Quantity"",T1.""WhsCode"",T1.""LineTotal"" "
                                    strsql += vbCrLf + " from OINV T0 join INV1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""='" & Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String & "' "
                                Else
                                    strsql = "select T0.DocEntry,Format(T0.DocDate,'dd/MM/yy') DocDate,T1.AcctCode ItemCode, "
                                    strsql += vbCrLf + " (Select AcctName from OACT where AcctCode= T1.AcctCode )Dscription, T1.Quantity,T1.WhsCode,T1.LineTotal "
                                    strsql += vbCrLf + " from OINV T0 join INV1 T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry='" & Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String & "' "
                                End If
                                objrs.DoQuery(strsql)
                                objform.Freeze(True)
                                'Matrix1.Clear()
                                odbdsDetails.Clear()
                                If objrs.RecordCount > 0 Then
                                    ' AttDate = objaddon.objglobalmethods.GetDateTimeValue(objrs.Fields.Item("DocDate").Value.ToString)
                                    While Not objrs.EoF
                                        If Matrix1.Columns.Item("Code").Cells.Item(rowno).Specific.String <> "" Then
                                            'Matrix1.AddRow()
                                            objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                                        End If
                                        odbdsDetails.Clear()
                                        Matrix1.GetLineData(rowno)
                                        odbdsDetails.SetValue("LineId", 0, rowno)
                                        odbdsDetails.SetValue("U_DocDet", 0, CmbName.Selected.Value)
                                        odbdsDetails.SetValue("U_DocType", 0, cmbtype.Selected.Value)
                                        odbdsDetails.SetValue("U_DocNum", 0, objrs.Fields.Item("DocEntry").Value.ToString)
                                        odbdsDetails.SetValue("U_DocDate", 0, objrs.Fields.Item("DocDate").Value.ToString)
                                        ' odbdsDetails.SetValue("U_DocDate", 0, AttDate.ToString("dd/MM/yy"))
                                        odbdsDetails.SetValue("U_Itemcode", 0, objrs.Fields.Item("ItemCode").Value.ToString)
                                        odbdsDetails.SetValue("U_ItemDesc", 0, objrs.Fields.Item("Dscription").Value.ToString)
                                        odbdsDetails.SetValue("U_Qty", 0, objrs.Fields.Item("Quantity").Value.ToString)
                                        odbdsDetails.SetValue("U_WhsCode", 0, objrs.Fields.Item("WhsCode").Value.ToString)
                                        odbdsDetails.SetValue("U_LineTot", 0, objrs.Fields.Item("LineTotal").Value.ToString)
                                        Matrix1.SetLineData(rowno)
                                        rowno += 1
                                        objrs.MoveNext()
                                    End While
                                Else
                                    objaddon.objapplication.SetStatusBarMessage("No Records Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    Exit Sub
                                End If
                                objform.Freeze(False)
                            End If
                            objform.Freeze(False)
                        ElseIf CmbName.Selected.Value = "18" Then
                            If cmbtype.Selected.Value = "I" Then
                                If objaddon.HANA Then
                                    strsql = "select T0.""DocEntry"",TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"",T1.""WhsCode"",T1.""LineTotal"" "
                                    strsql += vbCrLf + " from OPCH T0 join PCH1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""='" & Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String & "'"
                                Else
                                    strsql = "select T0.DocEntry, Format(T0.DocDate,'dd/MM/yy') DocDate,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.LineTotal "
                                    strsql += vbCrLf + " from OPCH T0 join PCH1 T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry='" & Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String & "'"
                                End If
                                objrs.DoQuery(strsql)
                                objform.Freeze(True)
                                'Matrix1.Clear()
                                odbdsDetails.Clear()
                                If objrs.RecordCount > 0 Then
                                    ' AttDate = objaddon.objglobalmethods.GetDateTimeValue(objrs.Fields.Item("DocDate").Value.ToString)
                                    While Not objrs.EoF
                                        If Matrix1.Columns.Item("Code").Cells.Item(rowno).Specific.String <> "" Then
                                            'Matrix1.AddRow()
                                            objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                                        End If
                                        odbdsDetails.Clear()
                                        Matrix1.GetLineData(rowno)
                                        odbdsDetails.SetValue("LineId", 0, rowno)
                                        odbdsDetails.SetValue("U_DocDet", 0, CmbName.Selected.Value)
                                        odbdsDetails.SetValue("U_DocType", 0, cmbtype.Selected.Value)
                                        odbdsDetails.SetValue("U_DocNum", 0, objrs.Fields.Item("DocEntry").Value.ToString)
                                        odbdsDetails.SetValue("U_DocDate", 0, objrs.Fields.Item("DocDate").Value.ToString)
                                        'odbdsDetails.SetValue("U_DocDate", 0, AttDate.ToString("dd/MM/yy"))
                                        odbdsDetails.SetValue("U_Itemcode", 0, objrs.Fields.Item("ItemCode").Value.ToString)
                                        odbdsDetails.SetValue("U_ItemDesc", 0, objrs.Fields.Item("Dscription").Value.ToString)
                                        odbdsDetails.SetValue("U_Qty", 0, objrs.Fields.Item("Quantity").Value.ToString)
                                        odbdsDetails.SetValue("U_WhsCode", 0, objrs.Fields.Item("WhsCode").Value.ToString)
                                        odbdsDetails.SetValue("U_LineTot", 0, objrs.Fields.Item("LineTotal").Value.ToString)
                                        Matrix1.SetLineData(rowno)
                                        rowno += 1
                                        objrs.MoveNext()
                                    End While
                                Else
                                    objaddon.objapplication.SetStatusBarMessage("No Records Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    Exit Sub
                                End If
                                objform.Freeze(False)
                            ElseIf cmbtype.Selected.Value = "S" Then
                                If objaddon.HANA Then
                                    strsql = "select T0.""DocEntry"",TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"",T1.""AcctCode"" ""ItemCode"", "
                                    strsql += vbCrLf + " (Select ""AcctName"" from OACT where ""AcctCode""= T1.""AcctCode"" )""Dscription"", T1.""Quantity"",T1.""WhsCode"",T1.""LineTotal"" "
                                    strsql += vbCrLf + " from OPCH T0 join PCH1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""='" & Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String & "' "
                                Else
                                    strsql = "select T0.DocEntry,Format(T0.DocDate,'dd/MM/yy') DocDate,T1.AcctCode ItemCode, "
                                    strsql += vbCrLf + " (Select AcctName from OACT where AcctCode= T1.AcctCode )Dscription, T1.Quantity,T1.WhsCode,T1.LineTotal "
                                    strsql += vbCrLf + " from OPCH T0 join PCH1 T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry='" & Matrix1.Columns.Item("DocNum").Cells.Item(rowno).Specific.String & "' "
                                End If
                                objrs.DoQuery(strsql)
                                objform.Freeze(True)
                                'Matrix1.Clear()
                                odbdsDetails.Clear()
                                If objrs.RecordCount > 0 Then
                                    'AttDate = objaddon.objglobalmethods.GetDateTimeValue(objrs.Fields.Item("DocDate").Value.ToString)
                                    While Not objrs.EoF
                                        If Matrix1.Columns.Item("Code").Cells.Item(rowno).Specific.String <> "" Then
                                            'Matrix1.AddRow()
                                            objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                                        End If
                                        odbdsDetails.Clear()
                                        Matrix1.GetLineData(rowno)
                                        odbdsDetails.SetValue("LineId", 0, rowno)
                                        odbdsDetails.SetValue("U_DocDet", 0, CmbName.Selected.Value)
                                        odbdsDetails.SetValue("U_DocType", 0, cmbtype.Selected.Value)
                                        odbdsDetails.SetValue("U_DocNum", 0, objrs.Fields.Item("DocEntry").Value.ToString)
                                        odbdsDetails.SetValue("U_DocDate", 0, objrs.Fields.Item("DocDate").Value.ToString)
                                        ' odbdsDetails.SetValue("U_DocDate", 0, AttDate.ToString("dd/MM/yy"))
                                        odbdsDetails.SetValue("U_Itemcode", 0, objrs.Fields.Item("ItemCode").Value.ToString)
                                        odbdsDetails.SetValue("U_ItemDesc", 0, objrs.Fields.Item("Dscription").Value.ToString)
                                        odbdsDetails.SetValue("U_Qty", 0, objrs.Fields.Item("Quantity").Value.ToString)
                                        odbdsDetails.SetValue("U_WhsCode", 0, objrs.Fields.Item("WhsCode").Value.ToString)
                                        odbdsDetails.SetValue("U_LineTot", 0, objrs.Fields.Item("LineTotal").Value.ToString)
                                        Matrix1.SetLineData(rowno)
                                        rowno += 1
                                        objrs.MoveNext()
                                    End While
                                Else
                                    objaddon.objapplication.SetStatusBarMessage("No Records Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    Exit Sub
                                End If
                                objform.Freeze(False)
                            End If
                        End If
                    Else
                        objaddon.objapplication.StatusBar.SetText("Please select the invoice number...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                    Matrix1.AutoResizeColumns()
                    'objaddon.objapplication.Menus.Item("1300").Activate()
                    'If Matrix1.Columns.Item("DocNum").Cells.Item(Matrix1.VisualRowCount).Specific.String <> "" Then
                    '    objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "DocNum", "#")
                    'End If
                    'Matrix1.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    'Matrix1.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                End If
                objrs = Nothing
            Catch ex As Exception
                objform.Freeze(False)
            Finally
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Button3_PressedBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button3.PressedBefore
            Try
                'RemoveLastrow(Matrix1, "Code")
                Dim DocEntry, Type, Name As String
                Dim Row As Integer = 0
                Dim cmbtype, CmbName As SAPbouiCOM.ComboBox
                Dim flag As Boolean = False
                Try
                    DocEntry = Matrix1.Columns.Item("DocNum").Cells.Item(Matrix1.GetCellFocus().rowIndex).Specific.String
                    Type = Matrix1.Columns.Item("DocType").Cells.Item(Matrix1.GetCellFocus().rowIndex).Specific.Selected.Value
                    Name = Matrix1.Columns.Item("Docname").Cells.Item(Matrix1.GetCellFocus().rowIndex).Specific.Selected.Value
                    Row = Matrix1.GetCellFocus().rowIndex
                Catch ex As Exception
                    DocEntry = Matrix1.Columns.Item("DocNum").Cells.Item(Matrix1.VisualRowCount).Specific.String
                    Type = Matrix1.Columns.Item("DocType").Cells.Item(Matrix1.VisualRowCount).Specific.Selected.Value
                    Name = Matrix1.Columns.Item("Docname").Cells.Item(Matrix1.VisualRowCount).Specific.Selected.Value
                    Row = Matrix1.VisualRowCount
                End Try
                If Matrix1.VisualRowCount > 1 Then
                    For i As Integer = Matrix1.VisualRowCount To 1 Step -1
                        cmbtype = Matrix1.Columns.Item("DocType").Cells.Item(i).Specific
                        CmbName = Matrix1.Columns.Item("Docname").Cells.Item(i).Specific
                        If Row <> i And cmbtype.Selected.Value = Type And CmbName.Selected.Value = Name And Matrix1.Columns.Item("DocNum").Cells.Item(i).Specific.String = DocEntry Then
                            flag = True
                            Exit For
                        End If
                    Next
                End If
                If flag = True Then
                    If objaddon.objapplication.MessageBox(DocEntry & " already loaded. Do you want to load again?", 2, "Yes", "No") <> 1 Then Matrix1.ClearRowData(Row) : BubbleEvent = False : Exit Sub

                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix1_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix1.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim cmbtype As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("DocType").Cells.Item(pVal.Row).Specific
                Dim cmbdocname As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("Docname").Cells.Item(pVal.Row).Specific
                Dim ColItem As SAPbouiCOM.Column = Matrix1.Columns.Item("DocNum")
                If cmbdocname.Selected.Value = "13" Then
                    ColItem.ChooseFromListUID.Remove(pVal.Row)
                    ColItem.ChooseFromListUID = "CFL_ARDoc"
                    ColItem.ChooseFromListAlias = "DocEntry"
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_ARDoc")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()
                        oCond = oConds.Add()
                        oCond.Alias = "DocType"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = cmbtype.Selected.Value
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf cmbdocname.Selected.Value = "18" Then
                    ColItem.ChooseFromListUID.Remove(pVal.Row)
                    ColItem.ChooseFromListUID = "CFL_APDoc"
                    ColItem.ChooseFromListAlias = "DocEntry"
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_APDoc")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DocType"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = cmbtype.Selected.Value
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Matrix1_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.ChooseFromListAfter
            If pVal.ColUID = "DocNum" And pVal.ActionSuccess = True Then
                Try
                    If pVal.ActionSuccess = False Then Exit Sub
                    pCFL = pVal
                    If Not pCFL.SelectedObjects Is Nothing Then
                        Try
                            Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value
                        Catch ex As Exception
                            Matrix1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value
                        End Try
                    End If
                    Matrix1.AutoResizeColumns()
                    'objaddon.objapplication.Menus.Item("1300").Activate()
                Catch ex As Exception
                End Try
            End If

        End Sub

#End Region

#Region "Costing Tab"

        Private Sub Matrix4_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix4.ClickAfter
            Try
                If pVal.Row = 0 Then
                    Matrix4.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                Else
                    Matrix4.SelectRow(pVal.Row, True, False)
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix4_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix4.KeyDownAfter
            Try
                Dim Row As Integer = Matrix4.VisualRowCount - 1
                Row = Row Mod 2
                Select Case pVal.ColUID
                    Case "Debit"
                        If Row <> 0 Then
                            If CDbl(Matrix4.Columns.Item("Debit").Cells.Item(Matrix4.VisualRowCount - 1).Specific.String) = 0.0 Then
                                Matrix4.Columns.Item("Debit").Cells.Item(Matrix4.VisualRowCount - 1).Click()
                                objaddon.objapplication.SetStatusBarMessage("Please Enter the amount in Debit...It will auto calculate to Credit... Row : " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                Exit Sub
                            End If
                        End If
                    Case "Credit"
                        Dim CurrRow As Integer = pVal.Row
                        If CurrRow = 1 Then
                            If CDbl(Matrix4.Columns.Item("Credit").Cells.Item(CurrRow).Specific.String) = 0.0 Then
                                Matrix4.Columns.Item("Credit").Cells.Item(CurrRow).Click()
                                objaddon.objapplication.SetStatusBarMessage("Please Enter the amount in Debit...It will auto calculate to Credit...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                Exit Sub
                            End If
                        End If
                End Select
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix4_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix4.LostFocusAfter
            Try
                If pVal.ActionSuccess = True Then
                    Select Case pVal.ColUID
                        Case "Code"
                            Dim Total As Double = 0.0
                            Dim Row As Integer = Matrix4.VisualRowCount
                            Try
                                objform.Freeze(True)
                                If Matrix4.VisualRowCount <> pVal.Row Then
                                    For i As Integer = 1 To Matrix4.VisualRowCount - 1
                                        If CDbl(Matrix4.Columns.Item("Debit").Cells.Item(i).Specific.String) = 0.0 And CDbl(Matrix4.Columns.Item("Credit").Cells.Item(i).Specific.String) <> 0.0 And Matrix4.Columns.Item("Stat").Cells.Item(i).Specific.String = "O" Then
                                            Total += CDbl(Matrix4.Columns.Item("Credit").Cells.Item(i).Specific.String)
                                        Else
                                            If Matrix4.VisualRowCount = i Then
                                                If CDbl(Matrix4.Columns.Item("Credit").Cells.Item(i).Specific.String) = 0.0 Then
                                                    Matrix4.Columns.Item("Debit").Cells.Item(i).Specific.String = CStr(Total)
                                                End If
                                            End If
                                            Total = 0
                                        End If
                                    Next
                                End If
                                objform.Freeze(False)
                            Catch ex As Exception
                                objform.Freeze(False)
                            End Try
                            If Matrix4.Columns.Item("Branch").Cells.Item(pVal.Row - 1).Specific.String <> "" Then
                                Matrix4.Columns.Item("Branch").Cells.Item(pVal.Row).Specific.String = Matrix4.Columns.Item("Branch").Cells.Item(pVal.Row - 1).Specific.String
                            Else
                                Matrix4.Columns.Item("Branch").Cells.Item(pVal.Row).Specific.String = Matrix4.Columns.Item("Branch").Cells.Item(pVal.Row).Specific.String
                            End If
                        Case "Debit"
                            Dim Total As Double = 0.0
                            Try
                                objform.Freeze(True)
                                If Matrix4.VisualRowCount <> pVal.Row Then
                                    For i As Integer = 1 To Matrix4.VisualRowCount
                                        If CDbl(Matrix4.Columns.Item("Credit").Cells.Item(i).Specific.String) = 0.0 And CDbl(Matrix4.Columns.Item("Debit").Cells.Item(i).Specific.String) <> 0.0 And Matrix4.Columns.Item("Stat").Cells.Item(i).Specific.String = "O" Then
                                            Total += CDbl(Matrix4.Columns.Item("Debit").Cells.Item(i).Specific.String)
                                            Matrix4.CommonSetting.SetCellEditable(i, 5, False)
                                        Else
                                            If Matrix4.VisualRowCount <> i Then
                                                If CDbl(Matrix4.Columns.Item("Debit").Cells.Item(i).Specific.String) = 0.0 Then
                                                    Matrix4.Columns.Item("Credit").Cells.Item(i).Specific.String = CStr(Total)
                                                End If
                                            End If
                                            Total = 0
                                        End If
                                    Next
                                End If
                                objform.Freeze(False)
                            Catch ex As Exception
                                objform.Freeze(False)
                            End Try

                        Case "Credit"
                            If Matrix4.VisualRowCount <> pVal.Row Then
                                For i As Integer = 1 To Matrix4.VisualRowCount - 1
                                    If CDbl(Matrix4.Columns.Item("Debit").Cells.Item(i).Specific.String) = 0.0 And CDbl(Matrix4.Columns.Item("Credit").Cells.Item(i).Specific.String) <> 0.0 And Matrix4.Columns.Item("Stat").Cells.Item(i).Specific.String = "O" Then
                                        Matrix4.CommonSetting.SetCellEditable(i, 4, False)
                                    End If
                                Next
                            End If

                    End Select
                    'Matrix4.Columns.Item("Debit").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    'Matrix4.Columns.Item("Credit").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Folder4_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder4.PressedAfter
            Try
                objform.Freeze(True)
                'Matrix_Total()
                objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                Matrix4.Columns.Item("Date").Cells.Item(Matrix4.VisualRowCount).Specific.String = Now.Date.ToString("dd/MM/yy")
                'Matrix4.Columns.Item("Debit").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix4.Columns.Item("Credit").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objform.Settings.MatrixUID = "MtxCosting"
                Matrix4.AutoResizeColumns()
                'objaddon.objapplication.Menus.Item("1300").Activate()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub Matrix4_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix4.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                If pVal.ColUID = "Code" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFLGL")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()
                        oCond = oConds.Add()
                        oCond.Alias = "Postable"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "Y"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                    End Try

                ElseIf pVal.ColUID = "CostCent" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFLCostCent")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "1"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "CostCent1" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFLCosCent1")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "2"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "CostCent2" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFLCosCent2")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "3"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                    End Try

                ElseIf pVal.ColUID = "CostCent3" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFLCosCent3")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "4"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "CostCent4" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFLCosCent4")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "5"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix4_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix4.ChooseFromListAfter
            Try
                If pVal.ColUID = "Code" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix4.Columns.Item("Code").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                            Try
                                Matrix4.Columns.Item("Name").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                        End If
                        objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                        Matrix4.Columns.Item("Date").Cells.Item(pVal.Row).Specific.String = Now.Date.ToString("dd/MM/yy")
                        Matrix4.Columns.Item("Stat").Cells.Item(pVal.Row).Specific.String = "O"
                    Catch ex As Exception
                    Finally
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    End Try
                End If
                If pVal.ColUID = "CostCent" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix4.Columns.Item("CostCent").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.ColUID = "CostCent1" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix4.Columns.Item("CostCent1").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.ColUID = "CostCent2" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix4.Columns.Item("CostCent2").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.ColUID = "CostCent3" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix4.Columns.Item("CostCent3").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.ColUID = "CostCent4" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix4.Columns.Item("CostCent4").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.ColUID = "Proj" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix4.Columns.Item("Proj").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("PrjCode").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.ColUID = "Branch" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix4.Columns.Item("Branch").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("BPLName").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                Matrix4.AutoResizeColumns()
                'objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button10_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button10.ClickBefore
            Try
                Dim Branch As String = ""
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                'Matrix4.Columns.Item("Debit").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix4.Columns.Item("Credit").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                If objaddon.HANA Then
                    Branch = objaddon.objglobalmethods.getSingleValue("select ""MltpBrnchs"" from OADM")
                Else
                    Branch = objaddon.objglobalmethods.getSingleValue("select MltpBrnchs from OADM")
                End If

                If Branch = "Y" Then
                    For i As Integer = 1 To Matrix4.VisualRowCount
                        If Matrix4.Columns.Item("Branch").Cells.Item(i).Specific.String = "" And Matrix4.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix4.Columns.Item("Stat").Cells.Item(i).Specific.String = "O" Then
                            objaddon.objapplication.SetStatusBarMessage("Please Update the Branch to posting JE", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False : Exit Sub
                        End If
                    Next
                End If
                If Matrix4.Columns.Item("Debit").ColumnSetting.SumType <> Matrix4.Columns.Item("Credit").ColumnSetting.SumType Then
                    objaddon.objapplication.SetStatusBarMessage("UnBalanced Credit & Debit. Please Check", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button10_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button10.ClickAfter
            Try
                objaddon.objapplication.SetStatusBarMessage("Creating Journal Entry Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                JournalEntry()
            Catch ex As Exception
            End Try
        End Sub

#End Region

#Region "Functions"

        Private Function GetProcessQty() As Double
            Try
                Dim SQuery As String, Getseq As String
                Dim objRs As SAPbobsCOM.Recordset
                Dim ProcessQty As Double
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    Getseq = objaddon.objglobalmethods.getSingleValue("Select Top 1 T0.""U_Sequence"" as ""Seq"" from ""@MIPL_BOM2"" T0 where T0.""Code""='" & Trim(objform.Items.Item("txtbitem").Specific.String) & "' and T0.""U_Priority""='2' and T0.""U_Sequence""<'" & Trim(objform.Items.Item("tprocseq").Specific.String) & "' order by T0.""U_Sequence"" desc")

                    SQuery = "Select Case when A.""BOMPrior""='2' and A.""ReceiptQty"" >0 then  A.""ReceiptQty"" Else A.""SubItemQty""-A.""ReceiptQty"" End as ""Qty"",* from(Select (Select ifnull(Sum(T2.""Quantity""),0) from OIGN T3 left join IGN1 T2  On T3.""DocEntry""= T2.""DocEntry"""
                    SQuery += vbCrLf + "And T1.""U_SItemCode""=T2.""ItemCode"" And T3.""U_SubConNo""=T1.""DocEntry"" "
                    If Getseq = "" Then
                        SQuery += vbCrLf + "where T2.""U_Process"" ='" & Trim(objform.Items.Item("tprocseq").Specific.String) & "')"
                    Else
                        SQuery += vbCrLf + "where T2.""U_Process"" ='" & Getseq & "')"
                    End If
                    SQuery += vbCrLf + "as ""ReceiptQty"",T1.""U_SQty"" as ""SubItemQty"",(Select ""U_Priority"" from ""@MIPL_BOM2"" where ""Code""=T1.""U_BOMCode"""
                    If Getseq = "" Then
                        SQuery += vbCrLf + "and ""U_Sequence"" ='" & Getseq & "')"
                    Else
                        SQuery += vbCrLf + "and ""U_Sequence"" ='" & Trim(objform.Items.Item("tprocseq").Specific.String) & "')"
                    End If
                    SQuery += vbCrLf + " as ""BOMPrior"" from ""@MIPL_OPOR"" T1 where T1.""U_PONum""='" & objform.Items.Item("txtponum").Specific.String & "' "
                    If Getseq = "" Then
                        SQuery += vbCrLf + "And T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'"
                    Else
                        SQuery += vbCrLf + "And T1.""DocEntry""=(Select ""DocEntry"" from ""@MIPL_OPOR"" where ""U_PONum""='" & objform.Items.Item("txtponum").Specific.String & "' and ""U_Procseq""='" & Getseq & "')"
                    End If
                    SQuery += vbCrLf + ") A"
                Else
                    Getseq = objaddon.objglobalmethods.getSingleValue("Select Top 1 T0.U_Sequence as Seq from [@MIPL_BOM2] T0 where T0.Code='" & Trim(objform.Items.Item("txtbitem").Specific.String) & "' and T0.U_Priority='2' and T0.U_Sequence<'" & Trim(objform.Items.Item("tprocseq").Specific.String) & "' order by T0.U_Sequence desc")

                    SQuery = "Select Case when A.BOMPrior='2' and A.ReceiptQty >0 then  A.ReceiptQty Else A.SubItemQty-A.ReceiptQty End as Qty,* from(Select (Select isnull(Sum(T2.Quantity),0) from OIGN T3 left join IGN1 T2  On T3.DocEntry= T2.DocEntry"
                    SQuery += vbCrLf + "And T1.U_SItemCode=T2.ItemCode And T3.U_SubConNo=T1.DocEntry "
                    If Getseq = "" Then
                        SQuery += vbCrLf + "where T2.U_Process ='" & Trim(objform.Items.Item("tprocseq").Specific.String) & "')"
                    Else
                        SQuery += vbCrLf + "where T2.U_Process ='" & Getseq & "')"
                    End If
                    SQuery += vbCrLf + "as ReceiptQty,T1.U_SQty as SubItemQty,(Select U_Priority from [@MIPL_BOM2] where Code=T1.U_BOMCode"
                    If Getseq = "" Then
                        SQuery += vbCrLf + "and U_Sequence ='" & Getseq & "')"
                    Else
                        SQuery += vbCrLf + "and U_Sequence ='" & Trim(objform.Items.Item("tprocseq").Specific.String) & "')"
                    End If
                    SQuery += vbCrLf + " as BOMPrior from [@MIPL_OPOR] T1 where T1.U_PONum='" & objform.Items.Item("txtponum").Specific.String & "' "
                    If Getseq = "" Then
                        SQuery += vbCrLf + "And T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'"
                    Else
                        SQuery += vbCrLf + "And T1.DocEntry=(Select DocEntry from [@MIPL_OPOR] where U_PONum='" & objform.Items.Item("txtponum").Specific.String & "' and U_Procseq='" & Getseq & "')"
                    End If
                    SQuery += vbCrLf + ") A"

                End If
                objRs.DoQuery(SQuery)
                If objRs.RecordCount > 0 Then
                    ProcessQty = CDbl(objRs.Fields.Item(0).Value.ToString)
                End If
                Return ProcessQty
            Catch ex As Exception

            End Try
        End Function

        Private Sub CreateDynamicUDF()
            Try
                Dim StrQuery As String = ""
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim I, O, S, R, C As Integer
                I = Matrix0.Columns.Count
                O = Matrix3.Columns.Count
                S = Matrix2.Columns.Count
                R = Matrix1.Columns.Count
                C = Matrix4.Columns.Count

                If objaddon.HANA Then
                    StrQuery = "select Count(*), ""FieldID"",""AliasID"",""Descr"",""TableID"" from CUFD where ""TableID""='@MIPL_POR1' and ""U_Valid""='Y' group by ""FieldID"",""AliasID"",""Descr"",""TableID"" having ""FieldID"">" & I & ""
                    StrQuery += vbCrLf + " union all"
                    StrQuery += vbCrLf + "Select Count(*), ""FieldID"",""AliasID"",""Descr"",""TableID"" from CUFD where ""TableID""='@MIPL_POR2' and ""U_Valid""='Y' group by ""FieldID"",""AliasID"",""Descr"",""TableID"" having ""FieldID"">" & O & ""
                    StrQuery += vbCrLf + "union all"
                    StrQuery += vbCrLf + "select Count(*),""FieldID"",""AliasID"",""Descr"",""TableID"" from CUFD where ""TableID""='@MIPL_POR3' and ""U_Valid""='Y' group by ""FieldID"",""AliasID"",""Descr"",""TableID"" having ""FieldID"">" & S & ""
                    StrQuery += vbCrLf + "union all"
                    StrQuery += vbCrLf + "Select Count(*), ""FieldID"",""AliasID"",""Descr"",""TableID"" from CUFD where ""TableID""='@MIPL_POR4' and ""U_Valid""='Y' group by ""FieldID"",""AliasID"",""Descr"",""TableID"" having ""FieldID"">" & R & ""
                    StrQuery += vbCrLf + "union all"
                    StrQuery += vbCrLf + " select Count(*), ""FieldID"",""AliasID"",""Descr"",""TableID"" from CUFD where ""TableID""='@MIPL_POR5' and ""U_Valid""='Y' group by ""FieldID"",""AliasID"",""Descr"",""TableID"" having ""FieldID"">" & C & ""
                Else
                    StrQuery = "select Count(*), FieldID,AliasID,Descr,TableID from CUFD where TableID='@MIPL_POR1' and U_Valid='Y' group by FieldID,AliasID,Descr,TableID having FieldID>" & I & ""
                    StrQuery += vbCrLf + " union all"
                    StrQuery += vbCrLf + "Select Count(*), FieldID,AliasID,Descr,TableID from CUFD where TableID='@MIPL_POR2' and U_Valid='Y' group by FieldID,AliasID,Descr,TableID having FieldID>" & O & ""
                    StrQuery += vbCrLf + "union all"
                    StrQuery += vbCrLf + "select Count(*),FieldID,AliasID,Descr,TableID from CUFD where TableID='@MIPL_POR3' and U_Valid='Y' group by FieldID,AliasID,Descr,TableID having FieldID>" & S & ""
                    StrQuery += vbCrLf + "union all"
                    StrQuery += vbCrLf + "Select Count(*), FieldID,AliasID,Descr,TableID from CUFD where TableID='@MIPL_POR4' and U_Valid='Y' group by FieldID,AliasID,Descr,TableID having FieldID>" & R & ""
                    StrQuery += vbCrLf + "union all"
                    StrQuery += vbCrLf + " select Count(*), FieldID,AliasID,Descr,TableID from CUFD where TableID='@MIPL_POR5' and U_Valid='Y' group by FieldID,AliasID,Descr,TableID having FieldID>" & C & ""
                End If

                objRs.DoQuery(StrQuery)
                If objRs.RecordCount > 0 Then
                    For Rec As Integer = 0 To objRs.RecordCount - 1
                        Dynamic_LineUDF(getMatrixID(objRs.Fields.Item("TableID").Value.ToString), Trim("U_" & objRs.Fields.Item("AliasID").Value.ToString), objRs.Fields.Item("TableID").Value.ToString, objRs.Fields.Item("Descr").Value.ToString)
                        objRs.MoveNext()
                    Next
                End If

            Catch ex As Exception

            End Try
        End Sub

        Private Sub Dynamic_LineUDF(ByVal MatrixID As SAPbouiCOM.Matrix, ByVal UID As String, ByVal TableName As String, ByVal Descr As String)
            Try
                Dim strsql As String
                If objaddon.HANA Then
                    strsql = objaddon.objglobalmethods.getSingleValue("select distinct 1 as ""Status"" from  UFD1 T1 inner join CUFD T0 on T0.""TableID""=T1.""TableID"" and T0.""FieldID""=T1.""FieldID"" where T0.""TableID""='" & TableName & "' and T0.""Descr""='" & Descr & "'")
                Else
                    strsql = objaddon.objglobalmethods.getSingleValue("select distinct 1 as Status from  UFD1 T1 inner join CUFD T0 on T0.TableID=T1.TableID and T0.FieldID=T1.FieldID where T0.TableID='" & TableName & "' and T0.Descr='" & Descr & "'")
                End If
                If strsql <> "" Then
                    MatrixID.Columns.Add(UID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                    MatrixID.Columns.Item(UID).DisplayDesc = True
                Else
                    MatrixID.Columns.Add(UID, SAPbouiCOM.BoFormItemTypes.it_EDIT)
                End If
                MatrixID.Columns.Item(UID).DataBind.SetBound(True, TableName, UID)
                MatrixID.Columns.Item(UID).Editable = True
                MatrixID.Columns.Item(UID).TitleObject.Caption = Descr


            Catch ex As Exception
            End Try
        End Sub

        Private Function getMatrixID(ByVal TableName As String)
            Try
                Dim MatriID As SAPbouiCOM.Matrix
                If TableName = "@MIPL_POR1" Then
                    MatriID = Matrix0 '"MtxinputN"
                ElseIf TableName = "@MIPL_POR2" Then
                    MatriID = Matrix3
                ElseIf TableName = "@MIPL_POR3" Then
                    MatriID = Matrix2
                ElseIf TableName = "@MIPL_POR4" Then
                    MatriID = Matrix1
                ElseIf TableName = "@MIPL_POR5" Then
                    MatriID = Matrix4
                End If
                Return MatriID
            Catch ex As Exception

            End Try
        End Function

        Private Function CheckInputItem(ByVal ItemCode As String) As Boolean
            Dim Input As Boolean = False
            Try
                For j As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("Code").Cells.Item(j).Specific.String = ItemCode Then
                        Input = True
                        Exit For
                    End If
                Next
            Catch ex As Exception
            End Try
            Return Input
        End Function

        Private Sub CalcGRPrice()
            Try
                Dim Input As Boolean = False
                Dim Count As Integer = 0
                Dim JEAmount, GetValue, Unitprice, ByProductPrice As String
                Dim BYPPriceTotal As Double = 0.0
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If VendorSPlPrice = "1" Then
                    For i = 1 To Matrix3.VisualRowCount
                        If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix3.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" And Matrix3.Columns.Item("GRNo").Cells.Item(i).Specific.String = "" Then
                            If objaddon.HANA Then
                                GetValue = objaddon.objglobalmethods.getSingleValue("select ifnull(sum(T1.""Quantity"" * T1.""StockPrice""),0) as ""UnitPrice""  from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.string & "' group by T0.""DocEntry"" order by T0.""DocEntry"" desc")
                            Else
                                GetValue = objaddon.objglobalmethods.getSingleValue("select isnull(sum(T1.Quantity * T1.StockPrice),0) as UnitPrice  from OIGE T0 join IGE1 T1 on T0.DocEntry=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.string & "' group by T0.DocEntry order by T0.DocEntry desc")
                            End If
                            Unitprice = CDbl(CDbl(IIf(GetValue = "", 0, GetValue))) / CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                            If objaddon.HANA Then
                                GetValue = "select * from (Select distinct T0.""CardCode"",T0.""ItemCode"", "
                                GetValue += vbCrLf + "case when ifnull(T0.""EXPAND"",'')='Y' Then (Select ""Price"" from SPP1 where current_date between ""FromDate"" and ifnull(""ToDate"",current_date) and ""CardCode""='" & EditText0.Value & "' and ""ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "') Else T0.""Price"" End as ""SplPrice"" "
                                GetValue += vbCrLf + "from OSPP T0 left join SPP1 T1 on T0.""ItemCode""=T1.""ItemCode"" and T0.""CardCode""=T1.""CardCode"") as A "
                                GetValue += vbCrLf + "where A.""CardCode""='" & EditText0.Value & "' and A.""ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "';"
                            Else
                                GetValue = "select * from (Select distinct T0.CardCode,T0.ItemCode, "
                                GetValue += vbCrLf + "case when isnull(T0.EXPAND,'')='Y' Then (Select Price from SPP1 where GetDate() between FromDate and isnull(ToDate,GetDate()) and CardCode='" & EditText0.Value & "' and ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "') Else T0.Price End as SplPrice "
                                GetValue += vbCrLf + "from OSPP T0 left join SPP1 T1 on T0.ItemCode=T1.ItemCode and T0.CardCode=T1.CardCode) as A "
                                GetValue += vbCrLf + "where A.CardCode='" & EditText0.Value & "' and A.ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "'"
                            End If
                            objRS.DoQuery(GetValue)
                            If objRS.RecordCount > 0 Then
                                Unitprice = CDbl(Unitprice) + CDbl(objRS.Fields.Item("SplPrice").Value.ToString)
                                Matrix3.Columns.Item("ProCost").Cells.Item(i).Specific.String = CDbl(objRS.Fields.Item("SplPrice").Value.ToString)
                                Matrix3.Columns.Item("TProCost").Cells.Item(i).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String) * CDbl(objRS.Fields.Item("SplPrice").Value.ToString)
                            Else
                                Unitprice = CDbl(Unitprice)
                            End If
                            Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String = Unitprice
                            Matrix3.Columns.Item("Linetot").Cells.Item(i).Specific.String = CDbl(Unitprice) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                        End If
                    Next
                ElseIf VendorSPlPrice = "2" Then
                    If objaddon.HANA Then
                        GetValue = "Select T0.""TableName"" as ""Header"" ,T1.""TableName"" as ""Line"" from OUDO T0 join UDO1 T1 on T0.""Code""=T1.""Code"" where T0.""TableName""=(Select case when ""U_UDOName"" is null then 'PRICEHEAD' else ""U_UDOName"" end from ""@MIPL_GEN"")"
                    Else
                        GetValue = "Select T0.TableName as Header ,T1.TableName as Line from OUDO T0 join UDO1 T1 on T0.Code=T1.Code where T0.TableName=(Select case when U_UDOName is null then 'PRICEHEAD' else U_UDOName end from [@MIPL_GEN])"
                    End If
                    objRS.DoQuery(GetValue)
                    For i = 1 To Matrix3.VisualRowCount
                        If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix3.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" And Matrix3.Columns.Item("GRNo").Cells.Item(i).Specific.String = "" Then
                            If objaddon.HANA Then
                                GetValue = objaddon.objglobalmethods.getSingleValue("select ifnull(sum(T1.""Quantity"" * T1.""StockPrice""),0) as ""UnitPrice""  from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.string & "' group by T0.""DocEntry"" order by T0.""DocEntry"" desc")
                            Else
                                GetValue = objaddon.objglobalmethods.getSingleValue("select isnull(sum(T1.Quantity * T1.StockPrice),0) as UnitPrice  from OIGE T0 join IGE1 T1 on T0.DocEntry=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.string & "' group by T0.DocEntry order by T0.DocEntry desc")
                            End If
                            Unitprice = CDbl(CDbl(IIf(GetValue = "", 0, GetValue))) / CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                            If EditText28.Value <> "" Then
                                Dim PosDate As Date = Date.ParseExact(EditText6.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                If objaddon.HANA Then
                                    GetValue = "Select ifnull(T1.""U_Price"",0) as ""ProcessCost"",ifnull(T1.""U_Weight"",0) as ""Weight"" from ""@" & Trim(objRS.Fields.Item(0).Value.ToString) & """ T0 left join ""@" & Trim(objRS.Fields.Item(1).Value.ToString) & """ T1 on T0.""Code""=T1.""Code"" "
                                    GetValue += vbCrLf + "where T0.""Code""='" & EditText0.Value & "' and T1.""U_ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T1.""U_Process""='" & EditText28.Value & "'  " ' EditText20.Value
                                    GetValue += vbCrLf + "and T1.""U_EffDate""=(Select max(B.""U_EffDate"") from ""@" & Trim(objRS.Fields.Item(0).Value.ToString) & """ A left join ""@" & Trim(objRS.Fields.Item(1).Value.ToString) & """ B on A.""Code""=B.""Code"""
                                    GetValue += vbCrLf + "where A.""Code""=T0.""Code"" and B.""U_ItemCode""=T1.""U_ItemCode"" and B.""U_Process""=T1.""U_Process"" and B.""U_EffDate""<='" & PosDate.ToString("yyyyMMdd") & "')"
                                Else
                                    GetValue = "Select isnull(T1.U_Price,0) as ProcessCost,isnull(T1.U_Weight,0) as Weight from [@" & Trim(objRS.Fields.Item(0).Value.ToString) & "] T0 left join [@" & Trim(objRS.Fields.Item(1).Value.ToString) & "] T1 on T0.Code=T1.Code "
                                    GetValue += vbCrLf + "where T0.Code='" & EditText0.Value & "' and T1.U_ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T1.U_Process='" & EditText28.Value & "'  "
                                    GetValue += vbCrLf + "and T1.U_EffDate=(Select max(B.U_EffDate) from [@" & Trim(objRS.Fields.Item(0).Value.ToString) & "] A left join [@" & Trim(objRS.Fields.Item(1).Value.ToString) & "] B on A.Code=B.Code"
                                    GetValue += vbCrLf + "where A.Code=T0.Code and B.U_ItemCode=T1.U_ItemCode and B.U_Process=T1.U_Process and B.U_EffDate<='" & PosDate.ToString("yyyyMMdd") & "')"
                                End If
                                objRS.DoQuery(GetValue)
                                If objRS.RecordCount > 0 Then
                                    Unitprice = CDbl(Unitprice) + CDbl(objRS.Fields.Item("ProcessCost").Value.ToString)
                                    objcombo = Matrix3.Columns.Item("cmblttype").Cells.Item(i).Specific
                                    Matrix3.SetCellWithoutValidation(i, "ProCost", CDbl(objRS.Fields.Item("ProcessCost").Value.ToString))
                                    Matrix3.SetCellWithoutValidation(i, "actwt", CDbl(objRS.Fields.Item("Weight").Value.ToString))
                                    Matrix3.SetCellWithoutValidation(i, "tweight", CDbl(objRS.Fields.Item("Weight").Value.ToString) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String))

                                    'Matrix3.Columns.Item("ProCost").Cells.Item(i).Specific.String = CDbl(objRS.Fields.Item("ProcessCost").Value.ToString)
                                    'Matrix3.Columns.Item("actwt").Cells.Item(i).Specific.String = CDbl(objRS.Fields.Item("Weight").Value.ToString)
                                    'Matrix3.Columns.Item("tweight").Cells.Item(i).Specific.String = CDbl(objRS.Fields.Item("Weight").Value.ToString) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                                    If objcombo.Selected.Value = "1" Then
                                        Matrix3.SetCellWithoutValidation(i, "TProCost", CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString))
                                        ' Matrix3.Columns.Item("TProCost").Cells.Item(i).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString)
                                    Else
                                        If Val(Matrix0.Columns.Item("tweight").Cells.Item(i).Specific.String) = 0 Then
                                            Matrix3.SetCellWithoutValidation(i, "TProCost", CDbl(objRS.Fields.Item("Weight").Value.ToString) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString))
                                            'Matrix3.Columns.Item("TProCost").Cells.Item(i).Specific.String = CDbl(objRS.Fields.Item("Weight").Value.ToString) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString)
                                        Else
                                            Matrix3.SetCellWithoutValidation(i, "TProCost", CDbl(Matrix0.Columns.Item("tweight").Cells.Item(i).Specific.String) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString))
                                            'Matrix3.Columns.Item("TProCost").Cells.Item(i).Specific.String = CDbl(Matrix0.Columns.Item("tweight").Cells.Item(i).Specific.String) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString)
                                        End If

                                    End If
                                Else
                                    Unitprice = CDbl(Unitprice)
                                End If
                            End If
                            Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String = Unitprice
                            'If objaddon.HANA Then
                            '    GetValue = "Select ifnull(T1.""U_Price"",0) as ""SplPrice"" from ""@PRICEHEAD"" T0 left join ""@PRICELine"" T1 on T0.""Code""=T1.""Code"" "
                            '    GetValue += vbCrLf + "where T0.""Code""='" & EditText0.Value & "' and T1.""U_ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T1.""U_Process""='" & EditText20.Value & "'  "
                            '    GetValue += vbCrLf + "and T1.""U_EffDate""=(Select max(B.""U_EffDate"") from ""@PRICEHEAD"" A left join ""@PRICELine"" B on A.""Code""=B.""Code"""
                            '    GetValue += vbCrLf + "where A.""Code""=T0.""Code"" and B.""U_ItemCode""=T1.""U_ItemCode"" and B.""U_Process""=T1.""U_Process"" and B.""U_EffDate""<='" & PosDate.ToString("yyyyMMdd") & "')"
                            'Else
                            '    GetValue = "Select isnull(T1.U_Price,0) as SplPrice from [@PRICEHEAD] T0 left join [@PRICELine] T1 on T0.Code=T1.Code "
                            '    GetValue += vbCrLf + "where T0.Code='" & EditText0.Value & "' and T1.U_ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T1.U_Process='" & EditText20.Value & "'  "
                            '    GetValue += vbCrLf + "and T1.U_EffDate=(Select max(B.U_EffDate) from [@PRICEHEAD] A left join [@PRICELine] B on A.Code=B.Code"
                            '    GetValue += vbCrLf + "where A.Code=T0.Code and B.U_ItemCode=T1.U_ItemCode and B.U_Process=T1.U_Process and B.U_EffDate<='" & PosDate.ToString("yyyyMMdd") & "')"
                            'End If
                            'objRS.DoQuery(GetValue)
                            'If objRS.RecordCount > 0 Then
                            '    Unitprice = CDbl(Unitprice) + CDbl(objRS.Fields.Item("SplPrice").Value.ToString)
                            '    Matrix3.Columns.Item("ProCost").Cells.Item(i).Specific.String = CDbl(objRS.Fields.Item("SplPrice").Value.ToString)
                            '    Matrix3.Columns.Item("TProCost").Cells.Item(i).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String) * CDbl(objRS.Fields.Item("SplPrice").Value.ToString)
                            'Else
                            '    Unitprice = CDbl(Unitprice)
                            'End If
                            'Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String = Unitprice
                            'Matrix3.Columns.Item("Linetot").Cells.Item(i).Specific.String = CDbl(Unitprice) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                        End If
                    Next
                Else
                    For i = 1 To Matrix3.VisualRowCount
                        If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix3.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" And Matrix3.Columns.Item("GRNo").Cells.Item(i).Specific.String = "" Then
                            Dim ReceiptDate As Date = Date.ParseExact(Matrix3.Columns.Item("Date").Cells.Item(i).Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            'JEAmount = objaddon.objglobalmethods.getSingleValue("select  ifnull(Sum(T1.""Debit""),0)  from OJDT T0 join JDT1 T1 on T0.""TransId""=T1.""TransId"" where T0.""U_SubConNo""='" & objform.Items.Item("txtdocnum").Specific.string & "' and T0.""RefDate""='" & ReceiptDate.ToString("yyyyMMdd") & "' and T0.""U_Status""='O' ")
                            If objaddon.HANA Then
                                JEAmount = objaddon.objglobalmethods.getSingleValue("select  ifnull(Sum(T1.""Debit""),0)  from OJDT T0 join JDT1 T1 on T0.""TransId""=T1.""TransId"" join ""@MIPL_OPOR"" T2 on T2.""DocNum""=T0.""U_SubConNo"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.string & "'  and T0.""RefDate""='" & ReceiptDate.ToString("yyyyMMdd") & "' and T0.""U_Status""='O'")
                                GetValue = objaddon.objglobalmethods.getSingleValue("select ifnull(sum(T1.""Quantity"" * T1.""StockPrice""),0) as ""UnitPrice""  from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.string & "' group by T0.""DocEntry"" order by T0.""DocEntry"" desc")
                            Else
                                JEAmount = objaddon.objglobalmethods.getSingleValue("select  isnull(Sum(T1.Debit),0)  from OJDT T0 join JDT1 T1 on T0.TransId=T1.TransId join [@MIPL_OPOR] T2 on T2.DocNum=T0.U_SubConNo where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.string & "'  and T0.RefDate='" & ReceiptDate.ToString("yyyyMMdd") & "' and T0.U_Status='O'")
                                GetValue = objaddon.objglobalmethods.getSingleValue("select isnull(sum(T1.Quantity * T1.StockPrice),0) as UnitPrice  from OIGE T0 join IGE1 T1 on T0.DocEntry=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.string & "' group by T0.DocEntry order by T0.DocEntry desc")
                            End If
                            Unitprice = CDbl(CDbl(IIf(GetValue = "", 0, GetValue)) + CDbl(JEAmount)) / CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                            Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String = Unitprice
                            Matrix3.Columns.Item("Linetot").Cells.Item(i).Specific.String = CDbl(Unitprice) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                            Count += 1
                        End If
                    Next
                    If Count > 1 Then
                        For i = 1 To Matrix3.VisualRowCount
                            If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix3.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" And Matrix3.Columns.Item("GRNo").Cells.Item(i).Specific.String = "" Then
                                If CheckInputItem(Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String) Then
                                    If objaddon.HANA Then
                                        ByProductPrice = objaddon.objglobalmethods.getSingleValue("select ifnull( T1.""StockPrice"",0) as ""UnitPrice"" from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where  T0.""U_SubConNo""='" & objform.Items.Item("txtdocnum").Specific.string & "' and T1.""ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "'")
                                    Else
                                        ByProductPrice = objaddon.objglobalmethods.getSingleValue("select isnull( T1.StockPrice,0) as UnitPrice from OIGE T0 join IGE1 T1 on T0.DocEntry=T1.DocEntry where  T0.U_SubConNo='" & objform.Items.Item("txtdocnum").Specific.string & "' and T1.ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "'")
                                    End If
                                    Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String = ByProductPrice
                                    BYPPriceTotal = CDbl(ByProductPrice) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                                    Matrix3.Columns.Item("Linetot").Cells.Item(i).Specific.String = CStr(BYPPriceTotal)
                                End If
                                If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String = objform.Items.Item("txtsitem").Specific.string Then
                                    Dim ReceiptDate As Date = Date.ParseExact(Matrix3.Columns.Item("Date").Cells.Item(i).Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                    If objaddon.HANA Then
                                        JEAmount = objaddon.objglobalmethods.getSingleValue("select  ifnull(Sum(T1.""Debit""),0)  from OJDT T0 join JDT1 T1 on T0.""TransId""=T1.""TransId"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.string & "' and T0.""RefDate""='" & ReceiptDate.ToString("yyyyMMdd") & "' and T0.""U_Status""='O' ")
                                        GetValue = objaddon.objglobalmethods.getSingleValue("select ifnull(sum(T1.""Quantity"" * T1.""StockPrice""),0) as ""UnitPrice""  from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.string & "' group by T0.""DocEntry"" order by T0.""DocEntry"" desc")
                                    Else
                                        JEAmount = objaddon.objglobalmethods.getSingleValue("select  isnull(Sum(T1.Debit),0)  from OJDT T0 join JDT1 T1 on T0.TransId=T1.TransId where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.string & "' and T0.RefDate='" & ReceiptDate.ToString("yyyyMMdd") & "' and T0.U_Status='O' ")
                                        GetValue = objaddon.objglobalmethods.getSingleValue("select isnull(sum(T1.Quantity * T1.StockPrice),0) as UnitPrice  from OIGE T0 join IGE1 T1 on T0.DocEntry=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.string & "' group by T0.DocEntry order by T0.DocEntry desc")
                                    End If
                                    Unitprice = CDbl((CDbl(IIf(GetValue = "", 0, GetValue)) + CDbl(JEAmount)) - BYPPriceTotal) / CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                                    Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String = Unitprice
                                    Matrix3.Columns.Item("Linetot").Cells.Item(i).Specific.String = CDbl(Unitprice) * CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                                End If
                            End If
                        Next
                    Else
                        Exit Sub
                    End If
                End If

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

        End Sub

        Private Sub CreateGoodsReceipt(ByVal TabType As String)
            Try
                'Posting GR in Scrap Tab
                Dim objMatrix As SAPbouiCOM.Matrix
                Dim objGRform As SAPbouiCOM.Form
                Dim objrecset As SAPbobsCOM.Recordset
                Dim Lineflag As Boolean = False
                Dim WhsCode As String = ""
                objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                For i As Integer = 1 To Matrix2.VisualRowCount
                    objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                    If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And objcombo.Selected.Description = "Scrap" And Matrix2.Columns.Item("GRNo").Cells.Item(i).Specific.string = "" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" Then
                        Lineflag = True
                        WhsCode = Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string
                    End If
                Next
                If Lineflag = True Then
                    objaddon.objapplication.Menus.Item("3078").Activate()
                    objGRform = objaddon.objapplication.Forms.ActiveForm
                    objGRform = objaddon.objapplication.Forms.Item(objGRform.UniqueID)
                    objGRform.Visible = True
                    objMatrix = objGRform.Items.Item("13").Specific
                    objaddon.objapplication.StatusBar.SetText("Data Loading to Goods Receipt Screen Please wait ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    'Dim SubPOHeader As SAPbouiCOM.DBDataSource = objform.DataSources.DBDataSources.Item("@MIPL_OPOR")
                    Dim AcctCode As String = ""
                    Try
                        objform.Freeze(True)
                        objGRform.Freeze(True)
                        SubConTypeCount = objform.TypeCount
                        objGRform.Items.Item("11").Specific.String = "From Sub-Con Add-on" & Now.ToString
                        ' SubPOHeader.GetValue("DocEntry", 0) 'CStr(objform.Items.Item("txtdocnum").Specific.String)
                        Dim oUDFForm As SAPbouiCOM.Form
                        oUDFForm = objaddon.objapplication.Forms.Item(objGRform.UDFFormUID)
                        If oUDFForm.Items.Item("U_ScrapNum").Enabled = False Then oUDFForm.Items.Item("U_ScrapNum").Enabled = True
                        oUDFForm.Items.Item("U_ScrapNum").Specific.String = objform.Items.Item("txtentry").Specific.String
                        objGRform.Items.Item("U_SubConNo").Specific.String = objform.Items.Item("txtentry").Specific.String
                        objGRform.ActiveItem = "9"
                        oUDFForm.Items.Item("U_ScrapNum").Enabled = False
                        If BranchEnabled = "Y" Then
                            Dim cmbSeries As SAPbouiCOM.ComboBox
                            cmbSeries = objGRform.Items.Item("30").Specific
                            Dim series As String
                            If objaddon.HANA Then
                                series = objaddon.objglobalmethods.getSingleValue("select Top 1 ""Series"" From NNM1 where ""ObjectCode""='59' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                        " and ""BPLId""=(select Top 1 ""BPLid"" from OWHS where ""WhsCode""='" & WhsCode & "') order by ""CreateDate"" desc ")
                            Else
                                series = objaddon.objglobalmethods.getSingleValue("select Top 1 Series From NNM1 where ObjectCode='59' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                                                                                        " and BPLId=(select Top 1 BPLid from OWHS where WhsCode='" & WhsCode & "') order by CreateDate desc")
                            End If

                            If series <> "" Then
                                cmbSeries.Select(series, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        End If
                        If TabType = "Scrap" Then
                            Dim InvRow As Integer = 0
                            If objMatrix.Columns.Item("U_PlanQty").Editable = False Or objMatrix.Columns.Item("U_TabType").Editable = False Or objMatrix.Columns.Item("U_CardCode").Editable = False Then
                                objMatrix.Columns.Item("U_PlanQty").Editable = True
                                objMatrix.Columns.Item("U_TabType").Editable = True
                                objMatrix.Columns.Item("U_CardCode").Editable = True
                            End If
                            For i As Integer = 1 To Matrix2.VisualRowCount
                                objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                                If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And objcombo.Selected.Description = "Scrap" And Matrix2.Columns.Item("GRNo").Cells.Item(i).Specific.string = "" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" Then
                                    InvRow += 1
                                    objMatrix.Columns.Item("1").Cells.Item(InvRow).Specific.String = Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String
                                    'objMatrix.Columns.Item("2").Cells.Item(InvRow).Specific.String = Matrix2.Columns.Item("Desc").Cells.Item(i).Specific.String
                                    objMatrix.Columns.Item("9").Cells.Item(InvRow).Specific.String = Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.String
                                    objMatrix.Columns.Item("U_PlanQty").Cells.Item(InvRow).Specific.String = Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.String   'U_TabType
                                    objMatrix.Columns.Item("U_TabType").Cells.Item(InvRow).Specific.String = TabType 'Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.String
                                    objMatrix.Columns.Item("10").Cells.Item(InvRow).Specific.String = Matrix2.Columns.Item("Price").Cells.Item(i).Specific.String
                                    objMatrix.Columns.Item("U_CardCode").Cells.Item(InvRow).Specific.String = objform.Items.Item("txtcode").Specific.String
                                    objMatrix.Columns.Item("14").Cells.Item(InvRow).Specific.String = Matrix2.Columns.Item("Linetot").Cells.Item(i).Specific.String
                                    objMatrix.Columns.Item("15").Cells.Item(InvRow).Specific.String = Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.String
                                    If ItemProcess = "Y" Then
                                        If Trim(EditText20.Value) <> "" Then
                                            objMatrix.Columns.Item("U_Process").Cells.Item(InvRow).Specific.String = Trim(EditText20.Value)
                                        End If
                                    End If
                                    ' AcctCode = objaddon.objglobalmethods.getSingleValue(" select  ""U_GoodsReceipt"" from ""@MIPL_SBGL"" where ""U_ItemCode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String & "' or ""U_WhsCode""='" & Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")
                                    If objaddon.HANA Then
                                        AcctCode = objaddon.objglobalmethods.getSingleValue("select ""U_GRCodeS"" from ""@MIPL_GEN""")
                                    Else
                                        AcctCode = objaddon.objglobalmethods.getSingleValue("select U_GRCodeS from [@MIPL_GEN]")
                                    End If

                                    If AcctCode <> "" Then
                                        objMatrix.Columns.Item("59").Cells.Item(InvRow).Specific.String = AcctCode
                                    Else
                                        objaddon.objapplication.StatusBar.SetText("Please update the Accountcode in General Settings...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objform.Freeze(False)
                                        Exit Sub
                                    End If
                                    'objMatrix.Columns.Item("5").Cells.Item(i).Specific.String = Matrix3.Columns.Item("SubWhse").Cells.Item(i).Specific.String
                                End If
                            Next
                            objMatrix.Columns.Item("1").Cells.Item(1).Click()
                        End If

                        If objMatrix.VisualRowCount = 1 Then
                            If objMatrix.Columns.Item("1").Cells.Item(1).Specific.String = "" Then
                                objaddon.objapplication.StatusBar.SetText("No More Quantity to post the Goods Receipt ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                        objMatrix.Columns.Item("U_PlanQty").Editable = False
                        objMatrix.Columns.Item("U_TabType").Editable = False
                        objMatrix.Columns.Item("U_CardCode").Editable = False
                        objrecset = Nothing
                        objaddon.objapplication.StatusBar.SetText("Data Loaded to Goods Receipt Screen ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Finally
                        objform.Freeze(False)
                        objGRform.Freeze(False)
                    End Try
                Else
                    objaddon.objapplication.SetStatusBarMessage("No more Data for posting the Goods Receipt ...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Exit Sub
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub CalcPlanQty()
            Try
                Dim BaseQty, PlanQty, Price As Double
                If Val(EditText25.Value) > 0 Then
                    objform.Freeze(True)
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        'If CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(i).Specific.String) = 0 Then
                        If Val(Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String) <> 0 Then
                            BaseQty = Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String
                            Matrix0.Columns.Item("PlanQty").Cells.Item(i).Specific.String = CDbl(BaseQty * EditText25.Value)
                            Matrix0.Columns.Item("actplanqty").Cells.Item(i).Specific.String = CDbl(BaseQty * EditText25.Value)
                        End If
                        PlanQty = CDbl(Matrix0.Columns.Item("PlanQty").Cells.Item(i).Specific.String)
                        Price = CDbl(Matrix0.Columns.Item("Price").Cells.Item(i).Specific.String)
                        If Val(Matrix0.Columns.Item("Price").Cells.Item(i).Specific.String) > 0 Then
                            Matrix0.Columns.Item("Linetot").Cells.Item(i).Specific.String = PlanQty * Price
                        End If
                        'If WPrice = "Y" Then
                        '    If Val(Matrix0.Columns.Item("tweight").Cells.Item(i).Specific.String) > 0 Then
                        '        Matrix0.Columns.Item("Linetot").Cells.Item(i).Specific.String = CDbl(Matrix0.Columns.Item("tweight").Cells.Item(i).Specific.String) * CDbl(Matrix0.Columns.Item("Price").Cells.Item(i).Specific.String)
                        '    Else
                        '        Matrix0.Columns.Item("Linetot").Cells.Item(i).Specific.String = PlanQty * Price
                        '    End If
                        'Else
                        '    If Val(Matrix0.Columns.Item("Price").Cells.Item(i).Specific.String) > 0 Then
                        '        Matrix0.Columns.Item("Linetot").Cells.Item(i).Specific.String = PlanQty * Price
                        '    End If
                        'End If
                        'End If
                    Next
                    'objform.Update()
                End If
                'Matrix0.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("PlanQty").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objform.Update()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Matrix_Total()
            Try
                'If Folder0.Selected = True Then    'Input
                'Matrix0.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("PlanQty").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                Matrix0.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'ElseIf Folder1.Selected = True Then 'Output
                'Matrix3.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix3.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                Matrix3.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix3.Columns.Item("ProCost").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                Matrix3.Columns.Item("TProCost").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'ElseIf Folder2.Selected = True Then 'Scrap
                'Matrix2.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix2.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                Matrix2.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'ElseIf Folder3.Selected = True Then 'Related Doc
                'Matrix1.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                Matrix1.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'ElseIf Folder4.Selected = True Then 'Costing
                Matrix4.Columns.Item("Debit").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                Matrix4.Columns.Item("Credit").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub CreateMySimpleForm(ByVal FormID As String, ByVal FormTitle As String, ByVal Header As String, ByVal Line As String, ByVal LinkedID As String, ByVal InvTransfer As String)
            Dim oCreationParams As SAPbouiCOM.FormCreationParams
            Dim objTempForm As SAPbouiCOM.Form
            Dim objrs As SAPbobsCOM.Recordset
            Try
                objaddon.objapplication.Forms.Item(FormID).Visible = True
            Catch ex As Exception
                oCreationParams = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
                oCreationParams.UniqueID = FormID
                objTempForm = objaddon.objapplication.Forms.AddEx(oCreationParams)
                objTempForm.Title = FormTitle
                objTempForm.Left = 400
                objTempForm.Top = 100
                objTempForm.ClientHeight = 200 '335
                objTempForm.ClientWidth = 500
                objTempForm.Left = objform.Left + 100
                objTempForm.Top = objform.Top + 100
                objTempForm = objaddon.objapplication.Forms.Item(FormID)
                Dim oitm As SAPbouiCOM.Item

                Dim oGrid As SAPbouiCOM.Grid
                oitm = objTempForm.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
                oitm.Top = 30
                oitm.Left = 2
                oitm.Width = 490
                oitm.Height = 120
                oGrid = objTempForm.Items.Item("Grid").Specific
                objTempForm.DataSources.DataTables.Add("DataTable")
                oitm = objTempForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oitm.Top = objTempForm.Items.Item("Grid").Top + objTempForm.Items.Item("Grid").Height + 10
                oitm.Left = 2
                Dim str_sql As String = ""
                If EditText26.Value = "" Then objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Sub
                'Or EditText11.Value = ""
                'If AutoProduction = "Y" Then
                '    If objaddon.HANA Then
                '        str_sql = "select Distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from " & Header & " T0 join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""BaseType""=202  and T1.""BaseEntry""='" & EditText11.Value & "';"
                '    Else
                '        str_sql = "select Distinct T0.DocEntry,T0.DocNum,T0.DocDate from " & Header & " T0 join " & Line & " T1 on T0.DocEntry=T1.DocEntry where T1.BaseType=202  and T1.BaseEntry='" & EditText11.Value & "'"
                '    End If
                'End If
                If AutoProduction <> "Y" Then
                    If objaddon.HANA Then
                        'str_sql = "select ""DocEntry"",""DocNum"",""DocDate"" from " & Header & " where ""U_SubConNo""='" & EditText26.Value & "';"
                        str_sql = "select T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from " & Header & " T0 left join ""@MIPL_OPOR"" T1 on T0.""U_SubConNo""=T1.""DocEntry"" where T0.""U_SubConNo""='" & EditText26.Value & "';"
                    Else
                        'str_sql = "select DocEntry,DocNum,DocDate from " & Header & " where U_SubConNo='" & EditText26.Value & "';"
                        str_sql = "select T0.DocEntry,T0.DocNum,T0.DocDate from " & Header & " T0 left join [@MIPL_OPOR] T1 on T0.U_SubConNo=T1.DocEntry where T0.U_SubConNo='" & EditText26.Value & "'"
                    End If
                End If
                If InvTransfer = "Y" Then
                    If objaddon.HANA Then
                        str_sql = "select T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"",T1.""Quantity"" from OWTR T0 left join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join  ""@MIPL_OPOR"" T2 on (T2.""DocEntry""=T0.""U_SubConNo"" or T2.""DocEntry""=T1.""U_S_DocEntry"") where T2.""DocEntry""='" & EditText26.Value & "' "
                    Else
                        str_sql = "select T0.DocEntry,T0.DocNum,T0.DocDate,T1.Quantity from OWTR T0 left join WTR1 T1 on T0.DocEntry=T1.DocEntry left join  [@MIPL_OPOR] T2 on (T2.DocEntry=T0.U_SubConNo or T2.DocEntry=T1.U_S_DocEntry) where T2.DocEntry='" & EditText26.Value & "' "
                    End If
                End If

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(str_sql)
                If objrs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objrs = Nothing : Exit Sub
                Dim objDT As SAPbouiCOM.DataTable
                objDT = objTempForm.DataSources.DataTables.Item("DataTable")
                objDT.Clear()
                objDT.ExecuteQuery(str_sql)
                objTempForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(str_sql)

                oGrid.DataTable = objTempForm.DataSources.DataTables.Item("DataTable")

                For i As Integer = 0 To oGrid.Columns.Count - 1
                    oGrid.Columns.Item(i).TitleObject.Sortable = True
                    oGrid.Columns.Item(i).Editable = False
                Next

                oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                Dim col As SAPbouiCOM.EditTextColumn
                col = oGrid.Columns.Item(0)
                col.LinkedObjectType = LinkedID
                objTempForm.Visible = True
                objTempForm.Update()
                bModal = True
                FormName = "ViewD"
            End Try
        End Sub

        Private Sub JournalEntry()
            Try
                Dim DocEntry, Branch As String
                Dim Credit, Debit As Double
                Dim objrecset As SAPbobsCOM.Recordset
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                For j = 1 To Matrix4.VisualRowCount
                    If Matrix4.Columns.Item("Stat").Cells.Item(j).Specific.String = "O" Then
                        Credit += CDbl(Matrix4.Columns.Item("Credit").Cells.Item(j).Specific.String)
                        Debit += CDbl(Matrix4.Columns.Item("Debit").Cells.Item(j).Specific.String)
                    End If
                Next
                If Credit <> Debit Then objaddon.objapplication.SetStatusBarMessage("UnBalanced Credit & Debit Amount. Please Check", SAPbouiCOM.BoMessageTime.bmt_Medium, True) : Exit Sub
                If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                'Dim Series As String = ""
                'Series = objGM.GetSeries("30", CDate(dtst1.Tables(Header).Rows(0)("RefDate")).ToString("yyyy-MM-dd"), Branch)
                'objjournalentry.Series = Series
                Dim DocDate As Date = Date.ParseExact(Matrix4.Columns.Item("Date").Cells.Item(Matrix4.VisualRowCount - 1).Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objjournalentry.ReferenceDate = DocDate 'Matrix4.Columns.Item("Date").Cells.Item(Matrix4.VisualRowCount - 1).Specific.String
                objjournalentry.DueDate = DocDate ' Matrix4.Columns.Item("Date").Cells.Item(Matrix4.VisualRowCount - 1).Specific.String
                objjournalentry.TaxDate = DocDate ' Matrix4.Columns.Item("Date").Cells.Item(Matrix4.VisualRowCount - 1).Specific.String
                objjournalentry.UserFields.Fields.Item("U_SubConNo").Value = objform.Items.Item("txtentry").Specific.string
                objjournalentry.UserFields.Fields.Item("U_Status").Value = "O"
                objjournalentry.Reference = "Auto Posted from sub-con add-on"
                'objjournalentry.Reference2 = ""
                objjournalentry.Memo = "SubContracting Costing JE"
                For j = 1 To Matrix4.VisualRowCount
                    If Matrix4.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And Matrix4.Columns.Item("Stat").Cells.Item(j).Specific.String = "O" Then
                        objjournalentry.Lines.AccountCode = Matrix4.Columns.Item("Code").Cells.Item(j).Specific.String
                        'objjournalentry.Lines.ContraAccount = dtst1.Tables("Line").Rows(j)("ContCode").ToString
                        If CDbl(Matrix4.Columns.Item("Credit").Cells.Item(j).Specific.String) <> 0 Then
                            objjournalentry.Lines.Credit = CDbl(Matrix4.Columns.Item("Credit").Cells.Item(j).Specific.String)
                            objjournalentry.Lines.Debit = 0
                        Else
                            objjournalentry.Lines.Debit = CDbl(Matrix4.Columns.Item("Debit").Cells.Item(j).Specific.String)
                            objjournalentry.Lines.Credit = 0
                        End If
                        If Matrix4.Columns.Item("CostCent").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode = Matrix4.Columns.Item("CostCent").Cells.Item(j).Specific.String
                        If Matrix4.Columns.Item("CostCent1").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode2 = Matrix4.Columns.Item("CostCent1").Cells.Item(j).Specific.String
                        If Matrix4.Columns.Item("CostCent2").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode3 = Matrix4.Columns.Item("CostCent2").Cells.Item(j).Specific.String
                        If Matrix4.Columns.Item("CostCent3").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode4 = Matrix4.Columns.Item("CostCent3").Cells.Item(j).Specific.String
                        If Matrix4.Columns.Item("CostCent4").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode5 = Matrix4.Columns.Item("CostCent4").Cells.Item(j).Specific.String
                        If Matrix4.Columns.Item("Proj").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.ProjectCode = Matrix4.Columns.Item("Proj").Cells.Item(j).Specific.String
                        If Matrix4.Columns.Item("Remarks").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.Reference1 = Matrix4.Columns.Item("Remarks").Cells.Item(j).Specific.String
                        If Matrix4.Columns.Item("Branch").Cells.Item(j).Specific.String <> "" Then
                            If BranchEnabled = "Y" Then
                                If objaddon.HANA Then
                                    Branch = objaddon.objglobalmethods.getSingleValue("select ""BPLId"" from OBPL where ""BPLName""='" & Matrix4.Columns.Item("Branch").Cells.Item(j).Specific.String & "'")
                                Else
                                    Branch = objaddon.objglobalmethods.getSingleValue("select BPLId from OBPL where BPLName='" & Matrix4.Columns.Item("Branch").Cells.Item(j).Specific.String & "'")
                                End If
                                objjournalentry.Lines.BPLID = Branch
                            End If
                        End If
                        'objjournalentry.Lines.LocationCode = ""
                        objjournalentry.Lines.Add()
                    End If
                Next

                If objjournalentry.Add <> 0 Then
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("Journal:" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Else
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    DocEntry = objaddon.objcompany.GetNewObjectKey()
                    objaddon.objapplication.SetStatusBarMessage("Journal Entry added Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    For j = 1 To Matrix4.VisualRowCount
                        If Matrix4.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And Matrix4.Columns.Item("Stat").Cells.Item(j).Specific.String = "O" Then
                            Matrix4.Columns.Item("JENum").Cells.Item(j).Specific.String = DocEntry
                            'Matrix4.Columns.Item("Stat").Cells.Item(j).Specific.String = "C"
                            Matrix4.CommonSetting.SetRowEditable(j, False)
                        End If
                    Next
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                objrecset = Nothing
            Catch ex As Exception
                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage("JE Posting Error" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub GoodsReceiptNew()
            Try
                'Posting GR in Output Tab
                Dim objMatrix As SAPbouiCOM.Matrix
                Dim objGRform As SAPbouiCOM.Form
                'Dim objrecset As SAPbobsCOM.Recordset
                Dim AcctCode As String = ""
                Dim InvRow As Integer = 0
                objaddon.objapplication.Menus.Item("3078").Activate()
                objGRform = objaddon.objapplication.Forms.ActiveForm
                objGRform = objaddon.objapplication.Forms.Item(objGRform.UniqueID)
                objGRform.Visible = True
                objMatrix = objGRform.Items.Item("13").Specific
                objaddon.objapplication.StatusBar.SetText("Data Loading to Goods Receipt Screen Please wait ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Try
                    objGRform.Freeze(True)
                    SubConTypeCount = objform.TypeCount
                    objGRform.Items.Item("11").Specific.String = "From Sub-Con Add-on" & Now.ToString
                    odbdsHeader = objform.DataSources.DBDataSources.Item("@MIPL_OPOR")
                    objGRform.Items.Item("U_SubConNo").Specific.String = objform.Items.Item("txtentry").Specific.String 'SubPOHeader.GetValue("DocEntry", 0) 'CStr(objform.Items.Item("txtdocnum").Specific.String)
                    objGRform.Items.Item("21").Specific.String = "Output"
                    Dim oUDFForm As SAPbouiCOM.Form
                    oUDFForm = objaddon.objapplication.Forms.Item(objGRform.UDFFormUID)
                    'oUDFForm.Items.Item("U_SubInv").Specific.String = "-"
                    oUDFForm.Items.Item("U_SubConNo").Enabled = False
                    If objMatrix.Columns.Item("U_LineID").Editable = False Then objMatrix.Columns.Item("U_LineID").Editable = True
                    If objMatrix.Columns.Item("U_PlanQty").Editable = False Then objMatrix.Columns.Item("U_PlanQty").Editable = True
                    If objMatrix.Columns.Item("U_CardCode").Editable = False Then objMatrix.Columns.Item("U_CardCode").Editable = True
                    If objMatrix.Columns.Item("U_Price").Editable = False Then objMatrix.Columns.Item("U_Price").Editable = True
                    'If objMatrix.Columns.Item("U_PlanQty").Editable = False Or objMatrix.Columns.Item("U_LineID").Editable = False Or objMatrix.Columns.Item("U_CardCode").Editable = False Or objMatrix.Columns.Item("U_Price").Editable = False Then
                    '    objMatrix.Columns.Item("U_PlanQty").Editable = True
                    '    objMatrix.Columns.Item("U_LineID").Editable = True
                    '    objMatrix.Columns.Item("U_CardCode").Editable = True
                    '    objMatrix.Columns.Item("U_Price").Editable = True
                    'End If
                    If objaddon.HANA Then
                        AcctCode = objaddon.objglobalmethods.getSingleValue(" select ""U_GRCodeO"" from ""@MIPL_GEN""")
                    Else
                        AcctCode = objaddon.objglobalmethods.getSingleValue(" select U_GRCodeO from [@MIPL_GEN]")
                    End If
                    For i As Integer = 1 To Matrix3.VisualRowCount
                        If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix3.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" And Matrix3.Columns.Item("GRNo").Cells.Item(i).Specific.String = "" Then
                            InvRow += 1
                            objMatrix.Columns.Item("1").Cells.Item(InvRow).Specific.String = Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String
                            'objMatrix.Columns.Item("2").Cells.Item(InvRow).Specific.String = Matrix3.Columns.Item("Desc").Cells.Item(i).Specific.String
                            objMatrix.Columns.Item("9").Cells.Item(InvRow).Specific.String = Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String
                            objMatrix.Columns.Item("U_PlanQty").Cells.Item(InvRow).Specific.String = Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String   'U_TabType
                            objMatrix.Columns.Item("U_LineID").Cells.Item(InvRow).Specific.String = Matrix3.Columns.Item("#").Cells.Item(i).Specific.String
                            objMatrix.Columns.Item("U_Price").Cells.Item(InvRow).Specific.String = Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String
                            objMatrix.Columns.Item("10").Cells.Item(InvRow).Specific.String = Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String
                            objMatrix.Columns.Item("U_CardCode").Cells.Item(InvRow).Specific.String = objform.Items.Item("txtcode").Specific.String
                            objMatrix.Columns.Item("15").Cells.Item(InvRow).Specific.String = Matrix3.Columns.Item("Whse").Cells.Item(i).Specific.String
                            objMatrix.Columns.Item("14").Cells.Item(InvRow).Specific.String = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String) * CDbl(Matrix3.Columns.Item("Price").Cells.Item(i).Specific.String)
                            If ItemProcess = "Y" Then
                                If Trim(EditText20.Value) <> "" Then
                                    objMatrix.Columns.Item("U_Process").Cells.Item(InvRow).Specific.String = Trim(EditText20.Value)
                                End If
                            End If
                            If BranchEnabled = "Y" Then
                                Dim cmbSeries As SAPbouiCOM.ComboBox
                                cmbSeries = objGRform.Items.Item("30").Specific
                                Dim series As String
                                If objaddon.HANA Then
                                    series = objaddon.objglobalmethods.getSingleValue("select Top 1""Series"" From NNM1 where ""ObjectCode""='59' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                               " and ""BPLId""=(select Top 1 ""BPLid"" from OWHS where ""WhsCode""='" & Matrix3.Columns.Item("Whse").Cells.Item(i).Specific.String & "') order by ""CreateDate"" desc")
                                Else
                                    series = objaddon.objglobalmethods.getSingleValue("select Top 1 Series From NNM1 where ObjectCode='59' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                                                                                               " and BPLId=(select Top 1 BPLid from OWHS where WhsCode='" & Matrix3.Columns.Item("Whse").Cells.Item(i).Specific.String & "')") ' order by CreateDate desc ")
                                End If
                                If series <> "" Then
                                    cmbSeries.Select(series, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If
                            ' AcctCode = objaddon.objglobalmethods.getSingleValue(" select  ""U_GoodsReceipt"" from ""@MIPL_SBGL"" where ""U_ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(i).Specific.String & "' or ""U_WhsCode""='" & MatrixID.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")

                            If AcctCode <> "" Then
                                objMatrix.Columns.Item("59").Cells.Item(InvRow).Specific.String = AcctCode
                            Else
                                objaddon.objapplication.StatusBar.SetText("Please update the Accountcode in General Settings...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objform.Freeze(False)
                                Exit Sub
                            End If
                            'objMatrix.Columns.Item("5").Cells.Item(i).Specific.String = Matrix3.Columns.Item("SubWhse").Cells.Item(i).Specific.String
                        End If
                    Next
                    objMatrix.Columns.Item("1").Cells.Item(1).Click()
                    'objMatrix.Columns.Item("U_PlanQty").Editable = False
                    'objMatrix.Columns.Item("U_LineID").Editable = False
                    'objMatrix.Columns.Item("U_CardCode").Editable = False
                    'objMatrix.Columns.Item("U_Price").Editable = False
                    'objrecset = Nothing
                    objaddon.objapplication.StatusBar.SetText("Data Loaded to Goods Receipt Screen ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Catch ex As Exception
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Finally
                    objGRform.Freeze(False)
                    'objGRform.Close()
                    objform.Freeze(False)
                End Try
            Catch ex As Exception

            End Try
        End Sub

        Private Sub PostingGoodsIssue()
            'Posting GI in Output Tab (Getting data from input tab & also posting GR)
            Try
                Dim objGoodsIssue As SAPbobsCOM.Documents
                Dim Quantity, OutputQuantity As Double
                Dim Retval As Integer
                Dim objrs, objrecset As SAPbobsCOM.Recordset
                Dim StrSql As String = "", GIStatus As String = "", GRStatus As String = "", GIStatus1 As String = "", GRStatus1 As String = ""
                Dim Branch As String = "", Docentry As String = "", Batchs As String = "", Serial As String = ""
                Dim AcctCode As String = "", WhsCode As String

                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
                objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim txtDate As SAPbouiCOM.EditText
                txtDate = Matrix3.Columns.Item("Date").Cells.Item(Matrix3.VisualRowCount).Specific
                Matrix3.Columns.Item("Whse").Cells.Item(Matrix3.VisualRowCount).Click()
                Dim DocDate As Date = Date.ParseExact(txtDate.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'Dim DocDueDate As Date = Date.ParseExact(Matrix3.Columns.Item("Date").Cells.Item(Matrix3.VisualRowCount).Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'Dim TaxDate As Date = Date.ParseExact(Matrix3.Columns.Item("Date").Cells.Item(Matrix3.VisualRowCount).Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Dim ChkGRStat As String
                'ChkGRStat = objaddon.objglobalmethods.getSingleValue("Call ""MIPL_ValidateGRQty"" ('@MIPL_POR2'," & objform.Items.Item("txtentry").Specific.String & ",'" & EditText8.Value & "')")
                If objaddon.HANA Then
                    ChkGRStat = objaddon.objglobalmethods.getSingleValue("Select  Distinct T1.""U_SQty""-T4.""Qty"" as ""PendQty""from  ""@MIPL_OPOR"" T1 inner join ""@MIPL_POR2"" T0 on T0.""DocEntry""=T1.""DocEntry"" and T0.""U_Itemcode""=T1.""U_SItemCode"" inner join " &
                                                                       "(select T2.""U_SubConNo"",sum(T3.""Quantity"") as ""Qty"",T3.""ItemCode"" from  OIGN T2 join IGN1 T3 on T2.""DocEntry""=T3.""DocEntry"" group by T2.""U_SubConNo"",T3.""ItemCode"" ) as T4 " &
                                                                      "  on T4.""U_SubConNo""=T1.""DocEntry"" and T0.""U_Itemcode""=T4.""ItemCode"" where T1.""U_SItemCode""='" & EditText8.Value & "'  and T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " ")
                Else
                    ChkGRStat = objaddon.objglobalmethods.getSingleValue("Select  Distinct T1.U_SQty-T4.Qty as PendQty from  [@MIPL_OPOR] T1 inner join [@MIPL_POR2] T0 on T0.DocEntry=T1.DocEntry and T0.U_Itemcode=T1.U_SItemCode inner join " &
                                                                         "(select T2.U_SubConNo,sum(T3.Quantity) as Qty,T3.ItemCode from  OIGN T2 join IGN1 T3 on T2.DocEntry=T3.DocEntry group by T2.U_SubConNo,T3.ItemCode ) as T4 " &
                                                                        "  on T4.U_SubConNo=T1.DocEntry and T0.U_Itemcode=T4.ItemCode where T1.U_SItemCode='" & EditText8.Value & "'  and T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " ")
                End If
                If ChkGRStat = "0" Then
                    objaddon.objapplication.StatusBar.SetText(" Goods Receipt Created for the Planned Quantity...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                If objaddon.HANA Then
                    GIStatus = objaddon.objglobalmethods.getSingleValue("select Distinct 1 from ""@MIPL_POR2"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                          " left join OIGE T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGE1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T0.""U_Itemcode""=T2.""ItemCode"" " &
                          " where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Date""='" & DocDate.ToString("yyyyMMdd") & "' and T0.""U_Status""='O' and ifnull(T0.""U_GINo"",'')=''  ")

                    GIStatus1 = objaddon.objglobalmethods.getSingleValue("select Distinct 1 from ""@MIPL_POR2"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                " left join OIGE T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGE1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T0.""U_Itemcode""=T2.""ItemCode"" " &
                                " where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Date""='" & DocDate.ToString("yyyyMMdd") & "' and T0.""U_Status""='O'  and ifnull(T0.""U_GINo"",'')<> ''  ")

                    GRStatus = objaddon.objglobalmethods.getSingleValue("select Distinct 1 from ""@MIPL_POR2"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                " left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T0.""U_Itemcode""=T2.""ItemCode"" " &
                                " where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Date""='" & DocDate.ToString("yyyyMMdd") & "' and T0.""U_Status""='O'   and ifnull(T0.""U_GRNo"",'')=''  ")

                    GRStatus1 = objaddon.objglobalmethods.getSingleValue("select Distinct 1 from ""@MIPL_POR2"" T0  left join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                " left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T0.""U_Itemcode""=T2.""ItemCode"" " &
                                " where T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T0.""U_Date""='" & DocDate.ToString("yyyyMMdd") & "' and T0.""U_Status""='C'  and ifnull(T0.""U_GRNo"",'')<> ''  ")
                Else
                    GIStatus = objaddon.objglobalmethods.getSingleValue("select Distinct 1 from [@MIPL_POR2] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry " &
                          " left join OIGE T3 on T3.U_SubConNo=T1.DocEntry left join IGE1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode " &
                          " where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Date='" & DocDate.ToString("yyyyMMdd") & "' and T0.U_Status='O' and isnull(T0.U_GINo,'')=''  ")

                    GIStatus1 = objaddon.objglobalmethods.getSingleValue("select Distinct 1 from [@MIPL_POR2] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry " &
                                " left join OIGE T3 on T3.U_SubConNo=T1.DocEntry left join IGE1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode " &
                                " where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Date='" & DocDate.ToString("yyyyMMdd") & "' and T0.U_Status='O'  and isnull(T0.U_GINo,'')<> ''  ")

                    GRStatus = objaddon.objglobalmethods.getSingleValue("select Distinct 1 from [@MIPL_POR2] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry " &
                                " left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode " &
                                " where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Date='" & DocDate.ToString("yyyyMMdd") & "' and T0.U_Status='O'   and isnull(T0.U_GRNo,'')=''  ")

                    GRStatus1 = objaddon.objglobalmethods.getSingleValue("select Distinct 1 from [@MIPL_POR2] T0  left join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry " &
                                " left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry and T0.U_Itemcode=T2.ItemCode " &
                                " where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T0.U_Date='" & DocDate.ToString("yyyyMMdd") & "' and T0.U_Status='C'  and isnull(T0.U_GRNo,'')<> ''  ")
                End If

                If GIStatus = "1" And GRStatus = "1" Then
                    'If objaddon.HANA Then
                    '    StrSql = "select T0.""DocEntry"" from OWTR T0 left join ""@MIPL_OPOR"" T1 on T0.""U_SubConNo""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "';"
                    'Else
                    '    StrSql = "select T0.DocEntry from OWTR T0 left join [@MIPL_OPOR] T1 on T0.U_SubConNo=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "'"
                    'End If
                    If objaddon.HANA Then
                        StrSql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OWTR T0 left join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join  ""@MIPL_OPOR"" T2 on (T2.""DocEntry""=T0.""U_SubConNo"" or T2.""DocEntry""=T1.""U_S_DocEntry"") where T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' "
                    Else
                        StrSql = "select distinct T0.DocEntry,T0.DocNum,T0.DocDate from OWTR T0 left join WTR1 T1 on T0.DocEntry=T1.DocEntry left join  [@MIPL_OPOR] T2 on (T2.DocEntry=T0.U_SubConNo or T2.DocEntry=T1.U_S_DocEntry) where T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' "
                    End If
                    objrecset.DoQuery(StrSql)
                    If objrecset.RecordCount = 0 Then
                        objaddon.objapplication.StatusBar.SetText("Inventory Transfer not created for the Planned Quantity...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    objGoodsIssue = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                    objaddon.objapplication.StatusBar.SetText("Goods Issue Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Try
                        txtDate.Value = Now.Date.ToString("yyyyMMdd")
                        'objform.Update()
                        If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                        objform.Freeze(True)
                        objGoodsIssue.DocDate = Now.Date 'DocDate
                        ' objGoodsIssue.DocDueDate = DocDueDate
                        objGoodsIssue.TaxDate = Now.Date ' DocDate 'objaddon.objcompany.GetCompanyDate

                        If BranchEnabled = "Y" Then
                            Dim series As String
                            If objaddon.HANA Then
                                series = objaddon.objglobalmethods.getSingleValue("select ""Series"" From NNM1 where ""ObjectCode""='60' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                        " and ""BPLId""=(select Top 1 ""BPLid"" from OWHS where ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(1).Specific.String & "') ")
                                Branch = objaddon.objglobalmethods.getSingleValue("select Top 1 ""BPLid"" from OWHS where ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(1).Specific.String & "'")
                            Else
                                series = objaddon.objglobalmethods.getSingleValue("select Series From NNM1 where ObjectCode='60' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                                                                                        " and BPLId=(select Top 1 BPLid from OWHS where WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(1).Specific.String & "') ")
                                Branch = objaddon.objglobalmethods.getSingleValue("select Top 1 BPLid from OWHS where WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(1).Specific.String & "'")
                            End If
                            If series <> "" Then
                                objGoodsIssue.Series = series
                            End If
                            objGoodsIssue.BPL_IDAssignedToInvoice = Branch
                        End If
                        'If objaddon.HANA Then
                        '    Branch = objaddon.objglobalmethods.getSingleValue("select ""BPLId"" from OBPL where ifnull(""MainBPL"",'Y')='Y'")
                        'Else
                        '    Branch = objaddon.objglobalmethods.getSingleValue("select BPLId from OBPL where isnull(MainBPL,'Y')='Y'")
                        'End If

                        'If Branch <> "" Then
                        '    objGoodsIssue.BPL_IDAssignedToInvoice = Branch
                        'End If
                        'Dim val As String = ComboBox1.Selected.Value
                        'objGoodsIssue.AgentCode = ComboBox1.Selected.Value 'objform.Items.Item("txtctper").Specific.string
                        objGoodsIssue.Reference1 = objform.Items.Item("txtvrefno").Specific.string
                        objGoodsIssue.Comments = "Auto Posted from Sub-Con Addon ->" & Now.ToString 'objform.Items.Item("txtremark").Specific.string 
                        objGoodsIssue.UserFields.Fields.Item("U_SubConNo").Value = objform.Items.Item("txtentry").Specific.string

                        For i As Integer = 1 To Matrix3.VisualRowCount
                            If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String = objform.Items.Item("txtsitem").Specific.String And Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String <> "" And Matrix3.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" Then
                                OutputQuantity = CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.String)
                            End If
                        Next
                        If GIFIFO = "Y" Then
                        Else
                            For IDocEntry As Integer = 0 To objrecset.RecordCount - 1
                                If IDocEntry = 0 Then
                                    Docentry = "'" & objrecset.Fields.Item("DocEntry").Value & "'"
                                Else
                                    Docentry = Docentry & ",'" & objrecset.Fields.Item("DocEntry").Value & "'"
                                End If
                                objrecset.MoveNext()
                            Next
                        End If

                        If objaddon.HANA Then
                            AcctCode = objaddon.objglobalmethods.getSingleValue(" select ""U_GICode"" from ""@MIPL_GEN""")
                        Else
                            AcctCode = objaddon.objglobalmethods.getSingleValue(" select U_GICode from [@MIPL_GEN]")
                        End If
                        For i As Integer = 1 To Matrix0.VisualRowCount
                            If Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String <> "" Then
                                Quantity = CDbl(Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String) * OutputQuantity
                                objGoodsIssue.Lines.ItemCode = Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String
                                objGoodsIssue.Lines.Quantity = Quantity
                                'If objaddon.HANA Then
                                '    WhsCode = objaddon.objglobalmethods.getSingleValue("Select Top 1 ""U_WhsCode"" from ""@MIPL_POR3""  where ""U_Itemcode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' order by ""LineId"" desc")
                                'Else
                                '    WhsCode = objaddon.objglobalmethods.getSingleValue("Select Top 1 U_WhsCode from [@MIPL_POR3]  where U_Itemcode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' order by LineId desc")
                                'End If
                                objGoodsIssue.Lines.WarehouseCode = Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String
                                'If WhsCode <> "" Then
                                '    objGoodsIssue.Lines.WarehouseCode = WhsCode
                                'Else
                                '    objGoodsIssue.Lines.WarehouseCode = Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String
                                'End If
                                If ItemProcess = "Y" Then
                                    If Trim(EditText20.Value) <> "" Then
                                        objGoodsIssue.Lines.UserFields.Fields.Item("U_Process").Value = Trim(EditText20.Value)
                                    End If
                                End If
                                If Matrix0.Columns.Item("cc1").Cells.Item(i).Specific.String <> "" Then objGoodsIssue.Lines.CostingCode = Matrix0.Columns.Item("cc1").Cells.Item(i).Specific.String
                                If Matrix0.Columns.Item("cc2").Cells.Item(i).Specific.String <> "" Then objGoodsIssue.Lines.CostingCode2 = Matrix0.Columns.Item("cc2").Cells.Item(i).Specific.String
                                If Matrix0.Columns.Item("cc3").Cells.Item(i).Specific.String <> "" Then objGoodsIssue.Lines.CostingCode3 = Matrix0.Columns.Item("cc3").Cells.Item(i).Specific.String
                                If Matrix0.Columns.Item("cc4").Cells.Item(i).Specific.String <> "" Then objGoodsIssue.Lines.CostingCode4 = Matrix0.Columns.Item("cc4").Cells.Item(i).Specific.String
                                If Matrix0.Columns.Item("cc5").Cells.Item(i).Specific.String <> "" Then objGoodsIssue.Lines.CostingCode5 = Matrix0.Columns.Item("cc5").Cells.Item(i).Specific.String
                                If AcctCode <> "" Then
                                    objGoodsIssue.Lines.AccountCode = AcctCode
                                Else
                                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    objaddon.objapplication.StatusBar.SetText("Please update the Accountcode in GL UDO screen...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    objform.Freeze(False)
                                    Exit Sub
                                End If
                                Dim BatchQty As Double = 0, BatchQty1 As Double = 0
                                If objaddon.HANA Then
                                    Serial = objaddon.objglobalmethods.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "'")
                                    Batchs = objaddon.objglobalmethods.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "'")
                                Else
                                    Serial = objaddon.objglobalmethods.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "'")
                                    Batchs = objaddon.objglobalmethods.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "'")
                                End If
                                If Batchs = "Y" And Serial = "N" Then
                                    If GIFIFO = "Y" Then
                                        If objaddon.HANA Then
                                            StrSql = "Select T0.""BatchNum"" ""BatchSerial"",T0.""Quantity"" ""Qty"",T0.""ItemCode"", T0.""WhsCode"""
                                            StrSql += vbCrLf + "from OBTN T2 left join OIBT T0 on T2.""DistNumber""=T0.""BatchNum"" and T0.""ItemCode""=T2.""ItemCode"""
                                            StrSql += vbCrLf + "where T0.""Direction""=0 and T0.""Quantity"">0 and T0.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T0.""WhsCode""='" & Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String & "' order by T2.""InDate""  "

                                        Else
                                            StrSql = "Select T0.BatchNum BatchSerial,T0.Quantity Qty,T0.ItemCode, T0.WhsCode"
                                            StrSql += vbCrLf + "from OBTN T2 left join OIBT T0 on T2.DistNumber=T0.BatchNum and T0.ItemCode=T2.ItemCode"
                                            StrSql += vbCrLf + "where T0.Direction=0 and T0.Quantity>0 and T0.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T0.WhsCode='" & Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String & "' order by T2.InDate  "
                                        End If
                                    Else
                                        Dim GIEntry, InvEntry As String
                                        If objaddon.HANA Then
                                            GIEntry = GetTransactionsVal("OIGE", "@MIPL_POR2", "U_SubConNo")
                                            InvEntry = GetTransactionsVal("OWTR", "@MIPL_POR3", "U_ScrapNum")
                                        Else
                                            GIEntry = GetTransactionsVal("OIGE", "[@MIPL_POR2]", "U_SubConNo")
                                            InvEntry = GetTransactionsVal("OWTR", "[@MIPL_POR3]", "U_ScrapNum")
                                        End If
                                        If objaddon.HANA Then
                                            'StrSql = "Select * from ("
                                            'StrSql += vbCrLf + "SELECT distinct T0.""U_SubConNo"",I1.""BatchNum"" ""BatchSerial"",T4.""WhsCode"",T1.""DocEntry"",T1.""ItemCode"",T4.""Quantity"" as ""Qty"", I1.""Quantity"",T4.""Status"" from OWTR T0 left join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                                            'StrSql += vbCrLf + "left outer join IBT1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                                            'StrSql += vbCrLf + "left outer join OIBT T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""BatchNum""=T4.""BatchNum"" and I1.""WhsCode"" = T4.""WhsCode"""
                                            'StrSql += vbCrLf + ")A Where A.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "' and A.""DocEntry"" in (" & Docentry & ") and A.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and A.""WhsCode""='" & Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String & "' and A.""BatchSerial"" <>'' and A.""Status""=0 and A.""Qty"">0 "

                                            StrSql = "SELECT A.""BatchNum"" as ""BatchSerial"",  SUM(A.""Quantity"") as ""Qty"" FROM ("
                                            StrSql += vbCrLf + "select T.""BatchNum"" ,  T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                                            StrSql += vbCrLf + "inner join wtr1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                                            StrSql += vbCrLf + "inner join owtr T3 on T2.""DocEntry""=T3.""DocEntry"""
                                            StrSql += vbCrLf + "where T.""BaseType""='67' and T.""Direction""=0 and T3.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "' and T.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T.""BaseEntry""in (" & Docentry & ")"
                                            StrSql += vbCrLf + "UNION ALL"
                                            StrSql += vbCrLf + "select T.""BatchNum"", -T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                                            StrSql += vbCrLf + "inner join ige1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                                            StrSql += vbCrLf + "inner join oige T3 on T2.""DocEntry""=T3.""DocEntry"""
                                            StrSql += vbCrLf + "where T.""BaseType""='60' and T.""Direction""=1 and T3.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "' and T.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' "
                                            If GIEntry <> "" Then
                                                StrSql += vbCrLf + "and T.""BaseEntry"" in (" & GIEntry & ")"
                                            Else
                                                StrSql += vbCrLf + "and IFNULL(T.""BaseEntry"",'') is Null"
                                            End If
                                            StrSql += vbCrLf + "UNION ALL"
                                            StrSql += vbCrLf + "select T.""BatchNum"", -T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                                            StrSql += vbCrLf + "inner join WTR1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                                            StrSql += vbCrLf + "inner join OWTR T3 on T2.""DocEntry""=T3.""DocEntry"""
                                            StrSql += vbCrLf + "where T.""BaseType""='67' and T.""Direction""=1 and T.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' "
                                            If InvEntry <> "" Then
                                                StrSql += vbCrLf + "and T3.""DocEntry"" in (" & InvEntry & ")"
                                            Else
                                                StrSql += vbCrLf + "and IFNULL(T3.""DocEntry"",'') is Null"
                                            End If
                                            StrSql += vbCrLf + ")A "
                                            StrSql += vbCrLf + "GROUP BY A.""BatchNum"" having SUM(A.""Quantity"") >0"

                                        Else
                                            'StrSql = "Select * from ("
                                            'StrSql += vbCrLf + "SELECT distinct T0.U_SubConNo,I1.BatchNum BatchSerial,T4.WhsCode,T1.DocEntry,T1.ItemCode,T4.Quantity as Qty, I1.Quantity,T4.Status from OWTR T0 left join WTR1 T1 on T0.DocEntry=T1.DocEntry"
                                            'StrSql += vbCrLf + "left outer join IBT1 I1 on T1.ItemCode=I1.ItemCode   and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                                            'StrSql += vbCrLf + "left outer join OIBT T4 on T4.ItemCode=I1.ItemCode and I1.BatchNum=T4.BatchNum and I1.WhsCode = T4.WhsCode"
                                            'StrSql += vbCrLf + ")A Where A.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "' and A.DocEntry in (" & Docentry & ") and A.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and A.WhsCode='" & Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String & "' and A.BatchSerial <>'' and A.Status=0 and A.Qty>0 "

                                            StrSql = "SELECT A.BatchNum as BatchSerial,  SUM(A.Quantity) as Qty FROM ("
                                            StrSql += vbCrLf + "select T.BatchNum,  T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                                            StrSql += vbCrLf + "inner join wtr1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                                            StrSql += vbCrLf + "inner join owtr T3 on T2.DocEntry=T3.DocEntry"
                                            StrSql += vbCrLf + "where T.BaseType='67' and T.Direction=0 and T3.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "' and T.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T.BaseEntry in (" & Docentry & ")"
                                            StrSql += vbCrLf + "UNION ALL"
                                            StrSql += vbCrLf + "select T.BatchNum, -T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                                            StrSql += vbCrLf + "inner join ige1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                                            StrSql += vbCrLf + "inner join oige T3 on T2.DocEntry=T3.DocEntry"
                                            StrSql += vbCrLf + "where T.BaseType='60' and T.Direction=1 and T3.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "' and T.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' "
                                            If GIEntry <> "" Then
                                                StrSql += vbCrLf + "and T.BaseEntry in (" & GIEntry & ")"
                                            Else
                                                StrSql += vbCrLf + "and ISNULL(T.BaseEntry,'') is Null"
                                            End If
                                            StrSql += vbCrLf + "UNION ALL"
                                            StrSql += vbCrLf + "select T.BatchNum, -T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                                            StrSql += vbCrLf + "inner join WTR1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                                            StrSql += vbCrLf + "inner join OWTR T3 on T2.DocEntry=T3.DocEntry"
                                            StrSql += vbCrLf + "where T.BaseType='67' and T.Direction=1 and T.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' "
                                            If InvEntry <> "" Then
                                                StrSql += vbCrLf + "and T3.DocEntry in (" & InvEntry & ")"
                                            Else
                                                StrSql += vbCrLf + "and ISNULL(T3.DocEntry,'') is Null"
                                            End If
                                            StrSql += vbCrLf + ")A "
                                            StrSql += vbCrLf + "GROUP BY A.BatchNum having SUM(A.Quantity) >0"
                                        End If
                                    End If

                                    objrs.DoQuery(StrSql)
                                    Dim BQty As Double = 0, PendQty As Double = 0, TotBatchQty As Double = 0
                                    BQty = Quantity '2
                                    If objrs.RecordCount > 0 Then
                                        For j As Integer = 0 To objrs.RecordCount - 1
                                            If (BQty - TotBatchQty) - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                                                PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                                            Else
                                                PendQty = BQty - TotBatchQty
                                            End If
                                            'If BQty - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                                            '    PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                                            'Else
                                            '    PendQty = BQty - TotBatchQty
                                            'End If
                                            objGoodsIssue.Lines.BatchNumbers.BatchNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                            objGoodsIssue.Lines.BatchNumbers.Quantity = PendQty
                                            objGoodsIssue.Lines.BatchNumbers.Add()
                                            TotBatchQty += PendQty
                                            If BQty - TotBatchQty > 0 Then
                                                objrs.MoveNext()
                                            Else
                                                Exit For
                                            End If
                                        Next
                                    End If
                                ElseIf Batchs = "N" And Serial = "Y" Then
                                    If GIFIFO = "Y" Then
                                        If objaddon.HANA Then
                                            StrSql = "SELECT distinct T4.""IntrSerial"" ""BatchSerial"", T4.""Quantity"",I1.""DocDate"""
                                            StrSql += vbCrLf + "from SRI1 I1 join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"""
                                            StrSql += vbCrLf + "where T4.""Quantity"">0 and I1.""Direction""=0 and I1.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and I1.""WhsCode""='" & Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String & "' order by I1.""DocDate"" "
                                        Else
                                            StrSql = "SELECT distinct T4.IntrSerial BatchSerial, T4.Quantity,I1.DocDate"
                                            StrSql += vbCrLf + "from SRI1 I1 join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode"
                                            StrSql += vbCrLf + "where T4.Quantity>0 and I1.Direction=0 and I1.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and I1.WhsCode='" & Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String & "' order by I1.DocDate "
                                        End If
                                    Else
                                        If objaddon.HANA Then
                                            StrSql = "Select * from ("
                                            StrSql += vbCrLf + "SELECT distinct T0.""U_SubConNo"",T4.""IntrSerial"" ""BatchSerial"",T4.""WhsCode"",T1.""DocEntry"",T1.""ItemCode"", T4.""Quantity"",T4.""Status"" from OWTR T0 inner join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                                            StrSql += vbCrLf + "left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                                            StrSql += vbCrLf + "left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" ) A "
                                            StrSql += vbCrLf + " Where A.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "' and A.""DocEntry"" in (" & Docentry & ")  and A.""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and  A.""WhsCode""='" & Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String & "' and A.""BatchSerial"" <>'' and A.""Status""=0"
                                        Else
                                            StrSql = "Select * from ("
                                            StrSql += vbCrLf + "SELECT distinct T0.U_SubConNo,T4.IntrSerial BatchSerial,T4.WhsCode,T1.DocEntry,T1.ItemCode, T4.Quantity,T4.Status from OWTR T0 inner join WTR1 T1 on T0.DocEntry=T1.DocEntry"
                                            StrSql += vbCrLf + "left outer join SRI1 I1 on T1.ItemCode=I1.ItemCode   and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                                            StrSql += vbCrLf + "left outer join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode ) A "
                                            StrSql += vbCrLf + " Where A.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "' and A.DocEntry in (" & Docentry & ")  and A.ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and  A.WhsCode='" & Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String & "' and A.BatchSerial <>'' and A.Status=0"
                                        End If
                                    End If
                                    objrs.DoQuery(StrSql)
                                    Dim SQty As Double = 0, TotSerialQty As Double = 0
                                    SQty = Quantity
                                    If objrs.RecordCount > 0 Then
                                        For j As Integer = 0 To objrs.RecordCount - 1
                                            objGoodsIssue.Lines.SerialNumbers.InternalSerialNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                            objGoodsIssue.Lines.SerialNumbers.Quantity = CDbl(1)
                                            objGoodsIssue.Lines.SerialNumbers.Add()
                                            TotSerialQty += CDbl(1)  '2
                                            If SQty - TotSerialQty > 0 Then
                                                objrs.MoveNext()
                                            Else
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Else
                                End If
                            End If
                            objGoodsIssue.Lines.Add()
                        Next i

                        Retval = objGoodsIssue.Add()
                        If Retval <> 0 Then
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objaddon.objapplication.StatusBar.SetText("Goods Issue : " & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsIssue)
                        Else
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Docentry = objaddon.objcompany.GetNewObjectKey()
                            For j = 1 To Matrix3.VisualRowCount
                                If Matrix3.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And Matrix3.Columns.Item("GINo").Cells.Item(j).Specific.String = "" And Matrix3.Columns.Item("Colstat").Cells.Item(j).Specific.String = "O" Then
                                    Matrix3.Columns.Item("GINo").Cells.Item(j).Specific.String = Docentry
                                End If
                            Next
                            'Matrix3.FlushToDataSource()
                            'For j As Integer = 0 To odbdsDetails2.Size - 1
                            '    Dim cc As String = odbdsDetails2.GetValue("U_GINo", j)
                            '    If odbdsDetails2.GetValue("U_Itemcode", j) <> "" And odbdsDetails2.GetValue("U_GINo", j) = "" And odbdsDetails2.GetValue("U_Status", j) = "O" Then
                            '        odbdsDetails2.SetValue("U_GINo", j, Docentry)
                            '    End If
                            'Next
                            'Matrix3.LoadFromDataSource()
                            'objform.Update()
                            'objform.Refresh()
                            If Not TranList = "Y" Then
                                If objaddon.HANA Then
                                    Docentry = objaddon.objglobalmethods.getSingleValue("Select ""DocNum"" from OIGE where ""DocEntry""=" & Docentry & " ")
                                Else
                                    Docentry = objaddon.objglobalmethods.getSingleValue("Select DocNum from OIGE where DocEntry=" & Docentry & "")
                                End If
                            End If
                            'If EditText22.Value = "" Then
                            '    EditText22.Value = Docentry
                            'Else
                            '    EditText22.Value = EditText22.Value + "," + Docentry
                            'End If
                            If odbdsHeader.GetValue("U_GINo", 0) = "" Then
                                odbdsHeader.SetValue("U_GINo", 0, Docentry)
                            Else
                                odbdsHeader.SetValue("U_GINo", 0, odbdsHeader.GetValue("U_GINo", 0) & "," & Docentry)
                            End If


                            CalcGRPrice()
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                            Status_Closing_RowDisable()
                            objaddon.objapplication.StatusBar.SetText("Goods Issue Added Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsIssue)
                            GC.Collect()
                            GoodsReceiptNew()
                        End If
                        objform.Freeze(False)
                        objrs = Nothing
                    Catch ex As Exception
                        objform.Freeze(False)
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                ElseIf GIStatus1 = "1" And GRStatus = "1" Then
                    CalcGRPrice()
                    GoodsReceiptNew()
                ElseIf GIStatus1 = "" And GRStatus1 = "1" Then
                    objaddon.objapplication.SetStatusBarMessage("Already Created Goods Issue & Goods Receipt for this line", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Exit Sub
                End If

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try

        End Sub

        Private Function GetTransactionsVal(ByVal Table1 As String, ByVal Table2 As String, ByVal FieldName As String)
            Try
                Dim Query As String
                Dim objrs As SAPbobsCOM.Recordset
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    Query = "select Distinct T0.""DocEntry"" from " & Table1 & " T0 join """ & Table2 & """ T1 on T1.""DocEntry""=T0.""" & FieldName & """  where T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " Order by T0.""DocEntry"""
                Else
                    Query = "select Distinct T0.DocEntry from " & Table1 & " T0 join " & Table2 & " T1 on T1.DocEntry=T0." & FieldName & "  where T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " Order by T0.DocEntry"
                End If
                objrs.DoQuery(Query)
                If objrs.RecordCount > 0 Then
                    Dim GetValues As New List(Of String)
                    For Rec As Integer = 0 To objrs.RecordCount - 1
                        GetValues.Add(objrs.Fields.Item(0).Value.ToString)
                        objrs.MoveNext()
                    Next
                    Dim DocEntryList = (From gv In GetValues Select New String(gv)).ToList()
                    Query = String.Join(",", DocEntryList)
                Else
                    Query = ""
                End If
                objrs = Nothing
                Return Query
            Catch ex As Exception
                Return ""
            End Try
        End Function

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

        Private Sub Addrow()
            Try
                If Matrix0.VisualRowCount > 0 And EditText0.Value <> "" Then
                    If odbdsDetails.GetValue("U_Itemcode", Matrix0.VisualRowCount - 1) = "" Then Exit Sub
                    objform.Freeze(True)
                    odbdsDetails.InsertRecord(odbdsDetails.Size)
                    odbdsDetails.SetValue("LineId", Matrix0.VisualRowCount, Matrix0.VisualRowCount + 1)
                    Matrix0.LoadFromDataSource()
                    objform.Freeze(False)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub FieldDisableInMatrix()
            Try
                If Not ReceiptAutoLoad = "Y" Then
                    Matrix0.Columns.Item("ProcQty").Visible = False
                    Matrix0.Columns.Item("OpenQty").Visible = False
                    Matrix0.Columns.Item("RetQty").Visible = False
                End If
                If Folder1.Selected = True Then   'Output
                    If ReceiptAutoLoad = "Y" Then
                        Matrix3.Columns.Item("GINo").Visible = False
                        Matrix3.Columns.Item("Colstat").Visible = False
                        Matrix3.Columns.Item("HSN").Visible = False
                        Matrix3.Columns.Item("Instock").Visible = False
                        Matrix3.Columns.Item("Code").Editable = False
                        Matrix3.Columns.Item("Desc").Editable = False
                        Matrix3.Columns.Item("Quant").Editable = False
                        Matrix3.Columns.Item("Whse").Editable = False
                        Matrix3.Columns.Item("DocDate").Visible = False
                        Matrix3.Columns.Item("RefNo").Visible = False
                        Matrix3.Columns.Item("Remarks").Visible = False
                    Else
                        Matrix3.Columns.Item("InvNum").Visible = False
                        Matrix3.Columns.Item("SeyDC").Visible = False
                        Matrix3.Columns.Item("SupDC").Visible = False
                    End If
                    If Not (VendorSPlPrice = "1" Or VendorSPlPrice = "2") Then
                        Matrix3.Columns.Item("ProCost").Visible = False
                        Matrix3.Columns.Item("TProCost").Visible = False
                    End If
                ElseIf Folder2.Selected = True Then  'Scrap
                    'If AutoProduction = "Y" Then
                    '    Matrix2.Columns.Item("Price").Visible = False
                    '    Matrix2.Columns.Item("HSN").Visible = False
                    '    Matrix2.Columns.Item("Instock").Visible = False
                    '    Matrix2.Columns.Item("GRNo").Visible = False
                    '    Matrix2.Columns.Item("Linetot").Visible = False
                    '    'For j = 1 To Matrix2.VisualRowCount
                    '    '    If Matrix2.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And Matrix2.Columns.Item("Colstat").Cells.Item(j).Specific.String = "C" Then
                    '    '        Matrix2.CommonSetting.SetRowEditable(j, False)
                    '    '    End If
                    '    'Next
                    '    For i As Integer = 1 To Matrix2.VisualRowCount
                    '        If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String = "" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" Then
                    '            Matrix2.CommonSetting.SetRowEditable(i, True)
                    '            'Matrix2.Columns.Item("Code").Editable = True
                    '            'Matrix2.Columns.Item("Quant").Editable = True
                    '            'Matrix2.Columns.Item("Whse").Editable = True
                    '        End If
                    '    Next
                    'End If
                ElseIf Folder3.Selected = True Then  'Related Doc
                    If APAutoLoad = "Y" Then
                        EditText11.Item.Enabled = False
                        Matrix1.Columns.Item("Code").Editable = False
                        Matrix1.Columns.Item("DocNum").Editable = False
                        Matrix1.Columns.Item("DocDate").Editable = False
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub AutoProductionOrder()
            Try
                Dim oProductionorder As SAPbobsCOM.ProductionOrders = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                Dim objRs As SAPbobsCOM.Recordset
                Dim GetValue, WhsCode As String
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    GetValue = objaddon.objglobalmethods.getSingleValue("select 1 as ""Status"" from OWOR where ""U_SubPONum""='" & objform.Items.Item("txtentry").Specific.String & "'")
                Else
                    GetValue = objaddon.objglobalmethods.getSingleValue("select 1 as Status from OWOR where U_SubPONum='" & objform.Items.Item("txtentry").Specific.String & "'")
                End If

                If EditText11.Value = "" And GetValue = "" Then
                    objaddon.objapplication.StatusBar.SetText("Production Order Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                    Dim PosDate As Date = Date.ParseExact(objform.Items.Item("posdate").Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    Dim DocDate As Date = Date.ParseExact(objform.Items.Item("deldate").Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    oProductionorder.PostingDate = PosDate
                    oProductionorder.DueDate = DocDate
                    'oProductionorder.StartDate = DocDate
                    oProductionorder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                    oProductionorder.ItemNo = objform.Items.Item("txtsitem").Specific.String
                    oProductionorder.PlannedQuantity = CDbl(objform.Items.Item("SQty").Specific.String)
                    If objaddon.HANA Then
                        WhsCode = objaddon.objglobalmethods.getSingleValue("select ""U_WhseCode"" from ""@MIPL_OBOM"" where ""Code""='" & EditText19.Value & "'")
                    Else
                        WhsCode = objaddon.objglobalmethods.getSingleValue("select U_WhseCode from [@MIPL_OBOM] where Code='" & EditText19.Value & "'")
                    End If
                    oProductionorder.Warehouse = WhsCode ' Matrix0.Columns.Item("SubWhse").Cells.Item(1).Specific.String
                    oProductionorder.Remarks = "From Sub-Con Add-on " & Now.ToString
                    oProductionorder.UserFields.Fields.Item("U_SubPONum").Value = objform.Items.Item("txtentry").Specific.String

                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String <> "" Then
                            If objaddon.HANA Then
                                GetValue = objaddon.objglobalmethods.getSingleValue("select 1 from OITM where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "'")
                            Else
                                GetValue = objaddon.objglobalmethods.getSingleValue("select 1 from OITM where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "'")
                            End If
                            If GetValue = "1" Then
                                oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                oProductionorder.Lines.ItemNo = Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String
                            Else
                                oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                oProductionorder.Lines.ItemNo = Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String
                            End If
                            'oProductionorder.Lines.ItemNo = Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String
                            'oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                            oProductionorder.Lines.BaseQuantity = CDbl(Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String)
                            oProductionorder.Lines.Warehouse = Matrix0.Columns.Item("SubWhse").Cells.Item(i).Specific.String
                            oProductionorder.Lines.Add()
                        End If
                    Next
                    If oProductionorder.Add() <> 0 Then
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objaddon.objapplication.SetStatusBarMessage("Production Posting Error : " & objaddon.objcompany.GetLastErrorDescription)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oProductionorder)
                        GC.Collect()
                    Else
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        GetValue = objaddon.objcompany.GetNewObjectKey()
                        EditText11.Value = GetValue
                        If objaddon.HANA Then
                            EditText27.Value = objaddon.objglobalmethods.getSingleValue("select Distinct ""DocNum"" from OWOR where ""DocEntry""='" & GetValue & "'")
                        Else
                            EditText27.Value = objaddon.objglobalmethods.getSingleValue("select Distinct DocNum from OWOR where DocEntry='" & GetValue & "'")
                        End If
                        EditText11.Item.Enabled = False
                        objaddon.objapplication.StatusBar.SetText("Production Order Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                Else
                    objaddon.objapplication.StatusBar.SetText("Production Order Already Created...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    If objaddon.HANA Then
                        GetValue = objaddon.objglobalmethods.getSingleValue("select ""DocEntry"" from OWOR where ""U_SubPONum""='" & objform.Items.Item("txtentry").Specific.String & "'")
                    Else
                        GetValue = objaddon.objglobalmethods.getSingleValue("select DocEntry from OWOR where U_SubPONum='" & objform.Items.Item("txtentry").Specific.String & "'")
                    End If
                    EditText11.Value = GetValue
                End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            Catch ex As Exception
                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

            End Try
        End Sub

        Private Sub LoadProdOrderReceiptToOutputScreen()
            Try
                Dim objRs, objRS1 As SAPbobsCOM.Recordset
                Dim StrReceipt As String = "", GREntry As String = ""
                Dim DocDate As Date
                Dim edittext As SAPbouiCOM.EditText
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_POR2")
                If EditText11.Value <> "" Then
                    If objaddon.HANA Then
                        GREntry = "select ""U_GRNo"" from ""@MIPL_POR2"" where ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'"
                    Else
                        GREntry = "select U_GRNo from [@MIPL_POR2] where DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'"
                    End If
                    objRS1.DoQuery(GREntry)
                    If objRS1.RecordCount > 0 Then
                        GREntry = "'"
                        For GRNum As Integer = 0 To objRS1.RecordCount - 1
                            GREntry += objRS1.Fields.Item("U_GRNo").Value + "','"
                            objRS1.MoveNext()
                        Next
                        GREntry = GREntry.Remove(GREntry.Length - 2)
                    End If
                    If Not Matrix3.VisualRowCount = 1 Then
                        RemoveLastrow(Matrix3, "Code")
                    End If

                    If objaddon.HANA Then
                        StrReceipt = "select ROW_NUMBER() OVER () AS ""LineId"", TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') ""DocDate"",T0.""U_SeyDCNum"",T0.""U_SupDCNum"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"",T1.""Price"",T1.""LineTotal"",T1.""WhsCode"",T1.""DocEntry"""
                        StrReceipt += vbCrLf + "from OIGN T0 join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry""  "
                        StrReceipt += vbCrLf + " where T1.""BaseType""=202 and T1.""BaseEntry""='" & EditText11.Value & "' "
                        StrReceipt += vbCrLf + " and  T0.""DocEntry"" not in (Select distinct ""U_GRNo"" from ""@MIPL_POR2"" where ""U_GRNo"" in (" & GREntry & ") )"
                        If Matrix3.VisualRowCount > 1 Then
                            edittext = Matrix3.Columns.Item("Date").Cells.Item(Matrix3.VisualRowCount).Specific
                            DocDate = Date.ParseExact(edittext.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            StrReceipt += vbCrLf + "  and  T0.""DocDate"">='" & DocDate.ToString("yyyyMMdd") & "'  "
                        End If
                        StrReceipt += vbCrLf + " order by T0.""DocDate"";"
                    Else
                        StrReceipt = "select ROW_NUMBER() over(Order by T1.DocEntry) AS LineId,Format(T0.DocDate,'dd/MM/yy') DocDate,T0.U_SeyDCNum,T0.U_SupDCNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.Price,T1.LineTotal,T1.WhsCode,T1.DocEntry"
                        StrReceipt += vbCrLf + "from OIGN T0 join IGN1 T1 on T0.DocEntry=T1.DocEntry  "
                        StrReceipt += vbCrLf + " where T1.BaseType=202 and T1.BaseEntry='" & EditText11.Value & "' "
                        StrReceipt += vbCrLf + " and  T0.DocEntry not in (Select distinct U_GRNo from [@MIPL_POR2] where U_GRNo in (" & GREntry & ") )"
                        If Matrix3.VisualRowCount > 1 Then
                            edittext = Matrix3.Columns.Item("Date").Cells.Item(Matrix3.VisualRowCount).Specific
                            DocDate = Date.ParseExact(edittext.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            StrReceipt += vbCrLf + "  and  T0.DocDate>='" & DocDate.ToString("yyyyMMdd") & "' "
                        End If
                        StrReceipt += vbCrLf + " order by T0.DocDate"
                    End If
                    objRs.DoQuery(StrReceipt)
                    'Matrix3.Clear()
                    odbdsDetails.Clear()
                    'objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                    If objRs.RecordCount > 0 Then
                        For GRRec As Integer = 0 To objRs.RecordCount - 1
                            If Matrix3.Columns.Item("Code").Cells.Item(Matrix3.VisualRowCount).Specific.String <> "" Then
                                Matrix3.AddRow()
                            End If
                            Matrix3.GetLineData(Matrix3.VisualRowCount)
                            odbdsDetails.SetValue("LineId", 0, Matrix3.VisualRowCount) 'objRs.Fields.Item("LineId").Value.ToString)
                            odbdsDetails.SetValue("U_Date", 0, Format(objRs.Fields.Item("DocDate").Value.ToString, "yyyyMMdd"))
                            odbdsDetails.SetValue("U_Itemcode", 0, objRs.Fields.Item("ItemCode").Value.ToString)
                            odbdsDetails.SetValue("U_ItemDesc", 0, objRs.Fields.Item("Dscription").Value.ToString)
                            odbdsDetails.SetValue("U_Qty", 0, objRs.Fields.Item("Quantity").Value.ToString)
                            odbdsDetails.SetValue("U_Price", 0, objRs.Fields.Item("Price").Value.ToString)
                            odbdsDetails.SetValue("U_LineTot", 0, objRs.Fields.Item("LineTotal").Value.ToString)
                            odbdsDetails.SetValue("U_WhsCode", 0, objRs.Fields.Item("WhsCode").Value.ToString)
                            odbdsDetails.SetValue("U_GRNo", 0, objRs.Fields.Item("DocEntry").Value.ToString)
                            odbdsDetails.SetValue("U_SeyDCNum", 0, objRs.Fields.Item("U_SeyDCNum").Value.ToString)
                            odbdsDetails.SetValue("U_SupDCNum", 0, objRs.Fields.Item("U_SupDCNum").Value.ToString)
                            Matrix3.SetLineData(Matrix3.VisualRowCount)
                            objRs.MoveNext()
                        Next
                    End If
                    'Matrix3.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    'Matrix3.Columns.Item("Price").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    'Matrix3.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                End If
                objRs = Nothing
                objRS1 = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub LoadAPInvoiceToRelatedDocTab()
            Try
                Dim objRs As SAPbobsCOM.Recordset
                Dim StrReceipt As String
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_POR4")
                If EditText24.Value <> "" Then
                    If objaddon.HANA Then
                        StrReceipt = "select ROW_NUMBER() OVER () AS ""LineId"",T0.""DocEntry"",TO_VARCHAR(Cast(T0.""DocDate"" as Timestamp),'dd/MM/yy') ""DocDate"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"",T1.""WhsCode"",T1.""LineTotal"" "
                        StrReceipt += vbCrLf + " from OPCH T0 join PCH1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""BaseType""=22 and T1.""BaseEntry""='" & EditText24.Value & "' order by T0.""DocDate"""
                    Else
                        StrReceipt = "select ROW_NUMBER() over(Order by T0.DocEntry) AS LineId,T0.DocEntry,Format(T0.DocDate,'dd/MM/yy') DocDate,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.LineTotal "
                        StrReceipt += vbCrLf + " from OPCH T0 join PCH1 T1 on T0.DocEntry=T1.DocEntry where T1.BaseType=22 and T1.BaseEntry='" & EditText24.Value & "' order by T0.DocDate"
                    End If
                    objRs.DoQuery(StrReceipt)
                    'RemoveLastrow(Matrix1, "DocNum")
                    Matrix1.Clear()
                    odbdsDetails.Clear()
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                    Dim cmbtype As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("DocType").Cells.Item(Matrix1.VisualRowCount).Specific
                    Dim CmbName As SAPbouiCOM.ComboBox = Matrix1.Columns.Item("Docname").Cells.Item(Matrix1.VisualRowCount).Specific
                    If objRs.RecordCount > 0 Then
                        CmbName.Select("18", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        cmbtype.Select("I", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        For GRRec As Integer = 0 To objRs.RecordCount - 1
                            If Matrix1.Columns.Item("Code").Cells.Item(Matrix1.VisualRowCount).Specific.String <> "" Then
                                Matrix1.AddRow()
                            End If
                            Matrix1.GetLineData(Matrix1.VisualRowCount)
                            odbdsDetails.SetValue("LineId", 0, objRs.Fields.Item("LineId").Value.ToString)
                            odbdsDetails.SetValue("U_DocNum", 0, objRs.Fields.Item("DocEntry").Value.ToString)
                            odbdsDetails.SetValue("U_DocDate", 0, objRs.Fields.Item("DocDate").Value.ToString)
                            odbdsDetails.SetValue("U_Itemcode", 0, objRs.Fields.Item("ItemCode").Value.ToString)
                            odbdsDetails.SetValue("U_ItemDesc", 0, objRs.Fields.Item("Dscription").Value.ToString)
                            odbdsDetails.SetValue("U_Qty", 0, objRs.Fields.Item("Quantity").Value.ToString)
                            odbdsDetails.SetValue("U_LineTot", 0, objRs.Fields.Item("LineTotal").Value.ToString)
                            odbdsDetails.SetValue("U_WhsCode", 0, objRs.Fields.Item("WhsCode").Value.ToString)
                            Matrix1.SetLineData(Matrix1.VisualRowCount)
                            objRs.MoveNext()
                        Next
                        'Matrix1.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                        'Matrix1.Columns.Item("Linetot").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Auto_InventoryTransfer(ByVal MatrixID As SAPbouiCOM.Matrix, ByVal HeaderUDF As String)
            Try
                Dim objstocktransfer As SAPbobsCOM.StockTransfer
                Dim objrs As SAPbobsCOM.Recordset
                Dim Lineflag As Boolean = False
                Dim Batch As String, Serial As String, DocEntry As String, Whse As String = "", FrmWhse As String = ""
                objstocktransfer = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If ToWhseInvTransfer = "Y" Then
                    If objaddon.HANA Then
                        Whse = objaddon.objglobalmethods.getSingleValue(" select ""U_InvWCode"" from ""@MIPL_GEN""")
                    Else
                        Whse = objaddon.objglobalmethods.getSingleValue(" select U_InvWCode from [@MIPL_GEN]")
                    End If
                End If
                If Whse = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Please update the Inventory Warehouse in General Settings...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Exit Sub
                End If
                For i As Integer = 1 To MatrixID.VisualRowCount
                    If MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And MatrixID.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" Then
                        Lineflag = True
                        FrmWhse = MatrixID.Columns.Item("Whse").Cells.Item(i).Specific.string
                    End If
                Next
                If Lineflag = True Then
                    objaddon.objapplication.StatusBar.SetText("Creating Inventory Transfer Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                    objstocktransfer.CardCode = objform.Items.Item("txtcode").Specific.String
                    objstocktransfer.DocDate = Now.Date
                    objstocktransfer.FromWarehouse = FrmWhse
                    objstocktransfer.ToWarehouse = Whse
                    objstocktransfer.JournalMemo = "Auto Generated" & Now.Date.ToString
                    objstocktransfer.UserFields.Fields.Item(HeaderUDF).Value = objform.Items.Item("txtentry").Specific.String
                    If MatrixID Is Matrix2 Then
                        objstocktransfer.Comments = "SubCon Scrap -> " & objform.Items.Item("txtentry").Specific.string
                    Else
                        objstocktransfer.Comments = "SubCon Output -> " & objform.Items.Item("txtentry").Specific.string
                    End If

                    For i As Integer = 1 To MatrixID.VisualRowCount
                        If MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And MatrixID.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" And MatrixID.Columns.Item("InvNum").Cells.Item(i).Specific.string = "" Then
                            If objaddon.HANA Then
                                Serial = objaddon.objglobalmethods.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                                Batch = objaddon.objglobalmethods.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                            Else
                                Serial = objaddon.objglobalmethods.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                                Batch = objaddon.objglobalmethods.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                            End If
                            If Batch = "Y" And Serial = "N" Then

                            ElseIf Batch = "N" And Serial = "Y" Then

                            ElseIf Batch = "N" And Serial = "N" Then
                                If CDbl(MatrixID.Columns.Item("Quant").Cells.Item(i).Specific.string) > 0 Then
                                    objstocktransfer.Lines.ItemCode = MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string
                                    objstocktransfer.Lines.Quantity = CDbl(MatrixID.Columns.Item("Quant").Cells.Item(i).Specific.string)
                                    objstocktransfer.Lines.FromWarehouseCode = MatrixID.Columns.Item("Whse").Cells.Item(i).Specific.string
                                    objstocktransfer.Lines.WarehouseCode = Whse
                                    objstocktransfer.Lines.Add()
                                End If
                            End If
                        End If
                    Next
                    If objstocktransfer.Add() <> 0 Then
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objaddon.objapplication.SetStatusBarMessage("Inventory Transfer: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Else
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objaddon.objcompany.GetNewObjectKey()
                        For i As Integer = 1 To MatrixID.VisualRowCount
                            If MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And MatrixID.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" Then
                                If MatrixID.Columns.Item("InvNum").Cells.Item(i).Specific.string = "" Then
                                    MatrixID.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O"
                                End If
                                MatrixID.Columns.Item("InvNum").Cells.Item(i).Specific.string = DocEntry
                                MatrixID.CommonSetting.SetCellEditable(i, 3, False)
                                MatrixID.CommonSetting.SetCellEditable(i, 5, False)
                                MatrixID.CommonSetting.SetCellEditable(i, 8, False)
                                MatrixID.Columns.Item("Colstat").Cells.Item(i).Specific.string = "C"
                            End If
                            If MatrixID.Columns.Item("Colstat").Cells.Item(i).Specific.string = "C" Then
                                MatrixID.CommonSetting.SetCellEditable(i, 3, False)
                                MatrixID.CommonSetting.SetCellEditable(i, 5, False)
                                MatrixID.CommonSetting.SetCellEditable(i, 8, False)
                            End If
                        Next
                        objaddon.objapplication.StatusBar.SetText("Inventory Transfer Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                Else
                    objaddon.objapplication.SetStatusBarMessage("No more Goods Receipt for posting Inventory Transfer...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Exit Sub
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objstocktransfer)
                GC.Collect()
            Catch ex As Exception
                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub Creating_Return_Auto_InventoryTransfer()
            Try
                Dim objstocktransfer As SAPbobsCOM.StockTransfer
                Dim objrs, objRecset As SAPbobsCOM.Recordset
                Dim Lineflag As Boolean = False
                Dim Batch As String, Serial As String, DocEntry As String, Whse As String = "", FrmWhse As String = ""
                Dim Quantity, BQty, TotBatchQty, TotSerialQty, PendQty, SQty As Double
                objstocktransfer = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                For i As Integer = 1 To Matrix2.VisualRowCount
                    objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                    If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And objcombo.Selected.Description = "Return" And Matrix2.Columns.Item("InvNum").Cells.Item(i).Specific.string = "" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" Then
                        If objaddon.HANA Then
                            objRecset.DoQuery("Select ""U_WhsCode"",""U_SubWhse"" from ""@MIPL_POR1"" where ""U_Itemcode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "' and ""DocEntry""='" & objform.Items.Item("txtentry").Specific.string & "'")
                        Else
                            objRecset.DoQuery("Select U_WhsCode,U_SubWhse from [@MIPL_POR1] where U_Itemcode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "' and DocEntry='" & objform.Items.Item("txtentry").Specific.string & "'")
                        End If
                        FrmWhse = objRecset.Fields.Item("U_SubWhse").Value.ToString
                        'Whse = objRecset.Fields.Item("U_WhsCode").Value.ToString
                        If Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string = "" Then
                            Whse = objRecset.Fields.Item("U_WhsCode").Value.ToString
                        Else
                            Whse = Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string
                        End If
                        Lineflag = True
                    End If
                Next
                If Lineflag = True Then
                    objaddon.objapplication.StatusBar.SetText("Creating Inventory Transfer Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                    objstocktransfer.CardCode = objform.Items.Item("txtcode").Specific.String
                    objstocktransfer.DocDate = Now.Date 'DocDate
                    objstocktransfer.FromWarehouse = FrmWhse
                    objstocktransfer.ToWarehouse = Whse
                    objstocktransfer.JournalMemo = "Auto Generated " & Now.ToString("dd/MMM/yyyy HH:mm:ss") 'Now.Date.ToString
                    objstocktransfer.UserFields.Fields.Item("U_ScrapNum").Value = objform.Items.Item("txtentry").Specific.String
                    objstocktransfer.Comments = "SubCon Return -> " & objform.Items.Item("txtentry").Specific.string

                    For i As Integer = 1 To Matrix2.VisualRowCount
                        TotBatchQty = 0
                        TotSerialQty = 0
                        objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                        If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And objcombo.Selected.Description = "Return" And Matrix2.Columns.Item("InvNum").Cells.Item(i).Specific.string = "" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" Then
                            If objaddon.HANA Then
                                Serial = objaddon.objglobalmethods.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                                Batch = objaddon.objglobalmethods.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                                objRecset.DoQuery("Select ""U_WhsCode"",""U_SubWhse"" from ""@MIPL_POR1"" where ""U_Itemcode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "' and ""DocEntry""='" & objform.Items.Item("txtentry").Specific.string & "'")
                            Else
                                Serial = objaddon.objglobalmethods.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                                Batch = objaddon.objglobalmethods.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                                objRecset.DoQuery("Select U_WhsCode,U_SubWhse from [@MIPL_POR1] where U_Itemcode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "' and DocEntry='" & objform.Items.Item("txtentry").Specific.string & "'")
                            End If
                            If Batch = "Y" And Serial = "N" Then
                                If CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string) > 0 Then
                                    Quantity = CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string)
                                    objstocktransfer.Lines.ItemCode = Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string
                                    objstocktransfer.Lines.Quantity = Quantity 'CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string)
                                    objstocktransfer.Lines.FromWarehouseCode = objRecset.Fields.Item("U_SubWhse").Value.ToString
                                    objstocktransfer.Lines.WarehouseCode = Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string ' objRecset.Fields.Item("U_WhsCode").Value.ToString
                                    If ItemProcess = "Y" Then
                                        If Trim(EditText20.Value) <> "" Then
                                            objstocktransfer.Lines.UserFields.Fields.Item("U_Process").Value = Trim(EditText20.Value)
                                        End If
                                    End If
                                    If GIFIFO = "Y" Then
                                        If objaddon.HANA Then
                                            strQuery = "Select T0.""BatchNum"" ""BatchSerial"",T0.""Quantity"" ""Qty"",T0.""ItemCode"", T0.""WhsCode"""
                                            strQuery += vbCrLf + "from OBTN T2 left join OIBT T0 on T2.""DistNumber""=T0.""BatchNum"" and T0.""ItemCode""=T2.""ItemCode"""
                                            strQuery += vbCrLf + "where T0.""Direction""=0 and T0.""Quantity"">0 and T0.""ItemCode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T0.""WhsCode""='" & objRecset.Fields.Item("U_SubWhse").Value.ToString & "' order by T2.""InDate""  "

                                        Else
                                            strQuery = "Select T0.BatchNum BatchSerial,T0.Quantity Qty,T0.ItemCode, T0.WhsCode"
                                            strQuery += vbCrLf + "from OBTN T2 left join OIBT T0 on T2.DistNumber=T0.BatchNum and T0.ItemCode=T2.ItemCode"
                                            strQuery += vbCrLf + "where T0.Direction=0 and T0.Quantity>0 and T0.ItemCode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T0.WhsCode='" & objRecset.Fields.Item("U_SubWhse").Value.ToString & "' order by T2.InDate  "
                                        End If
                                        objrs.DoQuery(strQuery)
                                    Else
                                        objrs = GetBatch_Serial("N", Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string, Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string)
                                    End If
                                    BQty = Quantity  '10
                                    If objrs.RecordCount > 0 Then
                                        For j As Integer = 0 To objrs.RecordCount - 1
                                            If (BQty - TotBatchQty) - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                                                PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                                            Else
                                                PendQty = BQty - TotBatchQty
                                            End If
                                            'If BQty - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                                            '    PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                                            'Else
                                            '    PendQty = BQty - TotBatchQty
                                            'End If
                                            objstocktransfer.Lines.BatchNumbers.BatchNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                            objstocktransfer.Lines.BatchNumbers.Quantity = PendQty ' BQty ' Quantity
                                            objstocktransfer.Lines.BatchNumbers.Add()
                                            TotBatchQty += PendQty  '2
                                            If BQty - TotBatchQty > 0 Then
                                                objrs.MoveNext()
                                            Else
                                                Exit For
                                            End If
                                        Next
                                        objstocktransfer.Lines.Add()
                                    End If
                                End If
                            ElseIf Batch = "N" And Serial = "Y" Then
                                If CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string) > 0 Then
                                    Quantity = CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string)
                                    objstocktransfer.Lines.ItemCode = Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string
                                    objstocktransfer.Lines.Quantity = Quantity
                                    objstocktransfer.Lines.FromWarehouseCode = objRecset.Fields.Item("U_SubWhse").Value.ToString
                                    objstocktransfer.Lines.WarehouseCode = Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string 'objRecset.Fields.Item("U_WhsCode").Value.ToString
                                    If ItemProcess = "Y" Then
                                        If Trim(EditText20.Value) <> "" Then
                                            objstocktransfer.Lines.UserFields.Fields.Item("U_Process").Value = Trim(EditText20.Value)
                                        End If
                                    End If
                                    'objrs = GetBatch_Serial("Y", Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string, Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string)
                                    objrs = GetBatch_Serial("Y", Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string, objRecset.Fields.Item("U_SubWhse").Value.ToString)
                                    SQty = Quantity
                                    If objrs.RecordCount > 0 Then
                                        For j As Integer = 0 To objrs.RecordCount - 1
                                            objstocktransfer.Lines.SerialNumbers.InternalSerialNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                            objstocktransfer.Lines.SerialNumbers.Quantity = CDbl(1)
                                            objstocktransfer.Lines.SerialNumbers.Add()
                                            TotSerialQty += CDbl(1)  '2
                                            If SQty - TotSerialQty > 0 Then
                                                objrs.MoveNext()
                                            Else
                                                Exit For
                                            End If
                                        Next
                                        objstocktransfer.Lines.Add()
                                    End If
                                End If
                            ElseIf Batch = "N" And Serial = "N" Then
                                If CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string) > 0 Then
                                    objstocktransfer.Lines.ItemCode = Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string
                                    objstocktransfer.Lines.Quantity = CDbl(Matrix2.Columns.Item("Quant").Cells.Item(i).Specific.string)
                                    objstocktransfer.Lines.FromWarehouseCode = objRecset.Fields.Item("U_SubWhse").Value.ToString
                                    objstocktransfer.Lines.WarehouseCode = Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string ' objRecset.Fields.Item("U_WhsCode").Value.ToString
                                    If ItemProcess = "Y" Then
                                        If Trim(EditText20.Value) <> "" Then
                                            objstocktransfer.Lines.UserFields.Fields.Item("U_Process").Value = Trim(EditText20.Value)
                                        End If
                                    End If
                                    objstocktransfer.Lines.Add()
                                End If
                            End If
                        End If
                    Next
                    If objstocktransfer.Add() <> 0 Then
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objaddon.objapplication.SetStatusBarMessage("Inventory Transfer: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Else
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objaddon.objcompany.GetNewObjectKey()
                        For i As Integer = 1 To Matrix2.VisualRowCount
                            objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                            If Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And objcombo.Selected.Description = "Return" And Matrix2.Columns.Item("InvNum").Cells.Item(i).Specific.string = "" And Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.string = "O" Then
                                Matrix2.Columns.Item("InvNum").Cells.Item(i).Specific.string = DocEntry
                                'Matrix2.CommonSetting.SetRowEditable(i, False)
                                Matrix2.CommonSetting.SetCellEditable(i, 2, False)
                                Matrix2.CommonSetting.SetCellEditable(i, 3, False)
                                Matrix2.CommonSetting.SetCellEditable(i, 6, False)
                                Matrix2.CommonSetting.SetCellEditable(i, 9, False)
                                Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.string = "C"
                            End If
                            'If Matrix2.Columns.Item("Colstat").Cells.Item(i).Specific.string = "C" Then
                            '    Matrix2.CommonSetting.SetRowEditable(i, False)
                            'End If
                        Next
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                        objaddon.objapplication.StatusBar.SetText("Inventory Transfer to Returned Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                Else
                    objaddon.objapplication.SetStatusBarMessage("No more Data for posting Inventory Transfer...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Exit Sub
                End If
                objrs = Nothing
                objRecset = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objstocktransfer)
                GC.Collect()
            Catch ex As Exception
                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Function Validate_Batch_Serial() As Boolean
            Try
                Dim objBatch_serial As SAPbobsCOM.Recordset
                Dim ErrCount As Integer = 0
                Dim Serial As String = "", Batch As String = ""

                For i As Integer = 1 To Matrix2.VisualRowCount
                    objcombo = Matrix2.Columns.Item("Type").Cells.Item(i).Specific
                    If objcombo.Selected.Description = "Return" Then
                        If objaddon.HANA Then
                            Serial = objaddon.objglobalmethods.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                            Batch = objaddon.objglobalmethods.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                        Else
                            Serial = objaddon.objglobalmethods.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                            Batch = objaddon.objglobalmethods.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string & "'")
                        End If

                        If Batch = "Y" And Serial = "N" Then
                            objBatch_serial = GetBatch_Serial("N", Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string, Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string)
                            If objBatch_serial.RecordCount = 0 Then
                                ErrCount += 1
                                objaddon.objapplication.StatusBar.SetText("Batch or Serial not available. ItemCode " & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        ElseIf Batch = "N" And Serial = "Y" Then
                            objBatch_serial = GetBatch_Serial("Y", Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string, Matrix2.Columns.Item("Whse").Cells.Item(i).Specific.string)
                            If objBatch_serial.RecordCount = 0 Then
                                ErrCount += 1
                                objaddon.objapplication.StatusBar.SetText("Batch or Serial not available. ItemCode " & Matrix2.Columns.Item("Code").Cells.Item(i).Specific.string, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    End If
                Next
                If ErrCount > 0 Then
                    Return False
                Else
                    Return True
                End If

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End Try
        End Function

        Private Function GetBatch_Serial(ByVal BatchSerial As String, ByVal ItemCode As String, ByVal WhsCode As String) As SAPbobsCOM.Recordset
            Dim objrs As SAPbobsCOM.Recordset
            Dim strSQL As String
            objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If BatchSerial = "N" Then
                'strSQL = "SELECT distinct I1.""BatchNum"" ""BatchSerial"", I1.""Quantity"",I1.""DocDate"""
                'strSQL += vbCrLf + " from IBT1 I1 join OIBT T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""BatchNum""=T4.""BatchNum"" and I1.""WhsCode"" = T4.""WhsCode"""
                'strSQL += vbCrLf + " where T4.""Quantity"">0 and I1.""ItemCode""='" & Matrix2.Columns.Item("itemcode").Cells.Item(1).Specific.String & "' and I1.""WhsCode""='' order by I1.""DocDate"" "
                'If objaddon.HANA Then
                '    strSQL = "Select * from ("
                '    strSQL += vbCrLf + "SELECT distinct I1.""BatchNum"" ""BatchSerial"",T0.""U_SubConNo"",T1.""ItemCode"",T4.""WhsCode"", T4.""Quantity"" as ""Qty"",I1.""Quantity"",T4.""Status"" from OWTR T0 left join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                '    strSQL += vbCrLf + "left outer join IBT1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                '    strSQL += vbCrLf + "left outer join OIBT T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""BatchNum""=T4.""BatchNum"" and I1.""WhsCode"" = T4.""WhsCode"""
                '    strSQL += vbCrLf + ")A Where  A.""U_SubConNo"" = '" & objform.Items.Item("txtentry").Specific.String & "' and A.""ItemCode""='" & ItemCode & "' and A.""WhsCode""='" & WhsCode & "' and A.""BatchSerial"" <>'' and A.""Status""=0 and A.""Qty"" >0 "
                'Else
                '    strSQL = "Select * from ("
                '    strSQL += vbCrLf + "SELECT distinct I1.BatchNum BatchSerial,T0.U_SubConNo,T1.ItemCode,T4.WhsCode, T4.Quantity as Qty,I1.Quantity,T4.Status from OWTR T0 left join WTR1 T1 on T0.DocEntry=T1.DocEntry"
                '    strSQL += vbCrLf + "left outer join IBT1 I1 on T1.ItemCode=I1.ItemCode   and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                '    strSQL += vbCrLf + "left outer join OIBT T4 on T4.ItemCode=I1.ItemCode and I1.BatchNum=T4.BatchNum and I1.WhsCode = T4.WhsCode"
                '    strSQL += vbCrLf + ")A Where  A.U_SubConNo = '" & objform.Items.Item("txtentry").Specific.String & "' and A.ItemCode='" & ItemCode & "' and A.WhsCode='" & WhsCode & "' and A.BatchSerial <>'' and A.Status=0 and A.Qty >0 "
                'End If
                Dim GIEntry, InvEntry As String
                If objaddon.HANA Then
                    GIEntry = GetTransactionsVal("OIGE", "@MIPL_POR2", "U_SubConNo")
                    InvEntry = GetTransactionsVal("OWTR", "@MIPL_POR3", "U_ScrapNum")
                Else
                    GIEntry = GetTransactionsVal("OIGE", "[@MIPL_POR2]", "U_SubConNo")
                    InvEntry = GetTransactionsVal("OWTR", "[@MIPL_POR3]", "U_ScrapNum")
                End If
                If objaddon.HANA Then
                    strSQL = "SELECT A.""BatchNum"" as ""BatchSerial"",  SUM(A.""Quantity"") as ""Qty"" FROM ("
                    strSQL += vbCrLf + "select T.""BatchNum"" ,  T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                    strSQL += vbCrLf + "inner join wtr1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                    strSQL += vbCrLf + "inner join owtr T3 on T2.""DocEntry""=T3.""DocEntry"""
                    strSQL += vbCrLf + "where T.""BaseType""='67' and T.""Direction""=0 and T3.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "' and T.""ItemCode""='" & ItemCode & "' "
                    strSQL += vbCrLf + "UNION ALL"
                    strSQL += vbCrLf + "select T.""BatchNum"", -T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                    strSQL += vbCrLf + "inner join ige1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                    strSQL += vbCrLf + "inner join oige T3 on T2.""DocEntry""=T3.""DocEntry"""
                    strSQL += vbCrLf + "where T.""BaseType""='60' and T.""Direction""=1 and T3.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "' and T.""ItemCode""='" & ItemCode & "' "
                    If GIEntry <> "" Then
                        strSQL += vbCrLf + "and T.""BaseEntry"" in (" & GIEntry & ")"
                    Else
                        strSQL += vbCrLf + "and IFNULL(T.""BaseEntry"",'') is Null"
                    End If
                    strSQL += vbCrLf + "UNION ALL"
                    strSQL += vbCrLf + "select T.""BatchNum"", -T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                    strSQL += vbCrLf + "inner join WTR1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                    strSQL += vbCrLf + "inner join OWTR T3 on T2.""DocEntry""=T3.""DocEntry"""
                    strSQL += vbCrLf + "where T.""BaseType""='67' and T.""Direction""=1 and T.""ItemCode""='" & ItemCode & "' "
                    If InvEntry <> "" Then
                        strSQL += vbCrLf + "and T3.""DocEntry"" in (" & InvEntry & ")"
                    Else
                        strSQL += vbCrLf + "and IFNULL(T3.""DocEntry"",'') is Null"
                    End If
                    strSQL += vbCrLf + ")A "
                    strSQL += vbCrLf + "GROUP BY A.""BatchNum"" having SUM(A.""Quantity"") >0"

                Else
                    strSQL = "SELECT A.BatchNum as BatchSerial,  SUM(A.Quantity) as Qty FROM ("
                    strSQL += vbCrLf + "select T.BatchNum,  T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                    strSQL += vbCrLf + "inner join wtr1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                    strSQL += vbCrLf + "inner join owtr T3 on T2.DocEntry=T3.DocEntry"
                    strSQL += vbCrLf + "where T.BaseType='67' and T.Direction=0 and T3.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "' and T.ItemCode='" & ItemCode & "'"
                    strSQL += vbCrLf + "UNION ALL"
                    strSQL += vbCrLf + "select T.BatchNum, -T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                    strSQL += vbCrLf + "inner join ige1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                    strSQL += vbCrLf + "inner join oige T3 on T2.DocEntry=T3.DocEntry"
                    strSQL += vbCrLf + "where T.BaseType='60' and T.Direction=1 and T3.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "' and T.ItemCode='" & ItemCode & "' "
                    If GIEntry <> "" Then
                        strSQL += vbCrLf + "and T.BaseEntry in (" & GIEntry & ")"
                    Else
                        strSQL += vbCrLf + "and ISNULL(T.BaseEntry,'') is Null"
                    End If
                    strSQL += vbCrLf + "UNION ALL"
                    strSQL += vbCrLf + "select T.BatchNum, -T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                    strSQL += vbCrLf + "inner join WTR1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                    strSQL += vbCrLf + "inner join OWTR T3 on T2.DocEntry=T3.DocEntry"
                    strSQL += vbCrLf + "where T.BaseType='67' and T.Direction=1 and T.ItemCode='" & ItemCode & "' "
                    If InvEntry <> "" Then
                        strSQL += vbCrLf + "and T3.DocEntry in (" & InvEntry & ")"
                    Else
                        strSQL += vbCrLf + "and ISNULL(T3.DocEntry,'') is Null"
                    End If
                    strSQL += vbCrLf + ")A "
                    strSQL += vbCrLf + "GROUP BY A.BatchNum having SUM(A.Quantity) >0"
                End If
            ElseIf BatchSerial = "Y" Then
                'strSQL = "SELECT distinct T4.""IntrSerial"" ""BatchSerial"", T4.""Quantity"",I1.""DocDate"""
                'strSQL += vbCrLf + " from SRI1 I1 join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"""
                'strSQL += vbCrLf + " where T4.""Quantity"">0 and I1.""ItemCode""='" & Matrix2.Columns.Item("itemcode").Cells.Item(1).Specific.String & "' and I1.""WhsCode""='' order by I1.""DocDate"" "
                If objaddon.HANA Then
                    strSQL = "Select * from ("
                    strSQL += vbCrLf + "SELECT distinct T4.""IntrSerial"" ""BatchSerial"",T0.""U_SubConNo"",T1.""ItemCode"", T4.""WhsCode"",T4.""Quantity"",T4.""Status"" from OWTR T0 inner join WTR1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                    strSQL += vbCrLf + "left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                    strSQL += vbCrLf + "left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" ) A "
                    strSQL += vbCrLf + " Where A.""U_SubConNo"" = '" & objform.Items.Item("txtentry").Specific.String & "' and A.""ItemCode""='" & ItemCode & "' and A.""WhsCode""='" & WhsCode & "' and A.""BatchSerial"" <>'' and A.""Status""=0"
                Else
                    strSQL = "Select * from ("
                    strSQL += vbCrLf + "SELECT distinct T4.IntrSerial BatchSerial,T0.U_SubConNo,T1.ItemCode, T4.WhsCode,T4.Quantity,T4.Status from OWTR T0 inner join WTR1 T1 on T0.DocEntry=T1.DocEntry"
                    strSQL += vbCrLf + "left outer join SRI1 I1 on T1.ItemCode=I1.ItemCode   and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                    strSQL += vbCrLf + "left outer join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode ) A "
                    strSQL += vbCrLf + " Where A.U_SubConNo = '" & objform.Items.Item("txtentry").Specific.String & "' and A.ItemCode='" & ItemCode & "' and A.WhsCode='" & WhsCode & "' and A.BatchSerial <>'' and A.Status=0"
                End If

            Else
                Return Nothing
            End If
            objrs.DoQuery(strSQL)
            Return objrs
        End Function

        Private Sub ManageAttributes()
            Try
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtdocnum", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtentry", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtstat", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "posdate", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "deldate", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "docdate", True, True, False)

                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtcode", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtname", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtctper", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtvrefno", True, True, False)
                'objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtsitem", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "SQty", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "SItemDesc", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "POLine", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtponum", True, True, True)
                'objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "MtxinputN", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "Series", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tInvUom", False, True, False)
                'objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "clstat", False, True, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "SubPoNum", True, True, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtbitem", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btngendoc", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btnOutput", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "InvScrap", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "BtnScrap", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btnload", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "BtnGIssue", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "BtnView", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "BtnInv", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btnbomref", True, False, True)


                If POItem = "Y" Then
                    'objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "SubPoNum", True, True, False)
                    objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtponum", False, True, False)
                End If
                If ItemProcess = "Y" Then
                    If objform.Items.Item("txtponum").Specific.String = "" Then
                        objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtponum", True, True, True)
                    Else
                        objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtponum", True, True, False)
                    End If
                    objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tproc", True, True, False)
                Else
                    objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tproc", False, True, False)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub SetTransEntryToSubPO(ByVal DocEntry As String, ByVal Header As String, ByVal UDFField As String)
            Try
                Dim GetVal As String = ""
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objform.Items.Item(UDFField).Specific.String <> "" Then Exit Sub
                If TranList = "Y" Then
                    If objaddon.HANA Then
                        GetVal = "select ""DocEntry"" from " & Header & " where ""U_SubConNo""='" & DocEntry & "'"
                    Else
                        GetVal = "select DocEntry from " & Header & " where U_SubConNo='" & DocEntry & "'"
                    End If
                Else
                    If objaddon.HANA Then
                        GetVal = "select ""DocNum"" from " & Header & " where ""U_SubConNo""='" & DocEntry & "'"
                    Else
                        GetVal = "select DocNum from " & Header & " where U_SubConNo='" & DocEntry & "'"
                    End If
                End If
                objRs.DoQuery(GetVal)
                If objRs.RecordCount > 0 Then
                    For i As Integer = 0 To objRs.RecordCount - 1
                        If i = 0 Then
                            GetVal = objRs.Fields.Item(0).Value.ToString
                        Else
                            GetVal += ", " + objRs.Fields.Item(0).Value.ToString
                        End If
                        objRs.MoveNext()
                    Next
                    objform.Items.Item(UDFField).Specific.String = GetVal
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Function GetOpen_VendorQty_Old(ByVal DocEntry As String)
            Dim GetValue As String = ""
            Dim VOBStock As Double, OpenQty As Double
            Try
                If DocEntry = "" Then Return 0

                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If EditText8.Value <> "" And objform.Items.Item("txtcode").Specific.String <> "" Then
                    If objaddon.HANA Then
                        'OpenQty = objaddon.objglobalmethods.getSingleValue(" select ifnull(sum(T3.""Quantity""),0) as ""Qty"" from OIGN T2 join IGN1 T3  on T2.""DocEntry""=T3.""DocEntry"" where T2.""U_SubConNo"" ='" & objform.Items.Item("txtentry").Specific.String & "' and T3.""ItemCode""='" & EditText8.Value & "';")
                        GetValue = "Select Case when Sum(T2.""Quantity"") >0 then T1.""U_SQty""-(Sum(T2.""Quantity"")+ "
                        GetValue += vbCrLf + "ifnull((select Top 1 sum(T1.""U_Qty"") from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""U_Type""='2' and T1.""U_Status""='C' and T0.""DocEntry""='" & DocEntry & "' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' "
                        GetValue += vbCrLf + " group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc),0)) else T1.""U_SQty""-"
                        GetValue += vbCrLf + "ifnull((select Top 1 sum(T1.""U_Qty"") from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" where T1.""U_Type""='2' and T1.""U_Status""='C' and T0.""DocEntry""='" & DocEntry & "' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' "
                        GetValue += vbCrLf + " group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc),0) end as ""OpenQty"" "
                        GetValue += vbCrLf + " from ""@MIPL_OPOR"" T1 left join OIGN T3 on T3.""U_SubConNo""=T1.""DocEntry"" left join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" and T1.""U_SItemCode""=T2.""ItemCode""  "
                        GetValue += vbCrLf + " where T1.""DocEntry""='" & DocEntry & "' group by T1.""U_SQty"" "

                        'GetValue = "Select MIN(((Case when A.""ReceiptQty"">0 then A.""U_PlanQty""-(A.""ReceiptQty""*A.""U_Qty"") Else A.""U_PlanQty"" End)-A.""ReturnQty"")/A.""U_Qty"") as ""OpenQty"",A.* "
                        'GetValue += vbCrLf + "from (Select T1.""DocEntry"",T0.""U_PlanQty"",T0.""U_Itemcode"",T0.""U_Qty"","
                        'GetValue += vbCrLf + "(Select ifnull((select top 1 sum(B.""U_Qty"") from ""@MIPL_OPOR"" A join ""@MIPL_POR3"" B on A.""DocEntry""=B.""DocEntry"" where B.""U_Type""='2' and B.""U_Status""='C' and A.""DocEntry""=T1.""DocEntry"" and B.""U_Itemcode""=T0.""U_Itemcode"""
                        'GetValue += vbCrLf + "and ifnull(B.""U_InvNo"",'') <>'' and ifnull(B.""U_GRNo"",'') ='' group by B.""U_Itemcode"" order by sum(B.""U_Qty"") ),0)) as ""ReturnQty"","
                        'GetValue += vbCrLf + "(Select Sum(T2.""Quantity"") from OIGN T3 join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" where T3.""U_SubConNo""=T1.""DocEntry"" and T1.""U_SItemCode""=T2.""ItemCode"") as ""ReceiptQty"""
                        'GetValue += vbCrLf + "from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"") A where A.""DocEntry""='" & DocEntry & "' group by A.""DocEntry"",A.""U_PlanQty"",A.""U_Itemcode"",A.""U_Qty"",A.""ReturnQty"",A.""ReceiptQty"""
                        'GetValue += vbCrLf + "order by MIN(((Case when A.""ReceiptQty"">0 then A.""U_PlanQty""-(A.""ReceiptQty""*A.""U_Qty"") Else A.""U_PlanQty"" End)-A.""ReturnQty"")/A.""U_Qty"")"
                        objRs.DoQuery(GetValue)
                        'OpenQty = objaddon.objglobalmethods.getSingleValue(GetValue)
                        OpenQty = CDbl(objRs.Fields.Item(0).Value.ToString)

                        'VOBStock = objaddon.objglobalmethods.getSingleValue("select ifnull(sum(T1.""U_SQty""),0)-ifnull(Sum(T3.""Quantity""),0) ""VOBStock""  from ""@MIPL_POR2"" T0 inner join ""@MIPL_OPOR"" T1 on T0.""DocEntry""=T1.""DocEntry""" & _
                        '                                                    " inner join OIGN T2 on T1.""DocEntry""=T2.""U_SubConNo"" inner join IGN1 T3 on T2.""DocEntry""=T3.""DocEntry"" and  T0.""U_Itemcode""=T3.""ItemCode""  and T1.""U_SItemCode"" =T3.""ItemCode""" & _
                        '                                                    " where T3.""ItemCode""='" & EditText8.Value & "' and T1.""Status""='O' and T1.""U_CardCode""='" & objform.Items.Item("txtcode").Specific.String & "'  ")
                        VOBStock = objaddon.objglobalmethods.getSingleValue("Select ifnull(Sum(T0.""U_SQty"")-(Select case when sum(T2.""Quantity"")>0 then sum(T2.""Quantity"") else 0 end from ""@MIPL_OPOR"" T0 join OIGN T1 on T1.""U_SubConNo""=T0.""DocEntry"" join IGN1 T2 on  T1.""DocEntry""=T2.""DocEntry""" &
                                                                         "and T0.""U_SItemCode""=T2.""ItemCode"" where T0.""Status""='O' and T0.""U_CardCode""='" & objform.Items.Item("txtcode").Specific.String & "' and T0.""U_SItemCode""='" & EditText8.Value & "'),0) as ""RecOpenQty"" " &
                                                                         " from ""@MIPL_OPOR"" T0 where T0.""U_CardCode""='" & objform.Items.Item("txtcode").Specific.String & "' and T0.""U_SItemCode""='" & EditText8.Value & "' and T0.""Status""='O'")

                    Else
                        'OpenQty = objaddon.objglobalmethods.getSingleValue(" select isnull(sum(T3.Quantity),0) as Qty from OIGN T2 join IGN1 T3  on T2.DocEntry=T3.DocEntry where T2.U_SubConNo ='" & objform.Items.Item("txtentry").Specific.String & "' and T3.ItemCode='" & EditText8.Value & "'")
                        'OpenQty = objaddon.objglobalmethods.getSingleValue("Select Case when Sum(T2.Quantity) >0 then T1.U_SQty-(Sum(T2.Quantity)+ (Select isnull(sum(U_Qty),0) from [@MIPL_POR3] where DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' " & _
                        '                                                     " and isnull(U_InvNo,'') <>'' and isnull(U_GRNo,'') ='' )) else (T1.U_SQty-(Select isnull(sum(U_Qty),0) from [@MIPL_POR3] where DocEntry='140' and isnull(U_InvNo,'') <>''and isnull(U_GRNo,'') ='' )) end as OpenQty " & _
                        '                                                     " from [@MIPL_OPOR] T1 left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry and T1.U_SItemCode=T2.ItemCode  " & _
                        '                                                      " where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T1.U_SQty ")

                        GetValue = "Select Case when Sum(T2.Quantity) >0 then T1.U_SQty-(Sum(T2.Quantity)+ "
                        GetValue += vbCrLf + "isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and T0.DocEntry='" & DocEntry & "' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' "
                        GetValue += vbCrLf + " group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0)) else T1.U_SQty-"
                        GetValue += vbCrLf + "isnull((select Top 1 sum(T1.U_Qty) from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry where T1.U_Type='2' and T1.U_Status='C' and T0.DocEntry='" & DocEntry & "' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' "
                        GetValue += vbCrLf + " group by T1.U_Itemcode order by sum(T1.U_Qty) desc),0) end as OpenQty "
                        GetValue += vbCrLf + " from [@MIPL_OPOR] T1 left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry and T1.U_SItemCode=T2.ItemCode  "
                        GetValue += vbCrLf + " where T1.DocEntry='" & DocEntry & "' group by T1.U_SQty "

                        'GetValue = "Select case when isnull((select Top 1 U_Roundoff from [@MI_ITMRND] where U_InvUOM=(SELECT InvntryUom FROM OITM where ItemCode=A.U_SItemCode and InvntryUom <>'') order by Code desc),'') ='Y' then "
                        'GetValue += vbCrLf + "cast(MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty) as decimal(18,0)) Else MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty) End As OpenQty,A.* "
                        'GetValue += vbCrLf + "from (Select T1.DocEntry,T0.U_PlanQty,T0.U_Itemcode,T0.U_Qty,T1.U_InvUom,T1.U_SItemCode,"
                        'GetValue += vbCrLf + "(Select isnull((Select top 1 sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B On A.DocEntry= B.DocEntry where B.U_Type='2' and B.U_Status='C' and A.DocEntry=T1.DocEntry and B.U_Itemcode=T0.U_Itemcode"
                        'GetValue += vbCrLf + "and isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' group by B.U_Itemcode order by sum(B.U_Qty) ),0)) as ReturnQty,"
                        'GetValue += vbCrLf + "(Select Sum(T2.Quantity) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T1.U_SItemCode=T2.ItemCode) as ReceiptQty"
                        'GetValue += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry) A where A.DocEntry='" & DocEntry & "' group by A.DocEntry,A.U_PlanQty,A.U_Itemcode,A.U_Qty,A.ReturnQty,A.ReceiptQty,A.U_SItemCode,A.U_InvUom"
                        'GetValue += vbCrLf + "order by MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty)"

                        objRs.DoQuery(GetValue)
                        'OpenQty = objaddon.objglobalmethods.getSingleValue(GetValue)
                        OpenQty = CDbl(objRs.Fields.Item(0).Value.ToString)
                        'VOBStock = objaddon.objglobalmethods.getSingleValue("select isnull(sum(T1.U_SQty),0)-isnull(Sum(T3.Quantity),0) VOBStock  from [@MIPL_POR2] T0 inner join [@MIPL_OPOR] T1 on T0.DocEntry=T1.DocEntry" & _
                        '                                                    " inner join OIGN T2 on T1.DocEntry=T2.U_SubConNo inner join IGN1 T3 on T2.DocEntry=T3.DocEntry and  T0.U_Itemcode=T3.ItemCode  and T1.U_SItemCode =T3.ItemCode" & _
                        '                                                    " where T3.ItemCode='" & EditText8.Value & "' and T1.Status='O' and T1.U_CardCode='" & objform.Items.Item("txtcode").Specific.String & "'  ")
                        VOBStock = objaddon.objglobalmethods.getSingleValue("Select isnull(Sum(T0.U_SQty)-(Select case when sum(T2.Quantity)>0 then sum(T2.Quantity) else 0 end from [@MIPL_OPOR] T0 join OIGN T1 on T1.U_SubConNo=T0.DocEntry join IGN1 T2 on  T1.DocEntry=T2.DocEntry" &
                                                                         " and T0.U_SItemCode=T2.ItemCode where T0.Status='O' and T0.U_CardCode='" & objform.Items.Item("txtcode").Specific.String & "' and T0.U_SItemCode='" & EditText8.Value & "'),0) as RecOpenQty " &
                                                                         " from [@MIPL_OPOR] T0 where T0.U_CardCode='" & objform.Items.Item("txtcode").Specific.String & "' and T0.U_SItemCode='" & EditText8.Value & "' and T0.Status='O'")

                    End If
                End If
            Catch ex As Exception

            End Try

            Return {OpenQty, VOBStock}
        End Function

        Private Function GetVendor_OpenReceiptQty(ByVal DocEntry As String)
            Try
                Dim GetValue As String = ""
                If DocEntry = "" Then Return 0
                If EditText8.Value <> "" And objform.Items.Item("txtcode").Specific.String <> "" Then
                    If objaddon.HANA Then
                        GetValue = "Select Sum(A.""U_SQty"")-sum(A.""RecQty"") as ""OpenVendorQty"" from (Select T0.""DocEntry"",(T0.""U_SQty""),(Select ifnull(sum(B.""Quantity""),0) from OIGN A join IGN1 B On A.""DocEntry""=B.""DocEntry"" and A.""U_SubConNo""=T0.""DocEntry"" "
                        GetValue += vbCrLf + "and B.""ItemCode""=T0.""U_SItemCode"" and B.""U_CardCode""=T0.""U_CardCode"") as ""RecQty"" from ""@MIPL_OPOR"" T0 where T0.""U_CardCode""='" & objform.Items.Item("txtcode").Specific.String & "' and "
                        GetValue += vbCrLf + "T0.""U_SItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and T0.""Status""='O') A  group by A.""DocEntry"" having Sum(A.""U_SQty"")-sum(A.""RecQty"")>0"

                    Else
                        GetValue = "Select Sum(A.U_SQty)-sum(A.RecQty) as OpenVendorQty from (Select T0.DocEntry,(T0.U_SQty),(Select isnull(sum(B.Quantity),0) from OIGN A join IGN1 B On A.DocEntry=B.DocEntry and A.U_SubConNo=T0.DocEntry "
                        GetValue += vbCrLf + "and B.ItemCode=T0.U_SItemCode and B.U_CardCode=T0.U_CardCode) as RecQty from [@MIPL_OPOR] T0 where T0.U_CardCode='" & objform.Items.Item("txtcode").Specific.String & "' and "
                        GetValue += vbCrLf + "T0.U_SItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and T0.Status='O') A  group by A.DocEntry having Sum(A.U_SQty)-sum(A.RecQty)>0"
                    End If
                    GetValue = objaddon.objglobalmethods.getSingleValue(GetValue)
                End If
                Return GetValue
            Catch ex As Exception

            End Try
        End Function

        Private Function GetStatusClosing() As String
            Try 'Closing the sub-contracting status
                Dim ReceiptStatus As String = "", Inventorystatus As String = ""
                If objform.Items.Item("txtsitem").Specific.String = "" Or objform.Items.Item("txtentry").Specific.String = "" Then Exit Function
                If objaddon.HANA Then
                    ReceiptStatus = "Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry"" left join ""@MIPL_OPOR"" T2 on T0.""U_SubConNo""=T2.""DocEntry"" and T2.""U_SItemCode""=  T1.""ItemCode"" "
                    ReceiptStatus += vbCrLf + "where T1.""ItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.""U_SQty"" "
                    ReceiptStatus += vbCrLf + "having T2.""U_SQty""- ifnull(sum(T1.""Quantity"")+ "
                    ReceiptStatus += vbCrLf + "ifnull((select Top 1 sum(T1.""U_Qty"")/ T2.""U_BOMQty"" from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" join ""@MIPL_POR1"" T2 on T1.""DocEntry""=T2.""DocEntry"" "
                    ReceiptStatus += vbCrLf + "where T1.""U_Type""='2' and T1.""U_Status""='C' and T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' "
                    ReceiptStatus += vbCrLf + " group by T1.""U_Itemcode"",T2.""U_BOMQty"" order by sum(T1.""U_Qty"")/T2.""U_BOMQty"" desc ),0),0) <=0"

                    ReceiptStatus = objaddon.objglobalmethods.getSingleValue(ReceiptStatus)

                    Inventorystatus = "SELECT DISTINCT 1 AS ""status"" FROM ""@MIPL_OPOR"" T1 LEFT OUTER JOIN ""@MIPL_POR1"" T0 ON T1.""DocEntry"" = T0.""DocEntry"" LEFT OUTER JOIN OWTR T3 ON T3.""U_SubConNo"" = T1.""DocEntry"""
                    Inventorystatus += vbCrLf + "LEFT OUTER JOIN WTR1 T2 ON T3.""DocEntry"" = T2.""DocEntry"" AND T2.""ItemCode"" = T0.""U_Itemcode"" WHERE T1.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "'"
                    Inventorystatus += vbCrLf + "AND T0.""U_Itemcode"" IN (SELECT ""U_Itemcode"" FROM ""@MIPL_POR1"" WHERE ""DocEntry"" = T0.""DocEntry"") GROUP BY T1.""DocEntry"", T0.""U_Itemcode"", T0.""U_PlanQty"""
                    Inventorystatus += vbCrLf + "having ifnull(sum(T2.""Quantity""), 0) - ifnull((Select Sum(B.""Quantity"") from OIGE A join IGE1 B on A.""DocEntry""=B.""DocEntry"" where A.""U_SubConNo""=T1.""DocEntry"" and B.""ItemCode""=T0.""U_Itemcode""),0) +"
                    Inventorystatus += vbCrLf + "ifnull((select sum(B.""U_Qty"") from ""@MIPL_OPOR"" A join ""@MIPL_POR3"" B on A.""DocEntry""=B.""DocEntry"" where B.""U_Type""='2' and B.""U_Status""='C'"
                    Inventorystatus += vbCrLf + "And A.""DocEntry""=T1.""DocEntry"" And B.""U_Itemcode""=T0.""U_Itemcode"" And ifnull(B.""U_InvNo"",'') <>'' and ifnull(B.""U_GRNo"",'') ='' ),0)<=0"

                    Inventorystatus = objaddon.objglobalmethods.getSingleValue(Inventorystatus)
                Else
                    ReceiptStatus = "Select distinct 1 from OIGN T0 inner join IGN1 T1 on T0.DocEntry=T1.DocEntry left join [@MIPL_OPOR] T2 on T0.U_SubConNo=T2.DocEntry and T2.U_SItemCode=  T1.ItemCode "
                    ReceiptStatus += vbCrLf + "where T1.ItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and T2.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T2.U_SQty "
                    ReceiptStatus += vbCrLf + "having T2.U_SQty- isnull(sum(T1.Quantity)+ "
                    ReceiptStatus += vbCrLf + "isnull((select Top 1 sum(T1.U_Qty)/ T2.U_BOMQty from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry join [@MIPL_POR1] T2 on T1.DocEntry=T2.DocEntry "
                    ReceiptStatus += vbCrLf + "where T1.U_Type='2' and T1.U_Status='C' and T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' "
                    ReceiptStatus += vbCrLf + " group by T1.U_Itemcode,T2.U_BOMQty order by sum(T1.U_Qty)/T2.U_BOMQty desc ),0),0) <=0"

                    ReceiptStatus = objaddon.objglobalmethods.getSingleValue(ReceiptStatus)

                    Inventorystatus = "Select distinct 1 As status from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 On T1.DocEntry=T0.DocEntry left join OWTR T3 On T3.U_SubConNo=T1.DocEntry"
                    Inventorystatus += vbCrLf + "Left Join WTR1 T2  on T3.DocEntry=T2.DocEntry And T2.ItemCode=T0.U_Itemcode where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "'"
                    Inventorystatus += vbCrLf + "And T0.U_Itemcode in (Select U_Itemcode from [@MIPL_POR1] where DocEntry=T0.DocEntry) group by T1.DocEntry,T0.U_Itemcode,T0.U_PlanQty "
                    Inventorystatus += vbCrLf + "having isnull(sum(T2.Quantity), 0) - (isnull((Select Sum(B.Quantity) from OIGE A join IGE1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T1.DocEntry and B.ItemCode=T0.U_Itemcode),0) +"
                    Inventorystatus += vbCrLf + "isnull((select sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B on A.DocEntry=B.DocEntry where B.U_Type='2' and B.U_Status='C'"
                    Inventorystatus += vbCrLf + "And A.DocEntry=T1.DocEntry And B.U_Itemcode=T0.U_Itemcode And isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' ),0))<=0"
                    Inventorystatus = objaddon.objglobalmethods.getSingleValue(Inventorystatus)

                    'Select Case distinct 1 As status from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 On T1.DocEntry=T0.DocEntry left join OWTR T3 On T3.U_SubConNo=T1.DocEntry 
                    'Left Join WTR1 T2  on T3.DocEntry=T2.DocEntry And T2.ItemCode=T0.U_Itemcode where T1.DocEntry='1542' 
                    'And T0.U_Itemcode in (Select U_Itemcode from [@MIPL_POR1] where DocEntry=T0.DocEntry) group by T1.DocEntry,T0.U_Itemcode,T0.U_PlanQty
                    'having isnull(sum(T2.Quantity), 0) - isnull(T0.U_PlanQty + isnull((select sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B on A.DocEntry=B.DocEntry where B.U_Type='2' and B.U_Status='C'
                    'And A.DocEntry=T1.DocEntry And B.U_Itemcode=T0.U_Itemcode And isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' ),0),0)  <=0

                End If
                If ReceiptStatus = "1" And Inventorystatus = "1" Then
                    Return "1"
                Else
                    Return "0"
                End If


            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Function StatusClosing_OpenQty() As Boolean
            Try
                If objform.Items.Item("txtsitem").Specific.String = "" Or objform.Items.Item("txtentry").Specific.String = "" Or ComboBox2.Selected.Value = "C" Then Exit Function

                'strQuery = "Select Top 1 B.* from (Select Case when A.InvTransferQty>0 and (A.InvTransferQty-A.IssueQty-A.ReturnQty)=0 then 'C' Else 'O' End ManualStatus,"
                'strQuery += vbCrLf + "Case when A.HeaderQty=A.ReceiptQty and (A.LineQty-(A.IssueQty+A.ReturnQty))=0 then 'C' Else 'O' End AutoStatus,"
                'strQuery += vbCrLf + "case when isnull((select Top 1 U_Roundoff from [@MI_ITMRND] where U_InvUOM=(SELECT InvntryUom FROM OITM where ItemCode=A.HeaderItem and InvntryUom <>'') "
                'strQuery += vbCrLf + "order by Code desc),'') ='Y' then cast(MIN(((Case when A.ReceiptQty>0 then A.LineQty-(A.ReceiptQty*A.BOMNorms) Else A.LineQty End)-A.ReturnQty)/A.BOMNorms) as decimal(18,0))"
                'strQuery += vbCrLf + "Else MIN(((Case when A.ReceiptQty>0 then A.LineQty-(A.ReceiptQty*A.BOMNorms) Else A.LineQty End)-A.ReturnQty)/A.BOMNorms) End As OpenQty,A.*"
                'strQuery += vbCrLf + "from (Select T1.DocEntry,T1.Status,T1.U_SItemCode as HeaderItem,T1.U_SQty as HeaderQty,T0.U_Itemcode as LineItem,T0.U_PlanQty as LineQty,"
                'strQuery += vbCrLf + "isnull(T0.U_ActPlanQty,0) as ActPlanQty,isnull((T0.U_PlanQty-T0.U_ActPlanQty),0) [Excess PlanQty],T0.U_Qty as BOMNorms,T1.U_InvUom,"
                'strQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T1.DocEntry and B.ItemCode=T0.U_Itemcode ) as InvTransferQty,"
                'strQuery += vbCrLf + "(Select isnull(Sum(T2.Quantity),0) from OIGE T3 join IGE1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T0.U_Itemcode=T2.ItemCode) as IssueQty,"
                'strQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A inner join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T0.DocEntry and B.ItemCode=T0.U_Itemcode) as ReturnQty,"
                'strQuery += vbCrLf + "(Select isnull(Sum(T2.Quantity),0) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T1.U_SItemCode=T2.ItemCode) as ReceiptQty"
                'strQuery += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry) A where A.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' "
                'strQuery += vbCrLf + "group by A.DocEntry,A.Status,A.LineQty,A.LineItem,A.BOMNorms,A.ReturnQty,A.ReceiptQty,A.HeaderItem,A.ActPlanQty,A.[Excess PlanQty],"
                'strQuery += vbCrLf + "A.U_InvUom,A.InvTransferQty,A.IssueQty,A.HeaderQty) B Order by B.OpenQty desc"

                strQuery = "Select Top 1 B.* from (Select Case when A.InvTransferQty>0 and (A.InvTransferQty-A.IssueQty-A.ReturnQty)=0 then 'C' Else 'O' End ManualStatus,"
                strQuery += vbCrLf + "Case when A.HeaderQty=A.ReceiptQty and (A.LineQty-(A.IssueQty+A.ReturnQty))=0 and isnull((C.[Tot Scrap Qty]-A.ScrapQty),0)=0 then 'C' Else 'O' End AutoStatus,"
                strQuery += vbCrLf + "case when isnull((select Top 1 U_Roundoff from [@MI_ITMRND] where U_InvUOM=(SELECT InvntryUom FROM OITM where ItemCode=A.HeaderItem and InvntryUom <>'') "
                strQuery += vbCrLf + "order by Code desc),'') ='Y' then cast(MIN(((Case when A.ReceiptQty>0 then A.LineQty-(A.ReceiptQty*A.BOMNorms) Else A.LineQty End)-A.ReturnQty)/A.BOMNorms) as decimal(18,0))"
                strQuery += vbCrLf + "Else MIN(((Case when A.ReceiptQty>0 then A.LineQty-(A.ReceiptQty*A.BOMNorms) Else A.LineQty End)-A.ReturnQty)/A.BOMNorms) End As OpenQty,"
                strQuery += vbCrLf + "isnull((C.[Tot Scrap Qty]-A.ScrapQty),0) [Open Scrap Qty],A.*,C.*"
                strQuery += vbCrLf + "from (Select T1.DocEntry,T1.Status,T1.U_SItemCode as HeaderItem,T1.U_BOMCode as BOMItem,T1.U_SQty as HeaderQty,T0.U_Itemcode as LineItem,T0.U_PlanQty as LineQty,"
                strQuery += vbCrLf + "isnull(T0.U_ActPlanQty,0) as ActPlanQty,isnull((T0.U_PlanQty-T0.U_ActPlanQty),0) [Excess PlanQty],T0.U_Qty as BOMNorms,T1.U_InvUom,"
                strQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T1.DocEntry and B.ItemCode=T0.U_Itemcode ) as InvTransferQty,"
                strQuery += vbCrLf + "(Select isnull(Sum(T2.Quantity),0) from OIGE T3 join IGE1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T0.U_Itemcode=T2.ItemCode) as IssueQty,"
                strQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A inner join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T0.DocEntry and B.ItemCode=T0.U_Itemcode) as ReturnQty,"
                strQuery += vbCrLf + "(Select isnull(Sum(T2.Quantity),0) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T1.U_SItemCode=T2.ItemCode) as ReceiptQty,"
                strQuery += vbCrLf + "(Select isnull(Sum(T2.Quantity),0) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_ScrapNum=T1.DocEntry) as ScrapQty"
                strQuery += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry) A"
                strQuery += vbCrLf + "left join (Select isnull(Sum(T0.U_Qty*T1.U_SQty),0) [Tot Scrap Qty],T1.DocEntry SubEntry,T1.U_BOMCode,T0.Code from [@MIPL_BOM1] T0 join [@MIPL_OPOR] T1 on"
                strQuery += vbCrLf + "T0.Code=T1.U_BOMCode where T0.U_SCType='S' group by T1.DocEntry,T1.U_BOMCode,T0.Code ) C on C.SubEntry=A.DocEntry and C.U_BOMCode=A.BOMItem"
                strQuery += vbCrLf + "where A.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' "
                strQuery += vbCrLf + "group by A.DocEntry,A.Status,A.LineQty,A.LineItem,A.BOMNorms,A.ReturnQty,A.ReceiptQty,A.HeaderItem,A.ActPlanQty,A.[Excess PlanQty],"
                strQuery += vbCrLf + "A.U_InvUom,A.InvTransferQty,A.IssueQty,A.HeaderQty,A.BOMItem,A.ScrapQty,C.SubEntry,C.[Tot Scrap Qty],C.U_BOMCode,C.Code) B Order by B.OpenQty desc"

                objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objRs.DoQuery(strQuery)
                Dim Flag As Boolean = False
                Dim OpenQty As String = ""
                If objRs.RecordCount > 0 Then
                    EditText13.Value = CDbl(objRs.Fields.Item("OpenQty").Value)
                    EditText5.Value = CDbl(objRs.Fields.Item("Open Scrap Qty").Value)
                    If objRs.Fields.Item("AutoStatus").Value = "C" Then
                        ComboBox2.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Matrix3.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        'Else
                        '    ComboBox2.Select("O", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Status_Closing: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End Try
        End Function

        Private Sub FieldEnabled_CommonSettings()
            Try
                objform.Freeze(True)
                ButtonCombo1.Item.Enabled = False
                Button4.Item.Enabled = False    'GR View
                Button5.Item.Enabled = False    'Inv Tran View
                Button6.Item.Enabled = False    'GI View
                Button9.Item.Enabled = False    'Scrap Return Inv Transfer
                ButtonCombo1.ValidValues.Add("Inventory Transfer", "")
                ButtonCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly   'Input tab Generate Document
                ButtonCombo1.Item.Enabled = False
                ButtonCombo2.ValidValues.Add("Goods Receipt", "")
                ButtonCombo2.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly  'Scrap tab Generate Document
                ButtonCombo2.Item.Enabled = False
                If UDF0 = "N" Then
                    Matrix2.Columns.Item("DocDate").Visible = False
                    Matrix3.Columns.Item("DocDate").Visible = False
                End If
                If UDF1 = "N" Then
                    Matrix2.Columns.Item("RefNo").Visible = False
                    Matrix3.Columns.Item("RefNo").Visible = False
                End If
                If Not WPrice = "Y" Then
                    Matrix0.Columns.Item("tweight").Visible = False
                    Matrix0.Columns.Item("cmblttype").Visible = False
                End If
                If Not RevPlanQty = "Y" Then
                    Matrix0.Columns.Item("revbase").Visible = False
                    Matrix0.Columns.Item("bomqty").Visible = False
                    Matrix0.Columns.Item("PlanQty").Editable = False
                End If
                'If PlanQtyUDFEditable = "N" Then
                '    Matrix0.Columns.Item("PlanQty").Editable = False
                'End If
                If AutoProduction = "Y" Then   'For enabled automatic production order
                    Button7.Item.Enabled = False
                    EditText11.Item.Enabled = False
                    Button8.Item.Visible = True
                    Matrix3.Columns.Item("Price").TitleObject.Caption = "Price"
                Else
                    Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Button7.Item.Visible = False
                    'Folder1.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Button8.Item.Visible = False
                    Matrix3.Columns.Item("Price").TitleObject.Caption = "Item Cost"
                End If
                If ReceiptAutoLoad = "Y" Then  'For enabled automatic Receipt loading in output tab
                    Folder1.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    ButtonCombo0.Item.Visible = False
                    Button8.Item.Enabled = False
                Else
                    Button8.Item.Visible = False
                    ButtonCombo0.ValidValues.Add("Issue & Receipt", "")
                    ButtonCombo0.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly  'Output tab Generate Document
                    ButtonCombo0.Item.Enabled = False
                End If
                If APAutoLoad = "Y" Then   'For enabled automatic AP Invoice loading in Related Doc tab
                    Button3.Item.Visible = False
                    Folder3.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Else
                    Button3.Item.Enabled = False
                End If
                If Costing = "Y" Then  'For disabled costing tab
                    Folder4.Item.Visible = False
                Else
                    'Folder4.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Button10.Item.Enabled = False
                End If
                If POItem = "Y" Then  'For enabled PO item should be a sub item
                    'objform.ActiveItem = "SubPoNum"
                    'EditText24.Item.Enabled = False
                Else
                    EditText17.Item.Visible = False
                    'objform.ActiveItem = "txtcode"
                    EditText24.Item.Enabled = True
                End If
                If ItemBOM = "Y" Then       'For enabled BOM item choosing manually
                    Matrix0.Columns.Item("Code").Editable = False
                    Matrix0.Columns.Item("Quant").Editable = False
                    'Matrix0.Columns.Item("Whse").Editable = False
                    ' Matrix0.Columns.Item("Price").Editable = False
                Else
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                End If

                'If Not TypeInScrap = "Y" Then
                '    Matrix2.Columns.Item("Type").Visible = False
                'End If
                Matrix0.Columns.Item("Item1").Visible = False
                Matrix3.Columns.Item("TabType").Visible = False
                Matrix2.Columns.Item("TabType").Visible = False
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Status_Closing_RowDisable()
            Try
                Dim DocEntry As String = ""
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim UpdateFlag As Boolean = False
                If Folder1.Selected = True Then
                    If ReceiptAutoLoad <> "Y" Then
                        For Code As Integer = 1 To Matrix3.VisualRowCount
                            If Matrix3.Columns.Item("Code").Cells.Item(Code).Specific.string <> "" And Matrix3.Columns.Item("GRNo").Cells.Item(Code).Specific.string <> "" And Matrix3.Columns.Item("GINo").Cells.Item(Code).Specific.string <> "" Then
                                'If Matrix3.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "O" Or Matrix3.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "C" Then
                                If Matrix3.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "O" Then Matrix3.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "C"
                                Matrix3.CommonSetting.SetCellEditable(Code, 2, False) 'ItemCode
                                Matrix3.CommonSetting.SetCellEditable(Code, 5, False) ' Quantity
                                Matrix3.CommonSetting.SetCellEditable(Code, 14, False) 'whse
                                Matrix3.CommonSetting.SetCellEditable(Code, 9, False)
                                Matrix3.CommonSetting.SetCellEditable(Code, 10, False)
                                'End If
                            ElseIf Matrix3.Columns.Item("Code").Cells.Item(Code).Specific.string <> "" And (Matrix3.Columns.Item("GRNo").Cells.Item(Code).Specific.string <> "" Or Matrix3.Columns.Item("GINo").Cells.Item(Code).Specific.string <> "") Then

                                Matrix3.CommonSetting.SetCellEditable(Code, 2, True)

                                'Matrix3.CommonSetting.SetCellEditable(Code, 5, True)
                                'Matrix3.CommonSetting.SetCellEditable(Code, 20, True)
                                'Matrix3.CommonSetting.SetCellEditable(Code, 21, True)
                            Else
                                Matrix3.CommonSetting.SetCellEditable(Code, 2, True)
                                Matrix3.CommonSetting.SetCellEditable(Code, 5, True)
                                Matrix3.CommonSetting.SetCellEditable(Code, 14, True)
                                Matrix3.CommonSetting.SetCellEditable(Code, 9, True)
                                Matrix3.CommonSetting.SetCellEditable(Code, 10, True)
                                'Matrix3.CommonSetting.SetCellEditable(Code, 20, True)
                                'Matrix3.CommonSetting.SetCellEditable(Code, 21, True)
                            End If
                            If objaddon.HANA Then
                                DocEntry = "select Distinct T0.""DocEntry"" from OIGE T0 join ""@MIPL_POR1"" T1 on T1.""DocEntry""=T0.""U_SubConNo""  where T1.""DocEntry""=" & objform.Items.Item("txtentry").Specific.String & " Order by T0.""DocEntry"""
                            Else
                                DocEntry = "select Distinct T0.DocEntry from OIGE T0 join [@MIPL_POR1] T1 on T1.DocEntry=T0.U_SubConNo  where T1.DocEntry=" & objform.Items.Item("txtentry").Specific.String & " Order by T0.DocEntry"
                            End If
                            objRS.DoQuery(DocEntry)
                            If objRS.RecordCount > 0 Then
                                For i As Integer = 0 To objRS.RecordCount - 1
                                    If Matrix3.Columns.Item("GINo").Cells.Item(i + 1).Specific.string = "" Then Matrix3.Columns.Item("GINo").Cells.Item(i + 1).Specific.string = objRS.Fields.Item("DocEntry").Value.ToString : UpdateFlag = True
                                    objRS.MoveNext()
                                Next
                            End If
                            If Matrix3.Columns.Item("GRNo").Cells.Item(Code).Specific.string = "" Then
                                If objaddon.HANA Then
                                    DocEntry = objaddon.objglobalmethods.getSingleValue("select T1.""DocEntry"" from OIGN T0 join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""U_SubConNo""='" & objform.Items.Item("txtentry").Specific.String & "' and T1.""ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(Code).Specific.string & "' and T1.""U_LineID""='" & Matrix3.Columns.Item("#").Cells.Item(Code).Specific.string & "'")
                                Else
                                    DocEntry = objaddon.objglobalmethods.getSingleValue("select T1.DocEntry from OIGN T0 join IGN1 T1 on T0.DocEntry=T1.DocEntry where T0.U_SubConNo='" & objform.Items.Item("txtentry").Specific.String & "' and T1.ItemCode='" & Matrix3.Columns.Item("Code").Cells.Item(Code).Specific.string & "' and T1.U_LineID='" & Matrix3.Columns.Item("#").Cells.Item(Code).Specific.string & "'")
                                End If
                                If Matrix3.Columns.Item("GRNo").Cells.Item(Code).Specific.string = "" Then Matrix3.Columns.Item("GRNo").Cells.Item(Code).Specific.string = DocEntry : UpdateFlag = True
                            End If
                            If UpdateFlag Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            If Matrix3.Columns.Item("Code").Cells.Item(Code).Specific.string <> "" And Matrix3.Columns.Item("GRNo").Cells.Item(Code).Specific.string <> "" Or Matrix3.Columns.Item("GINo").Cells.Item(Code).Specific.string <> "" Then
                                Matrix3.CommonSetting.SetCellEditable(Code, 5, False)
                                Matrix3.CommonSetting.SetCellEditable(Code, 14, False)
                                If Matrix3.Columns.Item("Whse").Cells.Item(Code).Specific.string = "" Then
                                    Matrix3.CommonSetting.SetCellEditable(Code, 14, True)
                                End If
                            End If
                        Next

                    End If
                ElseIf Folder2.Selected = True Then
                    For Code As Integer = 1 To Matrix2.VisualRowCount
                        objcombo = Matrix2.Columns.Item("Type").Cells.Item(Code).Specific
                        If TypeInScrap = "Y" Then
                            'If objCombo.Selected.Description = "Scrap" Then
                            '    Matrix2.CommonSetting.SetCellEditable(Code, 9, True)
                            'ElseIf objCombo.Selected.Description = "Return" Then
                            '    'Matrix2.CommonSetting.SetCellEditable(Code, 9, False)
                            'End If
                            If Matrix2.Columns.Item("Code").Cells.Item(Code).Specific.string <> "" And objcombo.Selected.Description = "Scrap" And Matrix2.Columns.Item("GRNo").Cells.Item(Code).Specific.string <> "" Then
                                If Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "O" Or Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "C" Then
                                    Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "C"
                                    Matrix2.CommonSetting.SetCellEditable(Code, 2, False) 'Type
                                    Matrix2.CommonSetting.SetCellEditable(Code, 3, False) 'ItemCode
                                    Matrix2.CommonSetting.SetCellEditable(Code, 6, False) 'Quantity
                                    Matrix2.CommonSetting.SetCellEditable(Code, 9, False) 'Whse
                                    'Matrix2.CommonSetting.SetCellEditable(Code, 16, False)
                                    'Matrix2.CommonSetting.SetCellEditable(Code, 17, False)
                                    'Matrix2.CommonSetting.SetCellEditable(Code, 9, False)
                                End If
                            ElseIf Matrix2.Columns.Item("Code").Cells.Item(Code).Specific.string <> "" And objcombo.Selected.Description = "Return" And Matrix2.Columns.Item("InvNum").Cells.Item(Code).Specific.string <> "" Then
                                If Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "O" Or Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "C" Then
                                    Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "C"
                                    Matrix2.CommonSetting.SetCellEditable(Code, 2, False)
                                    Matrix2.CommonSetting.SetCellEditable(Code, 3, False)
                                    Matrix2.CommonSetting.SetCellEditable(Code, 6, False)
                                    Matrix2.CommonSetting.SetCellEditable(Code, 9, False)
                                    'Matrix2.CommonSetting.SetCellEditable(Code, 16, False)
                                    'Matrix2.CommonSetting.SetCellEditable(Code, 17, False)
                                    'Matrix2.CommonSetting.SetCellEditable(Code, 9, False)
                                End If
                            Else
                                Matrix2.CommonSetting.SetCellEditable(Code, 2, True)
                                Matrix2.CommonSetting.SetCellEditable(Code, 3, True)
                                Matrix2.CommonSetting.SetCellEditable(Code, 6, True)
                                Matrix2.CommonSetting.SetCellEditable(Code, 9, True)
                                'Matrix2.CommonSetting.SetCellEditable(Code, 16, True)
                                'Matrix2.CommonSetting.SetCellEditable(Code, 17, True)
                            End If
                        ElseIf AutoProduction = "Y" Then
                            If Matrix2.Columns.Item("Code").Cells.Item(Code).Specific.string <> "" And Matrix2.Columns.Item("InvNum").Cells.Item(Code).Specific.string <> "" Then
                                If Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "O" Or Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "C" Then
                                    Matrix2.Columns.Item("Colstat").Cells.Item(Code).Specific.string = "C"
                                    Matrix2.CommonSetting.SetCellEditable(Code, 3, False)
                                    Matrix2.CommonSetting.SetCellEditable(Code, 6, False)
                                    'Matrix2.CommonSetting.SetCellEditable(Code, 9, False)
                                End If
                            Else
                                Matrix2.CommonSetting.SetCellEditable(Code, 3, True)
                                Matrix2.CommonSetting.SetCellEditable(Code, 6, True)
                                Matrix2.CommonSetting.SetCellEditable(Code, 9, True)
                            End If
                        End If
                        'Matrix2.CommonSetting.SetCellEditable(Code, 18, True)
                        'Matrix2.CommonSetting.SetCellEditable(Code, 16, True)
                        'Matrix2.CommonSetting.SetCellEditable(Code, 17, True)
                    Next
                ElseIf Folder4.Selected = True Then
                End If
                objRS = Nothing
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True And pVal.BeforeAction = False Then
                    If AddonVersion = "" Then Exit Sub
                    objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If objaddon.HANA Then
                        objRs.DoQuery("update ""@MIPL_OPOR"" set ""U_VersionNum""='" & AddonVersion & "'  where ""DocEntry""='" & odbdsHeader.GetValue("DocEntry", 0) & "'")
                    Else
                        objRs.DoQuery("update [@MIPL_OPOR] set U_VersionNum='" & AddonVersion & "'  where DocEntry='" & odbdsHeader.GetValue("DocEntry", 0) & "'")
                    End If
                    objRs = Nothing
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("SUBCTPO", pVal.FormTypeCount)
            Catch ex As Exception
            End Try

        End Sub

        Public Sub Payslip_AutoEmail_Test()
            Try
                Dim strsql As String = ""
                Dim objrs As SAPbobsCOM.Recordset
                'Dim objrsupdate As SAPbobsCOM.Recordset
                Dim Mailbody As String = ""

                Dim Email As New System.Net.Mail.MailMessage
                Dim MailServer As New System.Net.Mail.SmtpClient()

                Try
                    MailServer.Host = "smtp.office365.com" 'Mail_Host "smtp-mail.outlook.com"
                    MailServer.Port = "587" 'Mail_Port
                    MailServer.UseDefaultCredentials = False
                    MailServer.Credentials = New System.Net.NetworkCredential("saptech18@mukeshinfoserve.com", "tyntwdjfwlztgrns")
                    MailServer.EnableSsl = True
                    'MailServer.DeliveryMethod = SmtpDeliveryMethod.Network
                    Email.From = New System.Net.Mail.MailAddress("saptech18@mukeshinfoserve.com")

                    Email.To.Add(New System.Net.Mail.MailAddress("saptech18@mukeshinfoserve.com"))
                    Email.Subject = "Pay Slip - "

                    Mailbody = "Dear ,"
                    Mailbody += vbCrLf + " "
                    Mailbody += vbCrLf + " Please Find the Attached Payslip for the Month of April."
                    Mailbody += vbCrLf + " "
                    Mailbody += vbCrLf + "With Regards,"
                    Mailbody += vbCrLf + "HR Team"
                    Mailbody += vbCrLf + " "
                    Mailbody += vbCrLf + " "
                    Mailbody += "This is Auto generated E-Mail from SAP Business One . Please do not reply to this message. Thank you! "

                    Email.Body = Mailbody
                    Email.Priority = Net.Mail.MailPriority.High


                    MailServer.Send(Email)



                Catch ex As Exception
                    MsgBox("Mail: " + ex.Message.ToString)
                Finally
                    If Not Email Is Nothing Then Email.Dispose()
                    MailServer = Nothing
                End Try
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

        End Sub

        Private WithEvents ComboBox2 As SAPbouiCOM.ComboBox

        Private Sub Form_DataLoadBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    Dim cc As String = objform.Items.Item("txtentry").Specific.String
                    If cc = "CPPL-SD" Then
                        BubbleEvent = False
                    End If
                End If
                'BubbleEvent = False : objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Matrix0.AutoResizeColumns()
                Matrix1.AutoResizeColumns()
                Matrix2.AutoResizeColumns()
                Matrix3.AutoResizeColumns()
                Matrix4.AutoResizeColumns()
            Catch ex As Exception

            End Try
        End Sub

        Private WithEvents StaticText21 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText


        'Public Sub sendEMailThroughOUTLOOK()
        '    Try
        '        ' Create the Outlook application.
        '        Dim oApp As New Outlook.Application()
        '        ' Create a new mail item.
        '        Dim oMsg As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        '        oMsg.Sender.Address = "chitra.k@altrockstech.com"
        '        ' Set HTMLBody.
        '        'add the body of the email
        '        oMsg.HTMLBody = "Hello, Chitra your message body will go here!!"
        '        'Add an attachment.
        '        Dim sDisplayName As [String] = "MyAttachment"
        '        Dim iPosition As Integer = CInt(oMsg.Body.Length) + 1
        '        Dim iAttachType As Integer = CInt(Outlook.OlAttachmentType.olByValue)
        '        'now attached the file
        '        Dim oAttach As Outlook.Attachment = oMsg.Attachments.Add("E:\OneDrive - MIPL\E Drive\CDS\Deliverysample.pdf", iAttachType, iPosition, sDisplayName)
        '        'Subject line
        '        oMsg.Subject = "Your Subject will go here."
        '        ' Add a recipient.
        '        Dim oRecips As Outlook.Recipients = DirectCast(oMsg.Recipients, Outlook.Recipients)

        '        ' Change the recipient in the next line if necessary.
        '        Dim oRecip As Outlook.Recipient = DirectCast(oRecips.Add("chitra.k@altrockstech.com"), Outlook.Recipient)
        '        oRecip.Resolve()
        '        ' Send.
        '        oMsg.Send()
        '        ' Clean up.
        '        oRecip = Nothing
        '        oRecips = Nothing
        '        oMsg = Nothing
        '        oApp = Nothing
        '        'end of try block
        '    Catch ex As Exception
        '    End Try
        '    'end of catch
        'End Sub

#End Region

    End Class

End Namespace
