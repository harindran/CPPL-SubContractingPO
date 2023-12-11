Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO

    <FormAttribute("SUBBOM", "SubContractingPO/SubContractBOM.b1f")>
    Friend Class SubContract
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Private WithEvents objDBHeader As SAPbouiCOM.DBDataSource
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Private WithEvents odbdsDetails1 As SAPbouiCOM.DBDataSource
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText1 = CType(Me.GetItem("LQty").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("txtqty").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("ltype").Specific, SAPbouiCOM.StaticText)
            Me.StaticText3 = CType(Me.GetItem("lwhse").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("whse").Specific, SAPbouiCOM.EditText)
            Me.StaticText5 = CType(Me.GetItem("lprodno").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("txtcode").Specific, SAPbouiCOM.EditText)
            Me.StaticText6 = CType(Me.GetItem("lproddesc").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("txtname").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("mtxBOM").Specific, SAPbouiCOM.Matrix)
            Me.ComboBox0 = CType(Me.GetItem("Bomtype").Specific, SAPbouiCOM.ComboBox)
            Me.LinkedButton0 = CType(Me.GetItem("lnkcode").Specific, SAPbouiCOM.LinkedButton)
            Me.EditText2 = CType(Me.GetItem("txtprodno").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("fldritem").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("fldrproc").Specific, SAPbouiCOM.Folder)
            Me.Matrix1 = CType(Me.GetItem("mtxproc").Specific, SAPbouiCOM.Matrix)
            Me.CheckBox0 = CType(Me.GetItem("chkact").Specific, SAPbouiCOM.CheckBox)
            Me.EditText4 = CType(Me.GetItem("txtproc").Specific, SAPbouiCOM.EditText)
            Me.EditText7 = CType(Me.GetItem("txtprocode").Specific, SAPbouiCOM.EditText)
            Me.StaticText0 = CType(Me.GetItem("lentry").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("txtentry").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                'objform = objaddon.objapplication.Forms.GetForm("SUBBOM", Me.FormCount)
                'objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                Dim objRs As SAPbobsCOM.Recordset
                objform.Items.Item("mtxBOM").Enabled = False
                objform.Items.Item("mtxproc").Enabled = False
                objform.Items.Item("lentry").Visible = False
                objform.Items.Item("Item_0").Visible = False
                objform.Items.Item("txtentry").Visible = False
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtcode", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtname", False, True, False)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    objRs.DoQuery("select * from ""@MIPL_GEN""")
                Else
                    objRs.DoQuery("select * from [@MIPL_GEN]")
                End If
                If objRs.RecordCount > 0 Then
                    ResourceEnable = objRs.Fields.Item("U_ResEn").Value.ToString
                    ItemProcess = objRs.Fields.Item("U_Process").Value.ToString
                    ProcTypefromUDO = objRs.Fields.Item("U_PType").Value.ToString
                End If
                'If objaddon.HANA Then
                '    ResourceEnable = objaddon.objglobalmethods.getSingleValue("select ""U_ResEn"" from ""@MIPL_GEN"" where ifnull(""U_ResEn"",'')='Y'")
                'Else
                '    ResourceEnable = objaddon.objglobalmethods.getSingleValue("select U_ResEn from [@MIPL_GEN] where isnull(U_ResEn,'')='Y'")
                'End If

                Dim cmbtype As SAPbouiCOM.Column = Matrix0.Columns.Item("Type")
                If ResourceEnable <> "Y" Then
                    cmbtype.ValidValues.Remove("290", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                If ItemProcess = "N" Then
                    Folder1.Item.Visible = False
                End If
                If ProcTypefromUDO = "Y" Then
                    ComboBox0.Item.Visible = False
                Else
                    EditText4.Item.Visible = False
                    ComboBox0.Item.Left = EditText6.Item.Left '157
                    ComboBox0.Item.Top = StaticText2.Item.Top '50
                    'ComboBox0.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
                EditText7.Item.Visible = False
                CheckBox0.Item.Height = CheckBox0.Item.Height + 2
                If Link_Value <> "" And Link_objtype = "SUBBOM" Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    EditText5.Item.Enabled = True
                    'EditText0.Item.Enabled = True
                    EditText5.Value = Link_Value
                    ' EditText0.Value = objaddon.objglobalmethods.getSingleValue("Select ""DocEntry"" from ""@MIPL_OBOM"" where ""Code""='" & Link_Value & "'")
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    objform.ActiveItem = "txtqty"
                    EditText5.Item.Enabled = False
                    'EditText0.Item.Enabled = False
                    objaddon.objapplication.Menus.Item("1300").Activate()
                    Link_Value = "-1" : Link_objtype = "-1"
                    Folder0.Item.Click()
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    Exit Sub
                End If
                objDBHeader = objform.DataSources.DBDataSources.Item("@MIPL_OBOM")
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_BOM1")
                odbdsDetails1 = objform.DataSources.DBDataSources.Item("@MIPL_BOM2")
                objform.ActiveItem = "txtcode"
                'If objaddon.HANA Then
                '    EditText0.Value = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 ""DocEntry"" from ""@MIPL_OBOM""")
                'Else
                '    EditText0.Value = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 DocEntry from [@MIPL_OBOM]")
                'End If

                'Matrix0.AddRow()
                'Matrix0.Columns.Item("LineId").Cells.Item(1).Specific.String = 1
                objform.Settings.Enabled = True
                objform.Freeze(False)
                Matrix0.AutoResizeColumns()
                Matrix1.AutoResizeColumns()
                Folder0.Item.Click()
                'objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception
                'objform.Freeze(False)
            Finally
                objform.Freeze(False)
            End Try


        End Sub

#Region "Field Details"
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents Matrix1 As SAPbouiCOM.Matrix
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
#End Region

#Region "LostFocus Events"

        Private Sub EditText5_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText5.LostFocusAfter
            'Try
            '    If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
            '    Dim status As String, Itemcode As String
            '    If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '        If EditText5.Value <> "" Then
            '            If objaddon.HANA Then
            '                status = objaddon.objglobalmethods.getSingleValue("select 1 ""Status"" from ""@MIPL_OBOM"" where ""Code""='" & EditText5.Value & "' ")
            '            Else
            '                status = objaddon.objglobalmethods.getSingleValue("select 1 Status from [@MIPL_OBOM] where Code='" & EditText5.Value & "' ")
            '            End If
            '            If status = "1" Then
            '                objform.Freeze(True)
            '                Itemcode = EditText5.Value
            '                objaddon.objapplication.Menus.Item("1281").Activate()
            '                'objaddon.objapplication.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
            '                EditText5.Value = Itemcode
            '                objform.Items.Item("1").Click()
            '                objform.Freeze(False)
            '            Else
            '                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "ItemNo", "#")
            '            End If
            '        End If
            '    Else
            '        objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "ItemNo", "#")
            '    End If
            'Catch ex As Exception
            '    objform.Freeze(False)
            'End Try
        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try
                If pVal.ActionSuccess = True Then
                    Dim Quantity As Double, Price As Double, unitprice As Double
                    If Matrix0.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.String <> "" Then
                        Quantity = CDbl(Matrix0.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.String)
                    Else
                        Quantity = 0
                    End If
                    If Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String <> "" Then
                        Price = CDbl(Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String)
                    Else
                        Price = 0
                    End If
                    Select Case pVal.ColUID
                        Case "ItemNo"
                            If Matrix0.Columns.Item("ItemNo").Cells.Item(pVal.Row).Specific.String <> "" Then
                                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "ItemNo", "#")
                                If objaddon.HANA Then
                                    unitprice = objaddon.objglobalmethods.getSingleValue("select ""Price"" from  itm1 where ""ItemCode""='" & Matrix0.Columns.Item("ItemNo").Cells.Item(pVal.Row).Specific.String & "' and ""PriceList""=1")
                                Else
                                    unitprice = objaddon.objglobalmethods.getSingleValue("select Price from  itm1 where ItemCode='" & Matrix0.Columns.Item("ItemNo").Cells.Item(pVal.Row).Specific.String & "' and PriceList=1")
                                End If
                                If unitprice <> "" Then
                                    Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String = unitprice
                                Else
                                    Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String = 0
                                End If
                            End If
                        Case "Qty"
                            objform.Freeze(True)
                            If Matrix0.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.String <> "" Then
                                Matrix0.Columns.Item("Total").Cells.Item(pVal.Row).Specific.String = Quantity * Price
                            End If
                            objaddon.objapplication.Menus.Item("1300").Activate()
                            objform.Freeze(False)
                        Case "Price"
                            objform.Freeze(True)
                            If Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String <> "" Then
                                Matrix0.Columns.Item("Total").Cells.Item(pVal.Row).Specific.String = Quantity * Price
                            End If
                            objaddon.objapplication.Menus.Item("1300").Activate()
                            objform.Freeze(False)
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub EditText6_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText6.LostFocusAfter
            Dim status As String, Itemcode As String

            If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then Exit Sub
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If EditText5.Value <> "" Then
                        If objaddon.HANA Then
                            status = objaddon.objglobalmethods.getSingleValue("select 1 ""Status"" from ""@MIPL_OBOM"" where ""Code""='" & EditText5.Value & "' ")
                        Else
                            status = objaddon.objglobalmethods.getSingleValue("select 1 Status from [@MIPL_OBOM] where Code='" & EditText5.Value & "' ")
                        End If

                        If status = "1" Then
                            objform.Freeze(True)
                            Itemcode = EditText5.Value
                            objaddon.objapplication.Menus.Item("1281").Activate()
                            objaddon.objapplication.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                            EditText5.Value = Itemcode
                            objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objform.Freeze(False)
                        Else
                            Exit Sub
                        End If
                    End If
                    'RemoveLastrow(Matrix0, "ItemNo")
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub
#End Region

#Region "ChooseFromList Events"
        Private Sub EditText5_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText5.ChooseFromListBefore
            '            If pVal.ActionSuccess = True Then Exit Sub
            'Try
            '    Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_Item")
            '    Dim oConds As SAPbouiCOM.Conditions
            '    Dim oCond As SAPbouiCOM.Condition
            '    Dim oEmptyConds As New SAPbouiCOM.Conditions
            '    oCFL.SetConditions(oEmptyConds)
            '    oConds = oCFL.GetConditions()

            '    oCond = oConds.Add()
            '    oCond.Alias = "U_SubCont"
            '    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCond.CondVal = "Y"
            '    ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        

            '    oCFL.SetConditions(oConds)
            'Catch ex As Exception
            '    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'End Try

        End Sub

        Private Sub EditText5_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText5.ChooseFromListAfter
            'Try
            '    If pVal.ActionSuccess = False Then Exit Sub
            '    pCFL = pVal
            '    If Not pCFL.SelectedObjects Is Nothing Then
            '        Try
            '            EditText5.Value = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
            '        Catch ex As Exception
            '        End Try
            '        Try
            '            EditText6.Value = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
            '        Catch ex As Exception
            '        End Try
            '        Try
            '            EditText1.Value = "1"
            '        Catch ex As Exception
            '        End Try
            '    End If
            'Catch ex As Exception
            'Finally
            '    GC.Collect()
            '    GC.WaitForPendingFinalizers()
            'End Try

        End Sub

        Private Sub Matrix0_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ChooseFromListAfter
            Try
                If pVal.ColUID = "ItemNo" And pVal.ActionSuccess = True Then
                    Try
                        Dim cmbtype As SAPbouiCOM.ComboBox = Matrix0.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                        Dim UnitPrice As String = ""
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal

                        If Not pCFL.SelectedObjects Is Nothing Then
                            If cmbtype.Selected.Value = "4" Then
                                Try
                                    Matrix0.Columns.Item("ItemNo").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                                Catch ex As Exception
                                End Try
                                Try
                                    Matrix0.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                                Catch ex As Exception
                                End Try
                                Try
                                    Matrix0.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                                Catch ex As Exception
                                End Try

                            ElseIf cmbtype.Selected.Value = "290" Then
                                Try
                                    Matrix0.Columns.Item("ItemNo").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ResCode").Cells.Item(0).Value
                                Catch ex As Exception
                                End Try
                                Try
                                    Matrix0.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ResName").Cells.Item(0).Value
                                Catch ex As Exception
                                End Try
                                Try
                                    Matrix0.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("InvntryUom").Cells.Item(0).Value
                                Catch ex As Exception
                                End Try
                                If objaddon.HANA Then
                                    UnitPrice = objaddon.objglobalmethods.getSingleValue("select (""StdCost1""+""StdCost2""+""StdCost3""+""StdCost4""+""StdCost5""+""StdCost6""+""StdCost7""+""StdCost8""+""StdCost9""+""StdCost10"") as ""Cost"" from ORSC where ""ResCode""='" & pCFL.SelectedObjects.Columns.Item("ResCode").Cells.Item(0).Value & "'")
                                Else
                                    UnitPrice = objaddon.objglobalmethods.getSingleValue("select (StdCost1+StdCost2+StdCost3+StdCost4+StdCost5+StdCost6+StdCost7+StdCost8+StdCost9+StdCost10) as Cost from ORSC where ResCode='" & pCFL.SelectedObjects.Columns.Item("ResCode").Cells.Item(0).Value & "'")
                                End If

                                Try
                                    Matrix0.Columns.Item("Price").Cells.Item(pVal.Row).Specific.String = UnitPrice
                                Catch ex As Exception
                                End Try
                            End If
                            Matrix0.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.String = "1"
                            Matrix0.Columns.Item("Total").Cells.Item(pVal.Row).Specific.String = UnitPrice
                        End If

                    Catch ex As Exception
                    End Try

                ElseIf pVal.ColUID = "Whse" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("Whse").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                Matrix0.AutoResizeColumns()
            Catch ex As Exception
            End Try


        End Sub

        Private Sub EditText3_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText3.ChooseFromListAfter

            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText3.Value = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub Matrix0_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix0.ChooseFromListBefore
            Dim cmbtype As SAPbouiCOM.ComboBox = Matrix0.Columns.Item("Type").Cells.Item(pVal.Row).Specific
            Dim ColItem As SAPbouiCOM.Column = Matrix0.Columns.Item("ItemNo")
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Try
                If cmbtype.Selected.Value = "4" Then
                    ColItem.ChooseFromListUID.Remove(pVal.Row)
                    ColItem.ChooseFromListUID = "CFL_I"
                    ColItem.ChooseFromListAlias = "ItemCode"
                    oCFL = objform.ChooseFromLists.Item("CFL_I")
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()
                    oCond = oConds.Add()
                    oCond.Alias = "InvntItem"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "Y"
                    oCFL.SetConditions(oConds)
                ElseIf cmbtype.Selected.Value = "290" Then
                    ColItem.ChooseFromListUID.Remove(pVal.Row)
                    ColItem.ChooseFromListUID = "CFL_R"
                    ColItem.ChooseFromListAlias = "ResCode"
                End If
            Catch ex As Exception

            End Try

        End Sub

#End Region

        Private Sub Button0_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            If pVal.ActionSuccess = True And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                'If objaddon.HANA Then
                '    EditText0.Value = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 ""DocEntry"" from ""@MIPL_OBOM""")
                'Else
                '    EditText0.Value = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 DocEntry from [@MIPL_OBOM]")
                'End If
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "ItemNo", "#")
                objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "procc", "#")
            End If
        End Sub

        Private Sub Addrow()
            Try
                'If Matrix0.VisualRowCount > 0 And EditText0.Value <> "" Then
                '    'If odbdsDetails.GetValue("U_Itemcode", Matrix0.VisualRowCount - 1) = "" Then Exit Sub
                '    'objform.Freeze(True)
                '    'odbdsDetails.InsertRecord(odbdsDetails.Size)
                '    'odbdsDetails.SetValue("LineId", Matrix0.VisualRowCount, Matrix0.VisualRowCount + 1)
                '    'Matrix0.LoadFromDataSource()
                '    'objform.Freeze(False)
                'End If
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

                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Matrix0.FlushToDataSource()
                For i As Integer = 1 To Matrix0.VisualRowCount
                    Matrix0.GetLineData(i)
                    odbdsDetails.SetValue("Code", i - 1, EditText5.Value.ToString)
                    Matrix0.SetLineData(i)
                Next
                Matrix0.LoadFromDataSource()
                Matrix1.FlushToDataSource()
                For i As Integer = 1 To Matrix1.VisualRowCount
                    Matrix1.GetLineData(i)
                    odbdsDetails1.SetValue("Code", i - 1, EditText5.Value.ToString)
                    Matrix1.SetLineData(i)
                Next
                Matrix1.LoadFromDataSource()
                RemoveLastrow(Matrix0, "ItemNo")
                RemoveLastrow(Matrix1, "procc")
                'If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '    RemoveLastrow(Matrix0, "ItemNo")
                '    RemoveLastrow(Matrix1, "procc")
                'End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If EditText5.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("ItemCode is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    'If EditText1.Value = "" Then
                    '    EditText1.Value = "0"
                    'End If
                    If Val(EditText1.Value) <= 0 Then
                        objaddon.objapplication.SetStatusBarMessage("Quantity is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    If EditText3.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Warehouse is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If

                    If Matrix0.Columns.Item("ItemNo").Cells.Item(1).Specific.string = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Line Data is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If Matrix0.Columns.Item("ItemNo").Cells.Item(i).Specific.string <> "" Then
                            If Matrix0.Columns.Item("Qty").Cells.Item(i).Specific.string = "" Or Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.string = "" Or Matrix0.Columns.Item("Price").Cells.Item(i).Specific.string = "" Then
                                objaddon.objapplication.SetStatusBarMessage("One of the column value is missing.Please update", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                            If CDbl(Matrix0.Columns.Item("Qty").Cells.Item(i).Specific.string) = 0 Then
                                objaddon.objapplication.SetStatusBarMessage("Qty Field should not be zero. in Item Tab-Line " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next
                    If Matrix0.VisualRowCount <= 0 Then
                        objaddon.objapplication.SetStatusBarMessage("Please fill data in Row level", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Matrix0_LinkPressedBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix0.LinkPressedBefore
            Dim cmbtype As SAPbouiCOM.ComboBox = Matrix0.Columns.Item("Type").Cells.Item(pVal.Row).Specific
            Dim ColItem As SAPbouiCOM.Column = Matrix0.Columns.Item("ItemNo")
            Dim objlink As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
            Try
                If cmbtype.Selected.Value = "4" Then
                ElseIf cmbtype.Selected.Value = "290" Then
                    objlink.LinkedObjectType = "290"
                    objlink.Item.LinkTo = "ItemNo"
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                Dim IDFlag As Boolean = False
                If Matrix0.VisualRowCount > 0 Then
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If CInt(Matrix0.Columns.Item("#").Cells.Item(i).Specific.string) <> i Then
                            IDFlag = True
                            Matrix0.Columns.Item("#").Cells.Item(i).Specific.string = CStr(i)
                        End If
                    Next
                End If
                RemoveLastrow(Matrix0, "ItemNo")
                'objaddon.objapplication.Menus.Item("1300").Activate()
                If IDFlag Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
                Matrix0.AutoResizeColumns()
            Catch ex As Exception
            End Try
        End Sub

        Private Sub EditText2_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText2.LostFocusAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If EditText2.Value <> "" Then
                    If ProcTypefromUDO = "Y" Then
                        If EditText7.Value <> "" Then
                            EditText5.Value = EditText2.Value + "-" + EditText7.Value
                        Else
                            EditText5.Value = EditText2.Value
                        End If
                    Else
                        If ComboBox0.Selected.Value <> "-" Then
                            EditText5.Value = EditText2.Value + "-" + ComboBox0.Selected.Value
                        Else
                            EditText5.Value = EditText2.Value
                        End If
                    End If
                End If
                'Try
                '    Dim status As String, Itemcode As String
                '    If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                '        If EditText5.Value <> "" Then
                '            If objaddon.HANA Then
                '                status = objaddon.objglobalmethods.getSingleValue("select 1 ""Status"" from ""@MIPL_OBOM"" where ""Code""='" & EditText5.Value & "' ")
                '            Else
                '                status = objaddon.objglobalmethods.getSingleValue("select 1 Status from [@MIPL_OBOM] where Code='" & EditText5.Value & "' ")
                '            End If
                '            If status = "1" Then
                '                objform.Freeze(True)
                '                Itemcode = EditText5.Value
                '                objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                '                'objaddon.objapplication.Menus.Item("1281").Activate()
                '                EditText5.Value = Itemcode
                '                objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                '                objform.Freeze(False)
                '            Else
                '                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "ItemNo", "#")
                '            End If
                '        End If
                '    Else
                '        objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "ItemNo", "#")
                '    End If
                'Catch ex As Exception
                '    objform.Freeze(False)
                'End Try
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText2_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText2.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_Item")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "U_SubCont"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add()
                oCond.Alias = "InvntItem"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"

                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub EditText2_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText2.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        'EditText2.Value = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                        objDBHeader.SetValue("U_ItemCode", 0, pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        'EditText6.Value = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                        objDBHeader.SetValue("Name", 0, pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        EditText1.Value = "1"
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub ComboBox0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                'If ItemProcess = "Y" Then
                If EditText2.Value <> "" Then
                    'EditText5.Value = EditText2.Value + "-" + ComboBox0.Selected.Value
                    If ComboBox0.Selected.Value <> "-" Then
                        EditText5.Value = EditText2.Value + "-" + ComboBox0.Selected.Value
                    Else
                        EditText5.Value = EditText2.Value
                    End If
                End If
                'End If

            Catch ex As Exception

            End Try


        End Sub

        Private Sub Matrix0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ClickAfter
            Try
                If pVal.Row = 0 Then
                    Matrix0.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                Else
                    Matrix0.SelectRow(pVal.Row, True, False)
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix1_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.ChooseFromListAfter
            Try
                Dim cmbprocseq As SAPbouiCOM.ComboBox
                If pVal.ColUID = "procc" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix1.Columns.Item("procc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix1.Columns.Item("procc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value
                            End Try
                            Try
                                Matrix1.Columns.Item("procn").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix1.Columns.Item("procn").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value
                            End Try
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "procc", "#")
                            cmbprocseq = Matrix1.Columns.Item("procseq").Cells.Item(pVal.Row).Specific
                            cmbprocseq.Select(Matrix1.Columns.Item("#").Cells.Item(pVal.Row).Specific.String, SAPbouiCOM.BoSearchKey.psk_ByValue)

                        End If
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Folder1_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder1.PressedAfter
            Try
                objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "procc", "#")
                Matrix1.AutoResizeColumns()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix1_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.ClickAfter
            Try
                If pVal.Row = 0 Then
                    Matrix1.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                Else
                    Matrix1.SelectRow(pVal.Row, True, False)
                End If
            Catch ex As Exception
            End Try

        End Sub


        Private Sub EditText4_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText4.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText7.Value = pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                    Try
                        EditText4.Value = pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
                If EditText2.Value <> "" Then
                    If EditText7.Value <> "" Then
                        EditText5.Value = EditText2.Value + "-" + EditText7.Value
                    Else
                        EditText5.Value = EditText2.Value
                    End If
                End If

            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub Matrix1_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.LostFocusAfter
            Try
                If pVal.ActionSuccess = True Then
                    Select Case pVal.ColUID
                        Case "procc"
                            If Matrix1.Columns.Item("procc").Cells.Item(pVal.Row).Specific.String <> "" Then
                                objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "procc", "#")
                            End If
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                If ProcTypefromUDO = "Y" Then
                    ComboBox0.Item.Visible = False
                Else
                    EditText4.Item.Visible = False
                    ComboBox0.Item.Left = EditText6.Item.Left '157
                    ComboBox0.Item.Top = StaticText2.Item.Top '50
                End If
                Matrix0.AutoResizeColumns()
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_DoubleClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.DoubleClickBefore
            Try
                BubbleEvent = False
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("SUBBOM", pVal.FormTypeCount)
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Matrix0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.KeyDownAfter
            Try
                If pVal.ActionSuccess = True Then
                    Select Case pVal.ColUID
                        Case "Qty"
                            If CDbl(Matrix0.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.String) < 0 Then
                                objaddon.objapplication.StatusBar.SetText("In ""Qty"" (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Matrix0.Columns.Item("Qty").Cells.Item(pVal.Row).Click()
                                Matrix0.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.String = "1"
                            End If
                    End Select
                End If
            Catch ex As Exception
            End Try

        End Sub
    End Class

End Namespace