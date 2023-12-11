Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("GetList", "SubContractingPO/FrmGetList.b1f")>
    Friend Class FrmGetList
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Dim StrQuery As String
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("gridlist").Specific, SAPbouiCOM.Grid)
            Me.Button2 = CType(Me.GetItem("btnclr").Specific, SAPbouiCOM.Button)
            Me.Button3 = CType(Me.GetItem("btnsel").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("GetList", 0)

            Catch ex As Exception

            End Try

        End Sub

        Public Sub Load_SubContracting(ByVal CardCode As String)
            Try
                objaddon.objapplication.StatusBar.SetText("Loading Sub Contracting lists. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If objaddon.HANA Then
                    StrQuery = "Select * from "
                    StrQuery += vbCrLf + "(Select T0.""DocEntry"",T0.""DocNum"",T0.U_DocDate ""Document Date"",T0.""U_CardCode"" ""BP Code"",T0.""U_CardName"" ""BP Name"",T0.""U_SItemCode"" ""Parent Item Code"",T0.""U_SItemDesc"" ""Parent Item Desc"",T0.""U_Process"" ""Process"","
                    StrQuery += vbCrLf + "T0.""U_SQty"" ""Planned Qty"",T1.""U_Itemcode"" ""Input ItemCode"",T1.""U_ItemDesc ""Input ItemDesc"",T1.""U_PlanQty"" ""Input Planned Qty""," 'T1.""U_Qty"" ""Base Qty"",T1.""U_ActPlanQty"" ""Actual Planned Qty"",
                    StrQuery += vbCrLf + "(Select ifnull(Sum(B.""Quantity""),0) from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where (B.""U_S_DocEntry""=T1.""DocEntry"" or A.""U_SubConNo=T1.DocEntry"") and B.""ItemCode""=T1.""U_Itemcode"") as ""Transfer Qty"""
                    StrQuery += vbCrLf + ",(Select case when sum(B.""Quantity"")>0 then T1.""U_PlanQty""-ifnull(Sum(B.""Quantity""),0) else T1.""U_PlanQty"" end from OWTR A join WTR1 B on A.""DocEntry""=B.""DocEntry"" where (B.""U_S_DocEntry""=T1.""DocEntry"" or A.""U_SubConNo=T1.DocEntry"") and B.""ItemCode""=T1.""U_Itemcode"") as ""Pending Qty"","
                    StrQuery += vbCrLf + "T1.""LineId"" ""Input Line"",T1.""U_WhsCode"" ""From Whse"",(Select ""OnHand"" from OITW where ""ItemCode""=T1.""U_Itemcode"" and ""WhsCode""=T1.""U_WhsCode"") ""In Stock"",T1.""U_SubWhse"" ""To Whse"""
                    StrQuery += vbCrLf + "from ""@MIPL_OPOR"" T0 join ""@MIPL_POR1"" T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""Status""='O' and T0.""U_CardCode""='" & CardCode & "') A "
                    StrQuery += vbCrLf + "where A.""Pending Qty"">0 order by A.""DocEntry"" "
                Else
                    StrQuery = "Select * from "
                    StrQuery += vbCrLf + "(Select T0.DocEntry,T0.DocNum,T0.U_DocDate [Document Date],T0.U_CardCode [BP Code],T0.U_CardName [BP Name],T0.U_SItemCode [Parent Item Code],T0.U_SItemDesc [Parent Item Desc],T0.U_Process [Process],"
                    StrQuery += vbCrLf + "T0.U_SQty [Planned Qty],T1.U_Itemcode [Input ItemCode],T1.U_ItemDesc [Input ItemDesc],T1.U_PlanQty [Input Planned Qty]," 'T1.U_Qty [Base Qty],T1.U_ActPlanQty [Actual Planned Qty],
                    StrQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where (B.U_S_DocEntry=T1.DocEntry or A.U_SubConNo=T1.DocEntry) and B.ItemCode=T1.U_Itemcode) as [Transfer Qty]"
                    StrQuery += vbCrLf + ",(Select case when sum(B.Quantity)>0 then T1.U_PlanQty-isnull(Sum(B.Quantity),0) else T1.U_PlanQty end from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where (B.U_S_DocEntry=T1.DocEntry or A.U_SubConNo=T1.DocEntry) and B.ItemCode=T1.U_Itemcode) as [Pending Qty],"
                    StrQuery += vbCrLf + "T1.LineId [Input Line],T1.U_WhsCode [From Whse],(Select OnHand from OITW where ItemCode=T1.U_Itemcode and WhsCode=T1.U_WhsCode) [In Stock],T1.U_SubWhse [To Whse]"
                    StrQuery += vbCrLf + "from [@MIPL_OPOR] T0 join [@MIPL_POR1] T1 on T0.DocEntry=T1.DocEntry where T0.Status='O' and T0.U_CardCode='" & CardCode & "') A "
                    StrQuery += vbCrLf + "where A.[Pending Qty]>0 order by A.DocEntry "
                End If
                objform.DataSources.DataTables.Item("DT_List").ExecuteQuery(StrQuery)
                Grid0.DataTable = objform.DataSources.DataTables.Item("DT_List")
                Grid0.RowHeaders.TitleObject.Caption = "#"
                objform.Freeze(True)
                For i As Integer = 0 To Grid0.Columns.Count - 1
                    'If i = 0 Then Grid0.Columns.Item(i).Visible = False : Continue For
                    Grid0.Columns.Item("Input Line").Visible = False
                    Grid0.Columns.Item(i).Editable = False
                Next
                Grid0.Rows.SelectedRows.Add(0)
                Grid0.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                Dim col As SAPbouiCOM.EditTextColumn
                col = Grid0.Columns.Item(0)
                col.LinkedObjectType = "SUBPO"
                Grid0.AutoResizeColumns()
                objaddon.objapplication.StatusBar.SetText("Loaded Sub Contracting lists Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                LoadDetails_InventoryTransfer(OEForm.TypeCount)
            Catch ex As Exception

            End Try

        End Sub

        Private Sub LoadDetails_InventoryTransfer(ByVal FormTypeCount As String)
            Try
                Dim InvTfrForm As SAPbouiCOM.Form
                Dim InvTfrMatrix As SAPbouiCOM.Matrix
                If Grid0.Rows.SelectedRows.Count = 0 Then Exit Sub
                InvTfrForm = objaddon.objapplication.Forms.GetForm("940", FormTypeCount)
                InvTfrMatrix = InvTfrForm.Items.Item("23").Specific
                objaddon.objapplication.StatusBar.SetText("Loading Sub Contracting details. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                InvTfrMatrix.Clear()
                InvTfrMatrix.AddRow()
                InvTfrForm.Items.Item("U_SubConNo").Specific.String = "-1"
                InvTfrForm.Items.Item("22").Specific.String = "Created by " & objaddon.objcompany.UserName & " for Sub-Contracting bulk transfer on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")
                InvTfrForm.Items.Item("18").Specific.String = Grid0.DataTable.GetValue("From Whse", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                InvTfrForm.Items.Item("1470000101").Specific.String = Grid0.DataTable.GetValue("To Whse", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                If InvTfrMatrix.Columns.Item("U_LineID").Editable = False Then InvTfrMatrix.Columns.Item("U_LineID").Editable = True
                If InvTfrMatrix.Columns.Item("U_PlanQty").Editable = False Then InvTfrMatrix.Columns.Item("U_PlanQty").Editable = True
                If InvTfrMatrix.Columns.Item("U_S_DocEntry").Editable = False Then InvTfrMatrix.Columns.Item("U_S_DocEntry").Editable = True
                For i As Integer = 1 To Grid0.Rows.SelectedRows.Count 'To 1 Step -1
                    InvTfrMatrix.Columns.Item("1").Cells.Item(i).Specific.String = Grid0.DataTable.GetValue("Input ItemCode", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)) 'ItemCode
                    InvTfrMatrix.Columns.Item("1470001039").Cells.Item(i).Specific.String = Grid0.DataTable.GetValue("From Whse", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)) 'From Whse
                    InvTfrMatrix.Columns.Item("5").Cells.Item(i).Specific.String = Grid0.DataTable.GetValue("To Whse", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)) 'To Whse
                    InvTfrMatrix.Columns.Item("10").Cells.Item(i).Specific.String = Grid0.DataTable.GetValue("Pending Qty", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)) 'Quantity
                    InvTfrMatrix.Columns.Item("U_S_DocEntry").Cells.Item(i).Specific.String = Grid0.DataTable.GetValue("DocEntry", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                    InvTfrMatrix.Columns.Item("U_LineID").Cells.Item(i).Specific.String = Grid0.DataTable.GetValue("Input Line", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                    InvTfrMatrix.Columns.Item("U_PlanQty").Cells.Item(i).Specific.String = Grid0.DataTable.GetValue("Pending Qty", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)) 'Quantity
                Next
                InvTfrMatrix.Columns.Item("1").Cells.Item(1).Click()
                objaddon.objapplication.StatusBar.SetText("Loaded Sub Contracting details Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                InvTfrMatrix.AutoResizeColumns()
                objform.Close()
                OEForm = Nothing
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Grid0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.ClickAfter
            Try
                If pVal.Row <> -1 And pVal.ColUID <> "RowsHeader" Then
                    If Grid0.Rows.IsSelected(pVal.Row) = True Then
                        Grid0.Rows.SelectedRows.Remove(pVal.Row)
                    Else
                        Grid0.Rows.SelectedRows.Add(pVal.Row)
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents Button2 As SAPbouiCOM.Button

        Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try 'Clear All
                objform.Freeze(True)
                Grid0.Rows.SelectedRows.Clear()
                objform.Freeze(False)
                Grid0.RowHeaders.TitleObject.Click()
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Grid0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Grid0.ClickBefore
            Try
                If pVal.Row = -1 Then
                    BubbleEvent = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Grid0_LinkPressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.LinkPressedAfter
            Try
                Try
                    Link_Value = Grid0.DataTable.GetValue("DocEntry", pVal.Row)
                    Link_objtype = "SUBPO"
                    Dim activeform As New SubContractingPO
                    activeform.Show()
                Catch ex As Exception

                End Try
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Grid0.AutoResizeColumns()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If Grid0.Rows.SelectedRows.Count = 0 Then BubbleEvent = False : objaddon.objapplication.StatusBar.SetText("Select a Row!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents Button3 As SAPbouiCOM.Button

        Private Sub Button3_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try 'Select All
                objform.Freeze(True)
                Grid0.Rows.SelectedRows.AddRange(0, Grid0.Rows.Count)
                objform.Freeze(False)
                Grid0.RowHeaders.TitleObject.Click()
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
