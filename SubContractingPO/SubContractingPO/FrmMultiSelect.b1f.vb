Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("MULSEL", "SubContractingPO/FrmMultiSelect.b1f")>
    Friend Class FrmMultiSelect
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("Grdata").Specific, SAPbouiCOM.Grid)
            Me.Button2 = CType(Me.GetItem("btnclr").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("MULSEL", 0)
                bModal = True
                FormName = "MULSEL"
                If FieldName = "LCode" Or FieldName = "LCodeO" Then
                    objform.Title = "Select Location"
                ElseIf FieldName = "tprocseq" Then
                    objform.Title = "Select Process"
                ElseIf FieldName = "subinvoice" Then
                    objform.Title = "Open Receipt Entries"
                    objform.ClientWidth = objform.ClientWidth + 600
                    Grid0.Item.Width = Grid0.Item.Width + 550
                    'Grid0.AutoResizeColumns()
                End If
                LoadGrid(ProductQuery)

                objform.Update()
                objform.Refresh()
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid

        Private Sub LoadGrid(ByVal query As String)
            Try
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(query)
                Grid0.DataTable.ExecuteQuery(query)
                If objRs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No data found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objform.Close() : Exit Sub
                If FieldName = "LCode" Or FieldName = "LCodeO" Or FieldName = "tprocseq" Then
                    Grid0.Columns.Item("Select").Editable = False
                    Grid0.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                End If
                objform.Freeze(True)
                For i As Integer = 0 To Grid0.Columns.Count - 1
                    ' If i = 0 Then Continue For
                    Grid0.Columns.Item(i).Editable = False
                Next
                'Grid0.Columns.Item(1).Editable = False
                'Grid0.Columns.Item(2).Editable = False
                'Grid0.Columns.Item(3).Editable = False
                objform.Freeze(False)
                ProductQuery = ""
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If FieldName = "subinvoice" Then
                    Try
                        Dim OutwardMatrix As SAPbouiCOM.Matrix
                        OutwardMatrix = FrmMultiSel.Items.Item("mtxdata").Specific
                        FrmMultiSel.DataSources.DBDataSources.Item("@SUBINV1").Clear()

                        For i As Integer = Grid0.Rows.SelectedRows.Count To 1 Step -1
                            With FrmMultiSel.DataSources.DBDataSources.Item("@SUBINV1")
                                .InsertRecord(0)
                                .SetValue("LineId", 0, i)
                                .SetValue("U_SDocNum", 0, Grid0.DataTable.GetValue("Sub Con Doc#", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_SDocEntry", 0, Grid0.DataTable.GetValue("Sub Con DocEntry", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_SDate", 0, Grid0.DataTable.GetValue("Sub Con Doc Date", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_PoEntry", 0, Grid0.DataTable.GetValue("PO Entry", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_PoLine", 0, Grid0.DataTable.GetValue("PO Line", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_PoDate", 0, Grid0.DataTable.GetValue("PO Date", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_GrpoNum", 0, Grid0.DataTable.GetValue("GRPO Entry", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_GrpoDat", 0, Grid0.DataTable.GetValue("GRPO Date", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_ItemCode", 0, Grid0.DataTable.GetValue("Sub Item Code", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_ItemName", 0, Grid0.DataTable.GetValue("Sub Item Description", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_SQty", 0, Grid0.DataTable.GetValue("Plan Qty", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_GRNNum", 0, Grid0.DataTable.GetValue("GRN Entry", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_GRNDate", 0, Grid0.DataTable.GetValue("GRN Date", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))
                                .SetValue("U_GRNQty", 0, Grid0.DataTable.GetValue("GRN Qty", Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))

                            End With
                        Next
                        OutwardMatrix.LoadFromDataSourceEx()
                        OutwardMatrix.AutoResizeColumns()
                    Catch ex As Exception
                    End Try
                ElseIf FieldName = "LCode" Or FieldName = "LCodeO" Or FieldName = "tprocseq" Then
                    Dim code As String = "", Name As String = ""
                    For i As Integer = 0 To Grid0.Rows.Count - 1
                        If Grid0.DataTable.GetValue("Select", i).ToString = "Y" Then
                            code = code + (Grid0.DataTable.GetValue(1, i).ToString + ",")
                            Name = Name + (Grid0.DataTable.GetValue(2, i).ToString + ",")
                        End If
                    Next
                    code = code.Remove(code.Length - 1)
                    Name = Name.Remove(Name.Length - 1)
                    If code <> "" Then
                        If FieldName = "LCode" Or FieldName = "LCodeO" Then
                            FrmMultiSel.Items.Item(FieldName).Specific.string = ""
                            FrmMultiSel.Items.Item(FieldName).Specific.string = code
                        ElseIf FieldName = "tprocseq" Then
                            FrmMultiSel.Items.Item(FieldName).Specific.string = ""
                            FrmMultiSel.Items.Item(FieldName).Specific.string = code
                            FrmMultiSel.Items.Item("tproc").Specific.string = ""
                            FrmMultiSel.Items.Item("tproc").Specific.string = Name
                        End If

                    End If
                End If
                FrmMultiSel = Nothing
                FieldName = Nothing
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Grid0_DoubleClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.DoubleClickAfter
            Try
                Select Case pVal.ColUID
                    Case "Select"
                        If FieldName = "LCode" Or FieldName = "LCodeO" Then
                            If pVal.Row = -1 Then
                                'objform.Freeze(True)
                                If Grid0.DataTable.GetValue("Select", 0).ToString = "Y" Then
                                    For i As Integer = 0 To Grid0.Rows.Count - 1
                                        Grid0.Columns.Item("Select").Check(i, False)
                                    Next
                                Else
                                    For i As Integer = 0 To Grid0.Rows.Count - 1
                                        Grid0.Columns.Item("Select").Check(i, True)
                                    Next
                                End If

                                'objform.Freeze(False)
                            End If
                        End If

                    Case "Code"
                        If pVal.Row = -1 Then
                            Grid0.Columns.Item("Code").TitleObject.Sortable = True
                        End If
                    Case "Name"
                        If pVal.Row = -1 Then
                            Grid0.Columns.Item("Name").TitleObject.Sortable = True
                            'Grid0.Columns.Item("Name").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
                        End If
                End Select
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Grid0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.PressedAfter
            Try
                Select Case pVal.ColUID
                    Case "Select"
                        If FieldName = "tprocseq" Then
                            For i As Integer = 0 To Grid0.Rows.Count - 1
                                Grid0.Columns.Item("Select").Check(i, False)
                            Next
                            Grid0.Columns.Item("Select").Check(pVal.Row, True)
                        End If
                    Case Else
                        If pVal.Row <> -1 And pVal.ColUID <> "RowsHeader" Then
                            If Grid0.Rows.IsSelected(pVal.Row) = True Then
                                Grid0.Rows.SelectedRows.Remove(pVal.Row)
                            Else
                                Grid0.Rows.SelectedRows.Add(pVal.Row)
                            End If
                        End If
                End Select
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Grid0.AutoResizeColumns()
            Catch ex As Exception
            End Try

        End Sub

        Private WithEvents Button2 As SAPbouiCOM.Button

        Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                Grid0.Rows.SelectedRows.Clear()
            Catch ex As Exception
            End Try

        End Sub
    End Class

End Namespace
