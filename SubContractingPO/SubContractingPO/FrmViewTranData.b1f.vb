Option Strict Off
Option Explicit On

Imports System.Drawing
Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("TRANVIEW", "SubContractingPO/FrmViewTranData.b1f")>
    Friend Class FrmViewTranData
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("gridData").Specific, SAPbouiCOM.Grid)
            Me.Button2 = CType(Me.GetItem("Item_3").Specific, SAPbouiCOM.Button)
            Me.Button3 = CType(Me.GetItem("Item_4").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("TRANVIEW", 0)
                objform = objaddon.objapplication.Forms.ActiveForm
                bModal = True
                FormName = "TRANVIEW"

            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid

        Public Sub ViewTransactionData(ByVal Header As String, ByVal Line As String, ByVal LinkedID As String, ByVal InvTransfer As String, ByVal DocEntry As String)
            Dim objrs As SAPbobsCOM.Recordset
            Try
                Dim str_sql As String = ""
                If AutoProduction = "Y" Then
                    If InvTransfer = "Y" Then
                        If objaddon.HANA Then
                            str_sql = "select Distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"" ,T0.""Comments"""
                            str_sql += vbCrLf + "from " & Header & " T0 join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry""  where T0.""U_SubConNo""='" & DocEntry & "';"
                        Else
                            str_sql = "select Distinct T0.DocEntry,T0.DocNum,T0.DocDate,T1.ItemCode,T1.Dscription,T1.Quantity ,T0.Comments"
                            str_sql += vbCrLf + "from " & Header & " T0 join " & Line & " T1 on T0.DocEntry=T1.DocEntry  where T0.U_SubConNo='" & DocEntry & "'"
                        End If
                    Else
                        If objaddon.HANA Then
                            str_sql = "select Distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"" ,T0.""Comments"""
                            str_sql += vbCrLf + "from " & Header & " T0 join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry""  where T1.""BaseType""=202  and T1.""BaseEntry""='" & DocEntry & "';"
                        Else
                            str_sql = "select Distinct T0.DocEntry,T0.DocNum,T0.DocDate,T1.ItemCode,T1.Dscription,T1.Quantity ,T0.Comments"
                            str_sql += vbCrLf + "from " & Header & " T0 join " & Line & " T1 on T0.DocEntry=T1.DocEntry  where T1.BaseType=202  and T1.BaseEntry='" & DocEntry & "'"
                        End If
                    End If
                End If
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(str_sql)
                If objrs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objform.Close() : objrs = Nothing : Exit Sub
                Dim objDT As SAPbouiCOM.DataTable
                If objform.DataSources.DataTables.Count = 0 Then
                    objform.DataSources.DataTables.Add("DT_VIEW")
                End If

                objDT = objform.DataSources.DataTables.Item("DT_VIEW")
                objDT.ExecuteQuery(str_sql)
                objform.DataSources.DataTables.Item("DT_VIEW").ExecuteQuery(str_sql)

                Grid0.DataTable = objform.DataSources.DataTables.Item("DT_VIEW")
                Grid0.CollapseLevel = 1
                objform.Freeze(True)
                For i As Integer = 0 To Grid0.Columns.Count - 1
                    Grid0.Columns.Item(i).TitleObject.Sortable = True
                    Grid0.Columns.Item(i).Editable = False
                Next
                objform.Freeze(False)
                Grid0.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                Dim col As SAPbouiCOM.EditTextColumn
                col = Grid0.Columns.Item(0)
                col.LinkedObjectType = LinkedID
                objform.Visible = True
                objform.Update()

            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Public Sub View_ProcessWise_SubCon_Entries(ByVal OForm As SAPbouiCOM.Form, ByVal ProdEntry As String)
            Dim objrs As SAPbobsCOM.Recordset
            Try
                Dim str_sql As String = ""
                Dim GetSeq As String = ""
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objform.Left = OForm.Left + 100
                objform.Top = OForm.Top + 100
                If objaddon.HANA Then
                    str_sql = "Select ""U_ProcSeq"" from ""@MIPL_OPOR"" where ""U_PONum""='" & ProdEntry & "'"
                Else
                    str_sql = "Select U_ProcSeq from [@MIPL_OPOR] where U_PONum='" & ProdEntry & "'"
                End If
                objrs.DoQuery(str_sql)
                If objrs.RecordCount > 0 Then
                    For i As Integer = 0 To objrs.RecordCount - 1
                        If i = 0 Then
                            GetSeq = objrs.Fields.Item("U_ProcSeq").Value.ToString
                        Else
                            GetSeq += ", " + objrs.Fields.Item("U_ProcSeq").Value.ToString
                        End If
                        objrs.MoveNext()
                    Next
                End If

                If objaddon.HANA Then
                    'Select Case T0."DocEntry" AS "Sub Con Entry", T0."DocNum" AS "Sub Con Num", T0."U_CardCode" AS "Vendor Code", T0."U_CardName" AS "Card Name", T0."U_Process" AS "Process Name", (SELECT CASE WHEN "U_Priority" = '1' THEN 'Optional' ELSE 'Mandatory' END FROM "@MIPL_BOM2" WHERE "Code" = T0."U_BOMCode" AND "U_Sequence" = T0."U_ProcSeq") AS "Process Type", T0."U_SQty" AS "Header Planned Qty", (SELECT CASE WHEN SUM(B."Quantity") > 0 THEN SUM(B."Quantity") ELSE 0 END FROM OIGN A INNER JOIN IGN1 B ON A."DocEntry" = B."DocEntry" WHERE A."U_SubConNo" = T0."DocEntry") AS "Receipt Qty", (SELECT CASE WHEN SUM(B."Quantity") > 0 THEN T0."U_SQty" - SUM(B."Quantity") ELSE T0."U_SQty" END FROM OIGN A INNER JOIN IGN1 B ON A."DocEntry" = B."DocEntry" WHERE A."U_SubConNo" = T0."DocEntry") AS "Pending Qty" FROM "@MIPL_OPOR" T0 WHERE "U_PONum" = '162';

                    'str_sql = "Select  T0.""U_PONum"" as ""Production Order Entry"",T0.""DocEntry"" AS ""Sub Con Entry"", T0.""DocNum"" AS ""Sub Con Num"", T0.""U_CardCode"" AS ""Vendor Code"", T0.""U_CardName"" AS ""Card Name"", T0.""U_Process"" AS ""Process Name"","
                    'str_sql += vbCrLf + "(SELECT CASE WHEN ""U_Priority"" = '1' THEN 'Optional' ELSE 'Mandatory' END FROM ""@MIPL_BOM2"" WHERE ""Code"" = T0.""U_BOMCode"" AND ""U_Sequence"" = T0.""U_ProcSeq"") AS ""Process Type"","
                    'str_sql += vbCrLf + "T0.""U_SQty"" AS ""Header Planned Qty"", (SELECT CASE WHEN SUM(B.""Quantity"") > 0 THEN SUM(B.""Quantity"") ELSE 0 END FROM OIGN A INNER JOIN IGN1 B ON A.""DocEntry"" = B.""DocEntry"" "
                    'str_sql += vbCrLf + "WHERE A.""U_SubConNo"" = T0.""DocEntry"") AS ""Receipt Qty"", (SELECT CASE WHEN SUM(B.""Quantity"") > 0 THEN T0.""U_SQty"" - SUM(B.""Quantity"") ELSE T0.""U_SQty"" END"
                    'str_sql += vbCrLf + "FROM OIGN A INNER JOIN IGN1 B ON A.""DocEntry"" = B.""DocEntry"" WHERE A.""U_SubConNo"" = T0.""DocEntry"") AS ""Pending Qty"" FROM ""@MIPL_OPOR"" T0 WHERE ""U_PONum"" = '" & ProdEntry & "'"

                    str_sql = "With Process as ("
                    str_sql += vbCrLf + "Select T0.U_PONum as [Production Order Entry], (Select ItemCode from OWOR where DocEntry=T0.U_PoNum) as [Prod Order ItemCode],"
                    str_sql += vbCrLf + "(Select ItemName from OITM where ItemCode=(Select ItemCode from OWOR where DocEntry=T0.U_PoNum)) as [Prod Order ItemName],"
                    str_sql += vbCrLf + "(Select PlannedQty from OWOR where DocEntry=T0.U_PoNum) as [Prod Order Planned Qty],T0.DocEntry as [Sub-Con DocEntry],"
                    str_sql += vbCrLf + "T0.DocNum as [Sub-Con DocNum], T0.U_DocDate as [Document Date],T0.U_CardCode as [Vendor Code],T0.U_CardName as [Vendor Name], T0.U_Process as [Process Name],"
                    str_sql += vbCrLf + "T0.U_SItemCode as [Sub ItemCode],(Select ItemName from OITM where ItemCode=T0.U_SItemCode) as [Sub ItemName],"
                    str_sql += vbCrLf + "(Select case when U_Priority='1' then 'Optional' Else 'Mandatory' End From [@MIPL_BOM2] where Code=T0.U_BOMCode and U_Sequence=T0.U_ProcSeq) as [Process Type],T0.U_ProcSeq as [Sequence],T0.U_SQty as [Header Planned Qty],"
                    str_sql += vbCrLf + "(Select case when Sum(B.Quantity)>0 Then  Sum(B.Quantity) Else 0 End From OIGN A join IGN1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T0.DocEntry) as [Receipt Qty],"
                    str_sql += vbCrLf + "(Select case when Sum(B.Quantity)>0 Then T0.U_SQty- Sum(B.Quantity) Else T0.U_SQty End From OIGN A join IGN1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T0.DocEntry) as [Pending Qty]"
                    str_sql += vbCrLf + "from [@MIPL_OPOR] T0 where U_PONum='" & ProdEntry & "') ,"

                    str_sql += vbCrLf + "NProcess as ("
                    str_sql += vbCrLf + "Select distinct T1.U_PONum,(Select ItemCode from OWOR where DocEntry=T1.U_PoNum) [Production Order Entry],"
                    str_sql += vbCrLf + "(Select ItemName from OITM where ItemCode=(Select ItemCode from OWOR where DocEntry=T1.U_PoNum)) [Prod Order ItemName],(Select PlannedQty from OWOR where DocEntry=T1.U_PoNum) [Prod Order Planned Qty],"
                    str_sql += vbCrLf + "'' [Sub-Con DocEntry],'' [Sub-Con DocNum],null  [Document Date],'' [Vendor Code],'' [Vendor Name],T0.U_Procname, T1.U_SItemCode,(Select ItemName from OITM where ItemCode=T1.U_SItemCode) [Sub ItemName],"
                    str_sql += vbCrLf + "Case when U_Priority='1' then 'Optional' Else 'Mandatory' End  [Process Type] ,T0.U_Sequence,0 [Header Planned Qty],0 [Receipt Qty] ,0 [Pending Qty] "
                    str_sql += vbCrLf + "From [@MIPL_BOM2] T0 left join [@MIPL_OPOR]  T1 on T0.Code=T1.U_BOMCode left join Process C on C.[Process Name]=T1.U_Process"
                    str_sql += vbCrLf + "where T0.U_Sequence not in (" & GetSeq & ") and T1.U_PONum='" & ProdEntry & "')"

                    str_sql += vbCrLf + "Select * from Process "
                    str_sql += vbCrLf + "Union all"
                    str_sql += vbCrLf + "Select * from NProcess "
                Else
                    'str_sql = "Select T0.U_PONum as [Production Order Entry],(Select ItemCode from OWOR where DocEntry=T0.U_PoNum) as [Prod Order ItemCode],"
                    'str_sql += vbCrLf + "(Select ItemName from OITM where ItemCode=(Select ItemCode from OWOR where DocEntry=T0.U_PoNum)) as [Prod Order ItemName],"
                    'str_sql += vbCrLf + "(Select PlannedQty from OWOR where DocEntry=T0.U_PoNum) as [Prod Order Planned Qty],"
                    'str_sql += vbCrLf + "T0.DocEntry as [Sub-Con DocEntry],T0.DocNum as [Sub-Con DocNum], T0.U_CardCode as [Vendor Code],T0.U_CardName as [Vendor Name], T0.U_Process as [Process Name],"
                    'str_sql += vbCrLf + "T0.U_SItemCode as [Sub ItemCode],(Select ItemName from OITM where ItemCode=T0.U_SItemCode) as [Sub ItemName],"
                    'str_sql += vbCrLf + "(Select case when U_Priority='1' then 'Optional' Else 'Mandatory' End From [@MIPL_BOM2] where Code=T0.U_BOMCode and U_Sequence=T0.U_ProcSeq) as [Process Type],T0.U_SQty as [Header Planned Qty],"
                    'str_sql += vbCrLf + "(Select case when Sum(B.Quantity)>0 Then  Sum(B.Quantity) Else 0 End From OIGN A join IGN1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T0.DocEntry) as [Receipt Qty],"
                    'str_sql += vbCrLf + "(Select case when Sum(B.Quantity)>0 Then T0.U_SQty- Sum(B.Quantity) Else T0.U_SQty End From OIGN A join IGN1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T0.DocEntry) as [Pending Qty]"
                    'str_sql += vbCrLf + "from [@MIPL_OPOR] T0 where U_PONum='" & ProdEntry & "'"

                    str_sql = "With Process as ("
                    str_sql += vbCrLf + "Select T0.U_PONum as [Production Order Entry], (Select ItemCode from OWOR where DocEntry=T0.U_PoNum) as [Prod Order ItemCode],"
                    str_sql += vbCrLf + "(Select ItemName from OITM where ItemCode=(Select ItemCode from OWOR where DocEntry=T0.U_PoNum)) as [Prod Order ItemName],"
                    str_sql += vbCrLf + "(Select PlannedQty from OWOR where DocEntry=T0.U_PoNum) as [Prod Order Planned Qty],Convert(Varchar,T0.DocEntry) as [Sub-Con DocEntry],"
                    str_sql += vbCrLf + "Convert(Varchar,T0.DocNum) as [Sub-Con DocNum], Convert(Varchar,T0.U_DocDate) as [Document Date],T0.U_CardCode as [Vendor Code],T0.U_CardName as [Vendor Name], T0.U_Process as [Process Name],"
                    str_sql += vbCrLf + "T0.U_SItemCode as [Sub ItemCode],(Select ItemName from OITM where ItemCode=T0.U_SItemCode) as [Sub ItemName],"
                    str_sql += vbCrLf + "(Select case when U_Priority='1' then 'Optional' Else 'Mandatory' End From [@MIPL_BOM2] where Code=T0.U_BOMCode and U_Sequence=T0.U_ProcSeq) as [Process Type],T0.U_ProcSeq as [Sequence],T0.U_SQty as [Header Planned Qty],"
                    str_sql += vbCrLf + "(Select case when Sum(B.Quantity)>0 Then  Sum(B.Quantity) Else 0 End From OIGN A join IGN1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T0.DocEntry) as [Receipt Qty],"
                    str_sql += vbCrLf + "(Select case when Sum(B.Quantity)>0 Then T0.U_SQty- Sum(B.Quantity) Else T0.U_SQty End From OIGN A join IGN1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T0.DocEntry) as [Pending Qty]"
                    str_sql += vbCrLf + "from [@MIPL_OPOR] T0 where U_PONum='" & ProdEntry & "') ,"

                    str_sql += vbCrLf + "NProcess as ("
                    str_sql += vbCrLf + "Select distinct T1.U_PONum,(Select ItemCode from OWOR where DocEntry=T1.U_PoNum) [Production Order Entry],"
                    str_sql += vbCrLf + "(Select ItemName from OITM where ItemCode=(Select ItemCode from OWOR where DocEntry=T1.U_PoNum)) [Prod Order ItemName],(Select PlannedQty from OWOR where DocEntry=T1.U_PoNum) [Prod Order Planned Qty],"
                    str_sql += vbCrLf + "'' [Sub-Con DocEntry],'' [Sub-Con DocNum],'' [Document Date],'' [Vendor Code],'' [Vendor Name],T0.U_Procname, T1.U_SItemCode,(Select ItemName from OITM where ItemCode=T1.U_SItemCode) [Sub ItemName],"
                    str_sql += vbCrLf + "Case when U_Priority='1' then 'Optional' Else 'Mandatory' End  [Process Type] ,T0.U_Sequence,0 [Header Planned Qty],0 [Receipt Qty] ,0 [Pending Qty] "
                    str_sql += vbCrLf + "From [@MIPL_BOM2] T0 left join [@MIPL_OPOR]  T1 on T0.Code=T1.U_BOMCode left join Process C on C.[Process Name]=T1.U_Process"
                    str_sql += vbCrLf + "where T0.U_Sequence not in (" & GetSeq & ") and T1.U_PONum='" & ProdEntry & "')"

                    str_sql += vbCrLf + "Select * from Process "
                    str_sql += vbCrLf + "Union all"
                    str_sql += vbCrLf + "Select * from NProcess "

                End If
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(str_sql)
                If objrs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objform.Close() : objrs = Nothing : Exit Sub
                Dim objDT As SAPbouiCOM.DataTable
                If objform.DataSources.DataTables.Count = 0 Then
                    objform.DataSources.DataTables.Add("DT_VIEW")
                End If

                objDT = objform.DataSources.DataTables.Item("DT_VIEW")
                objDT.ExecuteQuery(str_sql)
                objform.DataSources.DataTables.Item("DT_VIEW").ExecuteQuery(str_sql)

                Grid0.DataTable = objform.DataSources.DataTables.Item("DT_VIEW")
                Grid0.CollapseLevel = 1
                objform.Freeze(True)
                For i As Integer = 0 To Grid0.Columns.Count - 1
                    Grid0.Columns.Item(i).TitleObject.Sortable = True
                    Grid0.Columns.Item(i).Editable = False
                    'Grid0.Columns.Item(i).TextStyle = FontStyle.Bold
                Next
                objform.Freeze(False)
                Grid0.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                Dim col As SAPbouiCOM.EditTextColumn
                col = Grid0.Columns.Item(0)
                col.LinkedObjectType = "202" 'LinkedID
                Grid0.AutoResizeColumns()
                objform.Visible = True
                objform.Update()
                'For i As Integer = 0 To Grid0.Rows.Count - 1
                '    'If Grid0.Rows(i).Then Then 'Grid0.DataTable.GetValue("Sub-Con DocEntry", i).ToString CInt(Grid0.Columns.Item("Sub-Con DocEntry").Cells.Item(i).Specific.String)
                '    '    Grid0.CommonSetting.SetRowFontColor(i, Color.Blue.B) 'Color.DarkRed.ToArgb
                '    '    'Grid0.CommonSetting.SetCellFontColor(i, 14, Color.Blue.B)

                '    'End If
                'Next

            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Button3 As SAPbouiCOM.Button

        Private Sub Button2_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                Grid0.Rows.ExpandAll()
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button3_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try
                Grid0.Rows.CollapseAll()
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Grid0.AutoResizeColumns()
            Catch ex As Exception
            End Try

        End Sub

    End Class
End Namespace
