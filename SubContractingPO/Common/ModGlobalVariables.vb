﻿Namespace SubContractingPO
    Module ModGlobalVariables
        Public NewLink As String = "-1"
        Public Link_Value As String = "-1"
        Public Link_objtype As String = "-1"
        Public POItem As Char = ""
        Public ItemBOM As Char = ""
        Public GetDatePO As Char = ""
        Public ResourceEnable As Char = ""
        Public Costing As Char = ""
        Public AutoProduction As Char = ""
        Public ReceiptAutoLoad As Char = ""
        Public APAutoLoad As Char = ""
        Public SubPOScreen As Char = ""
        Public AutoAssignItem As Char = ""
        Public TypeInScrap As Char = ""
        Public VendorSPlPrice As Char = ""
        Public ToLocationInOutput As Char = ""
        Public ToWhseInInput As Char = ""
        Public ToWhseInvTransfer As Char = ""
        Public TranList As Char = ""
        Public BranchEnabled As Char = ""
        Public SuperUser As String = ""
        Public FrmMultiSel As SAPbouiCOM.Form
        Public ProductQuery As String = ""
        Public FieldName As String = ""
        Public WPrice As Char = ""
        'Public PlanQtyUDFEditable As Char = ""
        Public UDF0 As Char = ""
        Public UDF1 As Char = ""
        Public ValUDF0 As Char = ""
        Public ValUDF1 As Char = ""
        Public BPWhse As Char = ""
        Public BomWhse As Char = ""
        Public ProdOrderStat_Close As Char = ""
        Public bModal As Boolean = False 'Cost Center
        Public ItemProcess As Char = ""
        Public BOMRefresh As Char = ""
        Public OEForm As SAPbouiCOM.Form
        Public Title As String = ""
        Public RowDelete As Char = ""
        Public ProcessValidation As Char = ""
        Public ProcTypefromUDO As Char = ""
        Public RevPlanQty As Char = ""
        Public FormName As String = ""
        Public AddonVersion As String = "-1"
        Public SubConTypeCount As Integer = 0
        Public GIFIFO As Char = ""
    End Module
End Namespace
