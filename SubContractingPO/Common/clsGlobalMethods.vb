Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.IO
Imports SAPbobsCOM

Namespace SubContractingPO

    Public Class clsGlobalMethods
        Dim strsql As String
        Dim objrs As SAPbobsCOM.Recordset

        Public Function GetDocNum(ByVal sUDOName As String, ByVal Series As Integer) As String
            Dim StrSQL As String
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'If objAddOn.HANA Then
            If Series = 0 Then
                StrSQL = " select  ""NextNumber""  from NNM1 where ""ObjectCode""='" & sUDOName & "'"
            Else
                StrSQL = " select  ""NextNumber""  from NNM1 where ""ObjectCode""='" & sUDOName & "' and ""Series"" = " & Series
            End If

            'Else
            'StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
            'End If
            objRS.DoQuery(StrSQL)
            objRS.MoveFirst()
            If Not objRS.EoF Then
                Return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString())
            Else
                GetDocNum = "1"
            End If
        End Function

        Public Function GetNextCode_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                If objaddon.HANA Then
                    strsql = "select IFNULL(Max(CAST(""Code"" As integer)),0)+1 from """ & Tablename.ToString & """"
                Else
                    strsql = "select ISNULL(Max(CAST(Code As integer)),0)+1 from " & Tablename.ToString & ""
                End If

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function

        Public Function GetNextDocNum_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                strsql = "select IFNULL(Max(CAST(""DocNum"" As integer)),0)+1 from """ & Tablename.ToString & """"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function

        Public Function GetNextDocEntry_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                strsql = "select IFNULL(Max(CAST(""DocEntry"" As integer)),0)+1 from """ & Tablename.ToString & """"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function

        Public Function GetDuration_BetWeenTime(ByVal strFrom As String, ByVal strTo As String)
            Dim Fromtime, Totime As DateTime
            Dim Duration As TimeSpan
            strFrom = Convert_String_TimeHHMM(strFrom) : strTo = Convert_String_TimeHHMM(strTo)
            Totime = New DateTime(2000, 1, 1, Left(strTo, 2), Right(strTo, 2), 0)
            Fromtime = New DateTime(2000, 1, 1, Left(strFrom, 2), Right(strFrom, 2), 0)
            If Totime < Fromtime Then Totime = New DateTime(2000, 1, 2, Left(strTo, 2), Right(strTo, 2), 0)
            Duration = Totime - Fromtime
            Return Duration.Hours.ToString + "." + Left((Duration.Minutes.ToString + "00"), 2).ToString
        End Function

        Public Function GetHours(ByVal FromHrs As String, ByVal ToHrs As String)
            Dim StartTime = New DateTime(2001, 1, 1, FromHrs, 0, 0)
            Dim EndTime = New DateTime(2001, 1, 1, ToHrs, 0, 0)
            Dim duration = EndTime - StartTime
            Dim durationhr = duration.TotalHours '+ "." + duration.TotalMinutes
            Return durationhr
        End Function

        Public Function Validation_From_To_Time(ByVal strFrom As String, ByVal strTo As String)
            Dim Fromtime, Totime As DateTime
            Dim Duration As TimeSpan
            strFrom = Convert_String_TimeHHMM(strFrom) : strTo = Convert_String_TimeHHMM(strTo)
            Totime = New DateTime(2000, 1, 1, Left(strTo, 2), Right(strTo, 2), 0)
            Fromtime = New DateTime(2000, 1, 1, Left(strFrom, 2), Right(strFrom, 2), 0)
            If Totime < Fromtime Then Totime = New DateTime(2000, 1, 2, Left(strTo, 2), Right(strTo, 2), 0)
            Duration = Totime - Fromtime
            If Duration.Hours < 0 Or Duration.Minutes < 0 Then Return False
            Return True
        End Function

        Public Function Convert_String_TimeHHMM(ByVal str As String)
            Return Right("0000" + Regex.Replace(str, "[^\d]", ""), 4)
        End Function

        Public Sub LoadCombo(ByVal objcombo As SAPbouiCOM.ComboBox, Optional ByVal strquery As String = "", Optional ByVal rs As SAPbobsCOM.Recordset = Nothing)
            objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If strquery.ToString = "" And rs Is Nothing Then Exit Sub
            If strquery.ToString <> "" Then objrs.DoQuery(strquery) Else objrs = rs
            If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

            If objcombo.ValidValues.Count > 0 Then
                For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1
                    objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If

            objrs.MoveFirst()
            For i As Integer = 0 To objrs.RecordCount - 1
                objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
                objrs.MoveNext()
            Next
        End Sub

        Public Sub LoadCombo_Series(ByVal objform As SAPbouiCOM.Form, ByVal comboname As String, ByVal objectid As String, ByVal docdate As Date)
            Try
                Dim objcombo As SAPbouiCOM.ComboBox
                If objectid.ToString = "" Then Exit Sub
                objcombo = objform.Items.Item(comboname).Specific
                'objcombo.ValidValues.LoadSeries(objectid, SAPbouiCOM.BoSeriesMode.sf_Add)

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'strsql = " Select Series,Seriesname from nnm1 where objectcode='" & objectid.ToString & "' and Indicator in (select Distinct Indicator  from OFPR where PeriodStat <>'Y') "
                strsql = " CALL ""MIPL_GetDefaultSeries"" ('" & objectid.ToString & "','" & objaddon.objcompany.UserName & "','" & docdate.ToString("yyyyMMdd") & "')"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

                If objcombo.ValidValues.Count > 0 Then
                    For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1 : objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index) : Next
                End If

                objrs.MoveFirst()
                For i As Integer = 0 To objrs.RecordCount - 1
                    objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
                    objrs.MoveNext()
                Next

                objrs.MoveFirst()
                objcombo.Select(objrs.Fields.Item("dflt").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

            Catch ex As Exception

            End Try
        End Sub

        Public Sub LoadCombo_SingleSeries_AfterFind(ByVal objform As SAPbouiCOM.Form, ByVal comboname As String, ByVal objectid As String, ByVal Seriesid As String)
            Try
                If objectid.ToString = "" Or Seriesid = "" Or comboname = "" Or objform Is Nothing Then Exit Sub

                Dim objcombo As SAPbouiCOM.ComboBox
                objcombo = objform.Items.Item(comboname).Specific
                'objcombo.ValidValues.LoadSeries(objectid, SAPbouiCOM.BoSeriesMode.sf_Add)

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strsql = " Select ""Series"",""SeriesName"" from nnm1 where ""ObjectCode""='" & objectid.ToString & "' and ""Series""='" & Seriesid.ToString & "'"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

                'If objcombo.ValidValues.Count > 0 Then
                '    For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1 : objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index) : Next
                'End If

                objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)

                objcombo.Select(Seriesid, SAPbouiCOM.BoSearchKey.psk_ByValue)

            Catch ex As Exception

            End Try
        End Sub

        Public Function default_series(ByVal objectid As String, ByVal docdate As Date)
            Try
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strsql = " CALL ""MIPL_GetDefaultSeries"" ('" & objectid.ToString & "','" & objaddon.objcompany.UserName & "','" & docdate & "')"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Return ""
                Return objrs.Fields.Item(0).Value
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Sub Matrix_Addrow(ByVal omatrix As SAPbouiCOM.Matrix, Optional ByVal colname As String = "", Optional ByVal rowno_name As String = "", Optional ByVal Error_Needed As Boolean = False)
            Try
                Dim addrow As Boolean = False

                If omatrix.VisualRowCount = 0 Then addrow = True : GoTo addrow
                If colname = "" Then addrow = True : GoTo addrow
                If omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific.string <> "" Then addrow = True : GoTo addrow

addrow:
                If addrow = True Then
                    omatrix.AddRow(1)
                    omatrix.ClearRowData(omatrix.VisualRowCount)
                    If rowno_name <> "" Then omatrix.Columns.Item("#").Cells.Item(omatrix.VisualRowCount).Specific.string = omatrix.VisualRowCount
                Else
                    If Error_Needed = True Then objaddon.objapplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub HeaderLabel_Color(ByRef item As SAPbouiCOM.Item, ByVal fontsize As Integer, ByVal forecolor As Integer, ByVal height As Integer, Optional ByVal width As Integer = 0)
            item.TextStyle = FontStyle.Bold
            item.FontSize = fontsize
            item.ForeColor = forecolor
            item.Height = height
            'If width <> 0 Then item.Width = width
        End Sub

        Public Sub RightClickMenu_Delete(ByVal MainMenu As String, ByVal NewMenuID As String)
            Try
                Dim omenuitem As SAPbouiCOM.MenuItem
                omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
                If omenuitem.SubMenus.Exists(NewMenuID) Then
                    objaddon.objapplication.Menus.RemoveEx(NewMenuID)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub SetAutomanagedattribute_Editable(ByVal oform As SAPbouiCOM.Form, ByVal fieldname As String, ByVal add As Boolean, ByVal find As Boolean, ByVal update As Boolean)

            If add = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If find = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If update Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If
        End Sub

        Public Sub SetAutomanagedattribute_Visible(ByVal oform As SAPbouiCOM.Form, ByVal fieldname As String, ByVal add As Boolean, ByVal find As Boolean, ByVal update As Boolean)

            If add = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If find = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If update Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

        End Sub

        Public Function GetDocnum_BaseonSeries(ByVal objectcode As String, ByVal Selected_seriescode As String)
            Try
                Dim strsql As String = "Select ""NextNumber"" from nnm1 where ""ObjectCode""='" & objectcode.ToString & "' and ""Series""='" & Selected_seriescode.ToString & "'"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount = 0 Then Return ""
                Return objrs.Fields.Item(0).Value.ToString
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Sub ChooseFromList_Before(ByVal OForm As SAPbouiCOM.Form, ByVal CFLID As String, ByVal SqlQuery_Condition As String, ByVal AliseID As String)
            Dim rsetCFL As SAPbobsCOM.Recordset
            rsetCFL = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = OForm.ChooseFromLists.Item(CFLID)
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                rsetCFL = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsetCFL.DoQuery(SqlQuery_Condition)
                rsetCFL.MoveFirst()
                If rsetCFL.RecordCount > 0 Then
                    For i As Integer = 1 To rsetCFL.RecordCount
                        If i = (rsetCFL.RecordCount) Then
                            oCond = oConds.Add()
                            oCond.Alias = AliseID
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        Else
                            oCond = oConds.Add()
                            oCond.Alias = AliseID
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        End If
                        rsetCFL.MoveNext()
                    Next
                Else
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()
                    oCond = oConds.Add()
                    oCond.Alias = AliseID
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NONE
                    oCond.CondVal = "-1"
                End If

                oCFL.SetConditions(oConds)
            Catch ex As Exception

            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsetCFL)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Public Function GetDateTimeValue(ByVal SBODaMIPLAGNTMASring As String) As DateTime
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            objBridge.Format_StringToDate("")
            Return objBridge.Format_StringToDate(SBODaMIPLAGNTMASring).Fields.Item(0).Value
        End Function

        Public Function getSingleValue(ByVal StrSQL As String) As String
            Try
                Dim rset As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strReturnVal As String = ""
                rset.DoQuery(StrSQL)
                Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(" Get Single Value Function Failed :  " & ex.Message + StrSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return ""
            End Try
        End Function

        Public Function GetSeries(ByVal Objcode As String, ByVal DocDate As String) As String
            Dim series As String = "", Indicator As String

            Indicator = getSingleValue("select ""Indicator""  from OFPR where '" & CDate(DocDate.ToString).ToString("yyyy-MM-dd") & "' between ""F_RefDate"" and ""T_RefDate""")
            If Objcode = "23" Then
                series = getSingleValue("select ""Series"" From  NNM1 where ""ObjectCode""='" & Objcode & "' and ""Indicator""='" & Indicator & "'")
            End If
            If series <> "" Then
                Return series
            Else
                Return ""
            End If
        End Function

        Public Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub SetCellEdit(ByVal Matrix0 As SAPbouiCOM.Matrix, ByVal EditFlag As Boolean)
            Matrix0.CommonSetting.SetCellEditable(Matrix0.VisualRowCount, 1, EditFlag)
            Matrix0.CommonSetting.SetCellEditable(Matrix0.VisualRowCount, 3, EditFlag)
            Matrix0.CommonSetting.SetCellEditable(Matrix0.VisualRowCount, 5, EditFlag)
            Matrix0.CommonSetting.SetCellEditable(Matrix0.VisualRowCount, 7, EditFlag)
        End Sub

        Public Sub SubConButtonEnabling(ByVal objform As SAPbouiCOM.Form)
            Try
                Dim StrQuery, POItemFlag As String
                Dim objRecSet As SAPbobsCOM.Recordset
                objRecSet = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objform.Items.Item("6").Specific.String <> "" Then
                    If objaddon.HANA Then
                        StrQuery = "update T0 set ""U_SubConBOM""='Y' From OWOR T0 inner join OITT T1 on T0.""ItemCode""=T1.""Code"" where T1.""U_SubConBOM""='Y'"
                    Else
                        StrQuery = "update T0 set U_SubConBOM='Y' From OWOR T0 inner join OITT T1 on T0.ItemCode=T1.Code where T1.U_SubConBOM='Y'"
                    End If
                    objRecSet.DoQuery(StrQuery)
                End If
                If objaddon.HANA Then
                    POItemFlag = objaddon.objglobalmethods.getSingleValue("select ""U_SubConBOM"" from OITT where ""Code""='" & objform.Items.Item("6").Specific.String & "'")
                Else
                    POItemFlag = objaddon.objglobalmethods.getSingleValue("select U_SubConBOM from OITT where Code='" & objform.Items.Item("6").Specific.String & "'")
                End If
                If POItemFlag = "Y" Then
                    objform.Items.Item("BtnSubCon").Enabled = True
                Else
                    objform.Items.Item("BtnSubCon").Enabled = False
                End If

            Catch ex As Exception
            End Try
        End Sub

        Public Sub LoadSeries(ByVal objform As SAPbouiCOM.Form, ByVal DBSource As SAPbouiCOM.DBDataSource, ByVal ObjectCode As String)
            Try
                Dim ComboBox0 As SAPbouiCOM.ComboBox
                Dim Flag As Boolean = False
                ComboBox0 = objform.Items.Item("Series").Specific
                ComboBox0.ValidValues.LoadSeries(ObjectCode, SAPbouiCOM.BoSeriesMode.sf_Add) '"SUBPO"
                Dim DfltSeries As String = ""
                If objaddon.HANA Then
                    DfltSeries = getSingleValue("select T0.""Series"" from NNM2 T0 join OUSR T1 on T0.""UserSign""=T1.""USERID"" where T0.""ObjectCode""='" & ObjectCode & "' and T1.USER_CODE='" & objaddon.objcompany.UserName & "'")
                Else
                    DfltSeries = getSingleValue("select T0.Series from NNM2 T0 join OUSR T1 on T0.UserSign=T1.USERID where T0.ObjectCode='" & ObjectCode & "' and T1.USER_CODE='" & objaddon.objcompany.UserName & "'")
                End If
                If DfltSeries <> "" Then
                    For i As Integer = 0 To ComboBox0.ValidValues.Count - 1
                        If ComboBox0.ValidValues.Item(i).Value = DfltSeries Then
                            Flag = True
                        End If
                    Next
                End If
                If Flag Then
                    ComboBox0.Select(DfltSeries, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    ComboBox0.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End If

                DBSource.SetValue("DocNum", 0, objaddon.objglobalmethods.GetDocNum(ObjectCode, CInt(ComboBox0.Selected.Value)))
            Catch ex As Exception

            End Try
        End Sub

        Public Function AutoAssign_SubItem(ByVal FolderID As SAPbouiCOM.Folder, ByVal MatrixID As SAPbouiCOM.Matrix) As Boolean
            'Dim RowFlag As Boolean = False
            Try
                If MatrixID.VisualRowCount = 1 Then
                    Return True
                End If
                If FolderID.Item.UniqueID = "flroutput" Then  'Output Folder
                    'If Matrix3.VisualRowCount = 0 Then GoTo addnewrow
                    If MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String <> "" Then
                        If MatrixID.Columns.Item("GRNo").Cells.Item(MatrixID.VisualRowCount).Specific.String <> "" Then
                            Return True
                        End If
                    End If
                ElseIf FolderID.Item.UniqueID = "flrscrap" Then  'Scrap Folder
                    Dim objCombo As SAPbouiCOM.ComboBox
                    'If Matrix3.VisualRowCount = 0 Then objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                    objCombo = MatrixID.Columns.Item("Type").Cells.Item(MatrixID.VisualRowCount).Specific
                    If objCombo.Selected.Description = "Scrap" Then
                        If MatrixID.Columns.Item("GRNo").Cells.Item(MatrixID.VisualRowCount).Specific.String <> "" Then
                            Return True
                        End If
                    ElseIf objCombo.Selected.Description = "Return" Then
                        If MatrixID.Columns.Item("InvNum").Cells.Item(MatrixID.VisualRowCount).Specific.String <> "" Then
                            Return True
                            'objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                        End If
                    End If
                End If
                'Return RowFlag
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Sub GetSubItemQty(ByVal objform As SAPbouiCOM.Form, ByVal MatrixID As SAPbouiCOM.Matrix)
            Try
                Dim Qty As Double = 0, InvQty As Double = 0
                Dim odbdsDetails2 As SAPbouiCOM.DBDataSource
                Dim objRs As SAPbobsCOM.Recordset
                odbdsDetails2 = objform.DataSources.DBDataSources.Item("@MIPL_POR2")
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                MatrixID.FlushToDataSource()
                'Dim iiii As Integer = odbdsDetails2.Size
                'Dim iij As Integer = MatrixID.VisualRowCount

                If objform.Items.Item("txtentry").Specific.String <> "" Then
                    If objaddon.HANA Then
                        'strsql = "SELECT IFNULL(CASE WHEN SUM(T2.""Quantity"") > 0 THEN (T0.""U_PlanQty"") - (SUM(T2.""Quantity"") * T0.""U_BOMQty"") ELSE T0.""U_PlanQty"" END"
                        'strsql += vbCrLf + "- IFNULL((SELECT TOP 1 SUM(B.""U_Qty"") FROM ""@MIPL_OPOR"" A INNER JOIN ""@MIPL_POR3"" B ON A.""DocEntry"" = B.""DocEntry"" INNER JOIN ""@MIPL_POR1"" C ON B.""DocEntry"" = C.""DocEntry"""
                        'strsql += vbCrLf + "WHERE B.""U_Type"" = '2' AND B.""U_Status"" = 'C' AND A.""DocEntry"" = T0.""DocEntry"" AND IFNULL(B.""U_InvNo"", '') <> '' AND IFNULL(B.""U_GRNo"", '') = '' GROUP BY B.""U_Itemcode"", C.""U_BOMQty"""
                        'strsql += vbCrLf + "ORDER BY SUM(B.""U_Qty"") DESC), 0), 0) / T0.""U_BOMQty"" AS ""OpenQty"" FROM ""@MIPL_OPOR"" T1 LEFT OUTER JOIN ""@MIPL_POR1"" T0 ON T1.""DocEntry"" = T0.""DocEntry"""
                        'strsql += vbCrLf + "LEFT OUTER JOIN OIGN T3 ON T3.""U_SubConNo"" = T1.""DocEntry"" LEFT OUTER JOIN IGN1 T2 ON T3.""DocEntry"" = T2.""DocEntry"" AND T1.""U_SItemCode"" = T2.""ItemCode"""
                        'strsql += vbCrLf + "WHERE T1.""DocEntry"" = '" & objform.Items.Item("txtentry").Specific.String & "' GROUP BY T0.""DocEntry"", T0.""U_PlanQty"", T0.""U_BOMQty"""

                        strsql = "Select MIN(((Case when A.""ReceiptQty"">0 then A.""U_PlanQty""-(A.""ReceiptQty""*A.""U_Qty"") Else A.""U_PlanQty"" End)-A.""ReturnQty"")/A.""U_Qty"") as ""OpenQty"",A.* "
                        strsql += vbCrLf + "from (Select T1.""DocEntry"",T0.""U_PlanQty"",T0.""U_Itemcode"",T0.""U_Qty"","
                        strsql += vbCrLf + "(Select ifnull((select top 1 sum(B.""U_Qty"") from ""@MIPL_OPOR"" A join ""@MIPL_POR3"" B on A.""DocEntry""=B.""DocEntry"" where B.""U_Type""='2' and B.""U_Status""='C' and A.""DocEntry""=T1.""DocEntry"" and B.""U_Itemcode""=T0.""U_Itemcode"""
                        strsql += vbCrLf + "and ifnull(B.""U_InvNo"",'') <>'' and ifnull(B.""U_GRNo"",'') ='' group by B.""U_Itemcode"" order by sum(B.""U_Qty"") ),0)) as ""ReturnQty"","
                        strsql += vbCrLf + "(Select Sum(T2.""Quantity"") from OIGN T3 join IGN1 T2  on T3.""DocEntry""=T2.""DocEntry"" where T3.""U_SubConNo""=T1.""DocEntry"" and T1.""U_SItemCode""=T2.""ItemCode"") as ""ReceiptQty"""
                        strsql += vbCrLf + "from ""@MIPL_OPOR"" T1 left join ""@MIPL_POR1"" T0 on T1.""DocEntry""=T0.""DocEntry"") A where A.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' group by A.""DocEntry"",A.""U_PlanQty"",A.""U_Itemcode"",A.""U_Qty"",A.""ReturnQty"",A.""ReceiptQty"""
                        strsql += vbCrLf + "order by MIN(((Case when A.""ReceiptQty"">0 then A.""U_PlanQty""-(A.""ReceiptQty""*A.""U_Qty"") Else A.""U_PlanQty"" End)-A.""ReturnQty"")/A.""U_Qty"")"
                        objRs.DoQuery(strsql)
                    Else
                        'strsql = "Select isnull(case when Sum(T2.Quantity)> 0 Then (T0.U_PlanQty)-(Sum(T2.Quantity)*T0.U_BOMQty) Else T0.U_PlanQty End"
                        'strsql += vbCrLf + "- isnull((select Top 1 sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B on A.DocEntry=B.DocEntry join [@MIPL_POR1] C on B.DocEntry=C.DocEntry where B.U_Type='2' and B.U_Status='C' and A.DocEntry=T0.DocEntry"
                        'strsql += vbCrLf + "and isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' group by B.U_Itemcode,C.U_BOMQty order by sum(B.U_Qty) desc),0),0)/T0.U_BOMQty  as OpenQty"
                        'strsql += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry left join OIGN T3 on T3.U_SubConNo=T1.DocEntry left join IGN1 T2  on T3.DocEntry=T2.DocEntry And T1.U_SItemCode=T2.ItemCode"
                        'strsql += vbCrLf + "where T1.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by T0.DocEntry,T0.U_PlanQty,T0.U_BOMQty"

                        strsql = "Select case when isnull((select Top 1 U_Roundoff from [@MI_ITMRND] where U_InvUOM=(SELECT InvntryUom FROM OITM where ItemCode=A.U_SItemCode and InvntryUom <>'') order by Code desc),'') ='Y' then "
                        strsql += vbCrLf + "cast(MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty) as decimal(18,0)) Else MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty) End As OpenQty,A.* "
                        strsql += vbCrLf + "from (Select T1.DocEntry,T0.U_PlanQty,T0.U_Itemcode,T0.U_Qty,T1.U_InvUom,T1.U_SItemCode,"
                        strsql += vbCrLf + "(Select isnull((Select top 1 sum(B.U_Qty) from [@MIPL_OPOR] A join [@MIPL_POR3] B On A.DocEntry= B.DocEntry where B.U_Type='2' and B.U_Status='C' and A.DocEntry=T1.DocEntry and B.U_Itemcode=T0.U_Itemcode"
                        strsql += vbCrLf + "and isnull(B.U_InvNo,'') <>'' and isnull(B.U_GRNo,'') ='' group by B.U_Itemcode order by sum(B.U_Qty) ),0)) as ReturnQty,"
                        strsql += vbCrLf + "(Select Sum(T2.Quantity) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T1.U_SItemCode=T2.ItemCode) as ReceiptQty"
                        strsql += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry) A where A.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' group by A.DocEntry,A.U_PlanQty,A.U_Itemcode,A.U_Qty,A.ReturnQty,A.ReceiptQty,A.U_SItemCode,A.U_InvUom"
                        strsql += vbCrLf + "order by MIN(((Case when A.ReceiptQty>0 then A.U_PlanQty-(A.ReceiptQty*A.U_Qty) Else A.U_PlanQty End)-A.ReturnQty)/A.U_Qty)"
                        objRs.DoQuery(strsql)
                    End If
                    Qty = CDbl(objRs.Fields.Item(0).Value.ToString)
                Else
                    Qty = CDbl(objform.Items.Item("SQty").Specific.String)
                End If

                'Dim txtDate As SAPbouiCOM.EditText
                'txtDate = MatrixID.Columns.Item("Date").Cells.Item(MatrixID.VisualRowCount).Specific
                'txtDate.Value = Now.Date.ToString("yyyyMMdd")
                odbdsDetails2.SetValue("U_Date", MatrixID.VisualRowCount - 1, Now.Date.ToString("yyyyMMdd"))
                If odbdsDetails2.GetValue("U_GINo", MatrixID.VisualRowCount - 1) <> "" Then Exit Sub 'MatrixID.Columns.Item("GINo").Cells.Item(MatrixID.VisualRowCount).Specific.String 
                'If CDbl(MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.String) <> 0 Then Exit Sub
                If objaddon.HANA Then
                    strsql = getSingleValue("Select 1 from OITM where ""ItemCode""='" & objform.Items.Item("txtsitem").Specific.String & "' and ifnull(""InvntItem"",'')='Y'")
                Else
                    strsql = getSingleValue("Select 1 from OITM where ItemCode='" & objform.Items.Item("txtsitem").Specific.String & "' and isnull(InvntItem,'')='Y' ")
                End If
                If strsql = "" Then Exit Sub
                Try
                    objform.Select()
                    'MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Click()
                    'MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.string = Trim(objform.Items.Item("txtsitem").Specific.String)
                    odbdsDetails2.SetValue("U_Itemcode", MatrixID.VisualRowCount - 1, Trim(objform.Items.Item("txtsitem").Specific.String))
                    'MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Code", Trim(objform.Items.Item("txtsitem").Specific.String))
                    If odbdsDetails2.GetValue("U_ItemDesc", MatrixID.VisualRowCount - 1) = "" Then 'MatrixID.Columns.Item("Desc").Cells.Item(MatrixID.VisualRowCount).Specific.String 
                        If objaddon.HANA Then
                            odbdsDetails2.SetValue("U_ItemDesc", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("select ""ItemName"" from  OITM where ""ItemCode""='" & odbdsDetails2.GetValue("U_Itemcode", MatrixID.VisualRowCount - 1) & "'"))
                            'strsql = objaddon.objglobalmethods.getSingleValue("select ""ItemName"" from  OITM where ""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.string & "'")
                        Else
                            odbdsDetails2.SetValue("U_ItemDesc", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("select ItemName from  OITM where ItemCode='" & odbdsDetails2.GetValue("U_Itemcode", MatrixID.VisualRowCount - 1) & "'"))
                            'strsql = objaddon.objglobalmethods.getSingleValue("select ItemName from  OITM where ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.string & "'")
                        End If
                        'MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Desc", Trim(strsql))
                    End If
                    If BomWhse = "Y" Then
                        If odbdsDetails2.GetValue("U_WhsCode", MatrixID.VisualRowCount - 1) = "" Then 'MatrixID.Columns.Item("Whse").Cells.Item(MatrixID.VisualRowCount).Specific.String
                            If objaddon.HANA Then
                                odbdsDetails2.SetValue("U_WhsCode", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("select ""U_WhseCode"" from  ""@MIPL_OBOM"" where ""Code""='" & objform.Items.Item("txtbitem").Specific.String & "'"))
                                'strsql = objaddon.objglobalmethods.getSingleValue("select ""U_WhseCode"" from  ""@MIPL_OBOM"" where ""Code""='" & objform.Items.Item("txtbitem").Specific.String & "'")
                            Else
                                odbdsDetails2.SetValue("U_WhsCode", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("select U_WhseCode from  [@MIPL_OBOM] where Code='" & objform.Items.Item("txtbitem").Specific.String & "'"))
                                'strsql = objaddon.objglobalmethods.getSingleValue("select U_WhseCode from  [@MIPL_OBOM] where Code='" & objform.Items.Item("txtbitem").Specific.String & "'")
                            End If
                            'MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Whse", Trim(strsql))
                        End If
                    End If
                Catch ex As Exception
                End Try

                odbdsDetails2.SetValue("U_Status", MatrixID.VisualRowCount - 1, "O")
                'MatrixID.Columns.Item("Colstat").Cells.Item(MatrixID.VisualRowCount).Specific.String = "O"
                If odbdsDetails2.GetValue("U_InvUom", MatrixID.VisualRowCount - 1) = "" Then 'MatrixID.Columns.Item("InvUom").Cells.Item(MatrixID.VisualRowCount).Specific.String 
                    If objaddon.HANA Then
                        odbdsDetails2.SetValue("U_InvUom", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & odbdsDetails2.GetValue("U_Itemcode", MatrixID.VisualRowCount - 1) & "'"))
                        'strsql = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                    Else
                        odbdsDetails2.SetValue("U_InvUom", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & odbdsDetails2.GetValue("U_Itemcode", MatrixID.VisualRowCount - 1) & "'"))
                        'strsql = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                    End If
                    'MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "InvUom", Trim(strsql))
                End If
                If odbdsDetails2.GetValue("U_InStock", MatrixID.VisualRowCount - 1) = "" Then 'MatrixID.Columns.Item("Instock").Cells.Item(MatrixID.VisualRowCount).Specific.String 
                    If objaddon.HANA Then
                        odbdsDetails2.SetValue("U_InStock", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from oitw where ""ItemCode""='" & odbdsDetails2.GetValue("U_Itemcode", MatrixID.VisualRowCount - 1) & "' and ""WhsCode""='" & odbdsDetails2.GetValue("U_WhsCode", MatrixID.VisualRowCount - 1) & "'"))
                        'strsql = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from oitw where ""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "' and ""WhsCode""='" & MatrixID.Columns.Item("Whse").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                    Else
                        odbdsDetails2.SetValue("U_InStock", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("select OnHand from oitw where ItemCode='" & odbdsDetails2.GetValue("U_Itemcode", MatrixID.VisualRowCount - 1) & "' and WhsCode='" & odbdsDetails2.GetValue("U_WhsCode", MatrixID.VisualRowCount - 1) & "'"))
                        'strsql = objaddon.objglobalmethods.getSingleValue("select OnHand from oitw where ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "' and WhsCode='" & MatrixID.Columns.Item("Whse").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                    End If
                    'MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Instock", Trim(strsql))
                End If
                If odbdsDetails2.GetValue("U_HSNCode", MatrixID.VisualRowCount - 1) = "" Then 'MatrixID.Columns.Item("HSN").Cells.Item(MatrixID.VisualRowCount).Specific.String 
                    If objaddon.HANA Then
                        odbdsDetails2.SetValue("U_HSNCode", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("Select T0.""ChapterID"" from OCHP T0 join OITM T1 on T0.""AbsEntry""=T1.""ChapterID"" where T1.""ItemCode""='" & odbdsDetails2.GetValue("U_Itemcode", MatrixID.VisualRowCount - 1) & "'"))
                        'strsql = objaddon.objglobalmethods.getSingleValue("Select T0.""ChapterID"" from OCHP T0 join OITM T1 on T0.""AbsEntry""=T1.""ChapterID"" where T1.""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                    Else
                        odbdsDetails2.SetValue("U_HSNCode", MatrixID.VisualRowCount - 1, objaddon.objglobalmethods.getSingleValue("Select T0.ChapterID from OCHP T0 join OITM T1 on T0.AbsEntry=T1.ChapterID where T1.ItemCode='" & odbdsDetails2.GetValue("U_Itemcode", MatrixID.VisualRowCount - 1) & "'"))
                        'strsql = objaddon.objglobalmethods.getSingleValue("Select T0.ChapterID from OCHP T0 join OITM T1 on T0.AbsEntry=T1.ChapterID where T1.ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                    End If
                    'MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "HSN", Trim(strsql))
                End If
                objform.Freeze(True)
                MatrixID.LoadFromDataSource()
                If odbdsDetails2.GetValue("U_Qty", MatrixID.VisualRowCount - 1) <> CDbl(0) Then Exit Sub
                odbdsDetails2.SetValue("U_Qty", MatrixID.VisualRowCount - 1, Qty)
                MatrixID.LoadFromDataSource()
                objform.Freeze(False)
                'MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Quant", Qty)
                If Qty < CDbl(objform.Items.Item("SQty").Specific.String) Then
                    'If odbdsDetails2.GetValue("U_Qty", MatrixID.VisualRowCount - 1) <> CDbl(0) Then Exit Sub 'Val(MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.string)
                    If ProcessValidation = "Y" Then
                        If ItemProcess = "Y" Then
                            If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                Dim SQuery As String, Getseq As String
                                If objform.Items.Item("txtponum").Specific.String <> "" And Trim(objform.Items.Item("tprocseq").Specific.String) <> "" Then
                                    If objaddon.HANA Then
                                        Getseq = objaddon.objglobalmethods.getSingleValue("Select Top 1 T0.""U_Sequence"" as ""Seq"" from ""@MIPL_BOM2"" T0 where T0.""Code""='" & Trim(objform.Items.Item("txtbitem").Specific.String) & "' and T0.""U_Priority""='2' and T0.""U_Sequence""<'" & Trim(objform.Items.Item("tprocseq").Specific.String) & "' order by T0."" U_Sequence"" desc")

                                        SQuery = " Select Case when A."" BOMPrior""='2' and A."" ReceiptQty"">0 then  A.""ReceiptQty"" Else A.""SubItemQty""-A.""ReceiptQty"" End as ""Qty"",* from(Select (Select ifnull(Sum(T2.""Quantity""),0) from OIGN T3 left join IGN1 T2  On T3.""DocEntry""= T2.""DocEntry"""
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

                                        SQuery = " Select Case when A.BOMPrior='2' and A.ReceiptQty >0 then  A.ReceiptQty Else A.SubItemQty-A.ReceiptQty End as Qty,* from(Select (Select isnull(Sum(T2.Quantity),0) from OIGN T3 left join IGN1 T2  On T3.DocEntry= T2.DocEntry"
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

                                    End If
                                    objRs.DoQuery(SQuery)
                                    If objRs.RecordCount > 0 Then
                                        If CDbl(objRs.Fields.Item(0).Value.ToString) > 0 Then
                                            If Getseq = "" Then
                                                MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Quant", CDbl(objform.Items.Item("SQty").Specific.String) - CDbl(Qty))
                                            Else
                                                'Dim vv As Double = CDbl(objRs.Fields.Item(0).Value.ToString)
                                                'Dim vv1 As Double = CDbl(MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.string)
                                                MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Quant", CDbl(objRs.Fields.Item(0).Value.ToString) - CDbl(Qty))
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Quant", CDbl(objform.Items.Item("SQty").Specific.String))
                                'MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.string = CDbl(objform.Items.Item("SQty").Specific.String)
                                'If CDbl(MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.string) = CDbl(0) Then
                                'End If
                            End If

                        Else
                            'If CDbl(MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.string) = CDbl(0) Then
                            '    'MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.string = CDbl((objform.Items.Item("SQty").Specific.String) - CDbl(CDbl(Qty) + InvQty))
                            '    MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Quant", CDbl((objform.Items.Item("SQty").Specific.String) - CDbl(CDbl(Qty) + InvQty)))
                            'End If
                            MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Quant", Qty) 'CDbl((objform.Items.Item("SQty").Specific.String) - (Qty + InvQty)
                        End If
                    Else
                        MatrixID.SetCellWithoutValidation(MatrixID.VisualRowCount, "Quant", Qty)
                    End If
                End If
                ProcessCost_Output(objform, MatrixID)
                objform.Update()
            Catch ex As Exception
                objform.Freeze(False)
                'objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                objform.Freeze(False)
            End Try
        End Sub

        Public Function Manual_StatusClose(ByVal objform As SAPbouiCOM.Form) As Boolean
            Try
                If objform.Items.Item("txtsitem").Specific.String = "" Or objform.Items.Item("txtentry").Specific.String = "" Then Exit Function
                Dim strQuery As String
                strQuery = "Select Top 1 B.* from (Select Case when A.InvTransferQty>0 and (A.InvTransferQty-A.IssueQty-A.ReturnQty)=0 then 'C' Else 'O' End ManualStatus,"
                strQuery += vbCrLf + "Case when A.HeaderQty=A.ReceiptQty and (A.LineQty-(A.IssueQty+A.ReturnQty))=0 then 'C' Else 'O' End AutoStatus,"
                strQuery += vbCrLf + "case when isnull((select Top 1 U_Roundoff from [@MI_ITMRND] where U_InvUOM=(SELECT InvntryUom FROM OITM where ItemCode=A.HeaderItem and InvntryUom <>'') "
                strQuery += vbCrLf + "order by Code desc),'') ='Y' then cast(MIN(((Case when A.ReceiptQty>0 then A.LineQty-(A.ReceiptQty*A.BOMNorms) Else A.LineQty End)-A.ReturnQty)/A.BOMNorms) as decimal(18,0))"
                strQuery += vbCrLf + "Else MIN(((Case when A.ReceiptQty>0 then A.LineQty-(A.ReceiptQty*A.BOMNorms) Else A.LineQty End)-A.ReturnQty)/A.BOMNorms) End As OpenQty,A.*"
                strQuery += vbCrLf + "from (Select T1.DocEntry,T1.Status,T1.U_SItemCode as HeaderItem,T1.U_SQty as HeaderQty,T0.U_Itemcode as LineItem,T0.U_PlanQty as LineQty,"
                strQuery += vbCrLf + "isnull(T0.U_ActPlanQty,0) as ActPlanQty,isnull((T0.U_PlanQty-T0.U_ActPlanQty),0) [Excess PlanQty],T0.U_Qty as BOMNorms,T1.U_InvUom,"
                strQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A join WTR1 B on A.DocEntry=B.DocEntry where A.U_SubConNo=T1.DocEntry and B.ItemCode=T0.U_Itemcode ) as InvTransferQty,"
                strQuery += vbCrLf + "(Select isnull(Sum(T2.Quantity),0) from OIGE T3 join IGE1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T0.U_Itemcode=T2.ItemCode) as IssueQty,"
                strQuery += vbCrLf + "(Select isnull(Sum(B.Quantity),0) from OWTR A inner join WTR1 B on A.DocEntry=B.DocEntry where A.U_ScrapNum=T0.DocEntry and B.ItemCode=T0.U_Itemcode) as ReturnQty,"
                strQuery += vbCrLf + "(Select isnull(Sum(T2.Quantity),0) from OIGN T3 join IGN1 T2  on T3.DocEntry=T2.DocEntry where T3.U_SubConNo=T1.DocEntry and T1.U_SItemCode=T2.ItemCode) as ReceiptQty"
                strQuery += vbCrLf + "from [@MIPL_OPOR] T1 left join [@MIPL_POR1] T0 on T1.DocEntry=T0.DocEntry) A where A.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' "
                strQuery += vbCrLf + "group by A.DocEntry,A.Status,A.LineQty,A.LineItem,A.BOMNorms,A.ReturnQty,A.ReceiptQty,A.HeaderItem,A.ActPlanQty,A.[Excess PlanQty],"
                strQuery += vbCrLf + "A.U_InvUom,A.InvTransferQty,A.IssueQty,A.HeaderQty) B Order by B.OpenQty desc"



                objrs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objrs.DoQuery(strQuery)
                Dim Flag As Boolean = False
                Dim OpenQty As String = ""
                If objrs.RecordCount > 0 Then
                    If objrs.Fields.Item("ManualStatus").Value = "C" Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Status_Closing: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End Try
        End Function

        Public Sub ProcessCost_Output(ByVal objform As SAPbouiCOM.Form, ByVal MatrixID As SAPbouiCOM.Matrix)
            Try
                Dim objcombo As SAPbouiCOM.ComboBox
                Dim objedit As SAPbouiCOM.EditText
                Dim GetValue, Unitprice As String
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objform.Items.Item("tproc").Specific.String = "" Then Exit Sub
                If objaddon.HANA Then
                    GetValue = "Select T0.""TableName"" as ""Header"" ,T1.""TableName"" as ""Line"" from OUDO T0 join UDO1 T1 on T0.""Code""=T1.""Code"" where T0.""TableName""=(Select case when ""U_UDOName"" is null then 'PRICEHEAD' else ""U_UDOName"" end from ""@MIPL_GEN"")"
                Else
                    GetValue = "Select T0.TableName as Header ,T1.TableName as Line from OUDO T0 join UDO1 T1 on T0.Code=T1.Code where T0.TableName=(Select case when U_UDOName is null then 'PRICEHEAD' else U_UDOName end from [@MIPL_GEN])"
                End If
                objRS.DoQuery(GetValue)
                For i = 1 To MatrixID.VisualRowCount
                    If MatrixID.Columns.Item("Code").Cells.Item(i).Specific.String <> "" And MatrixID.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" And MatrixID.Columns.Item("GRNo").Cells.Item(i).Specific.String = "" Then
                        objedit = objform.Items.Item("posdate").Specific
                        Dim PosDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If objaddon.HANA Then
                            GetValue = "Select ifnull(T1.""U_Price"",0) as ""ProcessCost"",ifnull(T1.""U_Weight"",0) as ""Weight"" from ""@" & Trim(objRS.Fields.Item(0).Value.ToString) & """ T0 left join ""@" & Trim(objRS.Fields.Item(1).Value.ToString) & """ T1 on T0.""Code""=T1.""Code"" "
                            GetValue += vbCrLf + "where T0.""Code""='" & objform.Items.Item("txtcode").Specific.String & "' and T1.""U_ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T1.""U_Process""='" & objform.Items.Item("tproc").Specific.String & "'  "
                            GetValue += vbCrLf + "and T1.""U_EffDate""=(Select max(B.""U_EffDate"") from ""@" & Trim(objRS.Fields.Item(0).Value.ToString) & """ A left join ""@" & Trim(objRS.Fields.Item(1).Value.ToString) & """ B on A.""Code""=B.""Code"""
                            GetValue += vbCrLf + "where A.""Code""=T0.""Code"" and B.""U_ItemCode""=T1.""U_ItemCode"" and B.""U_Process""=T1.""U_Process"" and B.""U_EffDate""<='" & PosDate.ToString("yyyyMMdd") & "')"
                        Else
                            GetValue = "Select isnull(T1.U_Price,0) as ProcessCost,isnull(T1.U_Weight,0) as Weight from [@" & Trim(objRS.Fields.Item(0).Value.ToString) & "] T0 left join [@" & Trim(objRS.Fields.Item(1).Value.ToString) & "] T1 on T0.Code=T1.Code "
                            GetValue += vbCrLf + "where T0.Code='" & objform.Items.Item("txtcode").Specific.String & "' and T1.U_ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(i).Specific.String & "' and T1.U_Process='" & objform.Items.Item("tproc").Specific.String & "'  "
                            GetValue += vbCrLf + "and T1.U_EffDate=(Select max(B.U_EffDate) from [@" & Trim(objRS.Fields.Item(0).Value.ToString) & "] A left join [@" & Trim(objRS.Fields.Item(1).Value.ToString) & "] B on A.Code=B.Code"
                            GetValue += vbCrLf + "where A.Code=T0.Code and B.U_ItemCode=T1.U_ItemCode and B.U_Process=T1.U_Process and B.U_EffDate<='" & PosDate.ToString("yyyyMMdd") & "')"
                        End If
                        objRS.DoQuery(GetValue)
                        If objRS.RecordCount > 0 Then
                            Unitprice = CDbl(Unitprice) + CDbl(objRS.Fields.Item("ProcessCost").Value.ToString)
                            objcombo = MatrixID.Columns.Item("cmblttype").Cells.Item(i).Specific
                            MatrixID.SetCellWithoutValidation(i, "ProCost", CDbl(objRS.Fields.Item("ProcessCost").Value.ToString))
                            MatrixID.SetCellWithoutValidation(i, "actwt", CDbl(objRS.Fields.Item("Weight").Value.ToString))
                            MatrixID.SetCellWithoutValidation(i, "tweight", CDbl(objRS.Fields.Item("Weight").Value.ToString) * CDbl(MatrixID.Columns.Item("Quant").Cells.Item(i).Specific.String))
                            If objcombo.Selected.Value = "1" Then
                                MatrixID.SetCellWithoutValidation(i, "TProCost", CDbl(MatrixID.Columns.Item("Quant").Cells.Item(i).Specific.String) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString))
                            Else
                                If Val(MatrixID.Columns.Item("tweight").Cells.Item(i).Specific.String) = 0 Then
                                    MatrixID.SetCellWithoutValidation(i, "TProCost", CDbl(objRS.Fields.Item("Weight").Value.ToString) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString))
                                Else
                                    MatrixID.SetCellWithoutValidation(i, "TProCost", CDbl(MatrixID.Columns.Item("tweight").Cells.Item(i).Specific.String) * CDbl(objRS.Fields.Item("ProcessCost").Value.ToString))
                                End If
                            End If
                        Else
                            Unitprice = CDbl(Unitprice)
                        End If
                        MatrixID.SetCellWithoutValidation(i, "Price", Unitprice)
                        MatrixID.SetCellWithoutValidation(i, "Linetot", CDbl(MatrixID.Columns.Item("Price").Cells.Item(i).Specific.String) * CDbl(MatrixID.Columns.Item("Quant").Cells.Item(i).Specific.String))

                    End If
                Next
            Catch ex As Exception

            End Try
        End Sub

        Public Sub addReport_Layouttype(ByVal FormType As String, ByVal AddonName As String)
            Dim rptTypeService As SAPbobsCOM.ReportTypesService
            Dim newType As SAPbobsCOM.ReportType
            Dim newtypeParam As SAPbobsCOM.ReportTypeParams
            Dim newReportParam As SAPbobsCOM.ReportLayoutParams
            Dim ReportExists As Boolean = False
            Try
                'For Changing add-on Layouts Name and Layout Menu ID 
                'update RTYP set Name='MCarriedOut'  where Name='CarriedOut'
                'update RDOC set DocName='MCarriedOut' where DocName='CarriedOut'
                Dim newtypesParam As SAPbobsCOM.ReportTypesParams
                rptTypeService = objaddon.objcompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
                newtypesParam = rptTypeService.GetReportTypeList

                Dim i As Integer
                For i = 0 To newtypesParam.Count - 1
                    If newtypesParam.Item(i).TypeName = FormType And newtypesParam.Item(i).MenuID = FormType Then
                        ReportExists = True
                        Exit For
                    End If
                Next i

                If Not ReportExists Then
                    rptTypeService = objaddon.objcompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
                    newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)

                    newType.TypeName = FormType 'clsJobCard.FormType
                    newType.AddonName = AddonName ' "Sub-Con Add-on"
                    newType.AddonFormType = FormType
                    newType.MenuID = FormType
                    newtypeParam = rptTypeService.AddReportType(newType)

                    Dim rptService As SAPbobsCOM.ReportLayoutsService
                    Dim newReport As SAPbobsCOM.ReportLayout
                    rptService = objaddon.objcompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
                    newReport = rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout)
                    newReport.Author = objaddon.objcompany.UserName
                    newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
                    newReport.Name = FormType
                    newReport.TypeCode = newtypeParam.TypeCode

                    newReportParam = rptService.AddReportLayout(newReport)

                    newType = rptTypeService.GetReportType(newtypeParam)
                    newType.DefaultReportLayout = newReportParam.LayoutCode
                    rptTypeService.UpdateReportType(newType)

                    Dim oBlobParams As SAPbobsCOM.BlobParams
                    oBlobParams = objaddon.objcompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
                    oBlobParams.Table = "RDOC"
                    oBlobParams.Field = "Template"
                    Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment
                    oKeySegment = oBlobParams.BlobTableKeySegments.Add
                    oKeySegment.Name = "DocCode"
                    oKeySegment.Value = newReportParam.LayoutCode

                    Dim oFile As FileStream
                    oFile = New FileStream(System.Windows.Forms.Application.StartupPath + "\Sample.rpt", FileMode.Open)
                    Dim fileSize As Integer
                    fileSize = oFile.Length
                    Dim buf(fileSize) As Byte
                    oFile.Read(buf, 0, fileSize)
                    oFile.Dispose()

                    Dim oBlob As SAPbobsCOM.Blob
                    oBlob = objaddon.objcompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob)
                    oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
                    objaddon.objcompany.GetCompanyService.SetBlob(oBlobParams, oBlob)
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(" addReport_Layouttype Method Failed :  " & ex.Message + strsql, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Public Sub setReport(ByVal FormType As String, ByVal FormCount As Integer)
            Try
                Dim objform As SAPbouiCOM.Form
                'objform = objaddon.objapplication.Forms.Item(FormUID)
                objform = objaddon.objapplication.Forms.GetForm("SUBCTPO", FormCount)
                Dim rptTypeService As SAPbobsCOM.ReportTypesService
                'Dim newType As SAPbobsCOM.ReportType
                Dim newtypesParam As SAPbobsCOM.ReportTypesParams
                rptTypeService = objaddon.objcompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
                newtypesParam = rptTypeService.GetReportTypeList
                Dim TypeCode As String
                If objaddon.HANA Then
                    TypeCode = getSingleValue("Select ""CODE"" from RTYP where ""NAME""='" & FormType & "'")
                Else
                    TypeCode = getSingleValue("Select CODE from RTYP where NAME='" & FormType & "'")
                End If
                objform.ReportType = TypeCode
                'Dim i As Integer
                'For i = 0 To newtypesParam.Count - 1
                '    If newtypesParam.Item(i).TypeName = FormType And newtypesParam.Item(i).MenuID = FormType Then
                '        objform.ReportType = newtypesParam.Item(i).TypeCode
                '        Exit For
                '    End If
                'Next i
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("setReport Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Public Sub WriteErrorLog(ByVal Str As String)
            Dim Foldername, Attachpath As String
            If objaddon.HANA Then
                Attachpath = getSingleValue("select ""AttachPath"" from OADP")
            Else
                Attachpath = getSingleValue("select AttachPath from OADP")
            End If

            Foldername = Attachpath + "Log\SubCon"
            If Directory.Exists(Foldername) Then
            Else
                Directory.CreateDirectory(Foldername)
            End If

            Dim fs As FileStream
            Dim chatlog As String = Foldername & "\Log_" & System.DateTime.Now.ToString("ddMMyyHHmmss") & ".txt"
            If File.Exists(chatlog) Then
            Else
                fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
                fs.Close()
            End If
            Dim sdate As String
            sdate = Now
            If System.IO.File.Exists(chatlog) = True Then
                Dim objWriter As New System.IO.StreamWriter(chatlog, True)
                objWriter.WriteLine(sdate & " : " & Str)
                objWriter.Close()
            Else
                Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            End If
        End Sub

        Public Sub AddToPermissionTree(ByVal Name As String, ByVal PermissionID As String, ByVal FormType As String, ByVal ParentID As String, ByVal AddPermission As Char)
            Try
                Dim RetVal As Long
                Dim ErrMsg As String = ""
                Dim oPermission As SAPbobsCOM.UserPermissionTree
                Dim objBridge As SAPbobsCOM.SBObob
                Dim objrs As SAPbobsCOM.Recordset
                If ParentID <> "" Then

                    If objaddon.HANA = True Then
                        strsql = objaddon.objglobalmethods.getSingleValue("Select 1 as ""Status"" from OUPT Where ""AbsId""='" & ParentID & "'")
                    Else
                        strsql = objaddon.objglobalmethods.getSingleValue("Select 1 as Status from OUPT Where AbsId='" & ParentID & "'")
                    End If

                    If strsql = "" Then Return
                End If

                oPermission = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
                objBridge = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs = objBridge.GetUserList()

                If oPermission.GetByKey(PermissionID) = False Then
                    oPermission.Name = Name
                    oPermission.PermissionID = PermissionID
                    oPermission.UserPermissionForms.FormType = FormType
                    If ParentID <> "" Then oPermission.ParentID = ParentID
                    oPermission.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
                    RetVal = oPermission.Add()
                    Dim temp_int As Integer = CInt((RetVal))
                    Dim temp_string As String = ErrMsg
                    objaddon.objcompany.GetLastError(temp_int, temp_string)

                    If RetVal <> 0 Then
                        objaddon.objapplication.StatusBar.SetText("AddToPermissionTree: " & temp_string, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        If AddPermission = "N"c Then Return
                        For i As Integer = 0 To objrs.RecordCount - 1
                            If objaddon.HANA = True Then
                                strsql = "Select ""USERID"" from OUSR Where ""USER_CODE""='" & Convert.ToString(objrs.Fields.Item(0).Value) & "'"
                            Else
                                strsql = "Select USERID from OUSR Where USER_CODE='" & Convert.ToString(objrs.Fields.Item(0).Value) & "'"
                            End If
                            strsql = objaddon.objglobalmethods.getSingleValue(strsql)
                            objaddon.objglobalmethods.AddPermissionToUsers(Convert.ToInt32(strsql), PermissionID)
                            objrs.MoveNext()
                        Next
                    End If
                    'Else
                    '    oPermission.Remove()
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Permission: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Sub AddPermissionToUsers(ByVal UserCode As Integer, ByVal PermissionID As String)
            Try
                Dim oUser As SAPbobsCOM.Users = Nothing
                Dim lRetCode As Integer
                Dim sErrMsg As String = ""
                oUser = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)

                If oUser.GetByKey(UserCode) = True Then
                    oUser.UserPermission.Add()
                    oUser.UserPermission.SetCurrentLine(0)
                    oUser.UserPermission.PermissionID = PermissionID
                    oUser.UserPermission.Permission = SAPbobsCOM.BoPermission.boper_Full
                    lRetCode = oUser.Update()
                    objaddon.objcompany.GetLastError(lRetCode, sErrMsg)

                    If lRetCode <> 0 Then
                        objaddon.objapplication.StatusBar.SetText("AddPermissionToUser: " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If

            Catch ex As Exception
            End Try
        End Sub

        Public Sub Update_UserFormSettings_UDF(ByVal form As SAPbouiCOM.Form, ByVal FormID As String, ByVal UserCode As Integer)
            Try
                Dim oCmpSrv As SAPbobsCOM.CompanyService
                Dim oFormPreferencesService As FormPreferencesService
                Dim oColsPreferences As ColumnsPreferences
                Dim oColPreferencesParams As ColumnsPreferencesParams

                strsql = objaddon.objglobalmethods.getSingleValue("Select count(*) from CPRF Where FormID='" & FormID & "' and UserSign=" & UserCode & " and TPLId<>0 and VisInForm='Y'")
                If strsql = "0" Then Exit Sub
                strsql = "Update T0 Set T0.TPLId=(Select TPLId from OUSR where USERID=T0.UserSign) from CPRF T0 Where T0.FormID='" & FormID & "' and T0.UserSign=" & UserCode & ""


                oCmpSrv = clsModule.objaddon.objcompany.GetCompanyService()
                oFormPreferencesService = oCmpSrv.GetBusinessService(ServiceTypes.FormPreferencesService)
                oColPreferencesParams = oFormPreferencesService.GetDataInterface(FormPreferencesServiceDataInterfaces.fpsdiColumnsPreferencesParams)
                oColPreferencesParams.FormID = FormID
                oColPreferencesParams.User = UserCode
                oColsPreferences = oFormPreferencesService.GetColumnsPreferences(oColPreferencesParams)

                For i As Integer = 0 To oColsPreferences.Count - 1

                    oColsPreferences.Item(i).EditableInForm = BoYesNoEnum.tNO
                    oColsPreferences.Item(i).VisibleInForm = BoYesNoEnum.tNO
                    'oColsPreferences.Item(i).TabsLayout = 0
                    'oColsPreferences.Item(i).FormID = FormID
                    strsql = oColsPreferences.Item(i).ItemNumber
                    strsql = strsql
                Next

                oFormPreferencesService.UpdateColumnsPreferences(oColPreferencesParams, oColsPreferences)
            Catch ex As Exception
                Throw
            End Try
        End Sub

    End Class

End Namespace
