Module modARCreditMemo

    Function AR_CreditMemo(ByRef oDICompany As SAPbobsCOM.Company, _
                              ByVal sDocEntry As String, ByVal dIncomeDate As Date, _
                               ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oARCreditmemo As SAPbobsCOM.Documents
        Dim oARInvoice As SAPbobsCOM.Documents
        Dim sPayDocEntry As String = String.Empty

        oARCreditmemo = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
        oARInvoice = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)


        Try
            sFuncName = "AR_CreditMemo"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If oARInvoice.GetByKey(sDocEntry) Then

                oARCreditmemo.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                oARCreditmemo.CardCode = oARInvoice.CardCode
                oARCreditmemo.NumAtCard = oARInvoice.NumAtCard
                oARCreditmemo.DocDate = oARInvoice.DocDate
                oARCreditmemo.TaxDate = oARInvoice.TaxDate
                oARCreditmemo.DocDueDate = oARInvoice.DocDueDate
                oARCreditmemo.DocType = oARInvoice.DocType

                For imjs As Integer = 0 To oARInvoice.Lines.Count - 1
                    oARInvoice.Lines.SetCurrentLine(imjs)
                    oARCreditmemo.Lines.BaseEntry = sDocEntry
                    oARCreditmemo.Lines.BaseLine = oARInvoice.Lines.LineNum
                    oARCreditmemo.Lines.BaseType = 13
                    oARCreditmemo.Lines.AccountCode = oARInvoice.Lines.AccountCode
                    '---------------------- Batch Information
                    For ibount As Integer = 0 To oARInvoice.Lines.BatchNumbers.Count - 1
                        oARCreditmemo.Lines.BatchNumbers.SetCurrentLine(ibount)
                        oARCreditmemo.Lines.BatchNumbers.BatchNumber = oARInvoice.Lines.BatchNumbers.BatchNumber
                        oARCreditmemo.Lines.BatchNumbers.Quantity = oARInvoice.Lines.BatchNumbers.Quantity
                        oARCreditmemo.Lines.BatchNumbers.Add()
                    Next
                    oARCreditmemo.Lines.Add()
                Next
                lRetCode = oARCreditmemo.Add()
                If lRetCode <> 0 Then
                    sErrDesc = oDICompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    AR_CreditMemo = RTN_ERROR
                Else
                    sErrDesc = String.Empty
                    Console.WriteLine("Completed with SUCCESS", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                    AR_CreditMemo = RTN_SUCCESS
                End If

            Else
                sErrDesc = "No matching records found in the AR Invoice " & sDocEntry
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                AR_CreditMemo = RTN_ERROR

            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AR_CreditMemo = RTN_ERROR
        Finally
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARCreditmemo)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
        End Try
    End Function

    Public Function AR_CreditMemo_Standalone(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, _
                                      ByRef sDocEntry As String, ByRef sDocNum As String, ByVal sCardCode As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oARInvoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        Dim oARInvoice_Doc As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
        Dim dIncomeDate As Date
        Dim tDocTime As DateTime
        Dim sWhsCode As String = String.Empty
        Dim sPOSNumber As String = String.Empty
        Dim sProductCode As String = String.Empty
        Dim sBOMCode As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim sQueryup As String = String.Empty
        Dim sManBatchItem As String = String.Empty
        Dim oBatchDT As DataTable = Nothing
        Dim dBatchQuantity As Double = 0
        Dim dRemBatchQuantity As Double = 0
        Dim dBatchNumber As String = String.Empty
        Dim dInvQuantity As Double
        Dim lRetCode As Integer
        Dim irow As Integer = 0
        Dim dDocTotal As Double = 0.0
        Dim oDV_BOM As DataView = New DataView(oDT_BOM)
        Dim oDT_Batch As DataTable = New DataTable
        Dim oDV_Batch As DataView = Nothing
        Dim oRow() As Data.DataRow = Nothing
        Dim SARDraft As String = String.Empty
        Dim dPostxdatetime As Date
        Dim oDT_Payamount As DataTable = New DataTable
        Dim dPayamount As Double = 0
        oDT_Payamount = oDVPayment.ToTable

        If oDT_Payamount.Rows.Count > 0 Then
            dPayamount = Convert.ToDecimal(oDT_Payamount.Compute("sum(PaymentAmount)", String.Empty).ToString)
        End If


        oDT_Batch.Columns.Add("ItemCode", GetType(String))
        oDT_Batch.Columns.Add("BatchNum", GetType(String))
        oDT_Batch.Columns.Add("Quantity", GetType(Decimal))
        '' oDT_Batch.Columns.Add("date", GetType(Date))


        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRset_Batch As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            sFuncName = "AR_CreditMemo_Standalone()"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)
            sWhsCode = CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AR Invoice dPostxdatetime " & dPostxdatetime, sFuncName)
            dPostxdatetime = oDVARInvoice.Item(0).Row("HPOSTxDatetime").ToString.Trim

            oARInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes

            tDocTime = tDocTime.AddHours(0)
            tDocTime = tDocTime.AddMinutes(0)
            oARInvoice.CardCode = sCardCode
            oARInvoice.DocDate = dIncomeDate
            oARInvoice.DocDueDate = dIncomeDate
            oARInvoice.TaxDate = dIncomeDate
            oARInvoice.NumAtCard = sWhsCode & " - " & sPOSNumber

            oARInvoice.UserFields.Fields.Item("U_POS_RefNo").Value = oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim
            oARInvoice.UserFields.Fields.Item("U_Date").Value = dIncomeDate
            oARInvoice.UserFields.Fields.Item("U_Time").Value = dPostxdatetime

            For Each dvr As DataRowView In oDVARInvoice
                oARInvoice.Lines.ItemCode = dvr("DItemCode").ToString.Trim
                oARInvoice.Lines.Quantity = CDbl(dvr("DQuantity").ToString.Trim)
                '' MsgBox(dvr("DPrice").ToString.Trim)
                '' oARInvoice.Lines.Price = CDbl(dvr("DPrice").ToString.Trim)
                oARInvoice.Lines.LineTotal = CDbl(dvr("DLineTotal").ToString.Trim)
                oARInvoice.Lines.WarehouseCode = sWhsCode
                If Not String.IsNullOrEmpty(p_oCompDef.p_sGLAccount) Then
                    oARInvoice.Lines.AccountCode = p_oCompDef.p_sGLAccount
                End If
                oARInvoice.Lines.VatGroup = dvr("VatGourpSa").ToString.Trim
                oARInvoice.Lines.Add()
            Next

            If dPayamount > 0 Then
                oARInvoice.DocTotal = dPayamount
            End If

            If oCompany.InTransaction = False Then oCompany.StartTransaction()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add Draft ", sFuncName)
            lRetCode = oARInvoice.Add()

            If lRetCode <> 0 Then
                sErrDesc = oCompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                Return RTN_ERROR

            Else
                '' System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                ''oARInvoice = Nothing
                '----------------- AR Invoice Draft Created Successfully
                oCompany.GetNewObjectCode(sDocEntry)
                Console.WriteLine("Draft Added Successfully " & sDocEntry, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Added Successfully  " & sDocEntry, sFuncName)
                Console.WriteLine("Assigning Batch   " & sDocEntry, sFuncName)

                If oARInvoice.GetByKey(sDocEntry) Then
                    dDocTotal = oARInvoice.DocTotal
                    sQuery = "SELECT T0.[LineNum], T0.[ItemCode], T0.[Quantity], T0.[WhsCode], T1.[ManBtchNum] FROM DRF1 T0 WITH (NOLOCK) INNER JOIN OITM T1 WITH (NOLOCK) ON T0.[ItemCode] = T1.[ItemCode] WHERE T0.[DocEntry] = '" & sDocEntry & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Details SQL " & sQuery, sFuncName)
                    oRset.DoQuery(sQuery)
                    For imjs As Integer = 0 To oRset.RecordCount - 1
                        sProductCode = oRset.Fields.Item("ItemCode").Value
                        sManBatchItem = oRset.Fields.Item("ManBtchNum").Value
                        irow = oRset.Fields.Item("LineNum").Value 'Row Number
                        dInvQuantity = CDbl(oRset.Fields.Item("Quantity").Value) 'Item Quantity
                        If sManBatchItem = "Y" Then
                            sQuery = "SELECT BatchNum ,Quantity , SysNumber  FROM OIBT WITH (NOLOCK) WHERE ItemCode ='" & sProductCode & "' and Quantity >0 " & _
                                              "AND WhsCode ='" & sWhsCode & "' ORDER BY InDate ASC "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Informations SQL " & sQuery, sFuncName)
                            oRset_Batch.DoQuery(sQuery)
                            For iloop As Integer = 0 To oRset_Batch.RecordCount - 1

                                dBatchQuantity = CDbl(oRset_Batch.Fields.Item("Quantity").Value) 'Batch Quantity
                                dBatchNumber = oRset_Batch.Fields.Item("BatchNum").Value 'Batch

                                oARInvoice.Lines.SetCurrentLine(irow)
                                oARInvoice.Lines.BatchNumbers.SetCurrentLine(iloop)

                                If dInvQuantity > 0 Then
                                    If oDT_Batch.Rows.Count = 0 Then
                                        oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                        If dInvQuantity > dBatchQuantity Then
                                            'If Balance Qty>Batch Qty, then get full Batch Qty
                                            oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                            'minus current qty with Batch Qty
                                            dInvQuantity = dInvQuantity - dBatchQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                        Else
                                            oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                            dInvQuantity = dInvQuantity - dInvQuantity
                                        End If

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                        oARInvoice.Lines.BatchNumbers.Add()
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)

                                        If dInvQuantity <= 0 Then Exit For
                                    Else
                                        oDV_Batch = New DataView(oDT_Batch)
                                        oDV_Batch.RowFilter = "ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'"
                                        If oDV_Batch.Count > 0 Then
                                            dRemBatchQuantity = oDV_Batch.Item(0).Row("Quantity")
                                            If dRemBatchQuantity > dInvQuantity Then
                                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                                oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                                oRow(0)("Quantity") = oDV_Batch.Item(0).Row("Quantity") - dInvQuantity
                                                oARInvoice.Lines.BatchNumbers.Add()
                                                Exit For
                                            End If
                                            oARInvoice.Lines.BatchNumbers.Quantity = dRemBatchQuantity
                                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                            oRow(0)("Quantity") = 0
                                            dInvQuantity = dInvQuantity - dRemBatchQuantity
                                            oARInvoice.Lines.BatchNumbers.Add()

                                        Else
                                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            If dInvQuantity > dBatchQuantity Then
                                                'If Balance Qty>Batch Qty, then get full Batch Qty
                                                oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                                'minus current qty with Batch Qty
                                                dInvQuantity = dInvQuantity - dBatchQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                            Else
                                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                                dInvQuantity = dInvQuantity - dInvQuantity
                                            End If

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                            oARInvoice.Lines.BatchNumbers.Add()
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)
                                            If dInvQuantity <= 0 Then Exit For
                                        End If
                                    End If
                                Else
                                    '-------------------------- -ve quantity
                                    oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                    oARInvoice.Lines.BatchNumbers.Quantity = Math.Abs(dInvQuantity)
                                    oARInvoice.Lines.BatchNumbers.Add()
                                    Exit For
                                End If
                                oRset_Batch.MoveNext()
                            Next iloop
                        End If
                        oRset.MoveNext()
                    Next imjs

                    ''Dim dblRoundAmt As Double = 0.0
                    ''If oDVPayment.Count > 0 Then
                    ''    dblRoundAmt = dPayamount - oARInvoice.DocTotal
                    ''    If dblRoundAmt <> 0 Then
                    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calculating Rounding Amount: " & dblRoundAmt, sFuncName)
                    ''        oARInvoice.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    ''        oARInvoice.RoundingDiffAmount = dblRoundAmt
                    ''    End If
                    ''End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Update the AR Credit Memo Draft with Batch Information", sFuncName)
                    lRetCode = oARInvoice.Update() 'Update the batch information
                    Console.WriteLine("Batch Updated Successfully " & sDocEntry, sFuncName)
                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Update AR Credit Memo Draft) ", sFuncName)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                        Return RTN_ERROR

                    End If

                    SARDraft = sDocEntry
                    Console.WriteLine("Attempting to Convert as a AR Invoice Document", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Update AR Credit Memo Draft) ", sFuncName)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Convert as a AR Credit Memo Document ", sFuncName)

                    lRetCode = oARInvoice.SaveDraftToDocument()

                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Convert as a AR Credit Memo Document) ", sFuncName)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                        Return RTN_ERROR
                    End If
                    oCompany.GetNewObjectCode(sDocEntry)
                    oARInvoice_Doc.GetByKey(sDocEntry)
                    sDocNum = oARInvoice_Doc.DocNum
                    Console.WriteLine("Converted To AR Credit Memo Successful " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Convert as a AR Credit Memo Document) " & sDocEntry, sFuncName)
                End If

                Return RTN_SUCCESS
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Console.WriteLine("Completed with ERROR", sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oARInvoice = Nothing
            Return RTN_ERROR
        Finally

            If Not String.IsNullOrEmpty(SARDraft) Then
                If oARInvoice.GetByKey(SARDraft) Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Remove te Draft ", sFuncName)
                    lRetCode = oARInvoice.Remove()
                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice_Doc)
        End Try
    End Function

End Module
