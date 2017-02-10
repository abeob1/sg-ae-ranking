Module modARInvoice


    Public Function AR_InvoiceCreation(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, _
                                       ByRef sDocEntry As String, ByRef sDocNum As String, ByVal sCardCode As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oARInvoice As SAPbobsCOM.Documents = Nothing
        Dim oARInvoice_Doc As SAPbobsCOM.Documents = Nothing
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
        Dim dHDocTotal As Double = 0
        oDT_Payamount = oDVPayment.ToTable
        Dim fBatch As Boolean = False

        If oDT_Payamount.Rows.Count > 0 Then
            dPayamount = Convert.ToDecimal(oDT_Payamount.Compute("sum(PaymentAmount)", String.Empty).ToString)
        End If


        oDT_Batch.Columns.Add("ItemCode", GetType(String))
        oDT_Batch.Columns.Add("BatchNum", GetType(String))
        oDT_Batch.Columns.Add("Quantity", GetType(Decimal))
        '' oDT_Batch.Columns.Add("date", GetType(Date))


        Dim oRset As SAPbobsCOM.Recordset = Nothing
        Dim oRset_Batch As SAPbobsCOM.Recordset = Nothing

        Try
            oARInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            oARInvoice_Doc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRset_Batch = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sFuncName = "AR_InvoiceCreation()"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AR Invoice dPostxdatetime " & dPostxdatetime, sFuncName)
            dPostxdatetime = oDVARInvoice.Item(0).Row("HPOSTxDatetime").ToString.Trim

            oARInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

            tDocTime = tDocTime.AddHours(0)
            tDocTime = tDocTime.AddMinutes(0)
            oARInvoice.CardCode = sCardCode
            oARInvoice.DocDate = dIncomeDate
            oARInvoice.DocDueDate = dIncomeDate
            oARInvoice.TaxDate = dIncomeDate
            oARInvoice.NumAtCard = sPOSNumber
            ''sWhsCode & " - " & sPOSNumber
            If Not String.IsNullOrEmpty(oDVARInvoice.Item(0).Row("DSalesman").ToString.Trim) Then
                oARInvoice.SalesPersonCode = oDVARInvoice.Item(0).Row("DSalesman").ToString.Trim
            End If
            oARInvoice.UserFields.Fields.Item("U_AB_POSTxNo").Value = oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim
            oARInvoice.UserFields.Fields.Item("U_AB_Date").Value = dIncomeDate
            oARInvoice.UserFields.Fields.Item("U_AB_Time").Value = dPostxdatetime
            If Not String.IsNullOrEmpty(oDVARInvoice.Item(0).Row("HCommEntitle").ToString.Trim) Then
                oARInvoice.UserFields.Fields.Item("U_AB_CommEntitle").Value = oDVARInvoice.Item(0).Row("HCommEntitle").ToString.Trim
            End If
            If Not String.IsNullOrEmpty(oDVARInvoice.Item(0).Row("HCommPercent").ToString.Trim) Then
                oARInvoice.UserFields.Fields.Item("U_AB_CommPercent").Value = oDVARInvoice.Item(0).Row("HCommPercent").ToString.Trim
            End If

            For Each dvr As DataRowView In oDVARInvoice
                oARInvoice.Lines.ItemCode = dvr("DItemCode").ToString.Trim
                oARInvoice.Lines.Quantity = Math.Abs(CDbl(dvr("DQuantity").ToString.Trim))
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item code " & dvr("DItemCode").ToString.Trim, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Quantity " & Math.Abs(CDbl(dvr("DQuantity").ToString.Trim)), sFuncName)
                oDT_Warehouse.DefaultView.RowFilter = "U_AB_POSLocCode='" & dvr("DOutlet").ToString.Trim & "'"
                If oDT_Warehouse.DefaultView.Count > 0 Then
                    sWhsCode = oDT_Warehouse.DefaultView(0)(0).ToString().Trim()
                Else
                    sErrDesc = "No matching records found in OWHS table " & CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)
                    Return RTN_ERROR
                End If
                oARInvoice.Lines.LineTotal = CDbl(dvr("DLineTotal").ToString.Trim) - CDbl(dvr("DTotalGST").ToString.Trim)
                oARInvoice.Lines.UnitPrice = CDbl(dvr("DPriceBefDi").ToString.Trim) / 1.07
                oARInvoice.Lines.TaxTotal = CDbl(dvr("DTotalGST").ToString.Trim)
                oARInvoice.Lines.WarehouseCode = sWhsCode
                oARInvoice.Lines.VatGroup = dvr("VatGourpSa").ToString.Trim
                If Not String.IsNullOrEmpty(dvr("DSalesman").ToString.Trim) Then
                    oARInvoice.Lines.SalesPersonCode = dvr("DSalesman").ToString.Trim
                End If

                oARInvoice.Lines.Add()
            Next

            ''If dPayamount > 0 Then
            ''    oARInvoice.DocTotal = dPayamount
            ''End If
            oARInvoice.DocTotal = oDVARInvoice.Item(0).Row("HDocTotal").ToString.Trim


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
                            fBatch = True
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

                    If fBatch = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Update the AR Invoice Draft with Batch Information", sFuncName)
                        lRetCode = oARInvoice.Update() 'Update the batch information
                        Console.WriteLine("Batch Updated Successfully " & sDocEntry, sFuncName)
                        If lRetCode <> 0 Then
                            sErrDesc = oCompany.GetLastErrorDescription
                            'If sErrDesc = "Internal error (-10) occurred" Then
                            '    sErrDesc = "Quantity falls into negative inventory  [INV1.ItemCode][line: 2]"
                            'End If
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Update AR Invoice Draft) ", sFuncName)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                            Return RTN_ERROR

                        End If
                        fBatch = False
                    End If
                  

                    SARDraft = sDocEntry
                    Console.WriteLine("Attempting to Convert as a AR Invoice Document", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Update AR Invoice Draft) ", sFuncName)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Convert as a AR Invoice Document ", sFuncName)

                    lRetCode = oARInvoice.SaveDraftToDocument()

                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Exception " & sErrDesc, sFuncName)
                        If Left(sErrDesc.ToUpper(), 14) = "INTERNAL ERROR" Then
                            sErrDesc = "Quantity falls into negative inventory  [INV1.ItemCode][line: 2]"
                        End If
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Convert as a AR Invoice Document) ", sFuncName)
                        Return RTN_ERROR
                    End If
                    oCompany.GetNewObjectCode(sDocEntry)
                    oARInvoice_Doc.GetByKey(sDocEntry)
                    sDocNum = oARInvoice_Doc.DocNum
                    Console.WriteLine("Converted To AR Invoice Successful " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Convert as a AR Invoice Document) " & sDocEntry, sFuncName)
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
            oARInvoice = Nothing
            oARInvoice_Doc = Nothing
            oRset = Nothing
            oRset_Batch = Nothing
        End Try
    End Function


    Public Function AR_InvoiceCreation_OLD1(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, ByRef oDTStatus As DataTable, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oARInvoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        Dim dIncomeDate As Date
        Dim tDocTime As DateTime
        Dim sWhsCode As String = String.Empty
        Dim sPOSNumber As String = String.Empty
        Dim sProductCode As String = String.Empty
        Dim sBOMCode As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim sManBatchItem As String = String.Empty
        Dim oBatchDT As DataTable = Nothing
        Dim dBatchQuantity As Double = 0
        Dim dBatchNumber As String = String.Empty
        Dim dInvQuantity As Double
        Dim sDocEntry As String = String.Empty
        Dim lRetCode As Integer
        Dim irow As Integer = 0

        Dim oDV_BOM As DataView = New DataView(oDT_BOM)

        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            sFuncName = "AR_InvoiceCreation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)
            sWhsCode = CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)

            '' oARInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

            tDocTime = tDocTime.AddHours(0)
            tDocTime = tDocTime.AddMinutes(0)

            oARInvoice.CardCode = p_oCompDef.p_sCardCode
            oARInvoice.DocDate = dIncomeDate
            oARInvoice.DocDueDate = dIncomeDate
            oARInvoice.TaxDate = dIncomeDate
            oARInvoice.NumAtCard = sWhsCode & " - " & sPOSNumber

            oARInvoice.UserFields.Fields.Item("U_AB_POSTxNo").Value = oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim
            oARInvoice.UserFields.Fields.Item("U_AB_Date").Value = dIncomeDate
            oARInvoice.UserFields.Fields.Item("U_AB_Time").Value = tDocTime

            '' oDV_BOM.RowFilter = "HeaderID = '" & oDVARInvoice.Item(0).Row("HTransID").ToString.Trim & "'"

            For Each dvr As DataRowView In oDVARInvoice
                '' sProductCode = dvr("ItemCode").ToString.Trim
                oARInvoice.Lines.ItemCode = dvr("DItemCode").ToString.Trim
                oARInvoice.Lines.Quantity = CDbl(dvr("DQuantity").ToString.Trim)
                oARInvoice.Lines.Price = CDbl(dvr("DPrice").ToString.Trim)
                oARInvoice.Lines.LineTotal = CDbl(dvr("DLineTotal").ToString.Trim)
                oARInvoice.Lines.WarehouseCode = sWhsCode
                '' oARInvoice.Lines.VatGroup = dvr("VatGourpSa").ToString.Trim

                sManBatchItem = dvr("ManBtchNum").ToString.Trim
                If sManBatchItem.ToUpper() = "Y" Then

                    sQuery = " SELECT BatchNum ,Quantity , SysNumber  FROM OIBT WITH (NOLOCK) WHERE ItemCode ='" & sProductCode & "' and Quantity >0 " & _
                                          "AND WhsCode ='" & sWhsCode & "' ORDER BY InDate ASC "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Query " & sQuery, sFuncName)
                    oRset.DoQuery(sQuery)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConvertRecordset() ", sFuncName)
                    oBatchDT = ConvertRecordset(oRset, sErrDesc)  ' Get_DataTable(sQuery, P_sSAPConString, sErrDesc)

                    For iBatchRow As Integer = 0 To oBatchDT.Rows.Count - 1

                        dBatchQuantity = CDbl(oBatchDT.Rows(iBatchRow)("Quantity").ToString().Trim())
                        dBatchNumber = oBatchDT.Rows(iBatchRow)("BatchNum").ToString().Trim()
                        dInvQuantity = CDbl(dvr("Quantity").ToString.Trim)

                        '' oARInvoice.Lines.BatchNumbers.InternalSerialNumber = oBatchDT.Rows(iBatchRow)("SysNumber").ToString().Trim()
                        ''oARInvoice.Lines.BatchNumbers.Location = irow
                        oARInvoice.Lines.BatchNumbers.SetCurrentLine(0)
                        oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                        If dInvQuantity > dBatchQuantity Then
                            'If Balance Qty>Batch Qty, then get full Batch Qty
                            oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                            'minus current qty with Batch Qty
                            dInvQuantity = dInvQuantity - dBatchQuantity
                        Else
                            oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                            dInvQuantity = dInvQuantity - dInvQuantity
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                        oARInvoice.Lines.BatchNumbers.Add()
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)
                        If dInvQuantity <= 0 Then Exit For
                    Next

                End If

                irow += 1
                oARInvoice.Lines.Add()
                If irow = 2 Then
                    Exit For
                End If
            Next



            oARInvoice.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
            oARInvoice.RoundingDiffAmount = CDbl(oDVARInvoice.Table.Rows(0).Item("HRounding").ToString.Trim)
            ''If oARInvoice.GetByKey(24216) Then
            ''    oARInvoice.SaveXML("E:\invoice1.xml")
            ''End If

            '' If oCompany.InTransaction = False Then oCompany.StartTransaction()

            oARInvoice.SaveXML("E:\Test123.xml")
            lRetCode = oARInvoice.Add()

            If lRetCode <> 0 Then
                sErrDesc = oCompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Update_Status() ", sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                Return RTN_ERROR

            Else
                '' System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                oARInvoice = Nothing
                oCompany.GetNewObjectCode(sDocEntry)

                If oDVARInvoice.Item(0).Row("HPOSTxType").ToString.Trim = "S" Then
                    '************************************ Incoming Payment Started ************************************************************************************

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_IncomingPayment() : AR Invoice DocEntry " & sDocEntry, sFuncName)

                    If AR_IncomingPayment(oDVPayment, oCompany, sDocEntry, dIncomeDate, sPOSNumber _
                                       , sWhsCode, p_oCompDef.p_sCardCode, sErrDesc) <> RTN_SUCCESS Then

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                        If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        oARInvoice = Nothing
                        Return RTN_ERROR
                    End If
                ElseIf oDVARInvoice.Item(0).Row("HPOSTxType").ToString.Trim = "V" Then
                    '************************************ AR Credit Memo ************************************************************************************

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_IncomingPayment() : AR Invoice DocEntry " & sDocEntry, sFuncName)

                    If AR_IncomingPayment(oDVPayment, oCompany, sDocEntry, dIncomeDate, sPOSNumber _
                                       , sWhsCode, p_oCompDef.p_sCardCode, sErrDesc) <> RTN_SUCCESS Then

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                        If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        oARInvoice = Nothing
                        Return RTN_ERROR
                    End If

                End If



                sErrDesc = ""

                ''  Update_Status(sTransID, sErrDesc, "SUCCESS", sDocEntry, "SalesTransHDR")
                oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString)
                If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Committed the Transaction Reference POSNumber : " & sPOSNumber, sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                '' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting the Company and Release the Object ", sFuncName)
                Return RTN_SUCCESS
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return RTN_ERROR
        End Try
    End Function

    Public Function AR_InvoiceCreation_OLD(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, ByRef oDTStatus As DataTable, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oARInvoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        Dim dIncomeDate As Date
        Dim tDocTime As DateTime
        Dim sWhsCode As String = String.Empty
        Dim sPOSNumber As String = String.Empty
        Dim sProductCode As String = String.Empty
        Dim sBOMCode As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim sManBatchItem As String = String.Empty
        Dim oBatchDT As DataTable = Nothing
        Dim dBatchQuantity As Double = 0
        Dim dBatchNumber As String = String.Empty
        Dim dInvQuantity As Double
        Dim sDocEntry As String = String.Empty
        Dim lRetCode As Integer
        Dim irow As Integer = 0

        Dim oDV_BOM As DataView = New DataView(oDT_BOM)

        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            sFuncName = "AR_InvoiceCreation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)
            sWhsCode = CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)

            '' oARInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

            tDocTime = tDocTime.AddHours(0)
            tDocTime = tDocTime.AddMinutes(0)

            oARInvoice.CardCode = p_oCompDef.p_sCardCode
            oARInvoice.DocDate = dIncomeDate
            oARInvoice.DocDueDate = dIncomeDate
            oARInvoice.TaxDate = dIncomeDate
            oARInvoice.NumAtCard = sWhsCode & " - " & sPOSNumber

            oARInvoice.UserFields.Fields.Item("U_AB_POSTxNo").Value = oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim
            oARInvoice.UserFields.Fields.Item("U_AB_Date").Value = dIncomeDate
            oARInvoice.UserFields.Fields.Item("U_AB_Time").Value = tDocTime

            oDV_BOM.RowFilter = "HeaderID = '" & oDVARInvoice.Item(0).Row("HTransID").ToString.Trim & "'"

            For Each dvr As DataRowView In oDV_BOM

                If sProductCode <> dvr("ItemCode").ToString.Trim Then
                    sProductCode = dvr("ItemCode").ToString.Trim
                    oARInvoice.Lines.ItemCode = sProductCode
                    oARInvoice.Lines.Quantity = CDbl(dvr("Quantity").ToString.Trim)
                    oARInvoice.Lines.Price = CDbl(dvr("Price").ToString.Trim)
                    oARInvoice.Lines.LineTotal = CDbl(dvr("LineTotal").ToString.Trim)
                    oARInvoice.Lines.WarehouseCode = sWhsCode
                    oARInvoice.Lines.VatGroup = dvr("VatGourpSa").ToString.Trim

                End If

                sManBatchItem = dvr("ManBtchNum").ToString.Trim
                If dvr("BOM").ToString.Trim = "BOM" Then
                    irow += 1
                    sBOMCode = dvr("Code").ToString.Trim

                    If sManBatchItem.ToUpper() = "Y" Then

                        sQuery = " SELECT BatchNum ,Quantity  FROM OIBT WITH (NOLOCK) WHERE ItemCode ='" & sBOMCode & "' and Quantity >0 " & _
                                              "AND WhsCode ='" & sWhsCode & "' ORDER BY InDate ASC "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Query " & sQuery, sFuncName)
                        oRset.DoQuery(sQuery)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConvertRecordset() ", sFuncName)
                        oBatchDT = ConvertRecordset(oRset, sErrDesc)  ' Get_DataTable(sQuery, P_sSAPConString, sErrDesc)

                        For iBatchRow As Integer = 0 To oBatchDT.Rows.Count - 1
                            dBatchQuantity = CDbl(oBatchDT.Rows(iBatchRow)("Quantity").ToString().Trim())
                            dBatchNumber = oBatchDT.Rows(iBatchRow)("BatchNum").ToString().Trim()
                            dInvQuantity = CDbl(dvr("Quantity").ToString.Trim) * CDbl(dvr("QuantityBOM").ToString.Trim)
                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                            If dInvQuantity > dBatchQuantity Then
                                'If Balance Qty>Batch Qty, then get full Batch Qty
                                oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                'minus current qty with Batch Qty
                                dInvQuantity = dInvQuantity - dBatchQuantity
                            Else
                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                dInvQuantity = dInvQuantity - dInvQuantity
                            End If
                            oARInvoice.Lines.SetCurrentLine(irow)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                            oARInvoice.Lines.BatchNumbers.Add()
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)
                            If dInvQuantity <= 0 Then Exit For
                        Next

                    End If
                ElseIf sManBatchItem.ToUpper() = "Y" Then
                    sQuery = " SELECT BatchNum ,Quantity  FROM OIBT WITH (NOLOCK) WHERE ItemCode ='" & sProductCode & "' and Quantity >0 " & _
                                              "AND WhsCode ='" & sWhsCode & "' ORDER BY InDate ASC "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Query " & sQuery, sFuncName)
                    oRset.DoQuery(sQuery)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConvertRecordset() ", sFuncName)
                    oBatchDT = ConvertRecordset(oRset, sErrDesc)  ' Get_DataTable(sQuery, P_sSAPConString, sErrDesc)

                    For iBatchRow As Integer = 0 To oBatchDT.Rows.Count - 1
                        dBatchQuantity = CDbl(oBatchDT.Rows(iBatchRow)("Quantity").ToString().Trim())
                        dBatchNumber = oBatchDT.Rows(iBatchRow)("BatchNum").ToString().Trim()
                        dInvQuantity = CDbl(dvr("Quantity").ToString.Trim)
                        oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                        If dInvQuantity > dBatchQuantity Then
                            'If Balance Qty>Batch Qty, then get full Batch Qty
                            oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                            'minus current qty with Batch Qty
                            dInvQuantity = dInvQuantity - dBatchQuantity
                        Else
                            oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                            dInvQuantity = dInvQuantity - dInvQuantity
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                        oARInvoice.Lines.SetCurrentLine(irow)
                        oARInvoice.Lines.BatchNumbers.Add()
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)
                        If dInvQuantity <= 0 Then Exit For
                    Next
                    irow += 1
                End If



                oARInvoice.Lines.Add()
            Next

            oARInvoice.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
            oARInvoice.RoundingDiffAmount = CDbl(oDVARInvoice.Table.Rows(0).Item("HRounding").ToString.Trim)
            ''If oARInvoice.GetByKey(24216) Then
            ''    oARInvoice.SaveXML("E:\invoice1.xml")
            ''End If

            '' If oCompany.InTransaction = False Then oCompany.StartTransaction()


            lRetCode = oARInvoice.Add()

            If lRetCode <> 0 Then
                sErrDesc = oCompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Update_Status() ", sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                Return RTN_ERROR

            Else
                '' System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                oARInvoice = Nothing
                oCompany.GetNewObjectCode(sDocEntry)

                '************************************ Incoming Payment Started ************************************************************************************

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_IncomingPayment() : AR Invoice DocEntry " & sDocEntry, sFuncName)

                If AR_IncomingPayment(oDVPayment, oCompany, sDocEntry, dIncomeDate, sPOSNumber _
                                   , sWhsCode, p_oCompDef.p_sCardCode, sErrDesc) <> RTN_SUCCESS Then

                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                    If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    oARInvoice = Nothing
                    Return RTN_ERROR
                End If

                sErrDesc = ""

                ''  Update_Status(sTransID, sErrDesc, "SUCCESS", sDocEntry, "SalesTransHDR")
                oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString)
                If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Committed the Transaction Reference POSNumber : " & sPOSNumber, sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                '' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting the Company and Release the Object ", sFuncName)
                Return RTN_SUCCESS
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return RTN_ERROR
        End Try
    End Function

    Function AR_Invoice_Cancel(ByRef oDICompany As SAPbobsCOM.Company, _
                               ByRef sInvoice As String, ByVal dDate As Date, ByVal spostdate As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oARInvoice As SAPbobsCOM.Documents = Nothing
        Dim oARInvoiceCancellation As SAPbobsCOM.Documents = Nothing
        Dim oRset As SAPbobsCOM.Recordset = Nothing
        oARInvoice = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        oRset = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim stime As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim sDocEntry As String = String.Empty

        Try
            

            sFuncName = "AR_Invoice_Cancel"
            Console.WriteLine("Starting Function", sFuncName)

            Dim sString() As String = spostdate.Split(" ")
            stime = Left(sString(1).Replace(":", ""), 4)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Before Time Split " & spostdate, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("After Time Split " & stime, sFuncName)
            If oARInvoice.GetByKey(sInvoice) Then
                oARInvoiceCancellation = oARInvoice.CreateCancellationDocument()
                ''oARInvoiceCancellation.DocDate = dDate
                ''oARInvoiceCancellation.DocDueDate = dDate
                oARInvoiceCancellation.UserFields.Fields.Item("U_AB_Date").Value = dDate
                ''  oARInvoiceCancellation.UserFields.Fields.Item("U_AB_Time").Value = "1132"

                lRetCode = oARInvoiceCancellation.Add()

                If lRetCode <> 0 Then
                    sErrDesc = oDICompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    AR_Invoice_Cancel = RTN_ERROR
                Else
                    Console.WriteLine("Completed with SUCCESS ", sFuncName)
                    oDICompany.GetNewObjectCode(sDocEntry)
                    '' sSQL = "Update OINV set [U_AB_Time] = '" & stime & "' FROM OINV T0 WHERE DocEntry = '" & sDocEntry & "'"
                    sSQL = "Update OINV set [U_AB_Time] = '" & stime & "' where DocEntry = '" & sDocEntry & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Time Change (Invoice) " & sSQL, sFuncName)
                    oRset.DoQuery(sSQL)
                    sErrDesc = String.Empty
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                    AR_Invoice_Cancel = RTN_SUCCESS
                    sErrDesc = String.Empty
                End If
            Else

                sErrDesc = "No matching records found in the AR Invoice " & sInvoice
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                AR_Invoice_Cancel = RTN_ERROR
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AR_Invoice_Cancel = RTN_ERROR

        Finally
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoiceCancellation)
            oARInvoice = Nothing
            oARInvoiceCancellation = Nothing
            oRset = Nothing
        End Try
    End Function
End Module
