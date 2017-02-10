Module OutgoingPayment

    Function Outgoing_Payment(ByRef oDVPayment As DataView, ByRef oDICompany As SAPbobsCOM.Company, _
                               ByVal sDocEntry As String, ByVal dIncomeDate As Date, _
                               ByVal sPOSNumber As String, ByVal sWhsCode As String, _
                              ByVal sCardCode As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oVendorPayment As SAPbobsCOM.Payments
        Dim sPayDocEntry As String = String.Empty

        Try
            sFuncName = "OutgoingPayment"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oVendorPayment = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

            Dim sCreditCard As String = String.Empty

            oVendorPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            oVendorPayment.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments
            oVendorPayment.CardCode = CStr(sCardCode)
            oVendorPayment.DocDate = dIncomeDate
            oVendorPayment.DueDate = dIncomeDate
            oVendorPayment.TaxDate = dIncomeDate
            oVendorPayment.CounterReference = Right(oDVPayment.Item(0).Row("POSTxNo").ToString.Trim, 8)

            If sDocEntry <> "" Then
                oVendorPayment.Invoices.DocEntry = sDocEntry
                oVendorPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote
                oVendorPayment.Invoices.Add()
            End If

            For Each drv In oDVPayment
                If drv("PaymentAmount").ToString.Trim = 0.0 Then Continue For

                oVendorPayment.CreditCards.CreditCard = drv("CreditCard").ToString.Trim
                oVendorPayment.CreditCards.CreditAcct = drv("AcctCode").ToString.Trim
                oVendorPayment.CreditCards.CreditSum = CDbl(drv("PaymentAmount").ToString.Trim)
                ' oVendorPayment.CreditCards.FirstPaymentDue = Now.Date
                oVendorPayment.CreditCards.CreditType = SAPbobsCOM.BoRcptCredTypes.cr_Regular
                oVendorPayment.CreditCards.VoucherNum = drv("CompanyId").ToString.Trim
                '' oVendorPayment.CreditCards.VoucherNum = sWhsCode & "-" & CDate(dIncomeDate).ToString("yyMMdd") & "-" & sPOSNumber
                oVendorPayment.CreditCards.Add()
            Next

            oVendorPayment.CashSum = 0

            Console.WriteLine("Attempting to Add ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
            lRetCode = oVendorPayment.Add()

            If lRetCode <> 0 Then
                sErrDesc = oDICompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

                Outgoing_Payment = RTN_ERROR
            Else

                Console.WriteLine("Completed with SUCCESS " & sDocEntry, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                Outgoing_Payment = RTN_SUCCESS

            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Outgoing_Payment = RTN_ERROR

        Finally
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oVendorPayment)
            oVendorPayment = Nothing
        End Try
    End Function

    Function AR_OutgoingPayment_Cancel(ByRef oDICompany As SAPbobsCOM.Company, _
                               ByRef sOutgoingPayment As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oOutgoingPayment As SAPbobsCOM.Payments
        oOutgoingPayment = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

        Try
            sFuncName = "AR_OutgoingPayment_Cancel"
            Console.WriteLine("Starting Function", sFuncName)

            If oOutgoingPayment.GetByKey(sOutgoingPayment) Then
                lRetCode = oOutgoingPayment.Cancel()

                If lRetCode <> 0 Then
                    sErrDesc = oDICompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    AR_OutgoingPayment_Cancel = RTN_ERROR
                Else
                    Console.WriteLine("Completed with SUCCESS " & oOutgoingPayment.DocNum, sFuncName)
                    sErrDesc = String.Empty
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                    AR_OutgoingPayment_Cancel = RTN_SUCCESS

                End If
            Else

                sErrDesc = "No matching records found in the Outgoing Payments " & sOutgoingPayment
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                AR_OutgoingPayment_Cancel = RTN_ERROR
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AR_OutgoingPayment_Cancel = RTN_ERROR

        Finally
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oOutgoingPayment)
            oOutgoingPayment = Nothing
        End Try
    End Function
End Module
