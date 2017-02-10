﻿Imports System.Data.SqlClient
Imports System.Configuration


Module modCommon

    Function ExecuteSQLQuery_DT(ByVal sConnectionString As String, ByVal sQuery As String) As DataTable

        Dim oDT_INTDBInformations As DataTable
        Dim sFuncName As String = String.Empty
        Dim oConnection As SqlConnection = Nothing
        Dim oSQLCommand As SqlCommand = Nothing
        Dim oSQLAdapter As SqlDataAdapter = New SqlDataAdapter

        Try
            sFuncName = "ExecuteSQLQuery_DT()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            oConnection = New SqlConnection(sConnectionString)

            If (oConnection.State = ConnectionState.Closed) Then
                oConnection.Open()
            End If

            oDT_INTDBInformations = New DataTable
            oSQLCommand = New SqlCommand(sQuery, oConnection)
            oSQLAdapter.SelectCommand = oSQLCommand
            oSQLCommand.CommandTimeout = 0
            oSQLAdapter.Fill(oDT_INTDBInformations)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Return oDT_INTDBInformations

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return Nothing
        Finally
            oSQLAdapter.Dispose()
            oSQLCommand.Dispose()
            oConnection.Close()
        End Try
    End Function

    Function ExecuteSQLQuery_DT(ByVal sConnectionString As String, ByVal sQuery As String, ByRef sErrDesc As String) As Long


        Dim oDT_INTDBInformations As DataTable
        Dim sFuncName As String = String.Empty
        Dim oConnection As SqlConnection = Nothing
        Dim oSQLCommand As SqlCommand = Nothing
        Dim oSQLAdapter As SqlDataAdapter = New SqlDataAdapter

        Try
            sFuncName = "ExecuteSQLQuery_DT()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            oConnection = New SqlConnection(sConnectionString)

            If (oConnection.State = ConnectionState.Closed) Then
                oConnection.Open()
            End If

            oDT_INTDBInformations = New DataTable
            oSQLCommand = New SqlCommand(sQuery, oConnection)
            oSQLAdapter.SelectCommand = oSQLCommand
            oSQLCommand.CommandTimeout = 0
            oSQLCommand.ExecuteNonQuery()
            ''Try
            ''    oSQLAdapter.Fill(oDT_INTDBInformations)
            ''Catch ex As Exception
            ''End Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Return RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return RTN_ERROR
        Finally
            oSQLAdapter.Dispose()
            oSQLCommand.Dispose()
            oConnection.Close()
        End Try
    End Function

    Function IntegrityValidation(ByVal oDT_Invoice As DataTable, ByVal oDT_Payments As DataTable, ByRef oDICompany As SAPbobsCOM.Company, _
                      ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sDocEntry As String = String.Empty
        Dim sTransID As String = String.Empty
        Dim sWhsCode As String = String.Empty
        Dim sPOSNumber As String = String.Empty
        Dim oDV_InvoiceInform As DataView = Nothing
        Dim oDV_PaymentsInform As DataView = Nothing
        Dim oDT_Distinct As DataTable = New DataTable
        Dim oDT_InvoiceStatus As DataTable = New DataTable
        Dim sProductCode As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim sErrDisplay As String = String.Empty
        Dim sManBatchItem As String = String.Empty
        Dim oBatchDT As DataTable = New DataTable
        Dim oARInvoice As SAPbobsCOM.Documents
        Dim sSQL As String = String.Empty


        Try
            sFuncName = "IntegrityValidation()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oDT_InvoiceStatus.Columns.Add("HID", GetType(String))
            oDT_InvoiceStatus.Columns.Add("LItem", GetType(String))
            oDT_InvoiceStatus.Columns.Add("Status", GetType(String))
            oDT_InvoiceStatus.Columns.Add("HErrorMsg", GetType(String))
            oDT_InvoiceStatus.Columns.Add("LErrorMsg", GetType(String))
            oDT_InvoiceStatus.Columns.Add("Time", GetType(String))
            oDT_InvoiceStatus.Columns.Add("Docentry", GetType(String))
            oDT_InvoiceStatus.Columns.Add("DocNum", GetType(String))
            oDT_InvoiceStatus.Columns.Add("POSTxType", GetType(String))

            oDV_InvoiceInform = New DataView(oDT_Invoice)
            oDV_PaymentsInform = New DataView(oDT_Payments)
            ' oDT_Distinct = oDV_InvoiceInform.Table.DefaultView.ToTable(True, "HTransID")
            oDT_Distinct = oDV_InvoiceInform.Table.DefaultView.ToTable(True, "HPOSTxNo", "HPOSTxType")
            For imjs As Integer = 0 To oDT_Distinct.Rows.Count - 1

                ''''''''''--------------------------------------
                '''''----------  Payment Code Validation
                ''''' -------------------------------------------

                Console.WriteLine("Calling Function AR_InvoiceCreation() POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function AR_InvoiceCreation() TransID " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)

                ' oDV_PaymentsInform.RowFilter = "HeaderID = '" & oDT_Distinct.Rows(imjs).Item("HTransID") & "'"
                oDV_PaymentsInform.RowFilter = "POSTxNo = '" & oDT_Distinct.Rows(imjs).Item("HPOSTxNo") & "'"
                If oDV_PaymentsInform.Count > 0 Then
                    If CInt(oDV_PaymentsInform.Item(0).Row("CreditCardCount").ToString.Trim) > 0 Then
                        For Each drv In oDV_PaymentsInform
                            If Not String.IsNullOrEmpty(drv("ErrMsg").ToString.Trim) Then
                                sErrDisplay = sErrDisplay & " " & drv("ErrMsg").ToString.Trim
                            End If
                        Next
                        oDT_InvoiceStatus.Rows.Add(oDV_PaymentsInform.Item(0).Row("POSTxNo").ToString.Trim, _
                                                                                                    "", "FAIL", _
                                                                         sErrDisplay, "", Now.ToShortTimeString, "", "", oDV_PaymentsInform.Item(0).Row("ID").ToString.Trim)
                        Console.WriteLine("Validation Fails POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Fails POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                    Else

                        ''''''''''--------------------------------------
                        '''''----------  Others Validation 
                        ''''' -------------------------------------------


                        oDV_InvoiceInform.RowFilter = "HPOSTxNo ='" & oDT_Distinct.Rows(imjs).Item("HPOSTxNo") & "' and Validation2Count = 0 and Validation3Count = 0 and HPOSTxType ='" & oDT_Distinct.Rows(imjs).Item("HPOSTxType") & "'"

                        If oDV_InvoiceInform.Count = 0 Then
                            oDV_InvoiceInform.RowFilter = "HPOSTxNo ='" & oDT_Distinct.Rows(imjs).Item("HPOSTxNo") & "'"
                            For Each drv As DataRowView In oDV_InvoiceInform
                                oDT_InvoiceStatus.Rows.Add(drv("HPOSTxNo").ToString.Trim, drv("DItemCode").ToString.Trim, "FAIL", _
                                                           "Validation Fails Pls find the line level error msg", drv("DetailsErrMsg").ToString.Trim, Now.ToShortTimeString, "", "", drv("HTransID").ToString.Trim)
                            Next
                            Console.WriteLine("Validation Fails POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Fails POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                        Else
                            ''''''''''--------------------------------------
                            '''''----------   Validation Succeed
                            ''''' -------------------------------------------

                            '' AR_InvoiceCreation 
                            ''  Console.WriteLine("Calling Function AR_InvoiceCreation() TransID " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                            Console.WriteLine("Validation SUCCESS POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation SUCCESS POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                            oDT_InvoiceStatus.Clear()
                            MarketingDocuments_Sync(oDV_InvoiceInform, oDV_PaymentsInform, p_oCompany, oDT_InvoiceStatus, sErrDesc)
                        End If
                    End If

                    If oDT_InvoiceStatus Is Nothing Then
                    Else
                        Dim sTrandID As String = String.Empty
                        Dim dSyncDatetime As DateTime

                        For imjd As Integer = 0 To oDT_InvoiceStatus.Rows.Count - 1

                            If sTrandID <> oDT_InvoiceStatus.Rows(imjd).Item("HID").ToString.Trim Then

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Date Time " & Now.Date & " " & oDT_InvoiceStatus.Rows(imjd).Item("Time").ToString.Trim, sFuncName)

                                sSQL += "UPDATE [AB_SalesTransHeader]" & _
    "SET [Status] = '" & oDT_InvoiceStatus.Rows(imjd).Item("Status").ToString.Trim & "' ,[ErrorMsg] = '" & oDT_InvoiceStatus.Rows(imjd).Item("HErrorMsg").ToString.Trim & "' , " & _
    "[SAPSyncDate] =  DATEADD(day,datediff(day,0,GETDATE()),0) ,[SAPSyncDateTime] = GETDATE() " & _
    "WHERE [ID] = '" & oDT_InvoiceStatus.Rows(imjd).Item("POSTxType").ToString.Trim & "' "

                                sSQL += "UPDATE [AB_SalesTransDetail] SET [ErrMsg] = '' " & _
  " WHERE [POSTxNo] = '" & oDT_InvoiceStatus.Rows(imjd).Item("HID").ToString.Trim & "'"

                                sTrandID = oDT_InvoiceStatus.Rows(imjd).Item("HID").ToString.Trim

                            End If
                            If Not String.IsNullOrEmpty(oDT_InvoiceStatus.Rows(imjd).Item("LErrorMsg").ToString.Trim) Then
                                sSQL += "UPDATE [AB_SalesTransDetail] SET [ErrMsg] = '" & oDT_InvoiceStatus.Rows(imjd).Item("LErrorMsg").ToString.Trim & "' " & _
        " WHERE [POSTxNo] = '" & oDT_InvoiceStatus.Rows(imjd).Item("HID").ToString.Trim & "' and [ItemCode] = '" & oDT_InvoiceStatus.Rows(imjd).Item("LItem").ToString.Trim & "' "

                            End If
                        Next imjd
                        oDT_InvoiceStatus.Clear()
                        sTrandID = String.Empty
                    End If

                    If sSQL.Length > 1 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Update SQL " & sSQL, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ExecuteSQLQuery_DT() ", sFuncName)
                        Console.WriteLine("Calling Function ExecuteSQLQuery_DT()")
                        If ExecuteSQLQuery_DT(P_sConString, sSQL, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                   
                Else

                    oDV_InvoiceInform.RowFilter = "HPOSTxNo ='" & oDT_Distinct.Rows(imjs).Item("HPOSTxNo") & "' and Validation2Count = 0  and Validation3Count = 0 and HPOSTxType ='" & oDT_Distinct.Rows(imjs).Item("HPOSTxType") & "'"

                    If oDV_InvoiceInform.Count = 0 Then
                        oDV_InvoiceInform.RowFilter = "HPOSTxNo ='" & oDT_Distinct.Rows(imjs).Item("HPOSTxNo") & "'"
                        For Each drv As DataRowView In oDV_InvoiceInform
                            oDT_InvoiceStatus.Rows.Add(drv("HPOSTxNo").ToString.Trim, drv("DItemCode").ToString.Trim, "FAIL", _
                                                       "Validation Fails Pls find the line level error msg", drv("DetailsErrMsg").ToString.Trim, Now.ToShortTimeString, "", "", drv("HTransID").ToString.Trim)
                        Next

                        Console.WriteLine("Validation Fails POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Fails POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)

                        If oDT_InvoiceStatus Is Nothing Then
                        Else
                            Dim sTrandID As String = String.Empty
                            Dim dSyncDatetime As DateTime

                            For imjd As Integer = 0 To oDT_InvoiceStatus.Rows.Count - 1

                                If sTrandID <> oDT_InvoiceStatus.Rows(imjd).Item("HID").ToString.Trim Then

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Date Time " & Now.Date & " " & oDT_InvoiceStatus.Rows(imjd).Item("Time").ToString.Trim, sFuncName)

                                    sSQL += "UPDATE [AB_SalesTransHeader]" & _
        "SET [Status] = '" & oDT_InvoiceStatus.Rows(imjd).Item("Status").ToString.Trim & "' ,[ErrorMsg] = '" & oDT_InvoiceStatus.Rows(imjd).Item("HErrorMsg").ToString.Trim & "' , " & _
        "[SAPSyncDate] =  DATEADD(day,datediff(day,0,GETDATE()),0) ,[SAPSyncDateTime] = GETDATE() " & _
        "WHERE [ID] = '" & oDT_InvoiceStatus.Rows(imjd).Item("POSTxType").ToString.Trim & "' "

                                    sSQL += "UPDATE [AB_SalesTransDetail] SET [ErrMsg] = '' " & _
      " WHERE [POSTxNo] = '" & oDT_InvoiceStatus.Rows(imjd).Item("HID").ToString.Trim & "'"

                                    sTrandID = oDT_InvoiceStatus.Rows(imjd).Item("HID").ToString.Trim

                                End If
                                If Not String.IsNullOrEmpty(oDT_InvoiceStatus.Rows(imjd).Item("LErrorMsg").ToString.Trim) Then
                                    sSQL += "UPDATE [AB_SalesTransDetail] SET [ErrMsg] = '" & oDT_InvoiceStatus.Rows(imjd).Item("LErrorMsg").ToString.Trim & "' " & _
            " WHERE [POSTxNo] = '" & oDT_InvoiceStatus.Rows(imjd).Item("HID").ToString.Trim & "' and [ItemCode] = '" & oDT_InvoiceStatus.Rows(imjd).Item("LItem").ToString.Trim & "' "

                                End If
                            Next imjd
                            oDT_InvoiceStatus.Clear()
                            sTrandID = String.Empty
                        End If

                        If sSQL.Length > 1 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Update SQL " & sSQL, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ExecuteSQLQuery_DT() ", sFuncName)
                            Console.WriteLine("Calling Function ExecuteSQLQuery_DT()")
                            If ExecuteSQLQuery_DT(P_sConString, sSQL, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                    Else
                        ''''''''''--------------------------------------
                        '''''----------   Validation Succeed
                        ''''' -------------------------------------------

                        '' AR_InvoiceCreation 
                        '' Console.WriteLine("Calling Function AR_InvoiceCreation() TransID " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                        Console.WriteLine("Validation SUCCESS POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation SUCCESS POSTxNo " & oDT_Distinct.Rows(imjs).Item("HPOSTxNo"), sFuncName)
                        oDT_InvoiceStatus.Clear()
                        MarketingDocuments_Sync(oDV_InvoiceInform, oDV_PaymentsInform, p_oCompany, oDT_InvoiceStatus, sErrDesc)
                    End If

                End If
            Next imjs

            Console.WriteLine("Completed with SUCCESS", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)

           
            Return RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Update_Status() Function" & sPOSNumber, sFuncName)
            ''  Update_Status(sTransID, sErrDesc, "FAIL", "", "SalesTransHDR")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the Transaction. POS Number : " & sPOSNumber, sFuncName)
            If oDICompany.InTransaction = True Then oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting the Company and Release the Object ", sFuncName)
            p_oCompany.Disconnect()
            oDICompany.Disconnect()
            oDICompany = Nothing
            p_oCompany = Nothing
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return RTN_ERROR

        End Try
    End Function


    Public Function MarketingDocuments_Sync(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, _
                                            ByVal oDTStatus As DataTable, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim dIncomeDate As Date
        Dim sPostxdatetime As String
        Dim dPostxdatetime As Date
        Dim tDocTime As DateTime
        Dim sWhsCode As String = String.Empty
        Dim sPOSNumber As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim sQueryup As String = String.Empty
        Dim sDocEntry As String = String.Empty
        Dim sDocNum As String = String.Empty
        Dim lRetCode As Integer
        Dim irow As Integer = 0
        Dim dDocTotal As Double = 0.0
        Dim sARInvoice As String = String.Empty
        oDTStatus.Clear()
        Dim sSql As String = String.Empty
        Dim sARInvoiceNo As String = String.Empty
        Dim sIncomingpaymentno As String = String.Empty
        Dim sARCreditnote As String = String.Empty
        Dim sOutgoingpayment As String = String.Empty

        Dim sCardCode As String = String.Empty
        Dim oDT_Payamount As DataTable = New DataTable
        Dim dPayamount As Double = 0
        If oDVPayment.Count > 0 Then
            oDT_Payamount = oDVPayment.ToTable
        End If


        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRset_Batch As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            sFuncName = "MarketingDocuments_Sync()"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)
            sWhsCode = CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)

            If Not String.IsNullOrEmpty(oDVARInvoice.Item(0).Row("HCardCode").ToString.Trim) Then ''CD0001              
                If oDVARInvoice.Item(0).Row("HCardCode").ToString.Trim.ToUpper() = "CASH" Then
                    sCardCode = p_oCompDef.p_sCardCode
                Else
                    sCardCode = oDVARInvoice.Item(0).Row("HCardCode").ToString.Trim
                End If
            Else
                sCardCode = p_oCompDef.p_sCardCode
            End If


            If oDVARInvoice.Item(0).Row("HPOSTxType").ToString.Trim = "S" Then '' AR Invoice & Incoming payments

                '************************************ AR Invoice Started ************************************************************************************

                If oDVARInvoice Is Nothing Then
                    Console.WriteLine("No matching records found in Sales Header Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Sales Header Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                Else
                    If oDVARInvoice.Count > 0 Then
                        Console.WriteLine("Calling Funcion AR_InvoiceCreation() " & sDocEntry, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_InvoiceCreation() : AR Invoice DocEntry " & sDocEntry, sFuncName)
                        If AR_InvoiceCreation(oDVARInvoice, oDVPayment, oCompany, sDocEntry, sDocNum, sCardCode, sErrDesc) <> RTN_SUCCESS Then

                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Console.WriteLine("Completed with ERROR", sFuncName)
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            MarketingDocuments_Sync = Nothing
                            GoTo ERRORDISPLAY

                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Payement Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                        Console.WriteLine("No matching records found in Payement Table " & sDocEntry, sFuncName)
                    End If
                End If



                '************************************ Incoming Payment Started ************************************************************************************
                If oDVPayment Is Nothing Then
                    Console.WriteLine("No matching records found in Payement Table " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Payment Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                    oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, sDocNum, oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                Else
                    If oDVPayment.Count > 0 Then
                        Console.WriteLine("Calling Funcion AR_IncomingPayment() " & sDocEntry, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_IncomingPayment() : AR Invoice DocEntry " & sDocEntry, sFuncName)
                        If AR_IncomingPayment(oDVPayment, oCompany, sDocEntry, dIncomeDate, sPOSNumber _
                                                                 , sWhsCode, sCardCode, sErrDesc) <> RTN_SUCCESS Then

                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Console.WriteLine("Completed with ERROR", sFuncName)
                            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            MarketingDocuments_Sync = Nothing
                            ''  Return RTN_ERROR
                        Else

                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, sDocNum, oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                            ''  Return RTN_ERROR
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Payement Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                        Console.WriteLine("No matching records found in Payement Table " & sDocEntry, sFuncName)
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, sDocNum, oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                    End If
                End If


            ElseIf oDVARInvoice.Item(0).Row("HPOSTxType").ToString.Trim = "V" Then '' Cancel Incomingpayments & create the AR Invoice cancellation


                '************************************ Identifing whether its a Void of sales / Void of return ************************************************************************************
                
                dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
                dPostxdatetime = oDVARInvoice.Item(0).Row("HPOSTxDatetime").ToString.Trim
                sPostxdatetime = oDVARInvoice.Item(0).Row("HPOSTxDatetime").ToString.Trim

                sSql = "select oinv.DocEntry [InvoiceNo] , rct2.DocNum [IncomingNo]  from oinv join RCT2 on oinv.DocEntry = rct2.DocEntry  where oinv.U_AB_POSTxNo = '" & sPOSNumber & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Identifing Incoming payments DocNum " & sSql, sFuncName)
                oRset.DoQuery(sSql)
                If oRset.Fields.Item("InvoiceNo").Value <> "0" Then
                    sARInvoiceNo = oRset.Fields.Item("InvoiceNo").Value
                End If

                If oRset.Fields.Item("IncomingNo").Value <> "0" Then
                    sIncomingpaymentno = oRset.Fields.Item("IncomingNo").Value
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AR Invoice No. " & sARInvoiceNo, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Incoming Payment No. " & sIncomingpaymentno, sFuncName)

                sSql = "select ORIN.DocEntry [Creditnote] , VPM2.DocNum [Outgoingno]  from ORIN join VPM2 on ORIN.DocEntry = VPM2.DocEntry  where ORIN.U_AB_POSTxNo = '" & sPOSNumber & "' and VPM2.InvType = '14'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Identifing Incoming payments DocNum " & sSql, sFuncName)
                oRset.DoQuery(sSql)
                If oRset.Fields.Item("Creditnote").Value <> "0" Then
                    sARCreditnote = oRset.Fields.Item("Creditnote").Value
                End If
                If oRset.Fields.Item("Outgoingno").Value <> "0" Then
                    sOutgoingpayment = oRset.Fields.Item("Outgoingno").Value
                End If
               
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AR Credit Note " & sARCreditnote, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Outgoing Payment No. " & sOutgoingpayment, sFuncName)

                '************************************ Void of Sales / Void of Exchange ************************************************************************************
                If Not String.IsNullOrEmpty(sARInvoiceNo) And Not String.IsNullOrEmpty(sIncomingpaymentno) Then

                    '************************************ Incoming Payment Cancellation Started ************************************************************************************
                    Console.WriteLine("Calling Funcion AR_IncomingPayment() " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_IncomingPayment() : AR Invoice DocEntry " & sDocEntry, sFuncName)

                    If oCompany.InTransaction = False Then oCompany.StartTransaction()
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SAP Transaction started successfully " & sDocEntry, sFuncName)
                    If AR_IncomingPayment_Cancel(oCompany, sIncomingpaymentno, sErrDesc) <> RTN_SUCCESS Then

                        If Left(sErrDesc, 19) = "No matching records" Then
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "Skip", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        Else
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        End If
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine("Completed with ERROR", sFuncName)
                        Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                        If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        MarketingDocuments_Sync = Nothing
                        GoTo ERRORDISPLAY

                        ''  Return RTN_ERROR
                    End If

                    '************************************ AR Invoice Cancel ************************************************************************************
                    Console.WriteLine("Calling Funcion AR_Invoice_Cancel() " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_Invoice_Cancel() : AR Invoice DocEntry " & sDocEntry, sFuncName)

                    ' If AR_CreditMemo(oCompany, sARInvoiceNo, dIncomeDate, sErrDesc) <> RTN_SUCCESS Then
                    If AR_Invoice_Cancel(oCompany, sARInvoiceNo, dIncomeDate, sPostxdatetime, sErrDesc) <> RTN_SUCCESS Then

                        If sErrDesc = "" Then
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "Skip", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        Else
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        End If
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine("Completed with ERROR", sFuncName)
                        Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                        If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        MarketingDocuments_Sync = Nothing
                        ''Return RTN_ERROR
                    Else
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sARInvoice, "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        ''  Return RTN_ERROR
                    End If
                    '************************************ Void of Refund ************************************************************************************
                ElseIf Not String.IsNullOrEmpty(sARCreditnote) And Not String.IsNullOrEmpty(sOutgoingpayment) Then

                    '************************************ Outgoing Payment Cancellation Started ************************************************************************************
                    Console.WriteLine("Calling Funcion AR_OutgoingPayment_Cancel() " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_OutgoingPayment_Cancel() : Outgoing Payment DocEntry " & sOutgoingpayment, sFuncName)

                    If oCompany.InTransaction = False Then oCompany.StartTransaction()
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SAP Transaction started successfully " & sDocEntry, sFuncName)
                    If AR_OutgoingPayment_Cancel(oCompany, sOutgoingpayment, sErrDesc) <> RTN_SUCCESS Then

                        If Left(sErrDesc, 19) = "No matching records" Then
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "Skip", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        Else
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        End If
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine("Completed with ERROR", sFuncName)
                        Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                        If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        MarketingDocuments_Sync = Nothing
                        GoTo ERRORDISPLAY

                        ''  Return RTN_ERROR
                    End If

                    '************************************ AR Credit Memo Cancellation ************************************************************************************
                    Console.WriteLine("Calling Funcion AR_Creditnote_Cancel() " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_Creditnote_Cancel() : AR Credit note DocEntry " & sARCreditnote, sFuncName)

                    If AR_Creditnote_Cancel(oCompany, sARCreditnote, dIncomeDate, sPostxdatetime, sErrDesc) <> RTN_SUCCESS Then

                        If sErrDesc = "" Then
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "Skip", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        Else
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        End If
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine("Completed with ERROR", sFuncName)
                        Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                        If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        MarketingDocuments_Sync = Nothing
                        ''Return RTN_ERROR
                    Else
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sARInvoice, "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                        ''  Return RTN_ERROR
                    End If

                Else
                    oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "Skip", "No matching records found", "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)

                End If


            ElseIf oDVARInvoice.Item(0).Row("HPOSTxType").ToString.Trim = "R" Then

                '************************************ AR Credit Memo Started ************************************************************************************

                If oDVARInvoice Is Nothing Then
                    Console.WriteLine("No matching records found in Sales Header Table : AR Credit Memo DocEntry " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Sales Header Table : AR Credit DocEntry " & sDocEntry, sFuncName)
                Else
                    If oDVARInvoice.Count > 0 Then
                        Console.WriteLine("Calling Funcion AR_CreditMemo_Standalone() " & sDocEntry, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_CreditMemo_Standalone() ", sFuncName)
                        If AR_CreditMemo_Standalone(oDVARInvoice, oDVPayment, oCompany, sDocEntry, sDocNum, sCardCode, sErrDesc) <> RTN_SUCCESS Then

                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Console.WriteLine("Completed with ERROR", sFuncName)
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "", oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            MarketingDocuments_Sync = Nothing
                            GoTo ERRORDISPLAY

                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Payement Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                        Console.WriteLine("No matching records found in Payement Table " & sDocEntry, sFuncName)
                    End If
                End If



                '************************************ Outgoing Payment Started ************************************************************************************
                If oDVPayment Is Nothing Then
                    Console.WriteLine("No matching records found in Payment Table " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Payment Table" & sDocEntry, sFuncName)
                    oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, sDocNum, oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                Else
                    If oDVPayment.Count > 0 Then
                        Console.WriteLine("Calling Funcion Outgoing_Payment() " & sDocEntry, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion Outgoing_Payment() : AR Credit Memo DocEntry " & sDocEntry, sFuncName)
                        If Outgoing_Payment(oDVPayment, oCompany, sDocEntry, dIncomeDate, sPOSNumber _
                                                                 , sWhsCode, sCardCode, sErrDesc) <> RTN_SUCCESS Then

                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Console.WriteLine("Completed with ERROR", sFuncName)
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            MarketingDocuments_Sync = Nothing
                            ''  Return RTN_ERROR
                        Else
                            oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, sDocNum, oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                            ''  Return RTN_ERROR
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Payement Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                        Console.WriteLine("No matching records found in Payement Table " & sDocEntry, sFuncName)
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, sDocNum, oDVARInvoice.Item(0).Row("HTransID").ToString.Trim)
                    End If
                End If

            End If
            sErrDesc = ""
            ''  oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString)

ERRORDISPLAY: If oDTStatus Is Nothing Then
            Else
                Dim sTrandID As String = String.Empty
                Dim dSyncDatetime As DateTime
                For imjs As Integer = 0 To oDTStatus.Rows.Count - 1

                    If sTrandID <> oDTStatus.Rows(imjs).Item("HID").ToString.Trim Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Date Time " & Now.Date & " " & oDTStatus.Rows(imjs).Item("Time").ToString.Trim, sFuncName)
                        dSyncDatetime = Now.Date & " " & oDTStatus.Rows(imjs).Item("Time").ToString.Trim
                        sQueryup += "UPDATE " & p_oCompDef.p_sIntDBName & ".. [AB_SalesTransHeader]" & _
"SET [Status] = '" & oDTStatus.Rows(imjs).Item("Status").ToString.Trim & "' ,[ErrorMsg] = '" & Replace(oDTStatus.Rows(imjs).Item("HErrorMsg").ToString.Trim, "'", "''") & "' , " & _
"[SAPSyncDate] =  DATEADD(day,datediff(day,0,GETDATE()),0) ,[SAPSyncDateTime] = GETDATE(), [ARDocEntry] = '" & oDTStatus.Rows(imjs).Item("DocEntry").ToString.Trim & "' " & _
"WHERE [ID] = '" & oDTStatus.Rows(imjs).Item("POSTxType").ToString.Trim & "'"
                        sTrandID = oDTStatus.Rows(imjs).Item("HID").ToString.Trim

                        sQueryup += "UPDATE " & p_oCompDef.p_sIntDBName & ".. [AB_SalesTransDetail] SET [ErrMsg] = '' " & _
" WHERE [POSTxNo] = '" & oDTStatus.Rows(imjs).Item("HID").ToString.Trim & "'"
                    End If

                    If Not String.IsNullOrEmpty(oDTStatus.Rows(imjs).Item("LErrorMsg").ToString.Trim) Then
                        sQueryup += "UPDATE " & p_oCompDef.p_sIntDBName & ".. [AB_SalesTransDetail] SET [ErrMsg] = '" & Replace(oDTStatus.Rows(imjs).Item("LErrorMsg").ToString.Trim, "'", "''") & "' " & _
    " WHERE [POSTxNo] = '" & oDTStatus.Rows(imjs).Item("HID").ToString.Trim & "' and [ItemCode] = '" & oDTStatus.Rows(imjs).Item("LItem").ToString.Trim & "' "
                    End If

                Next imjs
                oDTStatus.Clear()
                sTrandID = String.Empty

            End If

            If sQueryup.Length > 1 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Update SQL " & sQueryup, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing Query", sFuncName)
                oRset.DoQuery(sQueryup)
            End If

            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Console.WriteLine("Committed the Transaction for TransID " & oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Committed the Transaction Reference POSNumber : " & sPOSNumber, sFuncName)
            ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
            '' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting the Company and Release the Object ", sFuncName)

            Return RTN_SUCCESS


        Catch ex As Exception
            sErrDesc = ex.Message
            Console.WriteLine("Completed with ERROR", sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Return RTN_ERROR
        End Try
    End Function


    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing information about the initialing variables
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sQuery As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)


            oCompDef.p_sServerName = String.Empty
            oCompDef.p_sLicServerName = String.Empty
            oCompDef.p_sDBUserName = String.Empty
            oCompDef.p_sDBPassword = String.Empty

            oCompDef.p_sDataBaseName = String.Empty
            oCompDef.p_sSAPUserName = String.Empty
            oCompDef.p_sSAPPassword = String.Empty

            oCompDef.p_sLogDir = String.Empty
            oCompDef.p_sDebug = String.Empty
            oCompDef.p_sIntDBName = String.Empty
            oCompDef.p_sCardCode = String.Empty
            oCompDef.p_sX = String.Empty
            oCompDef.p_sGST = String.Empty
            oCompDef.p_sGLAccount = String.Empty

          
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.p_sServerName = ConfigurationManager.AppSettings("Server")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenseServer")) Then
                oCompDef.p_sLicServerName = ConfigurationManager.AppSettings("LicenseServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.p_sDataBaseName = ConfigurationManager.AppSettings("SAPDBName")
                ' AE_STUTTGART_DLL.P_sSAPDBName = oCompDef.p_sDataBaseName
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.p_sSAPUserName = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.p_sSAPPassword = ConfigurationManager.AppSettings("SAPPassword")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.p_sDBUserName = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.p_sDBPassword = ConfigurationManager.AppSettings("DBPwd")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SQLType")) Then
                oCompDef.p_sSQLType = ConfigurationManager.AppSettings("SQLType")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CardCode")) Then
                oCompDef.p_sCardCode = ConfigurationManager.AppSettings("CardCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("IntegrationDBName")) Then
                oCompDef.p_sIntDBName = ConfigurationManager.AppSettings("IntegrationDBName")
                ' AE_STUTTGART_DLL.P_sStagingDBName = oCompDef.p_sIntDBName
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("NumOfDays")) Then
                oCompDef.p_sX = ConfigurationManager.AppSettings("NumOfDays")
                ' AE_STUTTGART_DLL.P_sStagingDBName = oCompDef.p_sIntDBName
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GST")) Then
                oCompDef.p_sGST = ConfigurationManager.AppSettings("GST")
                ' AE_STUTTGART_DLL.P_sStagingDBName = oCompDef.p_sIntDBName
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GLAccount")) Then
                oCompDef.p_sGLAccount = ConfigurationManager.AppSettings("GLAccount")
                ' AE_STUTTGART_DLL.P_sStagingDBName = oCompDef.p_sIntDBName
            End If


            ' folder
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogDir")) Then
                oCompDef.p_sLogDir = ConfigurationManager.AppSettings("LogDir")
                'AE_STUTTGART_DLL.sLogFolderPath = oCompDef.p_sLogDir
            Else
                oCompDef.p_sLogDir = System.IO.Directory.GetCurrentDirectory()
                ' AE_STUTTGART_DLL.sLogFolderPath = oCompDef.p_sLogDir
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.p_sDebug = ConfigurationManager.AppSettings("Debug")
                If p_oCompDef.p_sDebug.ToUpper = "ON" Then
                    p_iDebugMode = 1
                Else
                    p_iDebugMode = 0
                End If
            Else
                p_iDebugMode = 0
            End If

            P_sConString = String.Empty
            P_sConString = "Data Source=" & p_oCompDef.p_sServerName & ";Initial Catalog=" & p_oCompDef.p_sIntDBName & ";User ID=" & p_oCompDef.p_sDBUserName & "; Password=" & p_oCompDef.p_sDBPassword

            sQuery = "[AE_SP001_GetINTDBInformation]'[" & p_oCompDef.p_sDataBaseName & "]'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching INT DB Query : " & sQuery, sFuncName)

            'Getting the Data from Invoice Table as DataSet
            Console.WriteLine("Calling ExecuteSQLQuery_DT() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            oDT_InvoiceData = ExecuteSQLQuery_DT(P_sConString, sQuery)


            sQuery = "select T0.* , T1.AcctCode , T1.CreditCard , T1.CompanyId,(select top 1 RCT3.CrCardNum  from [" & p_oCompDef.p_sDataBaseName & "].. RCT3 where RCT3.CreditCard = T1.CreditCard ) [CreditNumber]  " & _
"into #Payment from [AB_Payment] T0 left outer join [" & p_oCompDef.p_sDataBaseName & "].. OCRC T1 on T1.CardName = T0.PaymentCode " & _
"select #Payment.POSTxNo  , COUNT(isnull(#Payment.CreditCard,0 )) [CreditCardCount] into #Paycount from #Payment where isnull(#Payment.CreditCard,'') = '' group by #Payment.POSTxNo " & _
"select T0.*, isnull(T1.CreditCardCount,'') [CreditCardCount] " & _
 "into #PaymentFinal from #Payment T0 left outer join #Paycount T1 on T0.POSTxNo  = T1.POSTxNo " & _
 "select T4.*, case when isnull(T4.CreditCard,'') = '' then 'Payment Code {' + T4.PaymentCode  + '} does not exist in Credit Cards Setup' else '' end [ErrMsg]  from #PaymentFinal T4 " & _
"drop table #Payment " & _
"drop table #Paycount " & _
"drop table #PaymentFinal"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Payment Query : " & sQuery, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            oDT_PaymentData = ExecuteSQLQuery_DT(P_sConString, sQuery)

            sQuery = "SELECT T0.[WhsCode], T0.[WhsName], T0.[U_AB_POSLocCode] FROM [" & p_oCompDef.p_sDataBaseName & "].. OWHS T0"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Warehouse : " & sQuery, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            oDT_Warehouse = ExecuteSQLQuery_DT(P_sConString, sQuery)

            ' AE_STUTTGART_DLL.p_iDebugMode = p_iDebugMode

            'IntegrationDBName

            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

    Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
                                          ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ConnectToTargetCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2013 21
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet

        Try
            sFuncName = "ConnectToTargetCompany()"
            Console.WriteLine("Starting function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            Console.WriteLine("Initializing the Company Object", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)

            oCompany = New SAPbobsCOM.Company
            Console.WriteLine("Assigning the representing database name", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)

            oCompany.Server = p_oCompDef.p_sServerName
            oCompany.LicenseServer = p_oCompDef.p_sLicServerName
            oCompany.DbUserName = p_oCompDef.p_sDBUserName
            oCompany.DbPassword = p_oCompDef.p_sDBPassword
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            oCompany.UseTrusted = False

            If p_oCompDef.p_sSQLType = 2012 Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            ElseIf p_oCompDef.p_sSQLType = 2008 Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            End If

            oCompany.CompanyDB = p_oCompDef.p_sDataBaseName
            oCompany.UserName = p_oCompDef.p_sSAPUserName
            oCompany.Password = p_oCompDef.p_sSAPPassword

            Console.WriteLine("Connecting to the Company Database.", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)

            iRetValue = oCompany.Connect()

            If iRetValue <> 0 Then
                oCompany.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                    oCompany.CompanyDB, System.Environment.NewLine, _
                                vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)
            End If
            Console.WriteLine("Completed with SUCCESS " & p_oCompDef.p_sDataBaseName, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectToTargetCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToTargetCompany = RTN_ERROR
        End Try
    End Function

    Public Function GetSingleValue(ByVal Query As String, ByRef p_oDICompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As String

        ' ***********************************************************************************
        '   Function   :    GetSingleValue()
        '   Purpose    :    This function is handles - Return single value based on Query
        '   Parameters :    ByVal Query As String
        '                       sDate = Passing Query 
        '                   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany = Passing the Company which has been connected
        '                   ByRef sErrDesc As String
        '                       sErrDesc=Error Description to be returned to calling function
        '   Author     :    SRINIVASAN
        '   Date       :    15/08/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "GetSingleValue()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & Query, sFuncName)

            Dim objRS As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(Query)
            If objRS.RecordCount > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                GetSingleValue = RTN_SUCCESS

                Return objRS.Fields.Item(0).Value.ToString
            End If
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(ex.Message, sFuncName)
            GetSingleValue = RTN_SUCCESS
            Return ""
        End Try
        Return Nothing
    End Function

    Public Function ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset, ByRef sErrDesc As String) As DataTable

        '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
        '\ easily used ADO.NET datatable which can be used for data binding much easier.
        Dim sFuncName As String = String.Empty
        Dim dtTable As New DataTable
        Dim NewCol As DataColumn
        Dim NewRow As DataRow
        Dim ColCount As Integer

        Try
            sFuncName = "ConvertRecordset()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            For ColCount = 0 To SAPRecordset.Fields.Count - 1
                NewCol = New DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                dtTable.Columns.Add(NewCol)
            Next

            Do Until SAPRecordset.EoF

                NewRow = dtTable.NewRow
                'populate each column in the row we're creating
                For ColCount = 0 To SAPRecordset.Fields.Count - 1

                    NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = SAPRecordset.Fields.Item(ColCount).Value

                Next

                'Add the row to the datatable
                dtTable.Rows.Add(NewRow)


                SAPRecordset.MoveNext()
            Loop

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Return dtTable

        Catch ex As Exception

            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return Nothing

        End Try

    End Function

End Module
