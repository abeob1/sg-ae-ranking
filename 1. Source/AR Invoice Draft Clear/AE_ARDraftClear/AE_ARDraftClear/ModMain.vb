Module ModMain

    Public oDT_BOM As DataTable = Nothing
    Public oDT_InvoiceData As DataTable = Nothing
    Public oDT_PaymentData As DataTable = Nothing

    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        ' Dim oARInvoice As AE_STUTTGART_DLL.clsARInvoice = New AE_STUTTGART_DLL.clsARInvoice
        Dim orset As SAPbobsCOM.Recordset = Nothing
        Dim sQuery As String = String.Empty
        Dim oARDraft As SAPbobsCOM.Documents = Nothing
        Dim lRetCode As Integer = 0
        Try
            p_iDebugMode = DEBUG_ON
            sFuncName = "Main()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            'Getting the Parameter Values from App Cofig File
            Console.WriteLine("Calling GetSystemIntializeInfo() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If p_oCompany Is Nothing Then
                Console.WriteLine("Calling ConnectToTargetCompany() ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                If ConnectToTargetCompany(p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            orset = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("select DocEntry from odrf where ObjType = 13 and UserSign = 5 ")
            Dim oDTDraftkey As DataTable = Nothing
            oDTDraftkey = New DataTable()
            oDTDraftkey = ConvertRecordset(orset, sErrDesc)
            oARDraft = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            oARDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
            Console.WriteLine("Starting to Remove AR Invoice Draft ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting to Remove AR Invoice Draft ", sFuncName)
            For Each oDR As DataRow In oDTDraftkey.Rows
                If oARDraft.GetByKey(oDR(0)) Then
                    Console.WriteLine("Attempting to Remove AR Invoice Draft " & oDR(0), sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Remove AR Invoice Draft " & oDR(0), sFuncName)
                    lRetCode = oARDraft.Remove()
                    If lRetCode <> 0 Then
                        sErrDesc = p_oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine("Completed with ERROR " & oDR(0), sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & oDR(0), sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                    Else
                        Console.WriteLine("Completed with SUCCESS " & oDR(0), sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & oDR(0), sFuncName)
                    End If

                End If

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting to Remove AR Invoice Draft Completed Successfully", sFuncName)



            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR : " & sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

        End Try

    End Sub

End Module
