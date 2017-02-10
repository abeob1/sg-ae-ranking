Module ModMain

    Public oDT_BOM As DataTable = Nothing
    Public oDT_InvoiceData As DataTable = Nothing
    Public oDT_PaymentData As DataTable = Nothing

    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        ' Dim oARInvoice As AE_STUTTGART_DLL.clsARInvoice = New AE_STUTTGART_DLL.clsARInvoice

        Dim sQuery As String = String.Empty

        Try
            p_iDebugMode = DEBUG_ON
            sFuncName = "Main()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            'Getting the Parameter Values from App Cofig File
            Console.WriteLine("Calling GetSystemIntializeInfo() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            ''If Not oDT_InvoiceData Is Nothing Then
            ''    '' Function to connect the Company
            ''    If p_oCompany Is Nothing Then
            ''        Console.WriteLine("Calling ConnectToTargetCompany() ", sFuncName)
            ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            ''        If ConnectToTargetCompany(p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ''    End If

            ''    Console.WriteLine("Calling IntegrityValidation() ", sFuncName)
            ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IntegrityValidation()", sFuncName)
            ''    If IntegrityValidation(oDT_InvoiceData, oDT_PaymentData, p_oCompany, sErrDesc) <> RTN_SUCCESS Then
            ''        Call WriteToLogFile(sErrDesc, sFuncName)
            ''        Console.WriteLine("Completed with ERROR : " & sErrDesc, sFuncName)
            ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ''    End If
            ''Else

            ''    Console.WriteLine("There is No Pending Records Found in Integration DB", sFuncName)
            ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("There is No Pending Records Found in Integration DB", sFuncName)
            ''End If

            Console.WriteLine("Attempting SAP Master Data Sync to Integration DB", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting SAP Master Data Sync to Integration DB", sFuncName)

            Console.WriteLine("Executing Item Master Sync", sFuncName)
            sQuery = "[AE_SP003_ItemMasterSync]'[" & p_oCompDef.p_sDataBaseName & "]','" & p_oCompDef.p_sX & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Master Sync Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Price List Sync", sFuncName)
            sQuery = "[AE_SP004_PriceListSync]'[" & p_oCompDef.p_sDataBaseName & "]','" & p_oCompDef.p_sX & "','" & p_oCompDef.p_sGST & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Price List Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Promotion Price List Sync", sFuncName)
            sQuery = "[AE_SP005_PromotionPriceListSync]'[" & p_oCompDef.p_sDataBaseName & "]','" & p_oCompDef.p_sX & "','" & p_oCompDef.p_sGST & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Promotion Price List Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Warehouse Sync", sFuncName)
            sQuery = "[AE_SP006_WareHouseSync]'[" & p_oCompDef.p_sDataBaseName & "]','" & p_oCompDef.p_sX & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Warehouse Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Customer Sync", sFuncName)
            sQuery = "[AE_SP007_CustomerSync]'[" & p_oCompDef.p_sDataBaseName & "]','" & p_oCompDef.p_sX & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Customer Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing CustomerGroup Sync", sFuncName)
            sQuery = "[AE_SP008_CustomerGroupSync]'[" & p_oCompDef.p_sDataBaseName & "]','" & p_oCompDef.p_sX & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CustomerGroup Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Brand Sync", sFuncName)
            sQuery = "[AE_SP009_BrandSync]'[" & p_oCompDef.p_sDataBaseName & "]'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Brand Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Category Sync", sFuncName)
            sQuery = "[AE_SP010_CategorySync]'[" & p_oCompDef.p_sDataBaseName & "]'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Category Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Department Sync", sFuncName)
            sQuery = "[AE_SP011_DepartmentSync]'[" & p_oCompDef.p_sDataBaseName & "]'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Department Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Vendors Sync", sFuncName)
            sQuery = "[AE_SP012_VendorsSync]'[" & p_oCompDef.p_sDataBaseName & "]','" & p_oCompDef.p_sX & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Vendors Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Tender Sync", sFuncName)
            sQuery = "[AE_SP013_TenderSync]'[" & p_oCompDef.p_sDataBaseName & "]','" & p_oCompDef.p_sX & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tender Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing Sales Employee Sync", sFuncName)
            sQuery = "[AE_SP015_SalesmanSync]'[" & p_oCompDef.p_sDataBaseName & "]'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sales Employee Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)

            Console.WriteLine("Executing BarCode Sync", sFuncName)
            sQuery = "[AE_SP014_BarcodeSync]'[" & p_oCompDef.p_sDataBaseName & "]'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BarCode Query : " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)
            Console.WriteLine("Completed Successfully ", sFuncName)


            ''Console.WriteLine("Stock Checking Query :", sFuncName)
            ''sQuery = "[AE_SP002_GetNoStockItem]'[" & p_oCompDef.p_sDataBaseName & "]'"
            ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Stock Checking Query : " & sQuery, sFuncName)
            ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            ''ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)


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
