Imports System.Data.SqlClient
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
            Try
                oSQLAdapter.Fill(oDT_INTDBInformations)
            Catch ex As Exception
            End Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Return RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return RTN_ERROR
        Finally
            oSQLAdapter.Dispose()
            oSQLCommand.Dispose()
            oConnection.Close()
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
