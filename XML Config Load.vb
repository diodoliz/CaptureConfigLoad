Option Explicit

Global dict_CONN As Dictionary      'Dictionary to keep track of connections
Global dict_OPT As Dictionary       'Dictionary to keep track of options
Global fnLog_IndLevel As Long       'Logging Function - keeps track of indentation level
Global fnLog_LogLevel As Long       'Logging Function - keeps track of indentation level

' Cedar Project Module Script

Private Sub ScriptModule_BatchOpen(ByVal UserName As String, ByVal BatchDatabaseID As Long, ByVal ExternalGroupID As Long, ByVal ExternalBatchID As String, ByVal TransactionID As Long, ByVal WorkflowType As SCBCdrPROJLib.CDRDatabaseWorkflowTypes, ByVal BatchState As Long)
   fnLog_SetLogLevel(2)
   fnLog_SetIndentLevel(0)
   fnInitLoadConfig()
End Sub

Private Sub ScriptModule_Initialize(ByVal ModuleName As String)
   fnLog_SetLogLevel(2)
   fnLog_SetIndentLevel(0)
   fnInitLoadConfig()
End Sub

Private Sub ScriptModule_Terminate(ByVal ModuleName As String)
   fnCloseDBConnections()
End Sub


Public Function fnInitLoadConfig()
   'This function is used to load configuration from XML file

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   'Load The XML file
   Dim doc As New MSXML2.DOMDocument
   Dim Filename As String

   'Define the filename
   Filename = Replace(Project.Filename, ".sdp", ".xml")

   If Not doc.Load(Filename) Then
      fnLog(2, "fnInitLoadConfig() - ERROR - Loading Config file: " + Filename + " " + doc.parseError.reason)
      GoTo lbl_error
   Else
      fnLog(2, "fnInitLoadConfig() Loaded Config file: " + Filename )
      Dim nodeList As MSXML2.IXMLDOMNodeList

      '1. Read Through All of the conncetions
      Set nodeList = doc.selectNodes("/configuration/connectionStrings/connectionString")
      fnSetDBConnections(nodeList)

      '2. Read Through All of the options
      Set nodeList = doc.selectNodes("/configuration/options/option")
      fnSetOptions(nodeList)

      '2. Read Through All of the classes
      Set nodeList = doc.selectNodes("/configuration/classes/class")

      If Not nodeList Is Nothing Then
         Dim node As MSXML2.IXMLDOMNode

         For Each node In nodeList
            fnLog(2, "fnInitLoadConfig() Found Class: " + node.Attributes.getNamedItem("name").Text )
            fnSetClassSettings(node)
         Next node
      End If
   End If

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:

   fnLog(0, "fnInitLoadConfig() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function

Public Function fnSetClassSettings(classNode As MSXML2.IXMLDOMNode)
   'Function is called from fnInitLoadConfig to load individual class settings

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   fnLog(2, "fnSetClassSettings() Settings for class: " + classNode.Attributes.getNamedItem("name").Text )
   fnLog(2, "fnSetClassSettings() Settings for class: " + classNode.xml )

   Dim docClassName As String
   docClassName = classNode.Attributes.getNamedItem("name").Text

   'Ensure that class actually exists
   If Not Project.AllClasses.ItemExists(docClassName) Then
      fnLog(0, "fnSetClassSettings() - ERROR - Cannot find class: " + docClassName)
   End If

   'Get document Class onject
   Dim docClass As SCBCdrDocClass
   Set docClass = Project.AllClasses.ItemByName(docClassName)

   Dim nodeList As MSXML2.IXMLDOMNodeList

   'Read Through All of the classes
   Set nodeList = classNode.selectNodes("./fields/field")

   fnLog(2, "fnSetClassSettings() Found fields: " + CStr(nodeList.length))

   If Not nodeList Is Nothing Then
      Dim node As MSXML2.IXMLDOMNode

      For Each node In nodeList
         fnLog(2, "fnSetClassSettings() Found field: " + node.Attributes.getNamedItem("name").Text )
         fnSetFieldSettings(docClass, node)
      Next node
   End If

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:

   fnLog(0, "fnSetClassSettings() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function

Public Function fnSetFieldSettings(docClass As SCBCdrDocClass, fieldNode As MSXML2.IXMLDOMNode)
   'Function is called from to load individual field settings
   '  Properties:
   '     docClass    - Class on which the field needs to be configured
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   fnLog(2, "fnSetFieldSettings() Settings for field: " + fieldNode.Attributes.getNamedItem("name").Text )
   fnLog(2, "fnSetFieldSettings() Settings for field: " + fieldNode.xml )

   Dim fieldName As String
   fieldName = fieldNode.Attributes.getNamedItem("name").Text


   'Ensure that field actually exists
   If Not docClass.Fields.ItemExists(fieldName) Then
      fnLog(0, "fnSetClassSettings() - ERROR - Cannot find field: " + fieldNode.Attributes.getNamedItem("name").Text)
   End If

   'Get Field onject
   Dim Field As SCBCdrFieldDef
   Set Field = docClass.Fields(fieldName)

   Dim FieldType As String
   FieldType = fieldNode.Attributes.getNamedItem("type").Text

   Select Case FieldType
      Case "ASSA"
         fnLog(2, "fnSetFieldSettings() ASSA Field" )
         fnSetFieldSettings_ASSA(Field, fieldNode)
      Case "Format"
         fnLog(2, "fnSetFieldSettings() Format Field" )
         fnSetFieldSettings_Format(Field, fieldNode)
      Case Else
         fnLog(0, "fnSetFieldSettings() - ERROR - Unknown Field type - NOT LOADING CONFIG" )
   End Select

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:

   Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnSetFieldSettings() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function

Public Function fnSetFieldSettings_ASSA(Field As SCBCdrFieldDef, fieldNode As MSXML2.IXMLDOMNode)
   'Function is called from to load individual field settings for ASSA field
   '  Properties:
   '     docClass    - Class on which the field needs to be configured
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   fnLog(2, "fnSetFieldSettings_ASSA() Settings for ASSA field: " + fieldNode.Attributes.getNamedItem("name").Text )
   fnLog(2, "fnSetFieldSettings_ASSA() Settings for ASSA field: " + fieldNode.xml )

   Dim ASSAConfig As MSXML2.IXMLDOMNode
   Set ASSAConfig = fieldNode.selectSingleNode("./ASSAConfig")

   fnLog(2, "fnSetFieldSettings_ASSA() Settings for ASSA field: " + ASSAConfig.xml )

   Dim oPoolAnalysisSettings As SCBCdrSupExSettings
   Set oPoolAnalysisSettings = Field.AnalysisSetting("German")

   If oPoolAnalysisSettings Is Nothing Then
      fnLog(0, "fnSetFieldSettings_ASSA() - ERROR - Cannot get ASSA setting for field")
   End If

   'Alphanumeric settings
   Dim AlphaNum As String
   AlphaNum = ASSAConfig.selectSingleNode("./AlphaNum").Text
   fnLog(2, "fnSetFieldSettings_ASSA() AlphaNum:" + AlphaNum)
   If (UCase(AlphaNum) = "YES") Then
      oPoolAnalysisSettings.IsIDAlphNum = True
   Else
      oPoolAnalysisSettings.IsIDAlphNum = False
   End If

   'Alphanumeric settings
   Dim AutoImportOption As String
   AutoImportOption = ASSAConfig.selectSingleNode("./AutoImportOption").Text
   fnLog(2, "fnSetFieldSettings_ASSA() AutoImportOption:" + AutoImportOption)
   Select Case UCase(AutoImportOption)
    Case "ODBC"
      oPoolAnalysisSettings.AutomaticImportMethod = CdrAUTODBC
    Case "FILE"
      oPoolAnalysisSettings.AutomaticImportMethod = CdrAUTFile
    Case "NONE"
      oPoolAnalysisSettings.AutomaticImportMethod = CdrAUTNone
   End Select

   'Filename Import
   Dim FileRelative As String
   Dim ImportFilename As String
   Dim ImportPathFilename As String
   FileRelative = ASSAConfig.selectSingleNode("./FileRelative").Text
   ImportFilename = ASSAConfig.selectSingleNode("./ImportFilename").Text
   ImportPathFilename = ASSAConfig.selectSingleNode("./ImportPathFilename").Text
   fnLog(2, "fnSetFieldSettings_ASSA() FileRelative:" + FileRelative)
   fnLog(2, "fnSetFieldSettings_ASSA() ImportFilename:" + ImportFilename)
   fnLog(2, "fnSetFieldSettings_ASSA() ImportPathFilename:" + ImportPathFilename)

   If UCase(FileRelative) = "YES" Then
      oPoolAnalysisSettings.ImportFileNameRelative = True
      oPoolAnalysisSettings.ImportFileName = ImportFilename
   Else
      oPoolAnalysisSettings.ImportFileNameRelative = False
      oPoolAnalysisSettings.ImportFileName = ImportPathFilename
   End If

   'ODBC Import
   Dim ImportODBCSelect As String
   Dim ImportODBCDSN As String
   Dim ImportODBCUser As String
   Dim ImportODBCPWD As String
   ImportODBCSelect = ASSAConfig.selectSingleNode("./ImportODBCSelect").Text
   ImportODBCDSN = ASSAConfig.selectSingleNode("./ImportODBCDSN").Text
   ImportODBCUser = ASSAConfig.selectSingleNode("./ImportODBCUser").Text
   ImportODBCPWD = ASSAConfig.selectSingleNode("./ImportODBCPWD").Text
   fnLog(2, "fnSetFieldSettings_ASSA() ImportODBCSelect:" + ImportODBCSelect)
   fnLog(2, "fnSetFieldSettings_ASSA() ImportODBCDSN:" + ImportODBCDSN)
   fnLog(2, "fnSetFieldSettings_ASSA() ImportODBCUser:" + ImportODBCUser)
   fnLog(2, "fnSetFieldSettings_ASSA() ImportODBCPWD:" + ImportODBCPWD)

   If Not ImportODBCSelect = "" Then
      oPoolAnalysisSettings.SQLQuery = ImportODBCSelect
   End If
   If Not ImportODBCDSN = "" Then
      oPoolAnalysisSettings.ODBCName = ImportODBCDSN
   End If
   If Not ImportODBCUser = "" Then
      oPoolAnalysisSettings.UserName = ImportODBCUser
   End If
   If Not ImportODBCPWD = "" Then
      oPoolAnalysisSettings.Password = ImportODBCPWD
   End If

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:

   fnLog(0, "fnSetFieldSettings_ASSA() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function

Public Function fnSetFieldSettings_Format(Field As SCBCdrFieldDef, fieldNode As MSXML2.IXMLDOMNode)
   'Function is called from to load individual field settings for Format field
   '  Properties:
   '     docClass    - Class on which the field needs to be configured
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   fnLog(2, "fnSetFieldSettings_Format() Settings for ASSA field: " + fieldNode.Attributes.getNamedItem("name").Text )
   fnLog(2, "fnSetFieldSettings_Format() Settings for ASSA field: " + fieldNode.xml )

   Dim oFieldTemplate As SCBCdrFormatSettings
   Set oFieldTemplate = Field.AnalysisSetting("German")
   oFieldTemplate.DeleteAll()

   Dim formatList As MSXML2.IXMLDOMNodeList
   Set formatList = fieldNode.selectNodes("./formats/format")

   If Not formatList Is Nothing Then
      Dim node As MSXML2.IXMLDOMNode
      Dim lngCount As Long 'Keep track of number of formats
      lngCount = 0
      For Each node In formatList
         Dim formatString As String
         Dim CompareMethod As String
         Dim ignoreCharacters As String
         formatString = node.Attributes.getNamedItem("formatString").Text
         CompareMethod = node.Attributes.getNamedItem("compareMethod").Text
         ignoreCharacters = node.Attributes.getNamedItem("ignoreCharacters").Text
         fnLog(2, "fnSetFieldSettings_Format() " + CStr(lngCount) + " formatString: " + formatString )
         fnLog(2, "fnSetFieldSettings_Format() " + CStr(lngCount) + " compareMethod: " + CompareMethod )
         fnLog(2, "fnSetFieldSettings_Format() " + CStr(lngCount) + " ignoreCharacters: " + ignoreCharacters )

         'Add format
         oFieldTemplate.AddFormat(formatString)

         'Set analysis method
         oFieldTemplate.AnalysisMethod(lngCount) = CdrAnalysisMethod.CdrAnalysisString

         'Set compare method
         Select Case UCase(CompareMethod)
            Case UCase("CdrTypeLevenShtein")
               oFieldTemplate.CompareType(lngCount) = CdrCompareType.CdrTypeLevenShtein
            Case UCase("CdrTypeRegularExpression")
               oFieldTemplate.CompareType(lngCount) = CdrCompareType.CdrTypeRegularExpression
            Case UCase("CdrTypeSimpleExpression")
               oFieldTemplate.CompareType(lngCount) = CdrCompareType.CdrTypeSimpleExpression
            Case UCase("CdrTypeStringComp")
               oFieldTemplate.CompareType(lngCount) = CdrCompareType.CdrTypeStringComp
            Case UCase("CdrTypeTrigram")
               oFieldTemplate.CompareType(lngCount) = CdrCompareType.CdrTypeTrigram
            Case Else
               fnLog(0, "fnSetFieldSettings_Format() (" + CStr(lngCount) + ") - ERROR - FormatString IS INVALID: " + CompareMethod )
         End Select

         'Set ignore characters
         oFieldTemplate.IgnoreCharacters(lngCount) = ignoreCharacters

         lngCount = lngCount + 1
      Next node

   End If

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:

   fnLog(0, "fnSetFieldSettings_ASSA() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function

Public Function fnSetDBConnections(nodeList As MSXML2.IXMLDOMNodeList)
   'Function is called to populate connection objects
   '  Properties:
   '     docClass    - Class on which the field needs to be configured
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   'Close all connections and reset Connections Dictionary
   fnCloseDBConnections()

   fnLog(2, "fnSetDBConnections() Start" )
   fnLog(2, "fnSetDBConnections() Number of Connection strings found: " +  CStr(nodeList.length))

   If Not nodeList Is Nothing Then
      Dim connectionStringNode As MSXML2.IXMLDOMNode
      For Each connectionStringNode In nodeList
         Dim connectionStringName As String
         Dim connectionString As String
         connectionStringName = connectionStringNode.Attributes.getNamedItem("name").Text
         connectionString = connectionStringNode.Attributes.getNamedItem("connectionString").Text
         fnLog(2, "fnSetDBConnections() Found Connection: " + connectionStringName )
         fnLog(2, "fnSetDBConnections() Found Connection String: " + connectionString )
         If Not (dict_CONN.Exists(connectionStringName)) Then
            Dim objDBConn As ADODB.Connection
            Set objDBConn = New ADODB.Connection
            objDBConn.ConnectionString = connectionString
            dict_CONN.Add(connectionStringName, objDBConn)
            fnLog(2,"fnSetDBConnections() Connection Added: " + connectionStringName )
         End If
      Next connectionStringNode
   End If

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:

  fnLog(0, "fnSetDBConnections() - CRITICAL ERROR - " + Err.Description)
  Err.Clear()
  On Error GoTo 0
  fnLog_DecIndentLevel()

End Function

Public Function fnCloseDBConnections()
   'Function is called to close all of the active connections

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   fnLog(2,"fnCloseDBConnections() Start" )

   If Not dict_CONN Is Nothing Then
      fnLog(2,"fnCloseDBConnections() Number of connections: " + CStr(dict_CONN.Count) )
      Dim key As Variant
      Dim objDBConn As ADODB.Connection
      'For Each key In dict_CONN.Keys
         'objDBConn.Close()
      'Next key
      Dim lngCnt As Long
      For lngCnt = 0 To dict_CONN.Count - 1
         If dict_CONN.Items(lngCnt).State = 1 Then
            dict_CONN.Items(lngCnt).Close()
            fnLog(2, "fnCloseDBConnections() Connection closed: " + CStr(dict_CONN.Keys(lngCnt)) )
         Else
            fnLog(2, "fnCloseDBConnections() Connection already closed: " + CStr(dict_CONN.Keys(lngCnt)) )
         End If
      Next lngCnt
   End If

   Set dict_CONN = New Dictionary

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:

   Set dict_CONN = New Dictionary
   fnLog(0, "fnInitLoadConfig() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function

Public Function fnGetDBConnection(connectionName As String) As ADODB.Connection
   'Function is called to return connection for a ConnectionName

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   fnLog(2, "fnGetDBConnection() Start" )

   'First check if connection with this Name exists
   If Not dict_CONN.Exists(connectionName) Then
      fnLog(0, "fnGetDBConnection() - ERROR - Connection with name " + connectionName + " doesn't exist" )
      Exit Function
   Else
      fnLog(2, "fnGetDBConnection() Connection found" )
   End If

   'Get connection with this name
   Dim objDBConn As ADODB.Connection
   Set objDBConn = dict_CONN.Item(connectionName)

   If Not objDBConn.State = 1 Then
      objDBConn.open()
      fnLog(2, "fnGetDBConnection() Connection opened" )
   Else
      fnLog(2, "fnGetDBConnection() Connection already opened" )
   End If

   fnGetDBConnection = objDBConn

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:
   fnLog(0, "fnGetDBConnection() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function

Public Function fnSetOptions(nodeList As MSXML2.IXMLDOMNodeList)
   'Function is called to populate connection objects
   '  Properties:
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   fnLog(2, "fnSetOptions() Start" )
   fnLog(2, "fnSetOptions() Number of Option strings found: " +  CStr(nodeList.length))

   'Reset config options
   Set dict_OPT = New Dictionary

   If Not nodeList Is Nothing Then
      Dim optionStringNode As MSXML2.IXMLDOMNode
      For Each optionStringNode In nodeList
         Dim optionStringName As String
         Dim optionString As String
         optionStringName = optionStringNode.Attributes.getNamedItem("name").Text
         optionString = optionStringNode.Attributes.getNamedItem("value").Text
         fnLog(2, "fnSetOptions() Found Option: " + optionStringName )
         fnLog(2, "fnSetOptions() Found Option String: " + optionString )
         If Not (dict_OPT.Exists(optionStringName)) Then
            dict_OPT.Add(optionStringName, optionString)
            fnLog(2,"fnSetOptions() Option Added: " + optionStringName )
         End If
      Next optionStringNode
   End If

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:

  fnLog(0, "fnSetOptions() - CRITICAL ERROR - " + Err.Description)
  Err.Clear()
  On Error GoTo 0
  fnLog_DecIndentLevel()

End Function

Public Function fnGetOption(optionStringName As String) As String
   'Function is called to return connection for a ConnectionName

   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   fnLog(2, "fnGetOption() Start" )

   'First check if connection with this Name exists
   If Not dict_OPT.Exists(optionStringName) Then
      fnLog(0, "fnGetOption() - ERROR - Option with name " + optionStringName + " doesn't exist" )
      fnLog_DecIndentLevel()
      Exit Function
   Else
      fnLog(2, "fnGetOption() Option " + optionStringName + " found" )
   End If

   'Get connection with this name
   Dim optionString As String
   optionString = dict_OPT.Item(optionStringName)

   fnLog(2, "fnGetOption() Returning value: " + optionString )

   fnGetOption = optionString

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:
   fnLog(0, "fnGetOption() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function


Public Function fnLog(level As Integer, message As String)
   'Supported Levels are 0 - Error, 1 - Warning, 2 - Info
   'Global fnLog_IndLevel As Long       'Logging Function - keeps track of indentation level
   'Global fnLog_LogLevel As Long       'Logging Function - keeps track of indentation level
   On Error GoTo lbl_error
   fnLog_IncIndentLevel()

   If(level > fnLog_LogLevel) Then
      'Not logging this is more detailed message then we want
   Else
      Select Case level
         Case 0
            Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, Space(fnLog_IndLevel*2) + message)
         Case 1
            Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, Space(fnLog_IndLevel*2) + message)
         Case 2
            Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, Space(fnLog_IndLevel*2) + message)
         Case Else
            Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, Space(fnLog_IndLevel*2) + "fnLog() - ERROR (Invalid Logging level passed) - Message " + message)
      End Select
   End If

   fnLog_DecIndentLevel()
   Exit Function

lbl_error:
   Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnLog() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0
   fnLog_DecIndentLevel()

End Function

Public Function fnLog_SetLogLevel(logLevel As Integer )
   'Sets indentation level for logging
   Select Case logLevel
      Case 0
         fnLog_LogLevel = 0
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnLog_SetLogLevel() - Set Level to: 0")
      Case 1
         fnLog_LogLevel = 1
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnLog_SetLogLevel() - Set Level to: 1")
      Case 2
         fnLog_LogLevel = 2
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnLog_SetLogLevel() - Set Level to: 2")
      Case Else
         fnLog_LogLevel = 0
         Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, "fnLog_SetLogLevel() - ERROR (Invalid Logging level passed) - Set Level to: 0")
   End Select
End Function

Public Function fnLog_SetIndentLevel(indentLevel As Integer )
   'Sets indentation level for logging
   fnLog_IndLevel = indentLevel
End Function

Public Function fnLog_GetIndentLevel() As Integer
   'Gets indentation level for logging
   fnLog_GetIndentLevel = fnLog_IndLevel
End Function

Public Function fnLog_IncIndentLevel()
   'Increments indentation level for logging
   If(fnLog_IndLevel < 10) Then fnLog_IndLevel = fnLog_IndLevel + 1
End Function

Public Function fnLog_DecIndentLevel()
   'Increments indentation level for logging
   If(fnLog_IndLevel > 0) Then fnLog_IndLevel = fnLog_IndLevel - 1
End Function
