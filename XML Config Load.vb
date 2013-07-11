'Global Variable for connections
'Will be used to store named connections
Global dict_CONN As Dictionary     'Dictionary to keep track of connections

'Add  fnInitLoadConfig() on initialize
Private Sub ScriptModule_Initialize(ByVal ModuleName As String)
   fnInitLoadConfig()
End Sub

'Add  fnCloseDBConnections() to close all of the active connections
Private Sub ScriptModule_Terminate(ByVal ModuleName As String)
   fnCloseDBConnections()
End Sub

Public Function fnInitLoadConfig()
   'This function is used to load configuration from XML file

   On Error GoTo lbl_error

   'Load The XML file
   Dim doc As New MSXML2.DOMDocument
   Dim Filename As String

   'Define the filename
   Filename = Replace(Project.Filename, ".sdp", ".xml")

   If Not doc.Load(Filename) Then
      Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, "fnInitLoadConfig() - ERROR - Loading Config file: " + Filename + " " + doc.parseError.reason)
      GoTo lbl_error
   Else
      Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnInitLoadConfig() Loaded Config file: " + Filename )
      Dim nodeList As MSXML2.IXMLDOMNodeList

      '1. Read Through All of the conncetions
      Set nodeList = doc.selectNodes("/configuration/connectionStrings/connectionString")
      fnSetDBConnections(nodeList)

      '***TESTING****
      Exit Function
      '2. Read Through All of the classes
      Set nodeList = doc.selectNodes("/configuration/classes/class")

      If Not nodeList Is Nothing Then
         Dim node As MSXML2.IXMLDOMNode

         For Each node In nodeList
            Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, "fnInitLoadConfig() Found Class: " + node.Attributes.getNamedItem("name").Text )
            fnSetClassSettings(node)
         Next node
      End If
   End If

   Exit Function

lbl_error:

  Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnInitLoadConfig() - CRITICAL ERROR - " + Err.Description)
  Err.Clear()
  On Error GoTo 0

End Function

Public Function fnSetClassSettings(classNode As MSXML2.IXMLDOMNode)
   'Function is called from fnInitLoadConfig to load individual class settings

   On Error GoTo lbl_error

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetClassSettings() Settings for class: " + classNode.Attributes.getNamedItem("name").Text )
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetClassSettings() Settings for class: " + classNode.xml )

   Dim docClassName As String
   docClassName = classNode.Attributes.getNamedItem("name").Text

   'Ensure that class actually exists
   If Not Project.AllClasses.ItemExists(docClassName) Then
      Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, "fnSetClassSettings() - ERROR - Cannot find class: " + docClassName)
   End If

   'Get document Class onject
   Dim docClass As SCBCdrDocClass
   Set docClass = Project.AllClasses.ItemByName(docClassName)

   Dim nodeList As MSXML2.IXMLDOMNodeList

   'Read Through All of the classes
   Set nodeList = classNode.selectNodes("//class/fields/field")

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetClassSettings() Found fields: " + CStr(nodeList.length))

   If Not nodeList Is Nothing Then
      Dim node As MSXML2.IXMLDOMNode

      For Each node In nodeList
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetClassSettings() Found field: " + node.Attributes.getNamedItem("name").Text )
         fnSetFieldSettings(docClass, node)
      Next node
   End If

   Exit Function

lbl_error:

  Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnSetClassSettings() - CRITICAL ERROR - " + Err.Description)
  Err.Clear()
  On Error GoTo 0

End Function

Public Function fnSetFieldSettings(docClass As SCBCdrDocClass, fieldNode As MSXML2.IXMLDOMNode)
   'Function is called from to load individual field settings
   '  Properties:
   '     docClass    - Class on which the field needs to be configured
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings() Settings for field: " + fieldNode.Attributes.getNamedItem("name").Text )
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings() Settings for field: " + fieldNode.xml )

   Dim fieldName As String
   fieldName = fieldNode.Attributes.getNamedItem("name").Text


   'Ensure that field actually exists
   If Not docClass.Fields.ItemExists(fieldName) Then
      Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, "fnSetClassSettings() - ERROR - Cannot find field: " + fieldNode.Attributes.getNamedItem("name").Text)
   End If

   'Get Field onject
   Dim Field As SCBCdrFieldDef
   Set Field = docClass.Fields(fieldName)

   Dim FieldType As String
   FieldType = fieldNode.Attributes.getNamedItem("type").Text

   Select Case FieldType
      Case "ASSA"
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings() ASSA Field" )
         fnSetFieldSettings_ASSA(Field, fieldNode)
      Case "Format"
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings() Format Field" )
         fnSetFieldSettings_Format(Field, fieldNode)
      Case Else
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings() - ERROR - Unknown Field type - NOT LOADING CONFIG" )
   End Select

   Exit Function

lbl_error:

  Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnSetFieldSettings() - CRITICAL ERROR - " + Err.Description)
  Err.Clear()
  On Error GoTo 0

End Function

Public Function fnSetFieldSettings_ASSA(Field As SCBCdrFieldDef, fieldNode As MSXML2.IXMLDOMNode)
   'Function is called from to load individual field settings for ASSA field
   '  Properties:
   '     docClass    - Class on which the field needs to be configured
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() Settings for ASSA field: " + fieldNode.Attributes.getNamedItem("name").Text )
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() Settings for ASSA field: " + fieldNode.xml )

   Dim ASSAConfig As MSXML2.IXMLDOMNode
   Set ASSAConfig = fieldNode.selectSingleNode("//field/ASSAConfig")

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() Settings for ASSA field: " + ASSAConfig.xml )

   Dim oPoolAnalysisSettings As SCBCdrSupExSettings
   Set oPoolAnalysisSettings = Field.AnalysisSetting("German")

   If oPoolAnalysisSettings Is Nothing Then
      Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() - ERROR - Cannot get ASSA setting for field")
   End If

   'Alphanumeric settings
   Dim AlphaNum As String
   AlphaNum = ASSAConfig.selectSingleNode("//ASSAConfig/AlphaNum").Text
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() AlphaNum:" + AlphaNum)
   If (UCase(AlphaNum) = "YES") Then
      oPoolAnalysisSettings.IsIDAlphNum = True
   Else
      oPoolAnalysisSettings.IsIDAlphNum = False
   End If

   'Alphanumeric settings
   Dim AutoImportOption As String
   AutoImportOption = ASSAConfig.selectSingleNode("//ASSAConfig/AutoImportOption").Text
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() AutoImportOption:" + AutoImportOption)
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
   FileRelative = ASSAConfig.selectSingleNode("//ASSAConfig/FileRelative").Text
   ImportFilename = ASSAConfig.selectSingleNode("//ASSAConfig/ImportFilename").Text
   ImportPathFilename = ASSAConfig.selectSingleNode("//ASSAConfig/ImportPathFilename").Text
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() FileRelative:" + FileRelative)
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() ImportFilename:" + ImportFilename)
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() ImportPathFilename:" + ImportPathFilename)

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
   ImportODBCSelect = ASSAConfig.selectSingleNode("//ASSAConfig/ImportODBCSelect").Text
   ImportODBCDSN = ASSAConfig.selectSingleNode("//ASSAConfig/ImportODBCDSN").Text
   ImportODBCUser = ASSAConfig.selectSingleNode("//ASSAConfig/ImportODBCUser").Text
   ImportODBCPWD = ASSAConfig.selectSingleNode("//ASSAConfig/ImportODBCPWD").Text
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() ImportODBCSelect:" + ImportODBCSelect)
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() ImportODBCDSN:" + ImportODBCDSN)
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() ImportODBCUser:" + ImportODBCUser)
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() ImportODBCPWD:" + ImportODBCPWD)

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

   Exit Function

lbl_error:

  Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() - CRITICAL ERROR - " + Err.Description)
  Err.Clear()
  On Error GoTo 0

End Function

Public Function fnSetFieldSettings_Format(Field As SCBCdrFieldDef, fieldNode As MSXML2.IXMLDOMNode)
   'Function is called from to load individual field settings for Format field
   '  Properties:
   '     docClass    - Class on which the field needs to be configured
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_Format() Settings for ASSA field: " + fieldNode.Attributes.getNamedItem("name").Text )
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_Format() Settings for ASSA field: " + fieldNode.xml )

   Dim oFieldTemplate As SCBCdrFormatSettings
   Set oFieldTemplate = Field.AnalysisSetting("German")
   oFieldTemplate.DeleteAll()

   Dim formatList As MSXML2.IXMLDOMNodeList
   Set formatList = fieldNode.selectNodes("//field/formats/format")

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
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_Format() (" + CStr(lngCount) + " formatString: " + formatString )
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_Format() (" + CStr(lngCount) + " compareMethod: " + CompareMethod )
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetFieldSettings_Format() (" + CStr(lngCount) + " ignoreCharacters: " + ignoreCharacters )

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
               Project.LogScriptMessageEx(CDRTypeWarning, CDRSeverityLogFileOnly, "fnSetFieldSettings_Format() (" + CStr(lngCount) + ") - ERROR - FormatString IS INVALID: " + CompareMethod )
         End Select

         'Set ignore characters
         oFieldTemplate.IgnoreCharacters(lngCount) = ignoreCharacters

         lngCount = lngCount + 1
      Next node

   End If

   Exit Function

lbl_error:

  Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnSetFieldSettings_ASSA() - CRITICAL ERROR - " + Err.Description)
  Err.Clear()
  On Error GoTo 0

End Function

Public Function fnSetDBConnections(nodeList As MSXML2.IXMLDOMNodeList)
   'Function is called to populate connection objects
   '  Properties:
   '     docClass    - Class on which the field needs to be configured
   '     fieldNode   - XML node containing configuration

   On Error GoTo lbl_error

   'Close all connections and reset Connections Dictionary
   fnCloseDBConnections()

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetDBConnections() Start" )
   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetDBConnections() Number of Connection strings found: " +  CStr(nodeList.length))

   If Not nodeList Is Nothing Then
      Dim connectionStringNode As MSXML2.IXMLDOMNode
      For Each connectionStringNode In nodeList
         Dim connectionStringName As String
         Dim connectionString As String
         connectionStringName = connectionStringNode.Attributes.getNamedItem("name").Text
         connectionString = connectionStringNode.Attributes.getNamedItem("connectionString").Text
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetDBConnections() Found Connection: " + connectionStringName )
         Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetDBConnections() Found Connection String: " + connectionString )
         If Not (dict_CONN.Exists(connectionStringName)) Then
            Dim objDBConn As ADODB.Connection
            Set objDBConn = New ADODB.Connection
            objDBConn.ConnectionString = connectionString
            dict_CONN.Add(connectionStringName, objDBConn)
            Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnSetDBConnections() Connection Added: " + connectionStringName )
         End If
      Next connectionStringNode
   End If
   Exit Function

lbl_error:

  Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnSetDBConnections() - CRITICAL ERROR - " + Err.Description)
  Err.Clear()
  On Error GoTo 0

End Function

Public Function fnCloseDBConnections()
   'Function is called to close all of the active connections

   On Error GoTo lbl_error

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnCloseDBConnections() Start" )

   If Not dict_CONN Is Nothing Then
      Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnCloseDBConnections() Number of connections: " + CStr(dict_CONN.Count) )
      Dim key As Variant
      Dim objDBConn As ADODB.Connection
      'For Each key In dict_CONN.Keys
         'objDBConn.Close()
      'Next key
      Dim lngCnt As Long
      For lngCnt = 0 To dict_CONN.Count - 1
         If dict_CONN.Items(lngCnt).State = 1 Then
            dict_CONN.Items(lngCnt).Close()
            Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnCloseDBConnections() Connection closed: " + CStr(dict_CONN.Keys(lngCnt)) )
         Else
            Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnCloseDBConnections() Connection already closed: " + CStr(dict_CONN.Keys(lngCnt)) )
         End If
      Next lngCnt
   End If

   Set dict_CONN = New Dictionary

   Exit Function

lbl_error:
   Set dict_CONN = New Dictionary
   Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnInitLoadConfig() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0

End Function

Public Function fnGetDBConnection(connectionName As String) As ADODB.Connection
   'Function is called to return connection for a ConnectionName

   On Error GoTo lbl_error

   Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnGetDBConnection() Start" )

   'First check if connection with this Name exists
   If Not dict_CONN.Exists(connectionName) Then
      Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnGetDBConnection() - ERROR - Connection with name " + connectionName + " doesn't exist" )
      Exit Function
   Else
      Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnGetDBConnection() Connection found" )
   End If

   'Get connection with this name
   Dim objDBConn As ADODB.Connection
   Set objDBConn = dict_CONN.Item(connectionName)

   If Not objDBConn.State = 1 Then
      objDBConn.open()
      Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnGetDBConnection() Connection opened" )
   Else
      Project.LogScriptMessageEx(CDRTypeInfo, CDRSeverityLogFileOnly, "fnGetDBConnection() Connection already opened" )
   End If

   fnGetDBConnection = objDBConn

   Exit Function

lbl_error:
   Project.LogScriptMessageEx(CDRTypeError, CDRSeverityLogFileOnly, "fnGetDBConnection() - CRITICAL ERROR - " + Err.Description)
   Err.Clear()
   On Error GoTo 0

End Function
