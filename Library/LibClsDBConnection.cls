VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibClsDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Workbook_Open()
    Dim iConnection As New ADODB.Connection
    Dim iCommand As New ADODB.Command
    Dim iParameter As ADODB.Parameter
    Dim iRecordset As ADODB.Recordset

    ' Create and Open Connection Object.
    Call iConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\path\to\database.accdb;")
    ' Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\osama.hassanein\Documents\Database1.accdb;Persist Security Info=False;
    ' \\kedata\Data1\B2B_Business_Inteligence\Osama Hassanein\tmp

    ' Create Command Object.
    iCommand.ActiveConnection = iConnection
    iCommand.CommandText = "SELECT * FROM Table1 WHERE ID > ?"

    ' Create Parameter Object.
    Set iParameter = iCommand.CreateParameter(, adInteger, adParamInput, 5)
    iParameter.Value = 2
    Call iCommand.Parameters.Append(iParameter)
    Set iParameter = Nothing

    ' Open Recordset Object.
    Set iRecordset = iCommand.Execute()

    While Not iRecordset.EOF
        Call VBA.MsgBox(iRecordset.Fields(0).Name & ": " & iRecordset.Fields(0).Value)
        Call VBA.MsgBox(iRecordset.Fields(1).Name & ": " & iRecordset.Fields(1).Value)

        Call iRecordset.MoveNext()
    Wend

    'Terminate connection and variables
    Call iRecordset.Close()
    Set iRecordset = Nothing

    Call iConnection.Close()
    Set iConnection = Nothing
End Sub

'Option Explicit
'
'Private Const vDwhCredentials As String = ""
'
'Private vInternalConnection As ADODB.Connection
'Private vActiveCommands As Collection
'Private vActiveRecordsets As Collection
'
'Private Sub Class_Initialize()
'    Set vActiveCommands = New Collection
'    Set vActiveRecordsets = New Collection
'
'    On Error GoTo HandleRetry64:
'
'    Set vInternalConnection = New ADODB.Connection
'    vInternalConnection.ConnectionString = "Driver={Oracle in OraClient12Home1};" & vDwhCredentials
'    Call vInternalConnection.Open
'
'Terminate:
'    On Error GoTo 0
'    Call MException.Propagate
'    Exit Sub
'
'HandleError:
'    Set vActiveRecordsets = Nothing
'    Set vActiveCommands = Nothing
'    Set vInternalConnection = Nothing
'
'    Call MException.Store(MViewData.GetCaption("UnableToConnectToDatabase"))
'    Resume Terminate:
'
'HandleRetry64:
'    Resume Retry64:
'Retry64:
'    On Error GoTo HandleError:
'
'    Set vInternalConnection = New ADODB.Connection
'    vInternalConnection.ConnectionString = "Driver={Oracle in OraClient12Home1_32bit};" & vDwhCredentials
'    Call vInternalConnection.Open
'
'    GoTo Terminate:
'End Sub
'
'Private Sub Class_Terminate()
'    Dim vRecordset As ADODB.Recordset
'
'    'Release and clear allocated recordsets.
'    If Not (vActiveRecordsets Is Nothing) Then
'        For Each vRecordset In vActiveRecordsets
'            If Not (vRecordset Is Nothing) Then
'                If vRecordset.State = 1 Then
'                    Call vRecordset.Close
'                End If
'            End If
'        Next
'    End If
'
'    'Release and clear connection.
'    If Not (vInternalConnection Is Nothing) Then
'        If vInternalConnection.State = 1 Then
'            Call vInternalConnection.Close
'        End If
'    End If
'End Sub
'
'Public Function ExecuteQuery( _
'    ByRef vQueryText As String, _
'    ByRef vQueryParameters As Collection _
') As ADODB.Recordset
'    Dim vCommand As ADODB.command
'    Dim vQueryParameter As CQueryParameter
'    Dim vParameter As ADODB.parameter
'
'    On Error GoTo HandleError:
'
'    ' Create Command Object.
'    Set vCommand = New ADODB.command
'    vCommand.ActiveConnection = vInternalConnection
'    vCommand.CommandText = vQueryText
'    Call vActiveCommands.Add(vCommand)
'
'    ' Create Parameter Object.
'    For Each vQueryParameter In vQueryParameters
'        Select Case vQueryParameter.vDataType
'            Case EDataType.eBoolean
'                Set vParameter = vCommand.CreateParameter(Type:=adBoolean)
'                vParameter.Value = CBool(vQueryParameter.vValue)
'            Case EDataType.eInteger
'                Set vParameter = vCommand.CreateParameter(Type:=adInteger, Size:=64)
'                vParameter.Value = CLng(vQueryParameter.vValue)
'            Case EDataType.eFloat
'                Set vParameter = vCommand.CreateParameter(Type:=adDouble, Size:=128)
'                vParameter.Value = CDbl(vQueryParameter.vValue)
'            Case EDataType.eText
'                Set vParameter = vCommand.CreateParameter(Type:=adVarWChar, Size:=255)
'                vParameter.Value = CStr(vQueryParameter.vValue)
'            Case Else
'                Call MException.Throw
'        End Select
'
'        vParameter.Direction = adParamInput
'        Call vCommand.Parameters.Append(vParameter)
'    Next
'
'    ' Open Recordset Object.
'    Set ExecuteQuery = vCommand.Execute
'    Call vActiveRecordsets.Add(ExecuteQuery)
'
'Terminate:
'    On Error GoTo 0
'
'    Call MException.Propagate
'
'    Exit Function
'HandleError:
'    Call MException.Store(MViewData.GetCaption("UnableToExecuteQueryOnDatabase"))
'
'    Resume Terminate:
'End Function
'
''Option Explicit
''
''Private Const dwhCredentials As String = "Dbq=XXX;Uid=YYY;Pwd=ZZZ"
''
''Private internalConnection As ADODB.Connection
''Private activeCommands As Collection
''Private activeRecordsets As Collection
''
''Private Sub Initialize()
''    Set activeCommands = New Collection
''    Set activeRecordsets = New Collection
''
''    On Error GoTo HandleRetry64:
''
''    Set internalConnection = New ADODB.Connection
''    internalConnection.ConnectionString = "Driver={Oracle in OraClient12Home1};" & dwhCredentials
''    internalConnection.Open
''
''Terminate:
''    On Error GoTo 0
''    MException.Propagate
''    Exit Sub
''
''HandleError:
''    Set activeRecordsets = Nothing
''    Set activeCommands = Nothing
''    Set internalConnection = Nothing
''
''    MException.Store MViewData.GetCaption("UnableToConnectToDatabase")
''    Resume Terminate:
''
''HandleRetry64:
''    Resume Retry64:
''Retry64:
''    On Error GoTo HandleError:
''
''    Set internalConnection = New ADODB.Connection
''    internalConnection.ConnectionString = "Driver={Oracle in OraClient12Home1_32bit};" & dwhCredentials
''    internalConnection.Open
''
''    GoTo Terminate:
''End Sub
''
''Private Sub Terminate()
''    Dim iRecordset As ADODB.Recordset
''
''    'Release and clear allocated recordsets.
''    If Not (activeRecordsets Is Nothing) Then
''        For Each iRecordset In activeRecordsets
''            If Not (iRecordset Is Nothing) Then
''                If iRecordset.State = 1 Then
''                    iRecordset.Close
''                End If
''            End If
''        Next
''    End If
''    Set activeRecordsets = Nothing
''
''    'Release and clear allocated commands.
''    Set activeCommands = Nothing
''
''    'Release and clear connection.
''    If Not (internalConnection Is Nothing) Then
''        If internalConnection.State = 1 Then
''            internalConnection.Close
''        End If
''    End If
''    Set internalConnection = Nothing
''
''    Set iRecordset = Nothing
''End Sub
''
''Public Function ExecuteQuery( _
''    ByRef queryText As String, _
''    ByRef queryParameters As Collection _
'') As ADODB.Recordset
''    Dim command As ADODB.command
''    Dim iQueryParameter As CQueryParameter
''    Dim parameter As ADODB.parameter
''
''    On Error GoTo HandleError:
''
''    ' Create Command Object.
''    Set command = New ADODB.command
''    command.ActiveConnection = internalConnection
''    command.CommandText = queryText
''    activeCommands.Add command
''
''    ' Create Parameter Object.
''    For Each iQueryParameter In queryParameters
''        Select Case iQueryParameter.DataType
''            Case EDataType.eBoolean
''                Set parameter = command.CreateParameter(Type:=adBoolean)
''                parameter.Value = CBool(iQueryParameter.Value)
''            Case EDataType.eInteger
''                Set parameter = command.CreateParameter(Type:=adInteger, Size:=64)
''                parameter.Value = CLng(iQueryParameter.Value)
''            Case EDataType.eFloat
''                Set parameter = command.CreateParameter(Type:=adDouble, Size:=128)
''                parameter.Value = CDbl(iQueryParameter.Value)
''            Case EDataType.eText
''                Set parameter = command.CreateParameter(Type:=adVarWChar, Size:=255)
''                parameter.Value = CStr(iQueryParameter.Value)
''            Case Else
''                MException.Throw
''        End Select
''
''        parameter.Direction = adParamInput
''        command.Parameters.Append parameter
''    Next
''
''    ' Open Recordset Object.
''    Set ExecuteQuery = command.Execute
''    activeRecordsets.Add ExecuteQuery
''
''Terminate:
''    On Error GoTo 0
''
''    Set command = Nothing
''    Set iQueryParameter = Nothing
''    Set parameter = Nothing
''
''    MException.Propagate
''    Exit Function
''HandleError:
''    Set ExecuteQuery = Nothing
''
''    MException.Store MViewData.GetCaption("UnableToExecuteQueryOnDatabase")
''    Resume Terminate:
''End Function
''
''
''' Hints
''' Preloading ADODB.Field objects before iterating through a large recordset increases performance
''' Commands help prevent sql injection but parameters cannot be named so ordered application is the only feasible method for using parameters
''
''Sub DoSomeDatabaseStuff()
''    Dim iConnection As ADODB.Connection
''    Dim iCommand As ADODB.command
''    Dim iParameter As ADODB.parameter
''    Dim iRecordset As ADODB.Recordset
''
''    ' Create and Open Connection Object.
''    Set iConnection = New ADODB.Connection
''    iConnection.ConnectionString = "DSN=Biblio;UID=admin;PWD=xxx;"
''    iConnection.Open
''
''    ' Create Command Object.
''    Set iCommand = New ADODB.command
''    iCommand.ActiveConnection = iConnection
''    iCommand.CommandText = "SELECT * FROM authors WHERE au_id < ?"
''
''    ' Create Parameter Object.
''    Set iParameter = iCommand.CreateParameter(, adInteger, adParamInput, 5)
''    iParameter.Value = 5
''    iCommand.Parameters.Append iParameter
''    Set iParameter = Nothing
''
''    ' Open Recordset Object.
''    Set iRecordset = iCommand.Execute()
''
''    'Do something ...
''
''    'Terminate connection and variables
''End Sub
''
''
''Option Explicit
''Option Private Module
''
''Private Const strConnectionString As String = "Driver={Oracle in OraClient12Home1};Dbq=EWH9;Uid=MAX;Pwd=xx_4micx"
''Private Const strConnectionString64 As String = "Driver={Oracle in OraClient12Home1_32bit};Dbq=EWH9;Uid=MAX;Pwd=xx_4micx"
''Private objConnection As ADODB.Connection
''
''Public Sub Initialize()
''    ' Instantiate connection.
''    Set objConnection = New ADODB.Connection
''    objConnection.ConnectionString = strConnectionString
''    objConnection.CommandTimeout = 0
''
''    ' Open connection.
''    objConnection.Open
''End Sub
''
''Public Sub Terminate()
''    ' Close connection.
''    If Not (objConnection Is Nothing) Then
''        If objConnection.State = 1 Then
''            objConnection.Close
''        End If
''    End If
''
''    ' Dereference connection.
''    Set objConnection = Nothing
''End Sub
''
''Public Function CreateConnection() As ADODB.Connection
''    On Error GoTo HandleError:
''
''    Set CreateConnection = New ADODB.Connection
''    CreateConnection.Open strConnectionString
''Terminate:
''    On Error GoTo 0
''
''    CreateConnection.CommandTimeout = 0
''
''    Exit Function
''HandleError:
''    Set CreateConnection = New ADODB.Connection
''    CreateConnection.Open strConnectionString64
''
''    Resume Terminate:
''End Function
''
''Public Function Execute( _
''    ByRef strQuery As String, _
''    ByRef colParameterValues As Collection _
'') As ADODB.Recordset
''    'Create a populated command object.
''    Dim objCommand As ADODB.command
''    Set objCommand = New ADODB.command
''    objCommand.ActiveConnection = objConnection
''    objCommand.CommandText = strQuery
''
''    ' Set parameters.
''    Dim varValue As Variant
''    Dim objParameter As ADODB.parameter
''    For Each varValue In colParameterValues
''        'Create parameter object with a datatype based on the value.
''        Select Case VarType(varValue)
''            Case vbBoolean
''                Set objParameter = objCommand.CreateParameter(Type:=adBoolean)
''            Case vbInteger
''                Set objParameter = objCommand.CreateParameter(Type:=adInteger, Size:=64)
''            Case vbLong
''                Set objParameter = objCommand.CreateParameter(Type:=adInteger, Size:=64)
''            Case vbDouble
''                Set objParameter = objCommand.CreateParameter(Type:=adDouble, Size:=128)
''            Case vbString
''                Set objParameter = objCommand.CreateParameter(Type:=adVarWChar, Size:=255)
''            Case Else
''                ' MProcedure.ThrowError strDescription:="An unrecognised type has been passed to the query parameters object"
''                Err.Raise Number:=0, Description:="Nastala chyba pri komunik�cii s datab�zou"
''        End Select
''
''        'Set value and direction of the parameter object.
''        objParameter.Value = varValue
''        objParameter.Direction = adParamInput
''
''        'Append parameter object to the command object.
''        objCommand.Parameters.Append objParameter
''    Next
''    Set objParameter = Nothing
''    varValue = Null
''
''    ' Execute the command and retrieve the recordset object.
''    Set Execute = objCommand.Execute
''    Set objCommand = Nothing
''End Function
''
''Option Explicit
''Option Private Module
''
''Private Const strConnectionString As String = "Driver={Oracle in OraClient12Home1};Dbq=EWH9;"
''Private Const strConnectionString64 As String = "Driver={Oracle in OraClient12Home1_32bit};Dbq=EWH9;"
''Private objConnection As ADODB.Connection
''
''Public Sub Initialize()
''    Dim strUsername As String
''    Dim strPassword As String
''    Dim strSuffix As String
''
''    On Error GoTo HandleError:
''
''    ' Retrieve username and password
''    UFLogin.Show
''    strSuffix = "Uid=" & UFLogin.tbUsername & ";Pwd=" & UFLogin.tbPassword
''
''    ' Instantiate connection.
''    Set objConnection = New ADODB.Connection
''    objConnection.Open strConnectionString & strSuffix
''Terminate:
''    On Error GoTo 0
''
''    objConnection.CommandTimeout = 0
''
''    Exit Sub
''HandleError:
''    ' Retry using an alternate connection string.
''    Set objConnection = New ADODB.Connection
''    objConnection.Open strConnectionString64 & strSuffix
''
''    Resume Terminate:
''End Sub
''
''Public Sub Terminate()
''    ' Close connection.
''    If Not (objConnection Is Nothing) Then
''        If objConnection.State = 1 Then
''            objConnection.Close
''        End If
''    End If
''
''    ' Dereference connection.
''    Set objConnection = Nothing
''End Sub
''
''Public Function Execute( _
''    ByRef strQuery As String, _
''    Optional ByRef colParameterValues As Collection = Nothing _
'') As ADODB.Recordset
''    'Create a populated command object.
''    Dim objCommand As ADODB.command
''    Set objCommand = New ADODB.command
''    objCommand.ActiveConnection = objConnection
''    objCommand.CommandText = strQuery
''
''    ' Set parameters.
''    If Not colParameterValues Is Nothing Then
''        Dim varValue As Variant
''        Dim objParameter As ADODB.parameter
''        For Each varValue In colParameterValues
''            'Create parameter object with a datatype based on the value.
''            Select Case VarType(varValue)
''                Case vbBoolean
''                    Set objParameter = objCommand.CreateParameter(Type:=adBoolean)
''                Case vbInteger
''                    Set objParameter = objCommand.CreateParameter(Type:=adInteger, Size:=64)
''                Case vbLong
''                    Set objParameter = objCommand.CreateParameter(Type:=adInteger, Size:=64)
''                Case vbDouble
''                    Set objParameter = objCommand.CreateParameter(Type:=adDouble, Size:=128)
''                Case vbString
''                    Set objParameter = objCommand.CreateParameter(Type:=adVarWChar, Size:=255)
''                Case Else
''                    ' MProcedure.ThrowError strDescription:="An unrecognised type has been passed to the query parameters object"
''                    Err.Raise Number:=0, Description:="Nastala chyba pri komunik�cii s datab�zou"
''            End Select
''
''            'Set value and direction of the parameter object.
''            objParameter.Value = varValue
''            objParameter.Direction = adParamInput
''
''            'Append parameter object to the command object.
''            objCommand.Parameters.Append objParameter
''        Next
''
''        Set objParameter = Nothing
''        varValue = Null
''    End If
''
''    ' Execute the command and retrieve the recordset object.
''    Set Execute = objCommand.Execute
''    Set objCommand = Nothing
''End Function
''
''Public Function ExecuteFile( _
''    ByRef strFilePath As String, _
''    Optional ByRef colParameterValues As Collection = Nothing _
'') As ADODB.Recordset
''    Set ExecuteFile = MDatabase.Execute(MFile.ReadFileText(strFilePath), colParameterValues)
''End Function
