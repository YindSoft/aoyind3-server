VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMySQL"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
    Source As Any, ByVal Length As Long)

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private conn_main As Long
Private conn_exec As Long
Private conn_stored As Long

Private m_Hostname As String
Private m_Database As String
Private m_UserName As String
Private m_password As String

' # Procedure used to connect to MySQL #
Public Function SQLConnect(ByVal Hostname As String, ByVal Database As String, ByVal UserName As String, ByVal password As String) As Boolean
    
    On Error GoTo Error_Handler
    
    ' ** Connect to MySQL Server **
    conn_main = mysql_init(0)
    conn_exec = mysql_init(0)
    conn_stored = mysql_init(0)
        
    If mysql_real_connect(conn_main, Hostname, UserName, password, Database, 3306, vbNullString, 0) <> 0 And mysql_real_connect(conn_exec, Hostname, UserName, password, Database, 3306, vbNullString, 0) <> 0 _
        And mysql_real_connect(conn_stored, Hostname, UserName, password, Database, 3306, vbNullString, CLIENT_MULTI_RESULTS) <> 0 Then
                                                                    
        m_Hostname = Hostname
        m_Database = Database
        m_UserName = UserName
        m_password = password
        
        'Call SQLQuery_Simple("UPDATE `personajes` SET char_Online = '0'")
                
        SQLConnect = True 'Return True
            
    Else
        ' Error Unable to Connect!
        Call InternalMySQLError(vbNullString, conn_main)
        SQLConnect = False 'Return False
    End If
    
    Exit Function
    
Error_Handler:
        Call LogError("SQLConnect: " & Err.Description & " - " & Err.Number)
        Call SQLDisconnect
        SQLConnect = False
    
End Function

Public Sub SQLDisconnect()

If conn_main > 0 Then
    mysql_close (conn_main)
End If

If conn_exec > 0 Then
    mysql_close (conn_exec)
End If

If conn_stored > 0 Then
    mysql_close (conn_stored)
End If

End Sub

Public Function SQLQuery_Simple(ByRef strQuery As String) As Boolean

On Error GoTo Error_Handler

If mysql_send_query(conn_exec, strQuery, Len(strQuery)) Then
    Call InternalMySQLError(strQuery, conn_exec)
    Exit Function
End If

SQLQuery_Simple = True

Exit Function

Error_Handler:
    Call LogError("SQLQuery_Simple: " & Err.Description & " - " & Err.Number)

End Function

Public Function SQLQuery_Simple_Sync(ByRef strQuery As String) As Boolean

On Error GoTo Error_Handler

Dim lngMySQLRes As Long

If mysql_query(conn_main, strQuery) Then
    Call InternalMySQLError(strQuery, conn_main)
Else
    lngMySQLRes = mysql_store_result(conn_main)
    If lngMySQLRes <> 0 Then mysql_free_result (lngMySQLRes)
    SQLQuery_Simple_Sync = True
End If

Exit Function

Error_Handler:
    Call LogError("SQLQuery_Simple_Sync: " & Err.Description & " - " & Err.Number)

End Function

Public Function SQLQuery_Fast(ByRef strQuery As String, ByVal strField As String) As String

On Error GoTo Error_Handler

Dim staticRS As clsMySQLRecordSet

If SQLQuery(strQuery, staticRS) > 0 Then
    SQLQuery_Fast = staticRS(strField)
    If staticRS.State = adStateOpen Then staticRS.CloseRecordset
    Set staticRS = Nothing
End If

Exit Function

Error_Handler:
    Call LogError("SQLQuery_Fast: " & Err.Description & " - " & Err.Number)

End Function


Private Function Overflow_Check_Long(ByVal value As Currency, Optional ByVal default As Long = 0, Optional ByVal debugStr As String = "") As Long
On Error GoTo SetDefault
    Overflow_Check_Long = value
    Exit Function
SetDefault:
    Overflow_Check_Long = default
    LogError ("Overflow_Check_Long: " & Err.Description & " - " & Err.Number & " debugStr: " & debugStr)
End Function






Public Function Escape(strFrom As String) As String

On Error GoTo Error_Handler

Dim lngFromLength As String     'LENGTH OF SOURCE STRING
Dim strTo As String             'DESTINATION STRING COMING FROM FUNCTION
Dim lngToLength As Long         'LENGTH OF DESTINATION STRING
    
    lngFromLength = Len(strFrom)    'GET LENGTH OF INPUT
    
    strTo = Space$(lngFromLength * 2 + 1) 'ALLOCATE A BUFFER FOR OUTPUT OF FUNCTION
                                         '2 BYTES PER CHARACTER PLUS A BYTE FOR NULL
                                         'TERMINATOR USED BY FUNCTION
    
    If conn_main = 0 Then
        lngToLength = mysql_escape_string(strTo, strFrom, lngFromLength) 'CALL API
    Else
        lngToLength = mysql_real_escape_string(conn_main, strTo, strFrom, lngFromLength) 'CALL API
    End If
    
    Escape = left$(strTo, lngToLength) 'TRIM NULL TERMINATOR

Exit Function

Error_Handler:
    Call LogError("Escape: " & Err.Description & " - " & Err.Number)

End Function

Public Function SQLQuery(ByRef strQuery As String, ByRef cMySQLRecord As clsMySQLRecordSet) As Long

Dim lngRow As Long, lngRowOffset As Long, lngLenghts As Long, lngMyRes As Long
Dim i As Long, j As Long, lngFields As Long, arrLenghts() As Long, strData As String

Dim strFieldName As String, lngReturn As Long
Dim lngRecAffected As Long

Dim objMyField As MYSQL_FIELD

Dim tmpArr() As Byte

If (mysql_query(conn_main, strQuery) = 0) Then

    lngMyRes = mysql_store_result(conn_main)

    If (lngMyRes <> 0) Then
        
        lngRecAffected = mysql_num_rows(lngMyRes)
        lngFields = mysql_num_fields(lngMyRes)
        
        If lngRecAffected > 0 Then
        
            Set cMySQLRecord = New clsMySQLRecordSet
            
            Call cMySQLRecord.Initialize(lngRecAffected, lngFields)
            
            For j = 0 To lngFields - 1
                lngReturn = mysql_fetch_field(lngMyRes)
                
                If lngReturn <> 0 Then
                    CopyMemory objMyField, ByVal lngReturn, LenB(objMyField)
                    strFieldName = ptr2str(objMyField.Name)
                    cMySQLRecord.AddField (strFieldName)
                End If
                
            Next j
            
            ReDim arrLenghts(0 To lngFields - 1)
            
            For i = 0 To lngRecAffected - 1
            
                lngRow = mysql_fetch_row(lngMyRes)
                CopyMemory lngRowOffset, ByVal lngRow, 4

                lngLenghts = mysql_fetch_lengths(lngMyRes)
                CopyMemory arrLenghts(0), ByVal lngLenghts, 4 * lngFields
                
                For j = 0 To lngFields - 1
                    
                    If arrLenghts(j) > 0 Then
                        ReDim tmpArr(arrLenghts(j) - 1) As Byte
                        Call CopyMemory(tmpArr(0), ByVal lngRowOffset, arrLenghts(j))
                        strData = StrConv(tmpArr(), vbUnicode)
                    Else
                        strData = vbNullString
                    End If
                    
                    lngRowOffset = lngRowOffset + arrLenghts(j) + 1
                    Call cMySQLRecord.SetValue(j, i, strData, arrLenghts(j))
                Next j
                
            Next i

            SQLQuery = lngRecAffected
            
        End If
        
        mysql_free_result (lngMyRes)
        
    End If
Else
    Call InternalMySQLError(strQuery, conn_main)
    SQLQuery = -1
End If

End Function

Public Function SQLQuery_Multi(ByRef strQuery As String, ByRef cMySQLRecord As clsMySQLRecordSet) As Long

Dim lngRow As Long, lngRowOffset As Long, lngLenghts As Long, lngMyRes As Long
Dim i As Long, j As Long, lngFields As Long, arrLenghts() As Long, strData As String

Dim strFieldName As String, lngReturn As Long
Dim lngRecAffected As Long

Dim tmpArr() As Byte

Dim objMyField As MYSQL_FIELD

If (mysql_query(conn_stored, strQuery) = 0) Then

    lngMyRes = mysql_store_result(conn_stored)

    If (lngMyRes <> 0) Then
        
        lngRecAffected = mysql_num_rows(lngMyRes)
        lngFields = mysql_num_fields(lngMyRes)
        
        If lngRecAffected > 0 Then
        
            Set cMySQLRecord = New clsMySQLRecordSet
            
            Call cMySQLRecord.Initialize(lngRecAffected, lngFields)
            
            For j = 0 To lngFields - 1
                lngReturn = mysql_fetch_field(lngMyRes)
                
                If lngReturn <> 0 Then
                    CopyMemory objMyField, ByVal lngReturn, LenB(objMyField)
                    strFieldName = ptr2str(objMyField.Name)
                    cMySQLRecord.AddField (strFieldName)
                End If
                
            Next j
            
            ReDim arrLenghts(0 To lngFields - 1)
            
            For i = 0 To lngRecAffected - 1
            
                lngRow = mysql_fetch_row(lngMyRes)
                CopyMemory lngRowOffset, ByVal lngRow, 4

                lngLenghts = mysql_fetch_lengths(lngMyRes)
                CopyMemory arrLenghts(0), ByVal lngLenghts, 4 * lngFields
                
                For j = 0 To lngFields - 1
                    If arrLenghts(j) > 0 Then
                        ReDim tmpArr(arrLenghts(j) - 1) As Byte
                        Call CopyMemory(tmpArr(0), ByVal lngRowOffset, arrLenghts(j))
                        strData = StrConv(tmpArr(), vbUnicode)
                    Else
                        strData = vbNullString
                    End If
                    
                    lngRowOffset = lngRowOffset + arrLenghts(j) + 1
                    Call cMySQLRecord.SetValue(j, i, strData, arrLenghts(j))
                Next j
                
            Next i
            
            SQLQuery_Multi = lngRecAffected
            
        End If
        
        mysql_free_result (lngMyRes)
        
        lngMyRes = mysql_next_result(conn_stored)
        
        If lngMyRes <> 0 Then
            mysql_free_result (lngMyRes)
        End If
        
    End If
Else
    Call InternalMySQLError(strQuery, conn_stored)
    SQLQuery_Multi = -1
End If

End Function

Private Function ptr2str(ByVal ptr As Long) As String
On Error Resume Next
    If ptr = 0 Then Exit Function    'yeah..should never happen though
    Dim sout() As Byte
    Dim cchars As Long
     cchars = lstrlen(ptr)
     If cchars = 0 Then Exit Function
     ReDim sout(1 To cchars)                   'byte array
     sout = String$(cchars, " ")
     CopyMemory sout(1), ByVal (ptr), cchars
    ptr2str = StripNull(StrConv(sout, vbUnicode))  'convert byte array (lpstr) to unicode
End Function

Private Function StripNull(sName As String) As String
   Dim f As Long
   f = InStr(sName, vbNullChar)
   If f > 0 Then
    StripNull = Trim(left$(sName, f - 1))
   Else
    StripNull = sName
   End If
End Function

Private Sub InternalMySQLError(ByRef strQuery As String, ByVal curr_conn As Long)

Dim lngError As Long, strMessage As String

lngError = mysql_error(curr_conn)
strMessage = ptr2str(lngError)

Debug.Print "MySQL: " & strMessage

Call LogError(strMessage & " " & CStr(lngError) & " - " & strQuery)

If mysql_ping(curr_conn) <> 0 Then
    
    Call LogError("Database went down. Reconnecting. " & CStr(lngError) & " - " & strQuery)
    Call SQLReconnect
    Call SQLQuery_Simple_Sync(strQuery)
    
End If

End Sub

Private Sub SQLReconnect()

'Call ServerTCP.Send_Server_Message(Maintenance_Tasks, , , ToAll)
'Call ServerTCP.SendData(Main_Pause, ToAll, 0, 0)

Do While mysql_ping(conn_main) <> 0
    
    Call mysql_real_connect(conn_main, m_Hostname, m_UserName, m_password, m_Database, 0, vbNullString, 0)
    
    'Thread sleep
    Call Sleep(1)
    
    #If IDE = 1 Then
    DoEvents
    #End If
    
Loop

Do While mysql_ping(conn_exec) <> 0
    
    Call mysql_real_connect(conn_exec, m_Hostname, m_UserName, m_password, m_Database, 0, vbNullString, 0)
    
    'Thread sleep
    Call Sleep(1)
    
    #If IDE = 1 Then
    DoEvents
    #End If
    
Loop

Do While mysql_ping(conn_stored) <> 0
    
    Call mysql_real_connect(conn_stored, m_Hostname, m_UserName, m_password, m_Database, 0, vbNullString, CLIENT_MULTI_RESULTS)
    
    'Thread sleep
    Call Sleep(1)
    
    #If IDE = 1 Then
    DoEvents
    #End If
    
Loop

'Call ServerTCP.SendData(Main_Pause, ToAll, 0, 0)

End Sub

Public Function Get_Insert_ID() As Long

Get_Insert_ID = mysql_insert_id(conn_main)

End Function

Public Function Get_AffectedRows() As Long

Get_AffectedRows = mysql_affected_rows(conn_main)

End Function
