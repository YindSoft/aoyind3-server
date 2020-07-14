Attribute VB_Name = "modMySQL"
Option Explicit
Public MySQL_Host As String
Public MySQL_User As String
Public MySQL_Pass As String
Public MySQL_DB As String

Type USED_MEM
    next As Long            'pointer to a USED_MEM structure
    left As Long            'memory left in block
    size As Long            'size of block
End Type

Type MEM_ROOT
    free As Long            'pointer to a USED_MEM structure
    used As Long            'pointer to a USED_MEM structure
    pre_alloc As Long       'pointer to a USED_MEM structure
    min_malloc As Long
    block_size As Long
    Error_Handler As Long
End Type

Type MYSQL_FIELD
    Name As Long                  'Pointer to Name of column
    table As Long                 'Table of column if column was a field
    def As Long                   'Default value (set by mysql_list_fields)
    type As enum_field_types      'Type of field. Se mysql_com.h for types
    Length As Long                'Width of column
    max_length As Long            'Max width of selected set
    flags As Long                 'Div flags
    decimals As Long              'Number of decimals in field
End Type

Type MYSQL_ROWS
    next As Long            'pointer to a MYSQL_ROWS structure
    data As Long
End Type

Type MYSQL_DATA
    rows As Long
    fields As Long
    data As Long              'pointer to a MYSQL_ROWS structure
    alloc As MEM_ROOT
End Type

Type mysql_options
    connect_timeout As Long: client_flag As Long
    compress As Boolean: named_pipe As Boolean
    port As Long
    host As Long: init_command As Long: user As Long
    password As Long: unix_socket As Long: DB As Long
    my_cnf_file As Long: my_cnf_group As Long: charset_dir As Long: charset_Name As Long
    use_ssl As Boolean            'if to use SSL or not
    ssl_key As Long               'pointer to PEM key file
    ssl_cert As Long              'PEM cert file
    ssl_ca As Long                'PEM CA file
    ssl_capath As Long            'PEM directory of CA-s?
End Type

Enum mysql_option
    MYSQL_OPT_CONNECT_TIMEOUT
    MYSQL_OPT_COMPRESS
    MYSQL_OPT_NAMED_PIPE
    MYSQL_INIT_COMMAND
    MYSQL_READ_DEFAULT_FILE
    YSQL_READ_DEFAULT_GROUP
    MYSQL_SET_CHARSET_DIR
    MYSQL_SET_CHARSET_NAME
    MYSQL_OPT_LOCAL_INFILE
End Enum

Enum mysql_status
    MYSQL_STATUS_READY
    MYSQL_STATUS_GET_RESULT
    MYSQL_STATUS_USE_RESULT
End Enum

Type mysql_main
    net_param As NET            'Communication parameters
    connector_fd As Long      'ConnectorFd for SSL
    host As Long: user As Long: passwd As Long: unix_socket As Long
    server_version As Long: host_info As Long: info As Long: DB As Long
    port As Long: client_flag As Long: server_capabilities As Long
    PROTOCOL_VERSION As Long
    field_count As Long
    server_status As Long
    thread_id As Long         'Id for connection in server
    affected_rows As Long
    insert_id As Long         'id if insert on table with NEXTNR
    extra_info As Long        'Used by mysqlshow
    packet_length As Long
    Status As mysql_status
    fields As Long              'pointer to a MYSQL_FIELD structure
    field_alloc As MEM_ROOT
    free_me As Boolean          'If free in mysql_close
    reconnect As Boolean        'set to 1 if automatic reconnect
    options As mysql_options
    scramble_buff As String * 9
    charset As Long             'pointer to a CHARSET_INFO structure
    server_language As Long
End Type


Type MYSQL_RES
    row_count As Long
    field_count As Long: current_field As Long
    fields As Long                  'pointer to a MYSQL_FIELD structure
    data As Long                    'pointer to a MYSQL_DATA structure
    data_cursor As Long             'pointer to a MYSQL_ROWS structure
    field_alloc As MEM_ROOT
    row As Long                     'If unbuffered read */
    current_row As Long             'buffer to current row
    lengths As Long                 'pointer to a long : column lengths of current row
    handle As Long                  'for unbuffered reads (pointer to a MYSQL structure)
    EOF As Boolean                  'Used my mysql_fetch_row
End Type

'Functions to get information from the MYSQL and MYSQL_RES structures
'Should definitely be used if one uses shared libraries

'res : pointer to a MYSQL_RES structure
Public Declare Function mysql_num_rows Lib "libmySQL.dll" (ByVal res As Long) As Long
Public Declare Function mysql_num_fields Lib "libmySQL.dll" (ByVal res As Long) As Long
Public Declare Function mysql_eof Lib "libmySQL.dll" (ByVal res As Long) As Boolean

'the 2 following functions return a pointer to a MYSQL_FIELD structure
Public Declare Function mysql_fetch_field_direct Lib "libmySQL.dll" (ByVal res As Long, ByVal fieldnr As Long) As Long
Public Declare Function mysql_fetch_fields Lib "libmySQL.dll" (ByVal res As Long) As Long

'mysql_row_tell() returns a pointer to a MYSQL_ROWS structure
Public Declare Function mysql_row_tell Lib "libmySQL.dll" (ByVal res As Long) As Long
Public Declare Function mysql_field_tell Lib "libmySQL.dll" (ByVal res As Long) As Long

'pMysql is a pointer to a MYSQL structure
Public Declare Function mysql_field_count Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_affected_rows Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_insert_id Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_errno Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_error Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_info Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_thread_id Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_character_set_Name Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long

'mysql_init() returns a pointer to a MYSQL structure
Public Declare Function mysql_init Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long

'*** If you have openssl ***

'Public declare Function mysql_ssl_set Lib "libmysql.dll" (mysql As Long, key As String, _
'                                    cert As String, ca As string, capath As String) As Integer
'Public declare Function mysql_ssl_cipher Lib "libmysql.dll" (mysql As Long) As Long
'Public declare Function mysql_ssl_clear Lib "libmysql.dll" (mysql As Long) As Integer
        
'***   End of openssl    ***


Public Declare Function mysql_change_user Lib "libmySQL.dll" (ByVal pMySQL As Long, _
                                        ByVal user As String, _
                                        ByVal passwd As String, ByVal DB As String) As Boolean
                                                              
'mysql_connect() and mysql_real_connect() return a pointer to a MYSQL structure

Public Declare Function mysql_connect Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal host As String, _
                                                   ByVal user As String, ByVal passwd As String) As Long

Public Declare Function mysql_real_connect Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal host As String, _
                       ByVal user As String, ByVal passwd As String, ByVal DB As String, _
                       ByVal port As Integer, ByVal unix_socket As String, ByVal clientflag As Long) As Long

Declare Sub mysql_close Lib "libmySQL.dll" (ByVal pMySQL As Long)
Public Declare Function mysql_select_db Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal DB As String) As Integer
Public Declare Function mysql_query Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal q As String) As Integer
Public Declare Function mysql_send_query Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal q As String, _
                        ByVal Length As Long) As Integer
Public Declare Function mysql_read_query_result Lib "libmySQL.dll" (ByVal pMySQL As Long) As Integer
Public Declare Function mysql_real_query Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal q As String, _
                                                      ByVal Length As Long) As Integer
Public Declare Function mysql_create_db Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal DB As String) As Integer
Public Declare Function mysql_drop_db Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal DB As String) As Integer
Public Declare Function mysql_shutdown Lib "libmySQL.dll" (ByVal pMySQL As Long) As Integer
Public Declare Function mysql_dump_debug_info Lib "libmySQL.dll" (ByVal pMySQL As Long) As Integer
Public Declare Function mysql_refresh Lib "libmySQL.dll" (ByVal pMySQL As Long, _
                                                   ByVal refresh_options As Long) As Integer
Public Declare Function mysql_kill Lib "libmySQL.dll" (ByVal pMySQL As Long, pid As Long) As Integer
Public Declare Function mysql_ping Lib "libmySQL.dll" (ByVal pMySQL As Long) As Integer
Public Declare Function mysql_stat Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_get_server_info Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_get_client_info Lib "libmySQL.dll" () As Long
Public Declare Function mysql_get_host_info Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_get_proto_info Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long

'these 6 functions return a pointer to a MYSQL_RES structure
Public Declare Function mysql_list_dbs Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal wild As String) As Long
Public Declare Function mysql_list_tables Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal wild As String) As Long
Public Declare Function mysql_list_fields Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal table As String, _
                                                       ByVal wild As String) As Long
Public Declare Function mysql_list_processes Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_store_result Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_next_result Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long
Public Declare Function mysql_use_result Lib "libmySQL.dll" (ByVal pMySQL As Long) As Long

Public Declare Function mysql_options Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal opt As mysql_option, _
                                                   ByVal arg As String) As Integer

'result is a pointer to a MYSQL_RES structure
Declare Sub mysql_free_result Lib "libmySQL.dll" (ByVal result As Long)
Declare Sub mysql_data_seek Lib "libmySQL.dll" (ByVal result As Long, ByVal offset As Long)
 
'mysql_row_seek() returns a pointer to a MYSQL_ROWS stucture. offset is pointer to a MYSQL_ROWS structure
Public Declare Function mysql_row_seek Lib "libmySQL.dll" (ByVal result As Long, ByVal offset As Long) As Long

Public Declare Function mysql_field_seek Lib "libmySQL.dll" (ByVal result As Long, ByVal offset As Long) As Long
Public Declare Function mysql_fetch_row Lib "libmySQL.dll" (ByVal result As Long) As Long
'returns a pointer to a long
Public Declare Function mysql_fetch_lengths Lib "libmySQL.dll" (ByVal result As Long) As Long
'returns a pointer to a MYSQL_FIELD structure
Public Declare Function mysql_fetch_field Lib "libmySQL.dll" (ByVal result As Long) As Long
Public Declare Function mysql_escape_string Lib "libmySQL.dll" (ByVal strTo As String, ByVal from As String, _
                                                         ByVal from_length As Long) As Long
Public Declare Function mysql_real_escape_string Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal strTo As String, _
                                                 ByVal from As String, ByVal Length As Long) As Long
Declare Sub mysql_debug Lib "libmySQL.dll" (ByVal todebug As String)
'extend_buffer is a pointer to a function: nomFunction(long,long,string,long) as long
Public Declare Function mysql_odbc_escape_string Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal strTo As Long, _
                                                 ByVal to_length As Long, ByVal from As Long, _
                                                 ByVal from_length As Long, ByVal param As Long, _
                                                 ByVal extend_buffer As Long) As Long
Declare Sub myodbc_remove_escape Lib "libmySQL.dll" (ByVal pMySQL As Long, ByVal Name As String)
Public Declare Function mysql_thread_safe Lib "libmySQL.dll" () As Long
  
Public Function mysql_reload(pMySQL As Long) As Integer
mysql_reload = mysql_refresh(pMySQL, REFRESH_GRANT)
End Function

Public Function IS_PRI_KEY(N As Long) As Boolean
IS_PRI_KEY = ((N And PRI_KEY_FLAG) = PRI_KEY_FLAG)
End Function

Public Function IS_NOT_NULL(N As Long) As Boolean
IS_NOT_NULL = ((N And NOT_NULL_FLAG) = NOT_NULL_FLAG)
End Function

Public Function IS_BLOB(N As Long) As Boolean
IS_BLOB = ((N And BLOB_FLAG) = BLOB_FLAG)
End Function

Public Function IS_NUM(t As Long) As Boolean
IS_NUM = (t <= FIELD_TYPE_INT24) Or (t = FIELD_TYPE_YEAR)
End Function

Public Function IS_NUM_FIELD(f As MYSQL_FIELD) As Boolean
IS_NUM_FIELD = ((f.flags And NUM_FLAG) = NUM_FLAG)
End Function

Public Function INTERNAL_NUM_FIELD(f As MYSQL_FIELD) As Boolean
INTERNAL_NUM_FIELD = ((f.type <= FIELD_TYPE_INT24) And _
            (f.type <> FIELD_TYPE_TIMESTAMP Or f.Length = 14 Or f.Length = 8 Or f.type = FIELD_TYPE_YEAR))
End Function



Public Sub Execute(Query As String)
Call mySQL.SQLQuery_Simple_Sync(Query)
End Sub
Public Function GetByCampo(ByVal Query As String, ByVal Campo As String)
GetByCampo = mySQL.SQLQuery_Fast(Query, Campo)
End Function
Public Function Comillas(Valor As String) As String
Comillas = "'" & Replace(Valor, "'", "\'") & "'"
End Function

Function CIntNull(ByVal Var) As Integer
If Var = "" Then
    CIntNull = 0
Else
    CIntNull = Var
End If
End Function

Function CLngNull(ByVal Var) As Long
If Var = "" Then
    CLngNull = 0
Else
    CLngNull = Var
End If
End Function

Function CStrNull(ByVal Var) As String
If IsNull(Var) Then
    CStrNull = ""
Else
    CStrNull = Var
End If
End Function
