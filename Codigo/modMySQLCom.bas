Attribute VB_Name = "modMySQLCom"
' Api Mysql pour Visual Basic 6
' Traduction des fichiers mysql.h, mysql_com.h, and mysql_version.h
'
' Copyright (c) 2003 Romain Puyfoulhoux

'Ce programme est un logiciel libre ; vous pouvez le redistribuer et/ou
'le modifier conformément aux dispositions de la Licence Publique Générale GNU,
'telle que publiée par la Free Software Foundation ; version 2 de la licence,
'ou encore (à votre choix) toute version ultérieure.

'Ce programme est distribué dans l'espoir qu'il sera utile, mais
'SANS AUCUNE GARANTIE;sans même la garantie implicite de
'COMMERCIALISATION ou D'ADAPTATION A UN OBJET PARTICULIER.
'Pour plus de détail, voir la Licence Publique Générale GNU.

'Vous devez avoir reçu un exemplaire de la Licence Publique Générale GNU
'en même temps que ce programme ; si ce n'est pas le cas,
'écrivez à la Free Software Foundation
'Inc., 675 Mass Ave, Cambridge, MA 02139, Etats-Unis.

Option Explicit

Public Const NAME_LEN = 64      'Field/table Name length
Public Const HOSTNAME_LENGTH = 60
Public Const USERNAME_LENGTH = 16
Public Const SERVER_VERSION_LENGTH = 60

Public Const LOCAL_HOST = "localhost"
Public Const LOCAL_HOST_NAMEDPIPE = "."

Public Const MYSQL_NAMEDPIPE = "MySQL"
Public Const MYSQL_SERVICENAME = "MySql"

Enum enum_server_command
    COM_SLEEP
    COM_QUIT
    COM_INIT_DB
    COM_QUERY
    COM_FIELD_LIST
    COM_CREATE_DB
    COM_DROP_DB
    COM_REFRESH
    COM_SHUTDOWN
    COM_STATISTICS
    COM_PROCESS_INFO
    COM_CONNECT
    COM_PROCESS_KILL
    COM_DEBUG
    COM_PING
    COM_TIME
    COM_DELAYED_INSERT
    COM_CHANGE_USER
    COM_BINLOG_DUMP
    COM_TABLE_DUMP
    COM_CONNECT_OUT
End Enum

Public Const NOT_NULL_FLAG = 1       ' Field can't be NULL
Public Const PRI_KEY_FLAG = 2        ' Field is part of a primary key
Public Const UNIQUE_KEY_FLAG = 4     ' Field is part of a unique key
Public Const MULTIPLE_KEY_FLAG = 8   ' Field is part of a key
Public Const BLOB_FLAG = 16          ' Field is a blob
Public Const UNSIGNED_FLAG = 32      ' Field is unsigned
Public Const ZEROFILL_FLAG = 64      ' Field is zerofill
Public Const BINARY_FLAG = 128

'** The following are only sent to new clients **
Public Const ENUM_FLAG = 256             ' field is an enum
Public Const AUTO_INCREMENT_FLAG = 512   ' Field is a autoincrement field
Public Const TIMESTAMP_FLAG = 1024       ' Field is a timestamp
Public Const SET_FLAG = 2048             ' Field is a set
Public Const NUM_FLAG = 32768               ' Field is num (for clients)
Public Const PART_KEY_FLAG = 16384       ' Intern: Part of some key
Public Const GROUP_FLAG = 32768             ' Intern: Group field
Public Const UNIQUE_FLAG = 65536            ' Intern: Used by sql_yacc

Public Const REFRESH_GRANT = 1          ' Refresh grant tables
Public Const REFRESH_LOG = 2            ' Start on new log file
Public Const REFRESH_TABLES = 4         ' close all tables
Public Const REFRESH_HOSTS = 8          ' Flush host cache
Public Const REFRESH_STATUS = 16        ' Flush status variables
Public Const REFRESH_THREADS = 32       ' Flush status variables
Public Const REFRESH_SLAVE = 64         ' Reset master info and restart slave thread

Public Const REFRESH_MASTER = 128  ' Remove all bin logs in the index and truncate the index

' The following can't be set with mysql_refresh()
Public Const REFRESH_READ_LOCK = 16384      ' Lock tables for read
Public Const REFRESH_FAST = 32768           ' Intern flag

Public Const CLIENT_LONG_PASSWORD = 1       ' new more secure passwords
Public Const CLIENT_FOUND_ROWS = 2          ' Found instead of affected rows
Public Const CLIENT_LONG_FLAG = 4           ' Get all column flags
Public Const CLIENT_CONNECT_WITH_DB = 8     ' One can specify db on connect
Public Const CLIENT_NO_SCHEMA = 16          ' Don't allow database.table.column
Public Const CLIENT_COMPRESS = 32           ' Can use compression protocol
Public Const CLIENT_ODBC = 64               ' Odbc client
Public Const CLIENT_LOCAL_FILES = 128       ' Can use LOAD DATA LOCAL
Public Const CLIENT_IGNORE_SPACE = 256      ' Ignore spaces before '('
Public Const CLIENT_CHANGE_USER = 512       ' Support the mysql_change_user()
Public Const CLIENT_INTERACTIVE = 1024      ' This is an interactive client
Public Const CLIENT_SSL = 2048              ' Switch to SSL after handshake
Public Const CLIENT_IGNORE_SIGPIPE = 4096   ' IGNORE sigpipes
Public Const CLIENT_TRANSACTIONS = 8192     ' Client knows about transactions
Public Const CLIENT_MULTI_RESULTS = 131072     ' Client supports multi results

Public Const SERVER_STATUS_IN_TRANS = 1     ' Transaction has started
Public Const SERVER_STATUS_AUTOCOMMIT = 2   ' Server in auto_commit mode

Public Const MYSQL_ERRMSG_SIZE = 200
Public Const NET_READ_TIMEOUT = 30              ' Timeout on read
Public Const NET_WRITE_TIMEOUT = 60             ' Timeout on write
Public Const NET_WAIT_TIMEOUT = 8 * 60 * 60     ' Wait for new query

Type NET
    vio As Long
    fd As Long                 ' For Perl DBI/dbd
    fcntl As Integer
    buff As Integer: buff_end As Integer: write_pos As Integer: read_pos As Integer
    last_error As String * MYSQL_ERRMSG_SIZE
    last_errno As Long: max_packet As Long: timeout As Long: pkt_nr As Long
    error As Integer
    return_errno As Boolean: compress As Boolean
    no_send_ok As Boolean   'needed if we are doing several
                            'queries in one command ( as in LOAD TABLE ... FROM MASTER ),
                            'and do not want to confuse the client with OK at the wrong time
               
    remain_in_buf As Long: Length As Long: buf_length As Long: where_b As Long
    return_status As Long   'pointer to a long
    reading_or_writing As Integer
    save_char As Byte
End Type

Public Const packet_error As Long = -1

Enum enum_field_types
    FIELD_TYPE_DECIMAL
    FIELD_TYPE_TINY
    FIELD_TYPE_SHORT
    FIELD_TYPE_LONG
    FIELD_TYPE_FLOAT
    FIELD_TYPE_DOUBLE
    FIELD_TYPE_NULL
    FIELD_TYPE_TIMESTAMP
    FIELD_TYPE_LONGLONG
    FIELD_TYPE_INT24
    FIELD_TYPE_DATE
    FIELD_TYPE_TIME
    FIELD_TYPE_DATETIME
    FIELD_TYPE_YEAR
    FIELD_TYPE_NEWDATE
    FIELD_TYPE_ENUM = 247
    FIELD_TYPE_SET = 248
    FIELD_TYPE_TINY_BLOB = 249
    FIELD_TYPE_MEDIUM_BLOB = 250
    FIELD_TYPE_LONG_BLOB = 251
    FIELD_TYPE_BLOB = 252
    FIELD_TYPE_VAR_STRING = 253
    FIELD_TYPE_STRING = 254
End Enum

' For compability
Public Const FIELD_TYPE_CHAR = FIELD_TYPE_TINY
Public Const FIELD_TYPE_INTERVAL = FIELD_TYPE_ENUM

Public Const PROTOCOL_VERSION = 10
Public Const MYSQL_SERVER_VERSION = "3.23.49"
Public Const MYSQL_SERVER_SUFFIX = ""
Public Const FRM_VER = 6
Public Const MYSQL_VERSION_ID = 32349
Public Const MYSQL_PORT = 3306
Public Const MYSQL_UNIX_ADDR = "/tmp/mysql.sock"
