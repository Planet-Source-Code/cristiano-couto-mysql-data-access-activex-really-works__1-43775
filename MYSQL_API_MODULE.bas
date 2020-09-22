Attribute VB_Name = "MySQL_API_MODULE"
Public Declare Function lstrlen Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
'/****************************************************************************/'
Public Declare Function API_mysql_num_rows Lib "libmysql.dll" Alias "mysql_num_rows" (result As API_MYSQL_RES) As API_myulonglong 'don't forget to convert64
Public Declare Function API_mysql_num_fields Lib "libmysql.dll" Alias "mysql_num_fields" (result As API_MYSQL_RES) As Long
Public Declare Function API_mysql_eof Lib "libmysql.dll" Alias "mysql_eof" (result As API_MYSQL_RES) As Byte
Public Declare Function API_mysql_fetch_field_direct Lib "libmysql.dll" Alias "mysql_fetch_field_direct" (result As API_MYSQL_RES, fieldnr As Long) As Long
'/****************************************************************************/'
Public Declare Function API_mysql_fetch_fields Lib "libmysql.dll" Alias "mysql_fetch_fields" (result As API_MYSQL_RES) As Long
Public Declare Function API_mysql_row_tell Lib "libmysql.dll" Alias "mysql_row_tell" (result As API_MYSQL_RES) As Long
Public Declare Function API_mysql_field_tell Lib "libmysql.dll" Alias "mysql_field_tell" (result As API_MYSQL_RES) As Long
'/****************************************************************************/'
Public Declare Function API_mysql_field_count Lib "libmysql.dll" Alias "mysql_field_count" (TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_affected_rows Lib "libmysql.dll" Alias "mysql_affected_rows" (TMYSQL As API_MYSQL) As API_myulonglong 'don't forget to convert64
Public Declare Function API_mysql_insert_id Lib "libmysql.dll" Alias "mysql_insert_id" (TMYSQL As API_MYSQL) As API_myulonglong         'again..don't forget convert64
Public Declare Function API_mysql_errno Lib "libmysql.dll" Alias "mysql_errno" (TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_error Lib "libmysql.dll" Alias "mysql_error" (TMYSQL As API_MYSQL) As Long 'pointer to char *
'/****************************************************************************/'
Public Declare Function API_mysql_info Lib "libmysql.dll" Alias "mysql_info" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_thread_id Lib "libmysql.dll" Alias "mysql_thread_id" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_init Lib "libmysql.dll" Alias "mysql_init" (ByRef TMYSQL As API_MYSQL) As Long
'/****************************************************************************/'
Public Declare Function API_mysql_connect Lib "libmysql.dll" Alias "mysql_connect" (ByRef TMYSQL As API_MYSQL, ByVal host As Long, ByVal user As Long, ByVal Passwd As Long) As Long
Public Declare Function API_mysql_real_connect Lib "libmysql.dll" Alias "mysql_real_connect" (ByRef TMYSQL As API_MYSQL, ByVal host As Long, ByVal user As Long, ByVal Passwd As Long, ByVal DB As Long, ByVal Port As Long, ByVal Unix_Socket As Long, ByVal clientflag As Long) As Long
Public Declare Function API_mysql_close Lib "libmysql.dll" Alias "mysql_close" (TMYSQL As API_MYSQL) As Long
'/****************************************************************************/'
Public Declare Function API_mysql_select_db Lib "libmysql.dll" Alias "mysql_select_db" (ByRef TMYSQL As API_MYSQL, ByVal DB As Long) As Long
Public Declare Function API_mysql_query Lib "libmysql.dll" Alias "mysql_query" (ByRef TMYSQL As API_MYSQL, ByVal q As Long) As Long
Public Declare Function API_mysql_real_query Lib "libmysql.dll" Alias "mysql_real_query" (ByRef TMYSQL As API_MYSQL, ByVal q As Long, ByVal length As Long) As Long
Public Declare Function API_mysql_create_db Lib "libmysql.dll" Alias "mysql_create_db" (ByRef TMYSQL As API_MYSQL, ByVal DB As Long) As Long
Public Declare Function API_mysql_drop_db Lib "libmysql.dll" Alias "mysql_drop_db" (ByRef TMYSQL As API_MYSQL, ByVal DB As Long) As Long
Public Declare Function API_mysql_shutdown Lib "libmysql.dll" Alias "mysql_shutdown" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_dump_debug_info Lib "libmysql.dll" Alias "mysql_dump_debug_info" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_refresh Lib "libmysql.dll" Alias "mysql_refresh" (ByRef TMYSQL As API_MYSQL, ByVal refresh_options As Long) As Long
Public Declare Function API_mysql_kill Lib "libmysql.dll" Alias "mysql_kill" (ByRef TMYSQL As API_MYSQL, ByVal PID As Long) As Long
Public Declare Function API_mysql_ping Lib "libmysql.dll" Alias "mysql_ping" (ByRef TMYSQL As API_MYSQL) As Long
'/****************************************************************************/'
Public Declare Function API_mysql_stat Lib "libmysql.dll" Alias "mysql_stat" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_get_server_info Lib "libmysql.dll" Alias "mysql_get_server_info" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_get_client_info Lib "libmysql.dll" Alias "mysql_get_client_info" () As Long
Public Declare Function API_mysql_get_host_info Lib "libmysql.dll" Alias "mysql_get_host_info" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_get_proto_info Lib "libmysql.dll" Alias "mysql_get_proto_info" (ByRef TMYSQL As API_MYSQL) As Long
'/****************************************************************************/'
Public Declare Function API_mysql_list_dbs Lib "libmysql.dll" Alias "mysql_list_dbs" (ByRef TMYSQL As API_MYSQL, ByVal wild As Long) As Long
Public Declare Function API_mysql_list_tables Lib "libmysql.dll" Alias "mysql_list_tables" (ByRef TMYSQL As API_MYSQL, ByVal wild As Long) As Long
Public Declare Function API_mysql_list_fields Lib "libmysql.dll" Alias "mysql_list_fields" (ByRef TMYSQL As API_MYSQL, ByVal table As Long, ByVal wild As Long) As Long
Public Declare Function API_mysql_list_processes Lib "libmysql.dll" Alias "mysql_list_processes" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_store_result Lib "libmysql.dll" Alias "mysql_store_result" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_use_result Lib "libmysql.dll" Alias "mysql_use_result" (ByRef TMYSQL As API_MYSQL) As Long
Public Declare Function API_mysql_options Lib "libmysql.dll" Alias "mysql_options" (ByRef TMYSQL As API_MYSQL, TMYSQL_OPTION As API_st_mysql_options, ByVal arg As Long) As Long
Public Declare Function API_mysql_free_result Lib "libmysql.dll" Alias "mysql_free_result" (ByRef result As API_MYSQL_RES) As Long
Public Declare Function API_mysql_data_seek Lib "libmysql.dll" Alias "mysql_data_seek" (ByRef result As API_MYSQL_RES, ByVal offset As Double) As Long  'need byval 8 bytes
Public Declare Function API_mysql_row_seek Lib "libmysql.dll" Alias "mysql_row_seek" (ByRef result As API_MYSQL_RES, ByVal MYSQL_ROW_OFFSET As Long) As Long
Public Declare Function API_mysql_field_seek Lib "libmysql.dll" Alias "mysql_field_seek" (ByRef result As API_MYSQL_RES, offset)
Public Declare Function API_mysql_fetch_row Lib "libmysql.dll" Alias "mysql_fetch_row" (ByRef result As API_MYSQL_RES) As Long   'pointer to array
Public Declare Function API_mysql_fetch_lengths Lib "libmysql.dll" Alias "mysql_fetch_lengths" (result As API_MYSQL_RES) As Long 'returns * unsigned long
Public Declare Function API_mysql_fetch_field Lib "libmysql.dll" Alias "mysql_fetch_field" (result As API_MYSQL_RES) As Long
Public Declare Function API_mysql_escape_string Lib "libmysql.dll" Alias "mysql_escape_string" (ByRef TMYSQL As API_MYSQL, ByVal to_ As Long, ByVal from_ As Long, ByVal length As Long) As Long
Public Declare Function API_mysql_debug Lib "libmysql.dll" Alias "mysql_debug" (ByVal debug_ As Long) As Long
Public Declare Function API_mysql_thread_safe Lib "libmysql.dll" Alias "mysql_thread_safe" () As Long
'Public Declare Function API_mysql_character_set_name Lib "libmysql.dll" Alias "mysql_character_set_name" (ByRef TMYSQL As API_MYSQL) As Long
'Public Declare Function API_mysql_change_user Lib "libmysql.dll" Alias "mysql_change_user" (ByRef TMYSQL As API_MYSQL, ByVal user As Long) As Byte


'#################### TAKEN MYSQL.COM (description of API functions) ########################################
'mysql_affected_rows()  Returns the number of rows affected by the last UPDATE, DELETE, or INSERT query.
'mysql_close()  Closes a server connection.
'mysql_connect()  Connects to a MySQL server. This function is deprecated; use mysql_real_connect() instead.
'mysql_change_user()  Changes user and database on an open connection.
'mysql_character_set_name()  Returns the name of the default character set for the connection.
'mysql_create_db()  Creates a database. This function is deprecated; use the SQL command CREATE DATABASE instead.
'mysql_data_seek()  Seeks to an arbitrary row in a query result set.
'mysql_debug()  Does a DBUG_PUSH with the given string.
'mysql_drop_db()  Drops a database. This function is deprecated; use the SQL command DROP DATABASE instead.
'mysql_dump_debug_info()  Makes the server write debug information to the log.
'mysql_eof()  Determines whether or not the last row of a result set has been read. This function is deprecated; mysql_errno() or mysql_error() may be used instead.
'mysql_errno()  Returns the error number for the most recently invoked MySQL function.
'mysql_error()  Returns the error message for the most recently invoked MySQL function.
'mysql_real_escape_string()  Escapes special characters in a string for use in a SQL statement taking into account the current charset of the connection.
'mysql_escape_string()  Escapes special characters in a string for use in a SQL statement.
'mysql_fetch_field()  Returns the type of the next table field.
'mysql_fetch_field_direct()  Returns the type of a table field, given a field number.
'mysql_fetch_fields()  Returns an array of all field structures.
'mysql_fetch_lengths()  Returns the lengths of all columns in the current row.
'mysql_fetch_row()  Fetches the next row from the result set.
'mysql_field_seek()  Puts the column cursor on a specified column.
'mysql_field_count()  Returns the number of result columns for the most recent query.
'mysql_field_tell()  Returns the position of the field cursor used for the last mysql_fetch_field().
'mysql_free_result()  Frees memory used by a result set.
'mysql_get_client_info()  Returns client version information.
'mysql_get_host_info()  Returns a string describing the connection.
'mysql_get_proto_info()  Returns the protocol version used by the connection.
'mysql_get_server_info()  Returns the server version number.
'mysql_info()  Returns information about the most recently executed query.
'mysql_init()  Gets or initializes a MYSQL structure.
'mysql_insert_id()  Returns the ID generated for an AUTO_INCREMENT column by the previous query.
'mysql_kill()  Kills a given thread.
'mysql_list_dbs()  Returns database names matching a simple regular expression.
'mysql_list_fields()  Returns field names matching a simple regular expression.
'mysql_list_processes()  Returns a list of the current server threads.
'mysql_list_tables()  Returns table names matching a simple regular expression.
'mysql_num_fields()  Returns the number of columns in a result set.
'mysql_num_rows()  Returns the number of rows in a result set.
'mysql_options()  Sets connect options for mysql_connect().
'mysql_ping()  Checks whether or not the connection to the server is working, reconnecting as necessary.
'mysql_query()  Executes a SQL query specified as a null-terminated string.
'mysql_real_connect()  Connects to a MySQL server.
'mysql_real_query()  Executes a SQL query specified as a counted string.
'mysql_reload()  Tells the server to reload the grant tables.
'mysql_row_seek()  Seeks to a row in a result set, using value returned from mysql_row_tell().
'mysql_row_tell()  Returns the row cursor position.
'mysql_select_db()  Selects a database.
'mysql_shutdown()  Shuts down the database server.
'mysql_stat()  Returns the server status as a string.
'mysql_store_result()  Retrieves a complete result set to the client.
'mysql_thread_id()  Returns the current thread ID.
'mysql_thread_save()  Returns 1 if the clients are compiled as threadsafe.
'mysql_use_result()  Initiates a row-by-row result set retrieval.
'/****************************************************************************/'
'/*                        NOT EXPORTED BY DLL --
'/*                     If they are i need to look at the libmysql.dll again
'/****************************************************************************/'
'Public Declare Function API_mysql_real_escape_string Lib "libmysql.dll" Alias "mysql_real_escape_string" ()
'Public Declare Function API_mysql_ssl_set Lib "libmysql.dll" Alias "mysql_ssl_set" (ByRef TMYSQL As API_MYSQL, ByVal key As Long, ByVal cert As Long, ByVal ca As Long, ByVal capath As Long) As Long
'Public Declare Function API_mysql_ssl_cipher Lib "libmysql.dll" Alias "mysql_ssl_cipher" (ByRef TMYSQL As API_MYSQL) As Long
'Public Declare Function API_mysql_ssl_clear Lib "libmysql.dll" Alias "mysql_ssl_clear" (ByRef TMYSQL As API_MYSQL) As Long
'Public Declare Function API_mysql_odbc_escape_string Lib "libmysql.dll" Alias "mysql_odbc_escape_string" (ByRef TMYSQL As API_MYSQL, ByVal to_ As Long, ByVal to_length As Long, ByVal from_ As Long, ByVal from_length As Long,param as Any,byval extended_buffer as Long)
'Public Declare Function API_myodbc_remove_escape_string Lib "libmysql.dll" Alias "myodbc_remove_escape" (ByRef TMYSQL As API_MYSQL, ByVal name As Long) As Long
'/****************************************************************************/'
'----------------------------------------------------
'translated from mysql_com.h
'----------------------------------------------------
Public Const SIZE_OF_CHAR = 4
Public Const NAME_LEN = 64
Public Const HOSTNAME_LENGTH = 60
Public Const USERNAME_LENGTH = 16
Public Const LOCAL_HOST = "localhost"
Public Const LOCAL_HOST_NAMEDPIPE = "MySQL"
Public Const MYSQL_SERVICENAME = "MySql"
Public Enum enum_server_command
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
End Enum
Public Const NOT_NULL_FLAG = 1
Public Const PRI_KEY_FLAG = 2
Public Const UNIQUE_KEY_FLAG = 4
Public Const MULTIPLE_KEY_FLAG = 8
Public Const BLOB_FLAG = 16
Public Const UNSIGNED_FLAG = 32
Public Const ZEROFILL_FLAG = 64
Public Const BINARY_FLAG = 128
Public Const ENUM_FLAG = 256
Public Const AUTO_INCREMENT_FLAG = 512
Public Const TIMESTAMP_FLAG = 1024
Public Const SET_FLAG = 2048
Public Const PART_KEY_FLAG = 16384
Public Const GROUP_FLAG = 32768
Public Const UNIQUE_FLAG = 65536

Public Const REFRESH_GRANT = 1
Public Const REFRESH_LOG = 2
Public Const REFRESH_TABLES = 4
Public Const REFRESH_HOSTS = 8
Public Const REFRESH_STATUS = 16
Public Const REFRESH_THREADS = 32
Public Const REFRESH_SLAVE = 64
Public Const REFRESH_MASTER = 128
Public Const REFRESH_READ_LOCK = 256
Public Const REFRESH_FAST = 32768

Public Const CLIENT_LONG_PASSWORD = 1
Public Const CLIENT_FOUND_ROWS = 2
Public Const CLIENT_LONG_FLAG = 4
Public Const CLIENT_CONNECT_WITH_DB = 8
Public Const CLIENT_NO_SCHEMA = 16
Public Const CLIENT_COMPRESS = 32
Public Const CLIENT_ODBC = 64
Public Const CLIENT_LOCAL_FILES = 128
Public Const CLIENT_IGNORE_SPACE = 256
Public Const CLIENT_CHANGE_USER = 512
Public Const CLIENT_INTERACTIVE = 1024
Public Const CLIENT_SSL = 2048
Public Const CLIENT_IGNORE_SIGPIPE = 4096
Public Const CLIENT_TRANSACTIONS = 8196

Public Const SERVER_STATUS_IN_TRANS = 1
Public Const SERVER_STATUS_AUTOCOMMIT = 2

Public Const MYSQL_ERRMSG_SIZE = 200
Public Const NET_READ_TIMEOUT = 30
Public Const NET_WRITE_TIMEOUT = 60
Public Const NET_WAIT_TIMEOUT = 8 * 60 * 60

Public Const packet_error = -1

'Public Type vio
'End Type

'sizeof(NET)=272
Public Type API_NET
  vio As Long    'pointer->type vio
  fd As Long
  fcntl As Long
  buff As Long
  buff_end As Long
  write_pos As Long
  read_pos As Long
  last_error(1 To MYSQL_ERRMSG_SIZE) As Byte
  last_errno As Long
  max_packet As Long
  timeout As Long
  pkt_nr As Long
  error As Byte
  return_errno As Byte
  compress As Byte
  no_send_ok As Byte
  remain_in_buf As Long
  length As Long
  buf_length As Long
  where_b As Long
  return_status As Long
  reading_or_writing As Byte
  save_char As Byte
End Type
    ' adArray
    ' adBinary
    ' adBoolean
    ' adBSTR
    ' adChapter
    ' adChar
    ' adCurrency
    ' adDate
    ' adDBDate
    ' adDBTime
    ' adDBTimeStamp
    ' adEmpty
    ' adError
    ' adFileTime
    ' adGUID
    ' adIDispatch
    ' adSingle
    ' adIUnknown
    ' adLongVarBinary
    ' adLongVarChar
    ' adLongVarWChar
    ' adNumeric
    ' adPropVariant
    ' adSmallInt
    ' adUnsignedBigInt
    ' adUnsignedInt
    ' adUnsignedSmallInt
    ' adUnsignedTinyInt
    ' adVarBinary
    ' adVarChar
    ' adVariant
    ' adVarNumeric
    ' adVarWChar
    ' adWChar
Public Const FIELD_TYPE_CHAR = FIELD_TYPE_TINY
Public Const FIELD_TYPE_INTERVAL = FIELD_TYPE_ENUM

'----------------------------------------------------
'translated from mysql_version.h
'----------------------------------------------------
Public Const PROTOCOL_VERSION = 10
Public Const MYSQL_SERVER_VERSION = "3.23.33"
Public Const MYSQL_SERVER_SUFFIX = ""
Public Const FRM_VER = 6
Public Const MYSQL_VERSION_ID = 32333
Public Const MYSQL_PORT = 3306
Public Const MYSQL_UNIX_ADDR = "/tmp/mysql.sock"

'----------------------------------------------------
'translated from mysql.h
'----------------------------------------------------
'gptr is a long
'sizeof(USED_MEM)=12
Public Type API_myulonglong
   bytes(1 To 8) As Byte
End Type
Public Type API_USED_MEM
  next As Long
  Left As Long
  Size As Long
End Type
'sizeof(MEM_ROOT)=20
Public Type API_MEM_ROOT
  free As Long
  used As Long
  min_malloc As Long
  block_size As Long
  error_handler As Long 'pointer to an error handler
  'fix_mis_alignment where ever we don't land on a full word boundary (because this structure is 20 bytes in size)
End Type
'mysql_port is a long
'mysql_unix port is a long (pointer)
'sizeof(MYSQL_FIELD)=32
Public Type API_MYSQL_FIELD
  name As Long
  table As Long
  def As Long
  type As API_enum_field_types
  length As Long
  max_length As Long
  Flags As Long
  decimals As Long
End Type
'sizeof(mysql_options)=76
Public Type API_st_mysql_options
  connect_timeout As Long
  client_flag As Long
  compress As Byte
  named_pipe As Byte
  Port As Long
  host As Long
  init_command As Long
  user As Long
  Password As Long
  Unix_Socket As Long
  DB As Long
  my_cnf_file As Long
  my_cnf_group As Long
  charset_dir As Long
  charset_name As Long
  use_ssl As Byte               '/* if to use SSL or not */
  ssl_key As Long               '/* PEM key file */
  ssl_cert As Long              '/* PEM cert file */
  ssl_ca As Long                '/* PEM CA file */
  ssl_capath As Long            '/* PEM directory of CA-s? */
End Type
Public Enum API_mysql_option
        MYSQL_OPT_CONNECT_TIMEOUT
        MYSQL_OPT_COMPRESS
        MYSQL_OPT_NAMED_PIPE
        MYSQL_INIT_COMMAND
        MYSQL_READ_DEFAULT_FILE
        MYSQL_READ_DEFAULT_GROUP
        MYSQL_SET_CHARSET_DIR
        MYSQL_SET_CHARSET_NAME
End Enum
'sizeof(mysql_status)=4
Public Enum API_mysql_status
        MYSQL_STATUS_READY
        MYSQL_STATUS_GET_RESULT
        MYSQL_STATUS_USE_RESULT
End Enum
'sizeof(MYSQL)=496
Public Type API_MYSQL
  net_a As API_NET
  connector_fd As Long
  host As Long
  user As Long
  Passwd As Long
  Unix_Socket As Long
  server_version As Long
  host_info As Long
  info As Long
  DB As Long
  Port As Long
  client_flag As Long
  server_capabilities As Long
  protocol_ver As Long
  field_count As Long
  server_status As Long
  thread_id As Long
  affected_rows As API_myulonglong
  Insert_ID As API_myulonglong
  extra_info As API_myulonglong
  packet_length As Long
  status As API_mysql_status
  Fields As Long
  field_alloc As API_MEM_ROOT
  'we are 4 bytes short cause of mal-aligned mem_root...add 4
  FIX_MISALIGNMENT As Long
  free_me As Byte
  reconnect As Byte
  options As API_st_mysql_options
  scramble_buff(1 To 9) As Byte
  charset As Long
  server_language As Long
End Type
'sizeof(MYSQL_DATA)=40
Public Type API_MYSQL_DATA
  Rows As API_myulonglong
  Fields As Long
  Data As Long
  alloc As API_MEM_ROOT
  'again we seem to be 4 bytes short caused by mem_root mis-algined..add 4
  FIX_MISALIGNMENT As Long
End Type
'sizeof(MYSQL_ROWS)=8
Public Type API_MYSQL_ROWS
  next As Long
  Data As Long
End Type
'sizeof(MYSQL_RES)=72
Public Type API_MYSQL_RES
   row_count As API_myulonglong
   field_count As Long
   current_field As Long
   Fields As Long
   Data As Long
   data_cursor As Long
   field_alloc As API_MEM_ROOT
   'yet again we are not landing on those full word boundaries..add 4 :)
   FIX_MISALIGNMENT As Long
   row As Long
   current_row As Long
   lengths As Long
   handle As Long
   eof As Byte
End Type

'#####################################################################################
'#                                   MYSQL ERRORS AND DESCRIPTORS                                                     #
'#####################################################################################
Public Const CR_UNKNOWN_ERROR = 2000
Public Const CR_SOCKET_CREATE_ERROR = 2001
Public Const CR_CONNECTION_ERROR = 2002
Public Const CR_CONN_HOST_ERROR = 2003
Public Const CR_IPSOCK_ERROR = 2004
Public Const CR_UNKNOWN_HOST = 2005
Public Const CR_SERVER_GONE_ERROR = 2006
Public Const CR_VERSION_ERROR = 2007
Public Const CR_OUT_OF_MEMORY = 2008
Public Const CR_WRONG_HOST_INFO = 2009
Public Const CR_LOCALHOST_CONNECTION = 2010
Public Const CR_TCP_CONNECTION = 2011
Public Const CR_SERVER_HANDSHAKE_ERR = 2012
Public Const CR_SERVER_LOST = 2013
Public Const CR_COMMANDS_OUT_OF_SYNC = 2014
Public Const CR_NAMEDPIPE_CONNECTION = 2015
Public Const CR_NAMEDPIPEWAIT_ERROR = 2016
Public Const CR_NAMEDPIPEOPEN_ERROR = 2017
Public Const CR_NAMEDPIPESETSTATE_ERROR = 2018

Public Const ER_HASHCHK = 1000
Public Const ER_NISAMCHK = 1001
Public Const ER_NO = 1002
Public Const ER_YES = 1003
Public Const ER_CANT_CREATE_FILE = 1004
Public Const ER_CANT_CREATE_TABLE = 1005
Public Const ER_CANT_CREATE_DB = 1006
Public Const ER_DB_CREATE_EXISTS = 1007
Public Const ER_DB_DROP_EXISTS = 1008
Public Const ER_DB_DROP_DELETE = 1009
Public Const ER_DB_DROP_RMDIR = 1010
Public Const ER_CANT_DELETE_FILE = 1011
Public Const ER_CANT_FIND_SYSTEM_REC = 1012
Public Const ER_CANT_GET_STAT = 1013
Public Const ER_CANT_GET_WD = 1014
Public Const ER_CANT_LOCK = 1015
Public Const ER_CANT_OPEN_FILE = 1016
Public Const ER_FILE_NOT_FOUND = 1017
Public Const ER_CANT_READ_DIR = 1018
Public Const ER_CANT_SET_WD = 1019
Public Const ER_CHECKREAD = 1020
Public Const ER_DISK_FULL = 1021
Public Const ER_DUP_KEY = 1022
Public Const ER_ERROR_ON_CLOSE = 1023
Public Const ER_ERROR_ON_READ = 1024
Public Const ER_ERROR_ON_RENAME = 1025
Public Const ER_ERROR_ON_WRITE = 1026
Public Const ER_FILE_USED = 1027
Public Const ER_FILSORT_ABORT = 1028
Public Const ER_FORM_NOT_FOUND = 1029
Public Const ER_GET_ERRNO = 1030
Public Const ER_ILLEGAL_HA = 1031
Public Const ER_KEY_NOT_FOUND = 1032
Public Const ER_NOT_FORM_FILE = 1033
Public Const ER_NOT_KEYFILE = 1034
Public Const ER_OLD_KEYFILE = 1035
Public Const ER_OPEN_AS_READONLY = 1036
Public Const ER_OUTOFMEMORY = 1037
Public Const ER_OUT_OF_SORTMEMORY = 1038
Public Const ER_UNEXPECTED_EOF = 1039
Public Const ER_CON_COUNT_ERROR = 1040
Public Const ER_OUT_OF_RESOURCES = 1041
Public Const ER_BAD_HOST_ERROR = 1042
Public Const ER_HANDSHAKE_ERROR = 1043
Public Const ER_DBACCESS_DENIED_ERROR = 1044
Public Const ER_ACCESS_DENIED_ERROR = 1045
Public Const ER_NO_DB_ERROR = 1046
Public Const ER_UNKNOWN_COM_ERROR = 1047
Public Const ER_BAD_NULL_ERROR = 1048
Public Const ER_BAD_DB_ERROR = 1049
Public Const ER_TABLE_EXISTS_ERROR = 1050
Public Const ER_BAD_TABLE_ERROR = 1051
Public Const ER_NON_UNIQ_ERROR = 1052
Public Const ER_SERVER_SHUTDOWN = 1053
Public Const ER_BAD_FIELD_ERROR = 1054
Public Const ER_WRONG_FIELD_WITH_GROUP = 1055
Public Const ER_WRONG_GROUP_FIELD = 1056
Public Const ER_WRONG_SUM_SELECT = 1057
Public Const ER_WRONG_VALUE_COUNT = 1058
Public Const ER_TOO_LONG_IDENT = 1059
Public Const ER_DUP_FIELDNAME = 1060
Public Const ER_DUP_KEYNAME = 1061
Public Const ER_DUP_ENTRY = 1062
Public Const ER_WRONG_FIELD_SPEC = 1063
Public Const ER_PARSE_ERROR = 1064
Public Const ER_EMPTY_QUERY = 1065
Public Const ER_NONUNIQ_TABLE = 1066
Public Const ER_INVALID_DEFAULT = 1067
Public Const ER_MULTIPLE_PRI_KEY = 1068
Public Const ER_TOO_MANY_KEYS = 1069
Public Const ER_TOO_MANY_KEY_PARTS = 1070
Public Const ER_TOO_LONG_KEY = 1071
Public Const ER_KEY_COLUMN_DOES_NOT_EXITS = 1072
Public Const ER_BLOB_USED_AS_KEY = 1073
Public Const ER_TOO_BIG_FIELDLENGTH = 1074
Public Const ER_WRONG_AUTO_KEY = 1075
Public Const ER_READY = 1076
Public Const ER_NORMAL_SHUTDOWN = 1077
Public Const ER_GOT_SIGNAL = 1078
Public Const ER_SHUTDOWN_COMPLETE = 1079
Public Const ER_FORCING_CLOSE = 1080
Public Const ER_IPSOCK_ERROR = 1081
Public Const ER_NO_SUCH_INDEX = 1082
Public Const ER_WRONG_FIELD_TERMINATORS = 1083
Public Const ER_BLOBS_AND_NO_TERMINATED = 1084
Public Const ER_TEXTFILE_NOT_READABLE = 1085
Public Const ER_FILE_EXISTS_ERROR = 1086
Public Const ER_LOAD_INFO = 1087
Public Const ER_ALTER_INFO = 1088
Public Const ER_WRONG_SUB_KEY = 1089
Public Const ER_CANT_REMOVE_ALL_FIELDS = 1090
Public Const ER_CANT_DROP_FIELD_OR_KEY = 1091
Public Const ER_INSERT_INFO = 1092
Public Const ER_INSERT_TABLE_USED = 1093
Public Const ER_NO_SUCH_THREAD = 1094
Public Const ER_KILL_DENIED_ERROR = 1095
Public Const ER_NO_TABLES_USED = 1096
Public Const ER_TOO_BIG_SET = 1097
Public Const ER_NO_UNIQUE_LOGFILE = 1098
Public Const ER_TABLE_NOT_LOCKED_FOR_WRITE = 1099
Public Const ER_TABLE_NOT_LOCKED = 1100
Public Const ER_BLOB_CANT_HAVE_DEFAULT = 1101
Public Const ER_WRONG_DB_NAME = 1102
Public Const ER_WRONG_TABLE_NAME = 1103
Public Const ER_TOO_BIG_SELECT = 1104
Public Const ER_UNKNOWN_ERROR = 1105
Public Const ER_UNKNOWN_PROCEDURE = 1106
Public Const ER_WRONG_PARAMCOUNT_TO_PROCEDURE = 1107
Public Const ER_WRONG_PARAMETERS_TO_PROCEDURE = 1108
Public Const ER_UNKNOWN_TABLE = 1109
Public Const ER_FIELD_SPECIFIED_TWICE = 1110
Public Const ER_INVALID_GROUP_FUNC_USE = 1111
Public Const ER_UNSUPPORTED_EXTENSION = 1112
Public Const ER_TABLE_MUST_HAVE_COLUMNS = 1113
Public Const ER_RECORD_FILE_FULL = 1114
Public Const ER_UNKNOWN_CHARACTER_SET = 1115
Public Const ER_TOO_MANY_TABLES = 1116
Public Const ER_TOO_MANY_FIELDS = 1117
Public Const ER_TOO_BIG_ROWSIZE = 1118
Public Const ER_STACK_OVERRUN = 1119
Public Const ER_WRONG_OUTER_JOIN = 1120
Public Const ER_NULL_COLUMN_IN_INDEX = 1121
Public Const ER_CANT_FIND_UDF = 1122
Public Const ER_CANT_INITIALIZE_UDF = 1123
Public Const ER_UDF_NO_PATHS = 1124
Public Const ER_UDF_EXISTS = 1125
Public Const ER_CANT_OPEN_LIBRARY = 1126
Public Const ER_CANT_FIND_DL_ENTRY = 1127
Public Const ER_FUNCTION_NOT_DEFINED = 1128
Public Const ER_HOST_IS_BLOCKED = 1129
Public Const ER_HOST_NOT_PRIVILEGED = 1130
Public Const ER_PASSWORD_ANONYMOUS_USER = 1131
Public Const ER_PASSWORD_NOT_ALLOWED = 1132
Public Const ER_PASSWORD_NO_MATCH = 1133
Public Const ER_UPDATE_INFO = 1134
Public Const ER_CANT_CREATE_THREAD = 1135
Public Const ER_WRONG_VALUE_COUNT_ON_ROW = 1136
Public Const ER_CANT_REOPEN_TABLE = 1137
Public Const ER_INVALID_USE_OF_NULL = 1138
Public Const ER_REGEXP_ERROR = 1139
Public Const ER_MIX_OF_GROUP_FUNC_AND_FIELDS = 1140
Public Const ER_NONEXISTING_GRANT = 1141
Public Const ER_TABLEACCESS_DENIED_ERROR = 1142
Public Const ER_COLUMNACCESS_DENIED_ERROR = 1143
Public Const ER_ILLEGAL_GRANT_FOR_TABLE = 1144
Public Const ER_GRANT_WRONG_HOST_OR_USER = 1145
Public Const ER_NO_SUCH_TABLE = 1146
Public Const ER_NONEXISTING_TABLE_GRANT = 1147
Public Const ER_NOT_ALLOWED_COMMAND = 1148
Public Const ER_SYNTAX_ERROR = 1149
Public Const ER_DELAYED_CANT_CHANGE_LOCK = 1150
Public Const ER_TOO_MANY_DELAYED_THREADS = 1151
Public Const ER_ABORTING_CONNECTION = 1152
Public Const ER_NET_PACKET_TOO_LARGE = 1153
Public Const ER_NET_READ_ERROR_FROM_PIPE = 1154
Public Const ER_NET_FCNTL_ERROR = 1155
Public Const ER_NET_PACKETS_OUT_OF_ORDER = 1156
Public Const ER_NET_UNCOMPRESS_ERROR = 1157
Public Const ER_NET_READ_ERROR = 1158
Public Const ER_NET_READ_INTERRUPTED = 1159
Public Const ER_NET_ERROR_ON_WRITE = 1160
Public Const ER_NET_WRITE_INTERRUPTED = 1161
Public Const ER_TOO_LONG_STRING = 1162
Public Const ER_TABLE_CANT_HANDLE_BLOB = 1163
Public Const ER_TABLE_CANT_HANDLE_AUTO_INCREMENT = 1164
Public Const ER_DELAYED_INSERT_TABLE_LOCKED = 1165
Public Const ER_WRONG_COLUMN_NAME = 1166
Public Const ER_WRONG_KEY_COLUMN = 1167
Public Const ER_WRONG_MRG_TABLE = 1168
Public Const ER_DUP_UNIQUE = 1169
Public Const ER_BLOB_KEY_WITHOUT_LENGTH = 1170
Public Const ER_PRIMARY_CANT_HAVE_NULL = 1171
Public Const ER_TOO_MANY_ROWS = 1172
Public Const ER_REQUIRES_PRIMARY_KEY = 1173
Public Const ER_NO_RAID_COMPILED = 1174
Public Const ER_ERROR_MESSAGES = 175
Public Function convert642l(a As API_myulonglong) As Long
Dim res As Long
  CopyMemory res, a.bytes(1), 4
  convert642l = res  'return our result from computations
End Function

Public Function convertl264(l2convert As Long) As API_myulonglong
Dim a As API_myulonglong
  CopyMemory a.bytes(1), l2convert, 4
  convertl264 = a
End Function
Public Function ptr2str(ByVal ptr As Long) As String
On Error Resume Next
    If ptr = 0 Then Exit Function    'yeah..should never happen though
    Dim test As Byte
    Dim sout() As Byte
    Dim cchars As Long
     cchars = lstrlen(ptr)
     If cchars = 0 Then Exit Function
     ReDim sout(1 To cchars)                   'byte array
     sout = String$(cchars, " ")
     CopyMemory sout(1), ByVal (ptr), cchars
    ptr2str = StripNull(StrConv(sout, vbUnicode))  'convert byte array (lpstr) to unicode
End Function
Public Function StripNull(sName As String) As String
   Dim f As Long
   f = InStr(sName, vbNullChar)
   If f > 0 Then
    StripNull = Trim(Left$(sName, f - 1))
   Else
    StripNull = sName
   End If
End Function

