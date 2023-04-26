'------------------------------------------------------------------------------
' TRM.INC - Tsunami 记录管理器包含文件 - TRM.DLL
'------------------------------------------------------------------------------
%CASE_SENSITIVE = 1
%NO_DUPLICATES  = 2
%NO_COMPRESSION = 4
%BINARY_KEY     = 8
%KEY_ONLY       = 4096
%REBUILD_IDS    = 2
'------------------------------------------------------------------------------
' ADDED: Jan 18, 2004 - By Raymond L. King (Custom Software Designers).
' Tsunami 2.5 - Returned Error Code Constants.
'------------------------------------------------------------------------------
%TRM_SUCCESSFUL     = 0		' 无错误 SUCCESSFUL.
%TRM_NOTTSUNAMIFILE = 1		' 非 Tsunami 文件.
%TRM_IOERROR        = 2		' I/O 错误.
%TRM_FILENOTOPEN    = 3		' 文件未打开.
%TRM_KEYNOTFOUND    = 4		' 键未找到.
%TRM_DUPLICATEKEY   = 5		' 重复键.
%TRM_INVALIDKEY     = 6		' 无效键.
%TRM_FILECORRUPT    = 7		' 文件崩溃.
%TRM_NOCURRENTPOS   = 8		' 无当前位置.
%TRM_ENDOFFILE      = 9		' 文件末尾.
%TRM_INVPAGESIZE    = 10	' 无效的页大小.
%TRM_INVNOKEYSEGS   = 11	' 无效的键区值.
%TRM_INVFILEDEFSTR  = 12	' 无效文件定义字符串.
%TRM_INVKEYSEGPOS   = 13	' 无效键区位置.
%TRM_INVKEYSEGLEN   = 14	' 无效键区长度.
%TRM_INCKEYSEGDEF   = 15	' 歧义键区定义.
%TRM_INVRECORDLEN   = 20	' 无效记录长度.
%TRM_INVRECORDADDR  = 21	' 无效记录地址.
%TRM_LOSTRECORDPOS  = 22	' 记录位置丢失.
%TRM_ACCESSDENIED   = 30	' 拒绝访问.
%TRM_FILEEXISTS     = 31	' 文件已存在.
%TRM_NOMOREFILEHDLS = 32	' 无更多文件句柄.
%TRM_MAXFILESOPEN   = 33	' 已打开最多文件.
%TRM_ACCELACCDENIED = 40	' 拒绝快速访问.
%TRM_ACCELCACHEERR  = 41	' 快捷缓存错误.
%TRM_ACCESSFILEDEN  = 46	' 拒绝访问文件.
%TRM_DATABUFTOSMALL = 50	' 数据缓冲区太小.
%TRM_TIMEOUT        = 99	' 超时.
'------------------------------------------------------------------------------
' Tsunami 2.5 offers three APIs... a string-based API and two pointer-based APIs.
' The following 38 declarations represent the string-based API...
'------------------------------------------------------------------------------
DECLARE FUNCTION trm_Accelerate LIB "TRM.DLL" ALIAS "trm_Accelerate" ( FileHandle AS LONG, CacheSize AS LONG ) AS LONG
DECLARE FUNCTION trm_Close LIB "TRM.DLL" ALIAS "trm_Close" ( FileHandle AS LONG ) AS LONG
DECLARE FUNCTION trm_CloseAll LIB "TRM.DLL" ALIAS "trm_CloseAll" ( ) AS LONG
DECLARE FUNCTION trm_Count LIB "TRM.DLL" ALIAS "trm_Count" ( FileHandle AS LONG ) AS LONG
DECLARE FUNCTION trm_Create LIB "TRM.DLL" ALIAS "trm_Create" ( PathFileName AS STRING, FileDef AS STRING, OverWrite AS LONG ) AS LONG
DECLARE FUNCTION trm_CurrKeyPos LIB "TRM.DLL" ALIAS "trm_CurrKeyPos" ( FileHandle AS LONG ) AS LONG
DECLARE FUNCTION trm_Delete LIB "TRM.DLL" ALIAS "trm_Delete" ( FileHandle AS LONG ) AS LONG
DECLARE FUNCTION trm_FileIsOpen LIB "TRM.DLL" ALIAS "trm_FileIsOpen" ( FileHandle AS LONG ) AS LONG
DECLARE FUNCTION trm_FileSize LIB "TRM.DLL" ALIAS "trm_FileSize" ( FileHandle AS LONG ) AS LONG
DECLARE FUNCTION trm_Flush LIB "TRM.DLL" ALIAS "trm_Flush" ( ) AS LONG
DECLARE FUNCTION trm_GetByKeyPos LIB "TRM.DLL" ALIAS "trm_GetByKeyPos" ( FileHandle AS LONG, KeyPos AS LONG ) AS STRING
DECLARE FUNCTION trm_GetDirect LIB "TRM.DLL" ALIAS "trm_GetDirect" ( FileHandle AS LONG, RecordPtr AS STRING ) AS STRING
DECLARE FUNCTION trm_GetEqual LIB "TRM.DLL" ALIAS "trm_GetEqual" ( FileHandle AS LONG, KeyNo AS LONG, KeyVal AS STRING ) AS STRING
DECLARE FUNCTION trm_GetEqualOrGreater LIB "TRM.DLL" ALIAS "trm_GetEqualOrGreater" ( FileHandle AS LONG, KeyNo AS LONG, KeyVal AS STRING ) AS STRING
DECLARE FUNCTION trm_GetEqualOrLess LIB "TRM.DLL" ALIAS "trm_GetEqualOrLess" ( FileHandle AS LONG, KeyNo AS LONG, KeyVal AS STRING ) AS STRING
DECLARE FUNCTION trm_GetFileDef LIB "TRM.DLL" ALIAS "trm_GetFileDef" ( FileHandle AS LONG ) AS STRING
DECLARE FUNCTION trm_GetFileVer LIB "TRM.DLL" ALIAS "trm_GetFileVer" ( FileHandle AS LONG ) AS STRING
DECLARE FUNCTION trm_GetFirst LIB "TRM.DLL" ALIAS "trm_GetFirst" ( FileHandle AS LONG, KeyNo AS LONG ) AS STRING
DECLARE FUNCTION trm_GetGreater LIB "TRM.DLL" ALIAS "trm_GetGreater" ( FileHandle AS LONG, KeyNo AS LONG, KeyVal AS STRING ) AS STRING
DECLARE FUNCTION trm_GetLast LIB "TRM.DLL" ALIAS "trm_GetLast" ( FileHandle AS LONG, KeyNo AS LONG ) AS STRING
DECLARE FUNCTION trm_GetLess LIB "TRM.DLL" ALIAS "trm_GetLess" ( FileHandle AS LONG, KeyNo AS LONG, KeyVal AS STRING ) AS STRING
DECLARE FUNCTION trm_GetNext LIB "TRM.DLL" ALIAS "trm_GetNext" ( FileHandle AS LONG ) AS STRING
DECLARE FUNCTION trm_GetPosition LIB "TRM.DLL" ALIAS "trm_GetPosition" ( FileHandle AS LONG, Direction AS LONG ) AS STRING
DECLARE FUNCTION trm_GetPrev LIB "TRM.DLL" ALIAS "trm_GetPrev" ( FileHandle AS LONG ) AS STRING
DECLARE FUNCTION trm_Insert LIB "TRM.DLL" ALIAS "trm_Insert" ( FileHandle AS LONG, Record AS STRING ) AS LONG
DECLARE FUNCTION trm_Integrity LIB "TRM.DLL" ALIAS "trm_Integrity" ( FileHandle AS LONG ) AS LONG
DECLARE FUNCTION trm_Open LIB "TRM.DLL" ALIAS "trm_Open" ( PathFileName AS STRING, MultiUser AS LONG ) AS LONG
DECLARE FUNCTION trm_Rebuild LIB "TRM.DLL" ALIAS "trm_Rebuild" ( PathFileName AS STRING, SaveFile AS STRING, FileDef AS STRING ) AS LONG
DECLARE FUNCTION trm_Recover LIB "TRM.DLL" ALIAS "trm_Recover" ( PathFileName AS STRING, OutFile AS STRING, PageSize AS LONG, Compression AS LONG ) AS LONG
DECLARE FUNCTION trm_Result LIB "TRM.DLL" ALIAS "trm_Result" ( FileHandle AS LONG ) AS LONG
DECLARE FUNCTION trm_SetKeyPath LIB "TRM.DLL" ALIAS "trm_SetKeyPath" ( FileHandle AS LONG, KeyNo AS LONG ) AS LONG
DECLARE FUNCTION trm_StepFirst LIB "TRM.DLL" ALIAS "trm_StepFirst" ( FileHandle AS LONG ) AS STRING
DECLARE FUNCTION trm_StepLast LIB "TRM.DLL" ALIAS "trm_StepLast" ( FileHandle AS LONG ) AS STRING
DECLARE FUNCTION trm_StepNext LIB "TRM.DLL" ALIAS "trm_StepNext" ( FileHandle AS LONG ) AS STRING
DECLARE FUNCTION trm_StepPrev LIB "TRM.DLL" ALIAS "trm_StepPrev" ( FileHandle AS LONG ) AS STRING
DECLARE FUNCTION trm_TimeOut LIB "TRM.DLL" ALIAS "trm_TimeOut" ( Limit AS LONG ) AS LONG
DECLARE FUNCTION trm_Update LIB "TRM.DLL" ALIAS "trm_Update" ( FileHandle AS LONG, Record AS STRING ) AS LONG
DECLARE FUNCTION trm_Version LIB "TRM.DLL" ALIAS "trm_Version" ( ) AS LONG
' The following two declarations represent Tsunami's two pointer-based APIs.
' They both accomplish the same thing... one passes seven long integers as parameters,
' while the other passes a seven-element UDT (structured variable) as it's sole parameter.
DECLARE FUNCTION trm LIB "TRM.DLL" ALIAS "trm" ( Op AS LONG, File AS LONG, DataPtr AS LONG, DataLen AS LONG, KeyPtr AS LONG, KeyLen AS LONG, KeyNo AS LONG ) AS LONG
DECLARE FUNCTION trm_udt LIB "TRM.DLL" ALIAS "trm_udt" ( TRMudt AS ANY ) AS LONG
' The following User Defined Type (structured) variable is only used
' with the pointer-based API that passes parameters in the UDT...
TYPE TRMtype
	op AS LONG		' Tsunami operation number
	file AS LONG		' Tsunami file handle
	dataptr AS LONG		' Address of data buffer
	datalen AS LONG		' Length of data buffer
	keyptr AS LONG		' Address of key buffer
	keylen AS LONG		' Length of key buffer
	keyno AS LONG		' Key number
END TYPE
' The following EQUATES (constants) represent the Tsunami
' operation codes (Op) used with the pointer-based APIs...
%Tsu_Accelerate         = 32
%Tsu_Close              = 1
%Tsu_CloseAll           = 28
%Tsu_Count              = 17
%Tsu_Create             = 14
%Tsu_CurrKeyPos         = 45
%Tsu_Delete             = 4
%Tsu_FileIsOpen         = 16
%Tsu_FileSize           = 18
%Tsu_Flush              = 29
%Tsu_GetByKeyPos        = 44
%Tsu_GetDirect          = 23
%Tsu_GetEqual           = 5
%Tsu_GetEqualOrGreater  = 9
%Tsu_GetEqualOrLess     = 11
%Tsu_GetFileDef         = 36
%Tsu_GetFileVer         = 25
%Tsu_GetFirst           = 12
%Tsu_GetGreater         = 8
%Tsu_GetLast            = 13
%Tsu_GetLess            = 10
%Tsu_GetNext            = 6
%Tsu_GetPosition        = 22
%Tsu_GetPrev            = 7
%Tsu_Insert             = 2
%Tsu_Integrity          = 37
%Tsu_Open               = 0
%Tsu_Rebuild            = 38
%Tsu_Recover            = 39
%Tsu_SetKeyPath         = 30
%Tsu_StepFirst          = 33
%Tsu_StepLast           = 34
%Tsu_StepNext           = 24
%Tsu_StepPrev           = 35
%Tsu_TimeOut            = 31
%Tsu_Update             = 3
%Tsu_Version            = 26
