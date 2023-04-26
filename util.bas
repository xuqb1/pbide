'util.bas
'用来实现一些公用的函数

'DECLARE FUNCTION FileExist (BYVAL sfName AS STRING) AS LONG     '文件是否存在判断
'DECLARE FUNCTION FileNam   (BYVAL Src AS STRING) AS STRING      '获取文件名
'DECLARE FUNCTION FilePath  (BYVAL Src AS STRING) AS STRING      '获取文件路径
'DECLARE FUNCTION IsWin95() AS LONG                              '是否是Win95
'
TYPE BITMAPINFOTYPE
  bmiHeader    AS BITMAPINFOHEADER
  bmiColors(1) AS RGBQUAD
END TYPE
'------------------------------------------------------------------------------
' UDT (用户自定义类型及结构) UNIONS.
'------------------------------------------------------------------------------
' 保存和加载程序的参数 - 增加此UDT到你想要保存或重载的位置。
' 更新APPPROPS.INC的读写部分!
' 也要考虑本模块中的 AppInitialize() 和 AppTerminate() 函数!
'------------------------------------------------------------------------------
TYPE AppParametersTYPE
  sAppPath     AS STRING * %MAX_PATH      ' 程序路径.
  sAppIniPath  AS STRING * %MAX_PATH      ' 程序INI文件路径.
  sAppDbPath   AS STRING * %MAX_PATH      ' 程序数据库路径.
  sAppIdxPath  AS STRING * %MAX_PATH      ' 程序数据库目录路径.
  lSplashTime  AS LONG                    ' 程序闪屏计时.
  sbDateFmt    AS LONG                    ' 状态栏日期格式.
  sbTimeFmt    AS LONG                    ' 状态栏时间格式.
  rbRowCount   AS LONG                    ' Rebar 行数.
  rbBandCount  AS LONG                    ' Rebar 段数.
  rbBand0      AS LONG                    ' Rebar 0段是否被移动.
  rbBand1      AS LONG                    ' Rebar 0段是否被移动.
  sbTimePart   AS LONG                    ' 状态栏时间栏数.
  sbDatePart   AS LONG                    ' 状态栏日期栏数.
END TYPE
GLOBAL udtAp AS AppParametersTYPE         ' 全局程序属性UDT变量.

'------------------------------------------------------------------------------
' 函数 : IniRead () '
' 描述 : 从程序INI文件读取数据. '
' 用法 : sText = IniRead ("INIFILE", "SECTION", "KEY", "DEFAULT")
'      : lVal  = VAL (IniRead ("INIFILE", "SECTION", "KEY", "DEFAULT"))
'------------------------------------------------------------------------------
FUNCTION IniRead (BYVAL sIniFile AS STRING, _
          BYVAL sSection AS STRING, _
          BYVAL sKey     AS STRING, _
          BYVAL sDefault AS STRING) AS STRING

  LOCAL lResult   AS LONG
  LOCAL szSection AS ASCIIZ * 125
  LOCAL szKey     AS ASCIIZ * 125
  LOCAL szData    AS ASCIIZ * 150
  LOCAL szDefault AS ASCIIZ * 150
  LOCAL szIniFile AS ASCIIZ * %MAX_PATH
  szSection = sSection
  szKey     = sKey
  szIniFile = sIniFile
  szDefault = sDefault
  lResult = GetPrivateProfileString(szSection, szKey, szDefault, szData, _
                  SIZEOF(szData), szIniFile)
  'canLog=1
  'IF left$(szSection,6)="Editor" THEN
  '  RunLog "IniRead :" & str$(lResult) & "--" & szSection & "--" & szKey
    'msgbox "test"
  'END IF
  'canLog=0
  FUNCTION = szData
END FUNCTION

'------------------------------------------------------------------------------
' 函数 : IniWrite ()
'
' 描述 : 保存数据到程序INI文件.
'
' 用法 : lResult = IniWrite ("INIFILE", "SECTION", "KEY", "VALUE")
'------------------------------------------------------------------------------
FUNCTION IniWrite (BYVAL sIniFile AS STRING, _
           BYVAL sSection AS STRING, _
           BYVAL sKey     AS STRING, _
           BYVAL sValue   AS STRING) AS LONG

  LOCAL szSection AS ASCIIZ * 125
  LOCAL szKey     AS ASCIIZ * 125
  LOCAL szValue   AS ASCIIZ * 150
  LOCAL szIniFile AS ASCIIZ * %MAX_PATH
  szSection = sSection
  szKey     = sKey
  szIniFile = sIniFile
  szValue   = TRIM$(sValue)
  FUNCTION = WritePrivateProfileString(szSection, szKey, szValue, szIniFile)
END FUNCTION
'------------------------------------------------------------------------------
' 程序 : IniDelSection () '
' 描述 : 删除INI文件中的段.  '
' 用法 : lResult = IniDelSection ("INIFILE", "SECTION")
'------------------------------------------------------------------------------
FUNCTION IniDelSection (BYVAL sIniFile AS STRING, _
            BYVAL sSection AS STRING) AS LONG
  LOCAL szSection AS ASCIIZ * 125
  LOCAL szIniFile AS ASCIIZ * %MAX_PATH
  szSection = sSection
  szIniFile = sIniFile
  FUNCTION = WritePrivateProfileSection(szSection, BYVAL %NULL, szIniFile)
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : IniDelKey ()   '
' 描述 : 删除INI文件中的键值. '
' 用法 : lResult = IniDelKey ("INIFLE", "SECTION", "KEY")
'------------------------------------------------------------------------------
FUNCTION IniDelKey (BYVAL sIniFile AS STRING, _
          BYVAL sSection AS STRING, _
          BYVAL sKey     AS STRING) AS LONG
  LOCAL szSection AS ASCIIZ * 125
  LOCAL szKey     AS ASCIIZ * 125
  LOCAL szIniFile AS ASCIIZ * %MAX_PATH
  szSection = sSection
  szKey     = sKey
  szIniFile = sIniFile
  FUNCTION = WritePrivateProfileString(szSection, szKey, BYVAL %NULL, szIniFile)
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : AppPath ()   '
' 描述 : 返回程序路径.   '
' 用法 : sPath = AppPath ()
'------------------------------------------------------------------------------
FUNCTION AppPath () AS STRING
  LOCAL szTmpAsciiz AS ASCIIZ * %MAX_PATH
  GetModuleFileName GetModuleHandle(BYVAL %NULL), szTmpAsciiz, %MAX_PATH
  FUNCTION = LEFT$(szTmpAsciiz, INSTR(-1, szTmpAsciiz, ANY "\/:"))
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : GetProperty ()  '
' 描述 : 从config.INI文件读取属性值  '
' 用法 : sText = GetProperty ("KEY", "DEFAULT")
'      : lVal  = VAL (GetProperty ("KEY", "DEFAULT"))
'------------------------------------------------------------------------------
FUNCTION GetProperty (BYVAL sKey     AS STRING, _
            BYVAL sDefault AS STRING) AS STRING
  LOCAL lResult   AS LONG
  LOCAL szSection AS ASCIIZ * 125
  LOCAL szKey     AS ASCIIZ * 125
  LOCAL szData    AS ASCIIZ * 150
  LOCAL szDefault AS ASCIIZ * 150
  LOCAL szIniFile AS ASCIIZ * %MAX_PATH
  szSection = "Properties"
  szKey     = sKey
  szIniFile = AppPath & "Config.ini"
  szDefault = sDefault
  lResult = GetPrivateProfileString(szSection, szKey, szDefault, szData, _
                  SIZEOF(szData), szIniFile)
  FUNCTION = szData
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : SetProperty ()  '
' 描述 : 保存数据到程序INI文件.  '
' 用法  : lResult = SetProperty ("KEY", "VALUE")
'------------------------------------------------------------------------------
FUNCTION SetProperty (BYVAL sKey   AS STRING, _
            BYVAL sValue AS STRING) AS LONG
  LOCAL szSection AS ASCIIZ * 125
  LOCAL szKey     AS ASCIIZ * 125
  LOCAL szValue   AS ASCIIZ * 150
  LOCAL szIniFile AS ASCIIZ * %MAX_PATH
  szSection = "Properties"
  szKey     = sKey
  szIniFile = AppPath & "AppProps.ini"
  szValue   = TRIM$(sValue)
  FUNCTION = WritePrivateProfileString(szSection, szKey, szValue, szIniFile)
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : AppLoadParams ()   '
' 描述 : 从INI文件读取所有的程序参数.
'------------------------------------------------------------------------------
FUNCTION AppLoadParams () AS LONG
  ' INI 段 [APPLICATION].
  udtAp.sAppPath    = IniRead(udtAp.sAppIniPath, "APPLICATION PATHS", "APPPATH",    "")
  udtAp.sAppIniPath = IniRead(udtAp.sAppIniPath, "APPLICATION PATHS", "APPINIPATH", "")
  udtAp.sAppDbPath  = IniRead(udtAp.sAppIniPath, "APPLICATION PATHS", "APPDBPATH",  "")
  udtAp.sAppIdxPath = IniRead(udtAp.sAppIniPath, "APPLICATION PATHS", "APPIDXPATH", "")
  ' INI 段 [SPLASH TIMMING].
  udtAp.lSplashTime = VAL(IniRead(udtAp.sAppIniPath, "SPLASH TIMMING", "SPLASHTIME", ""))
  ' INI 段 [DATETIMEFORMAT].
  udtAp.sbDateFmt  = VAL(IniRead(udtAp.sAppIniPath, "DATE TIME FORMAT", "SBDATEFMT", ""))
  udtAp.sbTimeFmt  = VAL(IniRead(udtAp.sAppIniPath, "DATE TIME FORMAT", "SBTIMEFMT", ""))
  udtAp.sbDatePart = VAL(IniRead(udtAp.sAppIniPath, "DATE TIME FORMAT", "SBDATEPART", ""))
  udtAp.sbTimePart = VAL(IniRead(udtAp.sAppIniPath, "DATE TIME FORMAT", "SBTIMEPART", ""))
  ' INI 段 [REBAR POSITIONS].
  udtAp.rbRowCount  = VAL(IniRead(udtAp.sAppIniPath, "REBAR POSITIONS", "RBROWS",  ""))
  udtAp.rbBandCount = VAL(IniRead(udtAp.sAppIniPath, "REBAR POSITIONS", "RBBANDS", ""))
  udtAp.rbBand0     = VAL(IniRead(udtAp.sAppIniPath, "REBAR POSITIONS", "RBBAND0", ""))
  udtAp.rbBand1     = VAL(IniRead(udtAp.sAppIniPath, "REBAR POSITIONS", "RBBAND1", ""))
  FUNCTION = %FALSE
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : AppSaveParams ()'
' 描述 : 保存程序所有参数到INI文件.
'------------------------------------------------------------------------------
FUNCTION AppSaveParams () AS LONG
  ' INI 段 [Application].
  IniWrite udtAp.sAppIniPath, "APPLICATION PATHS", "APPPATH",    udtAp.sAppPath
  IniWrite udtAp.sAppIniPath, "APPLICATION PATHS", "APPINIPATH", udtAp.sAppIniPath
  IniWrite udtAp.sAppIniPath, "APPLICATION PATHS", "APPDBPATH",  udtAp.sAppDbPath
  IniWrite udtAp.sAppIniPath, "APPLICATION PATHS", "APPIDXPATH", udtAp.sAppIdxPath
  ' INI 段 [SPLASH TIMMING].
  IniWrite udtAp.sAppIniPath, "SPLASH TIMMING", "SPLASHTIME", STR$(udtAp.lSplashTime)
  ' INI 段 [DATETIMEFORMAT].
  IniWrite udtAp.sAppIniPath, "DATE TIME FORMAT", "SBDATEFMT",  STR$(udtAp.sbDateFmt)
  IniWrite udtAp.sAppIniPath, "DATE TIME FORMAT", "SBTIMEFMT",  STR$(udtAp.sbTimeFmt)
  IniWrite udtAp.sAppIniPath, "DATE TIME FORMAT", "SBDATEPART", STR$(udtAp.sbDatePart)
  IniWrite udtAp.sAppIniPath, "DATE TIME FORMAT", "SBTIMEPART", STR$(udtAp.sbTimePart)
  ' INI 段 [REBAR POSITIONS].
  IniWrite udtAp.sAppIniPath, "REBAR POSITIONS", "RBROWS",  STR$(udtAp.rbRowCount)
  IniWrite udtAp.sAppIniPath, "REBAR POSITIONS", "RBBANDS", STR$(udtAp.rbBandCount)
  IniWrite udtAp.sAppIniPath, "REBAR POSITIONS", "RBBAND0", STR$(udtAp.rbBand0)
  IniWrite udtAp.sAppIniPath, "REBAR POSITIONS", "RBBAND1", STR$(udtAp.rbBand1)
  FUNCTION = %FALSE
END FUNCTION
'===============================================================================
FUNCTION OnlyOneInstance(szName AS ASCIIZ) AS LONG
'------------------------------------------------------------------------------
  ' OnlyOneInstance - 只允许有一个程序实例运行
  '----------------------------------------------------------------------------
  LOCAL hMutex&
  hMutex& = CreateMutex(BYVAL %NULL, %TRUE, szName)
  IF GetLastError = %ERROR_ALREADY_EXISTS THEN
    CALL CloseHandle(hMutex&)
    FUNCTION = 0
  ELSE
    FUNCTION = 1
  END IF
END FUNCTION
'==============================================================================
FUNCTION FileExist(BYVAL sfName AS STRING) AS LONG
'------------------------------------------------------------------------------
  ' FileExist - 确认某个文件或文件是否存在
  '----------------------------------------------------------------------------
  LOCAL hRes AS DWORD, tWFD AS WIN32_FIND_DATA
  IF LEN(sfName) = 0 THEN EXIT FUNCTION       'no need to continue then..
  IF ASC(sfName, -1) = 92 THEN                'if a path with trailing backslash
      sfName = LEFT$(sfName, LEN(sfName) - 1) 'remove trailing backslash
  END IF
  '----------------------------------------------------------------------------
  hRes = FindFirstFile(BYVAL STRPTR(sfName), tWFD)
  IF hRes <> %INVALID_HANDLE_VALUE THEN
      FUNCTION = %TRUE
      FindClose hRes
  END IF
END FUNCTION
'==============================================================================
FUNCTION FileNam (BYVAL Src AS STRING) AS STRING
'------------------------------------------------------------------------------
  ' 获取给定路径中的文件名部分
  '----------------------------------------------------------------------------
  LOCAL x AS LONG
  x = INSTR(-1, Src, ANY ":/\")
  IF x THEN
    FUNCTION = MID$(Src, x + 1)
  ELSE
    FUNCTION = Src
  END IF
END FUNCTION
'==============================================================================
FUNCTION FilePath (BYVAL Src AS STRING) AS STRING
'------------------------------------------------------------------------------
  ' 获取给定路径中的路径部分
  '----------------------------------------------------------------------------
  LOCAL x AS LONG
  x = INSTR(-1, Src, ANY ":\/")
  IF x THEN
    FUNCTION = LEFT$(Src, x)
  ELSE
    FUNCTION = Src
  END IF
END FUNCTION
'==============================================================================
FUNCTION IsWin95() AS LONG
'------------------------------------------------------------------------------
  ' 是否是 Windows 95
  '----------------------------------------------------------------------------
  LOCAL vi AS OSVERSIONINFO
  vi.dwOsVersionInfoSize = SIZEOF(vi)
  GetVersionEx vi
  FUNCTION = ((vi.dwPlatformId = %VER_PLATFORM_WIN32_WINDOWS) AND (vi.dwMinorVersion = 0))
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : DrawBitMap () '
' 描述 : 绘制闪屏窗口位图.  '
'------------------------------------------------------------------------------
FUNCTION DrawBitMap (BYVAL hDC  AS DWORD, _
           BYVAL hBmp AS DWORD, _
           BYVAL x    AS LONG, _
           BYVAL y    AS LONG) AS LONG
  LOCAL bm     AS BITMAP
  LOCAL hDCmem AS DWORD
  LOCAL dwSize AS LONG
  LOCAL ptSize AS POINTAPI
  LOCAL ptOrg  AS POINTAPI
  hDCmem = CreateCompatibleDC(hDC)
  SelectObject hDCmem, hBmp
  SetMapMode hDCmem, GetMapMode(hDC)
  GetObject hBmp, SIZEOF(bm), bm
  ptSize.x = bm.bmWidth
  ptSize.y = bm.bmHeight
  DPtoLP hDC, ptSize, 1
  ptOrg.x = 0
  ptOrg.y = 0
  DPtoLP hDCmem, ptOrg, 1
  BitBlt hDC, x, y, ptSize.x, ptSize.y, hDCmem, ptOrg.x, ptOrg.y, %SRCCOPY
  DeleteDC hdcMem
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : CreateDibPalette ()  '
' 描述 : 使用模板创建DIB位图.  '
'------------------------------------------------------------------------------
FUNCTION CreateDibPalette (lpbmi   AS BITMAPINFOTYPE, _
               nColors AS LONG) AS LONG
  REGISTER i AS LONG
  LOCAL lpbi    AS BITMAPINFOHEADER PTR
  LOCAL lpPal   AS LOGPALETTE PTR
  LOCAL hLogPal AS DWORD
  LOCAL hPal    AS DWORD
  lpbi = VARPTR(lpbmi)
  IF @lpbi.biBitCount <= 8 THEN
    nColors = 1
    SHIFT LEFT nColors, @lpbi.biBitCount
  ELSE
    nColors = 0
  END IF
  IF nColors THEN
    hLogPal = GlobalAlloc(%GHND, LEN(LOGPALETTE) + LEN(PALETTEENTRY) * nColors)
    lpPal = GlobalLock(hLogPal)
    @lpPal.palVersion    = &H300
    @lpPal.palNumEntries = nColors
    FOR i = 0 TO nColors
      @lpPal.palPalEntry(i).peRed   = lpbmi.bmiColors(i).rgbRed
      @lpPal.palPalEntry(i).peGreen = lpbmi.bmiColors(i).rgbGreen
      @lpPal.palPalEntry(i).peBlue  = lpbmi.bmiColors(i).rgbBlue
      @lpPal.palPalEntry(i).peFlags = 0
    NEXT i
    hPal = CreatePalette(@lpPal)
    GlobalUnlock hLogPal
    GlobalFree hLogPal
  END IF
  FUNCTION = hPal
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : LoadBitmapRes ()  '
' 描述 : 从资源文件加载闪屏窗口位图.  '
'------------------------------------------------------------------------------
FUNCTION LoadBitmapRes (BYVAL hInstance AS DWORD, _
            ResBitmap       AS ASCIIZ, _
            hPalette        AS DWORD) AS LONG
  LOCAL hRes    AS LONG
  LOCAL hMem    AS DWORD
  LOCAL hBmp    AS DWORD
  LOCAL hDC     AS DWORD
  LOCAL nColors AS LONG
  LOCAL lpbi    AS BITMAPINFOHEADER PTR
  LOCAL Tmp     AS DWORD
  hRes = FindResource(hInstance, ResBitmap, BYVAL %RT_BITMAP)
  IF hRes = 0 THEN
    EXIT FUNCTION
  END IF
  hDC = GetDC(%NULL)
  IF GetDeviceCaps(hDC, %BITSPIXEL) = 3 THEN
    ReleaseDC %NULL, hDC
    FUNCTION = LoadBitmap(hInstance, ResBitmap)
  END IF
  hMem = LoadResource(hInstance, hRes)
  lpbi = LockResource(hMem)
  hPalette = CreateDIBPalette(BYVAL lpbi, nColors)
  IF hPalette THEN
    SelectPalette hDC, hPalette, %FALSE
    RealizePalette hDC
  END IF
  Tmp = lpbi + @lpbi.biSize + (nColors * LEN(RGBQUAD))
  FUNCTION = CreateDIBitmap(hDC, @lpbi, %CBM_INIT, BYVAL Tmp, BYVAL lpbi, %DIB_RGB_COLORS)
  ReleaseDC %NULL, hDC
  FreeResource hMem
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : CenterWindow ()'
' 描述 : 将给定hWnd句柄的任何窗口置中.
'------------------------------------------------------------------------------
FUNCTION CenterWindow (BYVAL hWnd AS DWORD) AS LONG
  LOCAL WndRect AS RECT
  LOCAL x       AS LONG
  LOCAL y       AS LONG
  CALL GetWindowRect(hWnd, WndRect)
  x = (GetSystemMetrics(%SM_CXSCREEN) - (WndRect.nRight - WndRect.nLeft)) \ 2
  y = (GetSystemMetrics(%SM_CYSCREEN) - (WndRect.nBottom - WndRect.nTop + _
            GetSystemMetrics(%SM_CYCAPTION))) \ 2
  CALL SetWindowPos(hWnd, %NULL, x, y, 0, 0, %SWP_NOSIZE OR %SWP_NOZORDER)
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : CenterWindowRel () '
' 描述 : 将给定hWnd句柄有关系的所有窗口置中.
'------------------------------------------------------------------------------
FUNCTION CenterWindowRel (BYVAL hWnd AS DWORD) AS LONG
  LOCAL hWndParent AS DWORD
  LOCAL rc         AS RECT
  LOCAL rcp        AS RECT
  LOCAL nWidth     AS LONG
  LOCAL nHeight    AS LONG
  LOCAL ScrWidth   AS LONG
  LOCAL ScrHeight  AS LONG
  LOCAL X          AS LONG
  LOCAL Y          AS LONG
  ' 使窗口关系到其父窗口.
  hWndParent = GetParent(hWnd)
  GetWindowRect(hWnd, rc)
  GetWindowRect(hWndParent, rcp)
  nWidth  = rc.nRight  - rc.nLeft
  nHeight = rc.nBottom - rc.nTop
  X = ((rcp.nRight  - rcp.nLeft) - nWidth)  / 2 + rcp.nLeft
  Y = ((rcp.nBottom - rcp.nTop)  - nHeight) / 2 + rcp.nTop
  ScrWidth  = GetSystemMetrics(%SM_CXSCREEN)
  ScrHeight = GetSystemMetrics(%SM_CYSCREEN)
  ' 确使对话框不会移动到屏幕外面.
  IF (X < 0) THEN X = 0
  IF (Y < 0) THEN Y = 0
  IF (X + nWidth  > ScrWidth)  THEN X = ScrWidth  - nWidth
  IF (Y + nHeight > ScrHeight) THEN Y = ScrHeight - nHeight
  MoveWindow(hWnd, X, Y, nWidth, nHeight, %FALSE)
  FUNCTION = %TRUE
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : MakeFont ()
' 描述 : 创建新字体，返回新字体句柄.
'------------------------------------------------------------------------------
FUNCTION MakeFont (BYVAL sFntName    AS STRING, _
           BYVAL fFntSize    AS LONG, _
           BYVAL lWeight     AS LONG, _
           BYVAL lUnderlined AS LONG, _
           BYVAL lItalic     AS LONG, _
           BYVAL lStrike     AS LONG, _
           BYVAL lCharSet    AS LONG)   AS DWORD
'------------------------------------------------------------------------------
' 使用示例:   '
'    hFont = MakeFont ("MS Sans Serif", 8, %FW_NORMAL, %FALSE, %FALSE, %FALSE,
'                      %DEFAULT_CHARSET)  '
'    hFont = MakeFont ("Courier New", 10, %FW_BOLD, %FALSE, %FALSE, %FALSE,
'                      %DEFAULT_CHARSET)
'    hFont = MakeFont ("Marlett", 8, %FW_NORMAL, %FALSE, %FALSE, %FALSE,
'                      %SYMBOL_CHARSET)  '
' 注意: 任何用MakeFont创建的字体在不再需要时应该使用DeleteObject来销毁它，以防止内存泄漏。
'       例如:    '
'        Case %WM_DESTROY         OnDestroy () 句柄.
'          DeleteObject hFont
'------------------------------------------------------------------------------
  DIM lf          AS LOGFONT
  DIM hDC         AS DWORD
  DIM lCyPixels   AS LONG

  hDC         = GetDC(%HWND_DESKTOP)
  lCyPixels   = GetDeviceCaps(hDC, %LOGPIXELSY)
  ReleaseDC %HWND_DESKTOP, hDC
  lf.lfHeight         = -(fFntSize * lCyPixels) \ 72
  lf.lfFaceName       = sFntName
  lf.lfPitchAndFamily = %FF_DONTCARE
  IF lWeight = %FW_DONTCARE THEN
    lf.lfWeight = %FW_NORMAL
  ELSE
    lf.lfWeight = lWeight
  END IF
  lf.lfUnderline  = lUnderlined
  lf.lfItalic     = lItalic
  lf.lfStrikeOut  = lStrike
  lf.lfCharSet    = lCharSet
  FUNCTION = CreateFontIndirect(lf)
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : ExecuteURL ()
'
' 描述 : 执行窗口程序，MAIL Web浏览器等。
'------------------------------------------------------------------------------
FUNCTION ExecuteURL (BYVAL hWnd AS DWORD, _
           BYVAL sText AS STRING) AS LONG
  LOCAL lResult AS LONG
  LOCAL szText  AS ASCIIZ * 128
  szText = sText
  lResult = ShellExecute(hWnd, "open", szText, BYVAL %NULL, _
             BYVAL %NULL, %SW_SHOWNORMAL)
  FUNCTION = lResult
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : ExeFile () '
' 描述 : 执行文件.
'------------------------------------------------------------------------------
FUNCTION ExeFile (BYVAL CmdString AS STRING, _
          BYVAL lStyle    AS LONG) AS LONG
  LOCAL lResult AS LONG
  LOCAL sErr    AS STRING
  LOCAL sExe    AS STRING
  ON ERROR RESUME NEXT
  sExe = AppPath & CmdString
  SHELL sExe, lStyle
  IF ERR THEN
    sErr = ERROR$
    lResult = MSGBOX(sErr, %MB_ICONERROR OR %MB_OKCANCEL, "程序出错")
    ERRCLEAR
    FUNCTION = %FALSE
    EXIT FUNCTION
  END IF
  FUNCTION = %TRUE
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : PbDoEvents () '
' 描述 : 为程序做事件响应缓冲。Do Events for Application.
'------------------------------------------------------------------------------
FUNCTION PbDoEvents () AS LONG
  LOCAL msg AS tagMSG
  IF PeekMessage(msg,0,0,0,1) THEN
    TranslateMessage(msg)
    DispatchMessage(msg)
  ELSE
    SLEEP (1)
  END IF
END FUNCTION
'------------------------------------------------------------------------------
' Sub  : PBCopyMemory (   '
' 描述 : 复制内存.
'------------------------------------------------------------------------------
SUB PBCopyMemory (BYVAL pDest   AS BYTE PTR, _
          BYVAL pSource AS BYTE PTR, _
          BYVAL cb      AS LONG)

  DO WHILE cb > 0
    @pDest = @pSource
    INCR pDest
    INCR pSource
    DECR cb
  LOOP
END SUB
'===========================================================================================================
' 判断文件是否存在
'-----------------------------------------------------------------------------------------------------------
FUNCTION Exist(sFilename AS STRING) AS LONG
  ' Checks to see if a given disk file exists
  LOCAL    dummy              AS LONG
  dummy = GETATTR(sFilename)
  FUNCTION = (ERRCLEAR = 0)
END FUNCTION
'创建数据库，创建其中的表
FUNCTION MakeDB(BYVAL dbName AS STRING)AS LONG '创建数据库
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  errStr=""
  IF dbName="" OR RIGHT$(dbName,3)<>".db" THEN
    errStr="数据库文件名为空或不标准,"
    GOTO abort
  END IF
  IF ISTRUE sqlOpen(dbName,hDB) THEN
    errStr="无法打开或创建数据库,"
    GOTO abort
  END IF
  '创建配置表
  sSql="CREATE TABLE "+$CONFIGTABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,name text,value text,remark text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$CONFIGTABLENAME+","
    GOTO abort
  END IF
  '创建默认配置表
  sSql="CREATE TABLE "+$DEFAULTCONFIGTABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,name text,value text,remark text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$DEFAULTCONFIGTABLENAME+","
    GOTO abort
  END IF
  '创建数据集表
  sSql="CREATE TABLE "+$DATASETTABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,type text,value text,remark text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$DATASETTABLENAME+","
    GOTO abort
  END IF
  '创建字符串语言表
  sSql="CREATE TABLE "+$STRINGLANGTABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,name text,english text,chinese text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$STRINGLANGTABLENAME+","
    GOTO abort
  END IF
  '创建控件样式表
  sSql="CREATE TABLE "+$CONTROLSTYLETABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,controlname text,stylename text,value text,remark text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$CONTROLSTYLETABLENAME+","
    GOTO abort
  END IF
  sqlClose hDB
  FUNCTION=1
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=0
END FUNCTION
' 初始化：默认配置表，配置表，数据集表的语言，字符串语言表
FUNCTION InitConfig() AS LONG
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="无法打开或创建数据库,"
    GOTO abort
  END IF
  ' 初始化默认配置表
  sSql="BEGIN TRANSACTION;"+$CRLF
  '常规
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showflashwindow','1');"+$CRLF          '启动时显示欢迎屏幕
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showguiderwindow','1');"+$CRLF         '启动后显示向导窗口
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowposopt','0');"+$CRLF         '主窗口位置配置:默认，固定，使用上一次
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowpos','300,200');"+$CRLF      '主窗口指定位置
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowsize','640,480');"+$CRLF     '主窗口指定尺寸
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowlastpos','300,200');"+$CRLF  '主窗口上次位置
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowlastsize','640,480');"+$CRLF '主窗口上次尺寸
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('language','English');"+$CRLF           '语言
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('clientrect','0,0,0,0');"+$CRLF         '主客户区矩形：左，上，宽，高
  'rebar和band
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarrowcount','2');"+$CRLF            'Rebar行数
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarbandcount','3');"+$CRLF           'Rebar的band数:基本工具栏，编辑工具栏，窗口工具栏，(方法列表工具栏)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband0','0,200');"+$CRLF           'Rebar的第一个(索引值)band信息：wID,cx(相应工具栏的ID,宽度)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband1','0,200');"+$CRLF           'Rebar的第二个(索引值)band信息：wID,cx(相应工具栏的ID,宽度)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband2','0,200');"+$CRLF           'Rebar的第三个(索引值)band信息：wID,cx(相应工具栏的ID,宽度)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband3','0,200');"+$CRLF                 'Rebar的第四个(索引值)band信息：wID,cx(相应工具栏的ID,宽度)(先占位)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband4','0,200');"+$CRLF                 'Rebar的第五个(索引值)band信息：wID,cx(相应工具栏的ID,宽度)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband5','0,200');"+$CRLF                 'Rebar的第六个(索引值)band信息：wID,cx(相应工具栏的ID,宽度)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarlocked','0');"+$CRLF              'rebar是否锁定
  '文件
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('sourcefiletype','.bas');"+$CRLF              '源文件类型扩展名
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('includefiletype','.inc');"+$CRLF             '头文件类型扩展名
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('resourcefiletype','.rc');"+$CRLF             '资源文件类型扩展名
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rcheadfiletype','.h');"+$CRLF                '资源头文件类型扩展名
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('projectfiletype','.pbj');"+$CRLF             '项目文件类型扩展名
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('registryfiletype','1');"+$CRLF               '注册文件类型扩展名
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autoreloadfile','0');"+$CRLF                 '自动重载最近编辑的文件
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfilenumber','5');"+$CRLF               '最近文件列表项最大数
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile1','');"+$CRLF                     '最近文件1
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile2','');"+$CRLF                     '最近文件2
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile3','');"+$CRLF                     '最近文件3
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile4','');"+$CRLF                     '最近文件4
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile5','');"+$CRLF                     '最近文件5
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile6','');"+$CRLF                     '最近文件6
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile7','');"+$CRLF                     '最近文件7
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile8','');"+$CRLF                     '最近文件8
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile9','');"+$CRLF                     '最近文件9
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile10','');"+$CRLF                    '最近文件10
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('startfolder','0');"+$CRLF                    '开始文件夹：0,以默认文件夹开始；1,以最后使用的文件夹开始
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('lastfolder','');"+$CRLF                      '最后使用的文件夹
  '代码编辑
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autoindent','1');"+$CRLF                     '自动缩进
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('undoaftersave','1');"+$CRLF                  '保存后仍可撤销编辑
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('spacepertab','2');"+$CRLF                    '每个Tab相当于的空格数
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autosave','1');"+$CRLF                       '自动备份
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autosaveminute','5');"+$CRLF                 '每#分钟自动备份一次
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showlinenumber','1');"+$CRLF                 '显示行号
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autocompletetip','1');"+$CRLF                '自动完成提示
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('keywordcase','1');"+$CRLF                    '关键字大小写：0,保持原来的大小写;1,大写;2,仅首字母大写;3,小写
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('font','Arial');"+$CRLF                       '字体
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('fontsize','12');"+$CRLF                      '字体大小
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('folderlevel','2');"+$CRLF                    '卷展级别：0,无卷展;1,关键字级;2,Sub/function级
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('foldericon','2');"+$CRLF                     '卷展图标：0,箭头;1,+/-;2,圆形;3,方形
  '语法颜色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('usecolorineditor','1');"+$CRLF               '在编辑器中使用语法颜色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('usecolorwhenprint','1');"+$CRLF              '在打印时使用语法颜色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('assembleforecolor','192,0,0');"+$CRLF        '汇编语句前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('assemblebackcolor','255,255,255');"+$CRLF    '汇编语句背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('commentforecolor','0,127,0');"+$CRLF         '注释语句前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('commentbackcolor','255,255,255');"+$CRLF     '注释语句背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('keywordforecolor','0,0,192');"+$CRLF         '关键字前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('keywordbackcolor','255,255,255');"+$CRLF     '关键字背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('pbformforecolor','192,100,0');"+$CRLF        '窗体前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('pbformbackcolor','255,255,255');"+$CRLF      '窗体背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('stringforecolor','192,32,192');"+$CRLF       '字符串前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('stringbackcolor','255,255,255');"+$CRLF      '字符串背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('textforecolor','0,0,0');"+$CRLF              '文本前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('textbackcolor','255,255,255');"+$CRLF        '文本背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('selectedforecolor','255,255,255');"+$CRLF    '选中语句前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('selectedbackcolor','49,106,197');"+$CRLF     '选中语句背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('breakpointforecolor','255,255,255');"+$CRLF  '断点语句前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('breakpointbackcolor','255,0,0');"+$CRLF      '断点语句背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('bookmarkforecolor','0,0,255');"+$CRLF        '书签语句前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('bookmarkbackcolor','0,255,255');"+$CRLF      '书签语句背景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('execpointforecolor','0,0,0');"+$CRLF         '执行点语句前景色
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('execpointbackcolor','255,255,0');"+$CRLF     '执行点语句背景色
  '窗口编辑
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showgrid','1');"+$CRLF                       '显示网格
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('snaptogrid','1');"+$CRLF                     '对齐到网格
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('gridsize','8x8');"+$CRLF                     '网格大小
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('snapbetweencontrol','1');"+$CRLF             '控件之间自动对齐
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('adjusthandlesize','1');"+$CRLF               '调整句柄尺寸：0,小;1,中;2,大
  '编译
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('includefilepath','D:\Program Files\PBWin\win32api');"+$CRLF        'PB头文件路径
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('resourcefilepath','D:\Program Files\PBWin\win32api');"+$CRLF       'RC头文件路径
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('compilefilepath','D:\Program Files\PBWin\bin\pbwin.exe');"+$CRLF   '编译文件路径
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('win32helpfilepath','D:\Program Files\PBWin\bin\win32.hlp');"+$CRLF 'win32.hlp文件路径
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('beepwhencomplete','0');"+$CRLF               '编译完成后使用提示音
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showcompileresult','1');"+$CRLF              '显示编译结果
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('createcompilelog','1');"+$CRLF               '创建编译日志
  '打印
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('pagenumberonfooter','1');"+$CRLF             '页脚处打印页数[#/#]
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('filenamedatatimeonheader','1');"+$CRLF       '页眉处包含文件名、日期和时间
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('topmargin','10');"+$CRLF                     '上边距
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('bottommargin','10');"+$CRLF                  '下边距
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('leftmargin','10');"+$CRLF                    '左边距
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rightmargin','10');"+$CRLF                   '右边距
  '热键
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Commands','New,Open,Save,Find,Find Next,Replace,Goto Line,Cut," & _
                                                    "Copy,Paste,Delete,Select All,Comment,UnComment,Tab Indent,Tab Outdent,Space Indent,Space Outdent');"+$CRLF '命令列表
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('New','Ctrl + N');"+$CRLF                    '新建
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Open','Ctrl + O');"+$CRLF                    '打开
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Save','Ctrl + S');"+$CRLF                    '保存
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Find','Ctrl + F');"+$CRLF                    '查找
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Find Next','F3');"+$CRLF                     '查找下一个
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Replace','Ctrl + H');"+$CRLF                 '替换
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Goto Line','Ctrl + G');"+$CRLF               '转到行
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Cut','Ctrl + X');"+$CRLF                     '剪切
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Copy','Ctrl + C');"+$CRLF                    '复制
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Paste','Ctrl + V');"+$CRLF                   '粘贴
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Delete','Del');"+$CRLF                       '删除
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Select All','Ctrl + A');"+$CRLF              '全选
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Comment','Ctrl + Q');"+$CRLF                 '注释
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('UnComment','Ctrl + Shift + Q');"+$CRLF       '反注释
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Tab Indent','Tab');"+$CRLF                   'Tab缩进
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Tab Outdent','Shift + Tab');"+$CRLF          'Tab反缩进
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Space Indent','Space');"+$CRLF               '空格缩进
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Space Outdent','Shift + Space');"+$CRLF      '空格反缩进
  sSql=sSql+"COMMIT;"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for insert "+$DEFAULTCONFIGTABLENAME+","
    GOTO abort
  END IF
  ' 初始化配置表
  REPLACE $DEFAULTCONFIGTABLENAME WITH $CONFIGTABLENAME IN sSql
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for insert "+$CONFIGTABLENAME+","
    GOTO abort
  END IF
  ' 数据集
  sSql="BEGIN TRANSACTION;"+$CRLF
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('language','English');"+$CRLF    '英语
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('language','Chinese');"+$CRLF    '中文
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','2x2');"+$CRLF    '网格大小
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','4x4');"+$CRLF    '网格大小
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','6x6');"+$CRLF    '网格大小
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','8x8');"+$CRLF    '网格大小
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','10x10');"+$CRLF    '网格大小
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','12x12');"+$CRLF    '网格大小
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','2');"+$CRLF    '水平制表符相当的空格数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','3');"+$CRLF    '水平制表符相当的空格数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','4');"+$CRLF    '水平制表符相当的空格数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','6');"+$CRLF    '水平制表符相当的空格数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','8');"+$CRLF    '水平制表符相当的空格数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','1');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','2');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','3');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','4');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','5');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','6');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','7');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','8');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','9');"+$CRLF'最近文件数
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','10');"+$CRLF'最近文件数
  sSql=sSql+"COMMIT;"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for insert "+$DATASETTABLENAME+","
    GOTO abort
  END IF
  ' 字符串语言表
  sSql="BEGIN TRANSACTION;"+$CRLF
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('language','language','语言');"+$CRLF    '语言
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('English','English','英语');"+$CRLF      '英文
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Chinese','Chinese','中文');"+$CRLF      '中文
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('The program is executing!','The program is executing!','程序正在运行!');"+$CRLF '程序运行提示
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Visual PowerBASIC','Visual PowerBASIC','Visual PowerBASIC');"+$CRLF '程序名
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New Project','New Project','新项目');"+$CRLF        '向导窗口标题
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Don''t show this dialog again','Don''t show this dialog again','启动时不再显示本对话框');"+$CRLF        '向导窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New','New','新建');"+$CRLF        '向导窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Exist','Exist','存在');"+$CRLF        '向导窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Recent','Recent','最近');"+$CRLF        '向导窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Search scope:','Search scope:','搜索范围:');"+$CRLF        '向导窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('File name:','File name:','文件名:');"+$CRLF        '向导窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('File type:','File type:','文件类型:');"+$CRLF        '向导窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Normal Window','Normal Window','一般窗口');"+$CRLF  '创建窗口菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('MDI Window','MDI Window','MDI窗口');"+$CRLF         '创建窗口菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Top Align','Top Align','上对齐');"+$CRLF            '控件对齐菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Horizontal Center Align','HorizontalCenter Align','水平中对齐');"+$CRLF  '控件对齐菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Bottom Align','Bottom Align','下对齐');"+$CRLF      '控件对齐菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Left Align','Left Align','左对齐');"+$CRLF          '控件对齐菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Vertical Center Align','Vertical Center Align','垂直中对齐');"+$CRLF  '控件对齐菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Right Align','Right Align','右对齐');"+$CRLF        '控件对齐菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('DDT Code','DDT Code','DDT代码');"+$CRLF             '代码格式菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('SDK Code','SDK Code','SDK代码');"+$CRLF             '代码格式菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Wait for develop...','Wait for develop...','待开发');"+$CRLF  '待开发提示
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('procedure','procedure','程序');"+$CRLF              'SUB/FUNCTION列表工具栏前提示
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Text File','Text File','文本文件');"+$CRLF              '打开与保存对话框用到的文件扩展名过滤字符串
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('All File','All File','所有文件');"+$CRLF                '打开与保存对话框用到的文件扩展名过滤字符串
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Source Code File','Source Code File','源代码文件');"+$CRLF '打开与保存对话框用到的文件扩展名过滤字符串
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Include File','Include File','包含文件');"+$CRLF        '打开与保存对话框用到的文件扩展名过滤字符串
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Resource File','Resource File','资源文件');"+$CRLF      '打开与保存对话框用到的文件扩展名过滤字符串
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Project File','Project File','项目文件');"+$CRLF        '打开与保存对话框用到的文件扩展名过滤字符串
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('MDI text editor, Visualize form editor, Project manager'," + _
                            "'MDI text editor, Visualize form editor, Project manager','MDI文本编辑器，可视化窗体编辑器，项目管理器');"+$CRLF        '用于关于对话框中的功能说明
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('ToolBar','ToolBar','工具栏');"+$CRLF
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Message','Message','消息');"+$CRLF
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show As','Show As','显示为');"+$CRLF
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Orgi String','Orgi String','原字符串');"+$CRLF
  '热键
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Hotkey','Hotkey','热键');"+$CRLF                          '热键
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Defined Hotkey:','Defined Hotkey:','已定义热键:');"+$CRLF '已定义热键
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Delete','Delete','删除');"+$CRLF                          '删除热键
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Command:','Command:','命令');"+$CRLF                      '命令
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Hotkey:','Hotkey:','热键');"+$CRLF                        '热键标签
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Update','Update','更新');"+$CRLF                          '更新


  'menu
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('File','File','文件');"+$CRLF                      '文件菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New','New','新建');"+$CRLF                        '新建菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Open','Open','打开');"+$CRLF                      '打开菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Reopen...','Reopen...','再次打开...');"+$CRLF     '再次打开菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Open Project','Open Project','打开项目');"+$CRLF  '打开项目菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save','Save','保存');"+$CRLF                      '保存菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save As','Save As','另存为');"+$CRLF              '另存为菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save All','Save All','保存全部');"+$CRLF          '保存全部菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save Project','Save Project','保存项目');"+$CRLF  '保存项目菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print Setting','Print Setting','打印设置');"+$CRLF'打印设置
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print Preview','Print Preview','打印预览');"+$CRLF'打印预览
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print...','Print...','打印...');"+$CRLF           '打印
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Close','Close','关闭');"+$CRLF                    '关闭
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Close All','Close All','关闭全部');"+$CRLF        '关闭全部
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Open Command','Open Command','打开命令行');"+$CRLF'打开命令行
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Exit','Exit','退出');"+$CRLF                      '退出

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Edit','Edit','编辑');"+$CRLF                      '编辑
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Undo','Undo','撤销');"+$CRLF                      '撤销
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Redo','Redo','重做');"+$CRLF                      '重做
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Cut','Cut','剪切');"+$CRLF                        '剪切
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Copy','Copy','复制');"+$CRLF                      '复制
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Paste','Paste','粘贴');"+$CRLF                    '粘贴
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Select all','Select all','全选');"+$CRLF          '全选
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Block','Block','块操作');"+$CRLF                  '块操作
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Comment','Comment','注释');"+$CRLF                '注释
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('UnComment','UnComment','反注释');"+$CRLF          '反注释
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tab Indent','Tab Indent','Tab缩进');"+$CRLF       'Tab缩进
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tab Outdent','Tab Outdent','Tab反缩进');"+$CRLF   'Tab反缩进
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Space Indent','Space Indent','空格缩进');"+$CRLF  '空格缩进
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Space Outdent','Space Outdent','Space反缩进');"+$CRLF         'Space反缩进
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Insert GUID','Insert GUID','插入GUID');"+$CRLF    '插入GUID
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Find...','Find...','查找...');"+$CRLF             '查找
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Find Next','Find Next','查找下一个');"+$CRLF      '查找下一个
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Replace...','Replace...','替换...');"+$CRLF       '替换
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Goto Line...','Goto Line...','转到行...');"+$CRLF '转到行
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Goto Bookmark...','Goto Bookmark...','转到书签...');"+$CRLF   '转到书签

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Run','Run','运行');"+$CRLF                        '运行
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compile','Compile','编译');"+$CRLF                '编译
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compile and Run','Compile and Run','编译并运行');"+$CRLF      '编译并运行
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Set Primary File','Set Primary File','设置主文件');"+$CRLF    '设置主文件
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Set Command Parameter','Set Command Parameter','设置命令行参数');"+$CRLF   '设置命令行参数

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('View','View','视图');"+$CRLF                      '视图
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Control ToolBox','Control ToolBox','控件箱');"+$CRLF          '控件箱
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Project ToolBox','Project ToolBox','项目窗口');"+$CRLF        '项目窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Property ToolBox','Property ToolBox','属性窗口');"+$CRLF      '属性窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Find Result Window','Find Result Window','查找结果窗口');"+$CRLF      '查找结果窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compile Info Window','Compile Info Window','编译信息窗口');"+$CRLF    '编译信息窗口

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tool','Tool','工具');"+$CRLF                  '工具
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Make Menu','Make Menu','制作菜单');"+$CRLF        '制作菜单
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Make Toolbar','Make Toolbar','制作工具栏');"+$CRLF            '制作工具栏
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Find Multi Dir','Find Multi Dir','多目录查找');"+$CRLF        '多目录查找
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Replace Multi Dir','Replace Multi Dir','多目录替换');"+$CRLF  '多目录替换
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Format code','Format code','代码格式化');"+$CRLF              '代码格式化
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('WinSpy','WinSpy','WinSpy');"+$CRLF                'WinSpy
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Code Counter','Code Counter','代码统计');"+$CRLF  '代码统计
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Code Manager','Code Manager','代码管理');"+$CRLF  '代码管理
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Custom Tool Setting','Custom Tool Setting','自定义工具设置');"+$CRLF   '自定义工具设置

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Project','Project','项目');"+$CRLF                '项目
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New Project...','New Project...','新项目');"+$CRLF                '项目
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save Project','Save Project','项目');"+$CRLF      '保存项目
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save as Project...','Save as Project...','另存项目...');"+$CRLF     '另存项目
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Achive...','Achive...','归档');"+$CRLF            '归档
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New Vision','New Vision','新建版本');"+$CRLF      '新建版本
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Set Vision for files','Set Vision for files','整理版本');"+$CRLF    '整理版本
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Package Vision...','Package Vision...','打包版本');"+$CRLF          '打包版本

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Cooperation','Cooperation','协同');"+$CRLF        '协同
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Coop Setting...','Coop Setting...','协同设置');"+$CRLF   '协同设置
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Send Invite','Send Invite','发送邀请');"+$CRLF    '发送邀请
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Send Visit','Send Visit','发送访问');"+$CRLF      '发送访问
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Download Files','Download Files','下载文件');"+$CRLF'下载文件
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Upload Files','Upload Files','上载文件');"+$CRLF    '上载文件
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Chat...','Chat...','会话...');"+$CRLF             '会话

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Window','Window','窗口');"+$CRLF               '窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Cascade','Cascade','层叠窗口');"+$CRLF            '层叠窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tile horizontally','Tile horizontally','水平优先排列窗口');"+$CRLF   '水平优先排列窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tile vertically','Tile vertically','纵向优先排列窗口');"+$CRLF   '纵向优先排列窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Arrange Icons','Arrange Icons','排列图标');"+$CRLF   '排列图标
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Close All','Close All','全部关闭');"+$CRLF        '全部关闭
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Maximize Window','Maximize Window','最大化窗口');"+$CRLF   '最大化窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Switch to Window','Switch to Window','切换窗口');"+$CRLF   '切换窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Option...','Option...','选项');"+$CRLF            '选项

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Help','Help','帮助');"+$CRLF                      '帮助
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('IDE Help','IDE Help','IDE 帮助');"+$CRLF          'IDE帮助
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB/WIN Content','PB/WIN Content','PB/WIN 目录');"+$CRLF   'PB/WIN 目录
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB/WIN Index','PB/WIN Index','PB/WIN 索引');"+$CRLF   'PB/WIN 索引
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB/WIN Search','PB/WIN Search','PB/WIN 检索');"+$CRLF   'PB/WIN 检索
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Windows SDK Help','Windows SDK Help','Windows SDK 帮助');"+$CRLF   'Windows SDK 帮助
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('About','About','关于');"+$CRLF      '关于
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Update','Update','更新');"+$CRLF    '更新
  '向导窗口
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New Project','New Project','新项目');"+$CRLF    '更新
  '选项窗口中的字符
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Option','Option','选项');"+$CRLF            '选项
  '按钮
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('OK','OK','确定');"+$CRLF    '确定
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Cancel','Cancel','取消');"+$CRLF    '取消
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Apply','Apply','应用');"+$CRLF    '应用
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Default','Default','默认');"+$CRLF    '默认
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('General','General','一般');"+$CRLF    '一般
  '选项列表
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('File','File','文件');"+$CRLF    '文件
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Code Editor','Code Editor','代码编辑器');"+$CRLF    '代码编辑器
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Color','Color','颜色');"+$CRLF    '颜色
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Window Editor','Window Editor','窗体编辑器');"+$CRLF    '窗体编辑器
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compile','Compile','编译');"+$CRLF    '编译
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print','Print','打印');"+$CRLF    '打印
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Hotkey','Hotkey','热键');"+$CRLF    '热键
  '常规选项
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show welcome window when startup','Show welcome window when startup','启动时显示欢迎窗口');"+$CRLF    '%IDC_SHOWWELCOM
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show guilder window after startup','Show guilder window after startup','启动后显示向导窗口');"+$CRLF    '%IDC_SHOWGUILD
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Main window pos && size','Main window pos && size','主窗口位置及大小');"+$CRLF    '%IDC_MAINWINDOWPOS
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use default position and size','Use default position and size','使用默认位置及大小');"+$CRLF    '%IDC_DEFAULTPOSSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use fix position and size','Use fix position and size','使用固定位置及大小');"+$CRLF    '%IDC_FIXPOSSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use the position and size for last','Use the position and size for last','使用上次的位置及大小');"+$CRLF    '%IDC_USELASTPOSSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Position','Position','位置');"+$CRLF    '%IDC_POSITION
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Center','Center','居中');"+$CRLF    '%IDC_SETCENTER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Special','Special','指定');"+$CRLF    '%IDC_SETPOSITION
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Size','Size','大小');"+$CRLF    '%IDC_MAINWINDOWSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Width:','Width:','宽度:');"+$CRLF    '%IDC_LABELWIDTH
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Height:','Height:','高度:');"+$CRLF    '%IDC_LABELHEIGHT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Please select language for this application:','Please select language for this application:','选择语言:');"+$CRLF    '%IDC_LANGUAGE
  '文件选项
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Type','Type','类型');"+$CRLF    '%IDC_FILETYPE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Source File:','PB Source File:','PB源文件:');"+$CRLF    '%IDC_LABELSOURCEFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Head File:','PB Head File:','PB头文件:');"+$CRLF    '%IDC_LABELINCLUDEFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('RC File:','RC File:','RC文件:');"+$CRLF    '%IDC_LABELRCFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('RC Head File:','RC Head File:','RC头文件:');"+$CRLF    '%IDC_LABELRCHEADFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Project File:','PB Project File:','PB项目文件:');"+$CRLF    '%IDC_LABELPROJECTFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Register these file extend in windows','Register these file extend in windows','在Windows中注册这些文件扩展名');"+$CRLF    '%IDC_SETFILEEXT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Reload previous file set at startup','Reload previous file set at startup','启动时重载之前的文件集');"+$CRLF    '%IDC_LOADRECENTFILES
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Maximum Recent Files:','Maximum Recent Files:','记录最近文件最大个数:');"+$CRLF    '%IDC_LABELMAXFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Initial startup folder','Initial startup folder','初始启动文件夹');"+$CRLF    '%IDC_STARTUPFOLDER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Start in Default folder','Start in Default folder','默认文件夹');"+$CRLF    '%IDC_DEFAULTFOLDER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Start in Last folder','Start in Last folder','最后使用的文件夹');"+$CRLF    '%IDC_LASTFOLDER
  '代码编辑器选项
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Auto indent','Auto indent','自动缩进');"+$CRLF '%IDC_AUTOINDENT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Keep Undo after save','Keep Undo after save','保存后仍可撤销');"+$CRLF '%IDC_KEEPUNDO
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tab size:','Tab size:','Tab大小:');"+$CRLF '%IDC_LABELTABSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Auto save when','Auto save when','自动保存每隔');"+$CRLF '%IDC_AUTOSAVE1
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('minutes','minutes','分钟');"+$CRLF '%IDC_AUTOSAVE2
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show line number','Show line number','显示行数');"+$CRLF '%IDC_SHOWLINENUMBER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show auto complete tip','Show auto complete tip','显示自动完成提示');"+$CRLF '%IDC_SHOWAUTOCOMPLETE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Keyword Case','Keyword Case','关键字大小写');"+$CRLF '%IDC_KEYWORDCASE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('No Case Change','No Case Change','不改变');"+$CRLF '%IDC_NOCASECHANGE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Upper Case','Upper Case','大写');"+$CRLF '%IDC_UPPERCASE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Mixed Case','Mixed Case','首字母大写');"+$CRLF '%IDC_MIXEDCASE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Lower Case','Lower Case','小写');"+$CRLF '%IDC_LOWERCASE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Editor Font','Editor Font','编辑器字体');"+$CRLF '%IDC_EDITORFONT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Font:','Font:','字体:');"+$CRLF '%IDC_LABELFONT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Size:','Size:','大小:');"+$CRLF '%IDC_LABELFONTSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Sample text','Sample text','示例文本');"+$CRLF '%IDC_LABELSAMPLETEXT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Collapse && Expand','Collapse && Expand','卷起 && 展开');"+$CRLF '%IDC_COLLAPSEEXPAND
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Foldding Level','Foldding Level','卷展层级');"+$CRLF '%IDC_COLLAPSELEVEL
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('No Collapse','No Collapse','不卷展');"+$CRLF '%IDC_NOCOLLAPSE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Keyword level','Keyword level','关键字级');"+$CRLF '%IDC_KEYWORDLEVEL
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Sub/Function level','Sub/Function level','Sub/Function 级');"+$CRLF '%IDC_SUBFUNLEVEL
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Folder icon','Folder icon','卷展图标');"+$CRLF '%IDC_FOLDERICON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Arrow','Arrow','箭头');"+$CRLF '%IDC_ARROWICON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Circle','Circle','圆形');"+$CRLF '%IDC_CIRCLEICON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Square','Square','方形');"+$CRLF '%IDC_SQUAREICON
  '颜色
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use Syntax Color in Editor','Use Syntax Color in Editor','在代码编辑器中使用语法颜色');"+$CRLF '%IDC_USECOLORINEDITOR
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use Syntax Color when Print','Use Syntax Color when Print','打印时使用语法颜色');"+$CRLF '%IDC_USECOLORINPRINTER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Load &Defaults','Load &Defaults','加载默认配色');"+$CRLF '%IDC_LOADDEFAULTCOLOR
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Assembler','Assembler','汇编');"+$CRLF '%IDC_ASSEMBLERBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Comments','Comments','注释');"+$CRLF '%IDC_COMMENTBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Keywords','Keywords','关键字');"+$CRLF '%IDC_KEYWORDBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Forms','PB Forms','PB窗体');"+$CRLF '%IDC_PBFORMSBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Strings','Strings','字符串');"+$CRLF '%IDC_STRINGBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Text','Text','文本');"+$CRLF '%IDC_TEXTBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Selection','Selection','选中');"+$CRLF '%IDC_SELECTIONBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Breakpoint','Breakpoint','断点');"+$CRLF '%IDC_BREAKPOINTBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Bookmark','Bookmark','书签');"+$CRLF '%IDC_BOOKMARKBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Exec point','Exec point','执行点');"+$CRLF '%IDC_EXECPOINTBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Reserved','Reserved','保留');"+$CRLF '%IDC_RESERVED1BUTTON %IDC_RESERVED2BUTTON
  '窗口编辑器
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show grid','Show grid','显示网格');"+$CRLF '%IDC_SHOWGRID
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Snap to grid','Snap to grid','对齐到网格');"+$CRLF '%IDC_GRIDSNAP
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Grid size:','Grid size:','网格大小:');"+$CRLF '%IDC_GRIDSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Snap between controls','Snap between controls','控件间对齐');"+$CRLF '%IDC_SNAPCONTROL
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Adjust handle size','Adjust handle size','调整块大小');"+$CRLF '%IDC_ADJUSTHANDLE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Small','Small','小');"+$CRLF '%IDC_SMALLHANDLE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Medium','Medium','中');"+$CRLF '%IDC_MEDIUMHANDLE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Large','Large','大');"+$CRLF '%IDC_LARGEHANDLE
  '编译器
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Paths','Paths','路径');"+$CRLF '%IDC_PATH
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Include:','PB Include:','PB包含:');"+$CRLF '%IDC_LABELPBINCLUDE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('RC Include:','RC Include:','RC包含:');"+$CRLF '%IDC_LABELRCINCLUDE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compiler File:','Compiler File:','编译器文件:');"+$CRLF '%IDC_LABELCOMPILER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Win32.hlp File:','Win32.hlp File:','Win32.hlp文件:');"+$CRLF '%IDC_LABELWIN32HELP
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compiler Options','Compiler Options','编译选项');"+$CRLF '%IDC_COMPILEOPTION
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Beep on compiletion','Beep on compiletion','编译结束时嘀一声');"+$CRLF '%IDC_BEEPCOMPILETION
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Display Results','Display Results','显示编译结果');"+$CRLF '%IDC_DISPLAYRESULT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Create Log file','Create Log file','创建日志文件');"+$CRLF '%IDC_CREATELOGFILE
  '打印
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use the color for code in editor when print','Use the color for code in editor when print','打印时使用代码在编辑器中的颜色');"+$CRLF '%IDC_USERCODECOLOR
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print page number on footer with the format [#/#]','Print page number on footer with the format [#/#]','在页脚以格式[#/#]打印页码');"+$CRLF '%IDC_PRINTPAGENUMBER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print filename, time and date of the file on header','Print filename, time and date of the file on header','在页眉打印文件名，时间和日期');"+$CRLF '%IDC_PRINTFILENAMETIME
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Margin','Margin','页边距');"+$CRLF '%IDC_PAGEMARGIN
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Top:','Top:','上边:');"+$CRLF '%IDC_LABELTOPMARGIN
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Bottom:','Bottom:','下边:');"+$CRLF '%IDC_LABELBOTTOMMARGIN
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Left:','Left:','左边:');"+$CRLF '%IDC_LABELLEFTMARGIN
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Right:','Right:','右边:');"+$CRLF '%IDC_LABELRIGHTMARGIN
  '热键
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Defined Hotkey:','Defined Hotkey:','已定义的热键:');"+$CRLF '%IDC_LABELHADHOTKEY
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Delete','Delete','删除');"+$CRLF '%IDC_DELETEHOTKEY
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Command:','Command:','命令:');"+$CRLF '%IDC_LABELHOTKEYCOMMAND
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Hotkey:','Hotkey:','热键:');"+$CRLF '%IDC_LABELHOTKEY
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Update','Update','更新');"+$CRLF '%IDC_UPDATEHOTKEY

  sSql=sSql+"COMMIT;"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for insert "+$DATASETTABLENAME+","
    GOTO abort
  END IF
  sqlClose hDB
  FUNCTION=1
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=0
END FUNCTION
' WriteConfig
' 写入配置项
' 输入配置项名，值，全局参数：表名t_config
FUNCTION WriteConfig(BYVAL ConfigName AS STRING,BYVAL value AS STRING)AS LONG
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  IF Exist($DATABASENAME)=0 THEN '如果数据库文件不存在，则创建数据库
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "数据异常 0001"
      FUNCTION=0
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="无法打开或创建数据库,"
    GOTO abort
  END IF
  sSql="update "+$CONFIGTABLENAME+" set value='"+value+"' where name='"+ConfigName+"'"
  IF sqlExe(hDB,sSql)<0 THEN
    sSql="insert "+$CONFIGTABLENAME+" (name,value)values('"+ConfigName+"','"+value+"')"
    IF sqlExe(hDB,sSql)<0 THEN
      errStr=sqlErrMsg(hDB) & " for insert "+$CONFIGTABLENAME+","
      GOTO abort
    END IF
  END IF
  sqlClose hDB
  FUNCTION=1
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=0
END FUNCTION
'读取配置
'传入参数为配置名
FUNCTION ReadConfig(BYVAL ConfigName AS STRING)AS STRING
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL rs() AS STRING
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '如果数据库文件不存在，则创建数据库
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "数据异常 0001"
      FUNCTION=""
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="无法打开或创建数据库,"
    GOTO abort
  END IF
  sSql="select value from "+$CONFIGTABLENAME+" where name='"+ConfigName+"'"
  IF sqlQuery(hDB,sSql,rs())<0 THEN
    errStr="数据库错误，查询" & $CONFIGTABLENAME & ","
    GOTO abort
  END IF
  tmpStr=""
  IF UBOUND(rs)<=0 THEN
    FUNCTION=tmpStr
    sqlClose hDB
    EXIT FUNCTION
  END IF
  tmpStr=rs(1,0)
  sqlClose hDB
  FUNCTION=tmpStr
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=""
END FUNCTION
'读取默认配置
FUNCTION ReadDefaultConfig(BYVAL ConfigName AS STRING)AS STRING
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL rs() AS STRING
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '如果数据库文件不存在，则创建数据库
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "数据异常 ReadDefaultConfig"
      FUNCTION=""
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="无法打开或创建数据库,"
    GOTO abort
  END IF
  sSql="select value from "+$DEFAULTCONFIGTABLENAME+" where name='"+ConfigName+"'"
  IF sqlQuery(hDB,sSql,rs())<0 THEN
    errStr="数据库错误，查询" & $DEFAULTCONFIGTABLENAME & ","
    GOTO abort
  END IF
  tmpStr=""
  IF UBOUND(rs)<=0 THEN
    FUNCTION=tmpStr
    sqlClose hDB
    EXIT FUNCTION
  END IF
  tmpStr=rs(1,0)
  sqlClose hDB
  FUNCTION=tmpStr
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=""
END FUNCTION
'读取数据集
FUNCTION ReadDataset(BYVAL datasetName AS STRING)AS STRING
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL rs() AS STRING
  LOCAL i AS INTEGER
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '如果数据库文件不存在，则创建数据库
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "数据异常 ReadDataset"
      FUNCTION=""
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="无法打开或创建数据库,"
    GOTO abort
  END IF
  sSql="select value from "+$DATASETTABLENAME+" where type='"+datasetName+"'"
  IF sqlQuery(hDB,sSql,rs())<0 THEN
    errStr="数据库错误，查询" & $DATASETTABLENAME & ","
    GOTO abort
  END IF
  tmpStr=""
  IF UBOUND(rs)<=0 THEN
    FUNCTION=tmpStr
    sqlClose hDB
    EXIT FUNCTION
  END IF
  FOR i=1 TO UBOUND(rs(1))
    IF tmpStr="" THEN
      tmpStr=rs(i,0)
    ELSE
      tmpStr=tmpStr & "|" & rs(i,0)
    END IF
  NEXT i
  'tmpStr=rs(1,0)
  sqlClose hDB
  FUNCTION=tmpStr
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=""
END FUNCTION
'读取某个原始字符串的相应的语言对应的字符串
FUNCTION ReadLang(BYVAL OriString AS STRING,BYVAL selLang AS STRING)AS STRING
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL rs() AS STRING
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '如果数据库文件不存在，则创建数据库
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "数据异常 ReadLang"
      FUNCTION=""
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="无法打开或创建数据库,"
    GOTO abort
  END IF
  sSql="select "+selLang+" from "+$STRINGLANGTABLENAME+" where name='"+OriString+"'"
  IF sqlQuery(hDB,sSql,rs())<0 THEN
    errStr="数据库错误，查询" & $STRINGLANGTABLENAME & ","
    GOTO abort
  END IF
  tmpStr=""
  IF UBOUND(rs)<=0 THEN
    FUNCTION=tmpStr
    sqlClose hDB
    EXIT FUNCTION
  END IF
  tmpStr=rs(1,0)
  sqlClose hDB
  FUNCTION=tmpStr
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=""
END FUNCTION
'写（增加或修改）某个原始字符串的相应的语言对应的字符串
FUNCTION WriteLang(BYVAL OriString AS STRING,BYVAL ShowString AS STRING)AS LONG
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL tmpLng AS LONG
  LOCAL rs()   AS STRING
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '如果数据库文件不存在，则创建数据库
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "数据异常 WriteLang"
      FUNCTION=0
      EXIT FUNCTION
    ELSE
      InitConfig
      MSGBOX "init..."
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="无法打开或创建数据库,"
    GOTO abort
  END IF
  sSql="select id from "+$STRINGLANGTABLENAME+" where name='"+OriString+"'"
  tmpLng=sqlQuery(hDB,sSql,rs())
  IF tmpLng<0 THEN
    errStr="数据库错误，查询" & sSql & ","
    GOTO abort
  ELSEIF tmpLng=0 THEN
    sSql="BEGIN TRANSACTION;"+$CRLF
    sSql=sSql & "insert into "+$STRINGLANGTABLENAME+" (name," + LCASE$(language) + ")values('"+OriString+"','"+ShowString+"');" & $CRLF
    sSql=sSql & "COMMIT;"
  ELSE
    sSql="BEGIN TRANSACTION;"+$CRLF
    sSql=sSql & "update "+$STRINGLANGTABLENAME+" set "+LCASE$(language)+"='" + ShowString+"' where name='"+OriString+"';" & $CRLF
    sSql=sSql & "COMMIT;"
  END IF
  tmpLng=sqlExe(hDB,sSql)
  sqlClose hDB
  'MSGBOX "test3 " & language & " " & GetLang(OriString) & $CRLF & sSql
  MSGBOX "已成功地将 " & $CRLF & OriString & $CRLF & " 的" & language & "语言定义为 " & $CRLF & GetLang(OriString),_
          %MB_ICONINFORMATION,GetLang("Message")
  RunLog $CRLF & sSql
  FUNCTION=1
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=0
END FUNCTION
'根据原始字符串返回当前语言相对应的字符串
'查找失败时，返回原始字符串
FUNCTION GetLang(BYVAL OriStr AS STRING)AS STRING
  LOCAL tmpStr AS STRING
  tmpStr=OriStr
  language=ReadConfig("language")
  IF language="" THEN
    language="English"
  END IF
  tmpStr=ReadLang(tmpStr,language)
  IF tmpStr="" THEN
    FUNCTION = OriStr
  ELSE
    FUNCTION = tmpStr
  END IF
END FUNCTION
FUNCTION RunLog1(BYVAL logStr AS STRING)AS LONG
  LOCAL fn AS LONG
  LOCAL st AS SYSTEMTIME
  IF RIGHT$(logStr,2)<>$CRLF THEN
    logStr=logStr & $CRLF
  END IF
  GetLocalTime st
  fn=FREEFILE
  OPEN EXE.PATH$ & "log.txt" FOR APPEND AS #fn
  PRINT #fn,FORMAT$(st.wYear,"0000") & "/" & FORMAT$(st.wMonth,"00") & "/" & FORMAT$(st.wDay,"00") & " " _
            & FORMAT$(st.wHour,"00") & ":" & FORMAT$(st.wMinute,"00") & ":" & FORMAT$(st.wSecond,"00") & "." _
            & FORMAT$(st.wMilliseconds,"000") & "  ";
  PRINT #fn,logStr;
  CLOSE #fn
END FUNCTION
' 记录日志，用于调试
FUNCTION RunLog(BYVAL logStr AS STRING)AS LONG
  LOCAL fn AS LONG
  LOCAL st AS SYSTEMTIME
  LOCAL timeStr AS STRING
  LOCAL filename AS STRING
  IF ISFALSE canLog THEN
    EXIT FUNCTION
  END IF
  GetLocalTime st
  fn=FREEFILE
  IF RIGHT$(logStr,2)<>$CRLF THEN
    logStr=logStr & $CRLF
  END IF
  filename  = FORMAT$(st.wYear,"0000") & FORMAT$(st.wMonth,"00") & FORMAT$(st.wDay,"00") & ".log"
  timeStr= FORMAT$(st.wYear,"0000") & "/" & FORMAT$(st.wMonth,"00") & "/" _
          & FORMAT$(st.wDay,"00") & "-" & FORMAT$(st.wHour,"00") & ":" _
          & FORMAT$(st.wMinute,"00") & ":" & FORMAT$(st.wSecond,"00") & "." _
          & FORMAT$(st.wMilliseconds,"000")
  OPEN EXE.PATH$ & filename FOR APPEND AS #fn
  PRINT #fn,timeStr & " -- " & logStr;
  CLOSE #fn
END FUNCTION
'-------------------------------------------------------------------------
'创建标准字体:
'-------------------------------------------------------------------------
FUNCTION MakeStdFont(BYVAL Font1 AS STRING, BYVAL PointSize AS LONG) AS LONG
  LOCAL hDC AS LONG, CyPixels AS LONG
  hDC = GetDC(%HWND_DESKTOP)
  CyPixels  = GetDeviceCaps(hDC, %LOGPIXELSY)
  ReleaseDC %HWND_DESKTOP, hDC
  PointSize = (PointSize * CyPixels) \ 72
  FUNCTION = CreateFont(0 - PointSize, 0, 0, 0, %FW_NORMAL, 0, 0, 0, _
                          %ANSI_CHARSET, %OUT_TT_PRECIS, %CLIP_DEFAULT_PRECIS, _
                          %DEFAULT_QUALITY, %FF_DONTCARE, BYCOPY Font1)
END FUNCTION
' 点变量值转字符串
FUNCTION PtToStr(BYVAL pt AS POINTAPI)AS STRING
  FUNCTION=FORMAT$(pt.x) & "," & FORMAT$(pt.y)
END FUNCTION
' 字符串转点变量
FUNCTION StrToPt(BYVAL ptStr AS STRING,pt AS POINTAPI)AS LONG
  LOCAL tmpStr AS STRING
  tmpStr=TRIM$(ptStr)
  IF tmpStr="" OR PARSECOUNT(tmpStr,",")<4 THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  pt.x=VAL(PARSE$(tmpStr,",",1))
  pt.y=VAL(PARSE$(tmpStr,",",2))
  FUNCTION=1
END FUNCTION
' 矩形值转字符串
FUNCTION RcToStr(BYVAL rc AS RECT)AS STRING
  FUNCTION=FORMAT$(rc.nLeft) & "," & FORMAT$(rc.nTop) & "," _
          & FORMAT$(rc.nRight) & "," & FORMAT$(rc.nBottom)
END FUNCTION
' 字符串转矩形
FUNCTION StrToRc(BYVAL str AS STRING,rc AS RECT)AS LONG
  LOCAL tmpStr AS STRING
  tmpStr=TRIM$(str)
  IF tmpStr="" OR PARSECOUNT(tmpStr,",")<4 THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  rc.nLeft    = VAL(PARSE$(tmpStr,",",1))
  rc.nTop     = VAL(PARSE$(tmpStr,",",2))
  rc.nRight   = VAL(PARSE$(tmpStr,",",3))
  rc.nBottom  = VAL(PARSE$(tmpStr,",",4))
  FUNCTION=1
END FUNCTION
' 保存主窗口位置宽高
FUNCTION SaveMainRc()AS LONG
  LOCAL rc AS RECT
  LOCAL tmpStr AS STRING
  GetWindowRect g_hWndMain,rc
  rc.nRight=rc.nRight-rc.nLeft 'nRight转成存宽度
  rc.nBottom=rc.nBottom-rc.nTop'nBottom转成存高度
  rc.nTop=IIF(rc.nTop<0,0,rc.nTop)
  rc.nLeft=IIF(rc.nLeft<0,0,rc.nLeft)
  tmpStr=RcToStr(rc)
  IniWrite EXE.PATH$ & "config.ini","main","rc",tmpStr
END FUNCTION
' 从配置文件获取主窗口矩形：位置宽高
FUNCTION GetMainRc(rc AS RECT) AS LONG
  LOCAL tmpStr AS STRING
  tmpStr=IniRead(EXE.PATH$ & "config.ini","main","rc","")
  FUNCTION=StrToRc(tmpStr,rc)
END FUNCTION
' 设置主窗口位置及尺寸
FUNCTION SetWindowPosSize(BYVAL hWnd AS DWORD)AS LONG
  LOCAL Zoomed AS LONG
  LOCAL WinPla AS WINDOWPLACEMENT
  LOCAL rcDeskTop AS RECT
  LOCAL rc AS RECT
  LOCAL rs AS STRING
  WinPla.Length = SIZEOF(WinPla)
  IF WinPla.showCmd = %SW_SHOWMAXIMIZED THEN
    Zoomed = %TRUE
  END IF
  rs = IniRead (g_zIni, "main", "Zoomed", "")
  IF LEN(rs) THEN
    Zoomed=VAL(rs)
  ELSE
    EXIT FUNCTION
  END IF
  GetMainRc WinPla.rcNormalPosition
  WinPla.showCmd=0
  SystemParametersInfo %SPI_GETWORKAREA, 0, rcDesktop, 0
  '确保窗口宽度不超出屏幕宽度
  IF WinPla.rcNormalPosition.nRight - WinPla.rcNormalPosition.nLeft > rcDesktop.nRight THEN
     WinPla.rcNormalPosition.nLeft  = 0
     WinPla.rcNormalPosition.nRight = rcDesktop.nRight
   END IF
  '确保窗口高度不超出屏幕高度
  IF WinPla.rcNormalPosition.nBottom - WinPla.rcNormalPosition.nTop > rcDesktop.nBottom THEN
     WinPla.rcNormalPosition.nTop    = 0
     WinPla.rcNormalPosition.nBottom = rcDesktop.nBottom
   END IF
  '确保窗口左侧边可见
  IF WinPla.rcNormalPosition.nLeft < 0 THEN
     WinPla.rcNormalPosition.nRight = WinPla.rcNormalPosition.nRight - WinPla.rcNormalPosition.nLeft
     WinPla.rcNormalPosition.nLeft = 0
  END IF
  '确保窗口右侧边可见
  IF WinPla.rcNormalPosition.nRight > rcDesktop.nRight THEN
     WinPla.rcNormalPosition.nLeft = WinPla.rcNormalPosition.nLeft - (WinPla.rcNormalPosition.nRight - rcDesktop.nRight)
     WinPla.rcNormalPosition.nRight = rcDesktop.nRight
  END IF
  '确保窗口上边可见
  IF WinPla.rcNormalPosition.nTop < 0 THEN
     WinPla.rcNormalPosition.nBottom = WinPla.rcNormalPosition.nBottom - WinPla.rcNormalPosition.nTop
     WinPla.rcNormalPosition.nTop = 0
  END IF
  '确保窗口下边可见
  IF WinPla.rcNormalPosition.nBottom > rcDesktop.nBottom THEN
     WinPla.rcNormalPosition.nTop = WinPla.rcNormalPosition.nTop - (WinPla.rcNormalPosition.nBottom - rcDesktop.nBottom)
     WinPla.rcNormalPosition.nBottom = rcDesktop.nBottom
  END IF
  SystemParametersInfo %SPI_GETWORKAREA, 0, BYVAL VARPTR(rc), 0
  IF WinPla.rcNormalPosition.nLeft = rc.nLeft AND  WinPla.rcNormalPosition.nRight = rc.nRight THEN WinPla.rcNormalPosition.nRight = WinPla.rcNormalPosition.nRight + 2
  IF WinPla.rcNormalPosition.nTop = rc.nTop AND  WinPla.rcNormalPosition.nBottom = rc.nBottom THEN WinPla.rcNormalPosition.nBottom = WinPla.rcNormalPosition.nBottom + 2
  WinPla.Length = SIZEOF(WinPla)
  SetWindowPlacement hWnd, WinPla
END FUNCTION
' 保存主窗口客户区实体窗口矩形：位置宽高，相对于主窗口
FUNCTION SaveMainClientRc()AS LONG
  LOCAL rc AS RECT
  LOCAL tmpStr AS STRING
  GetWindowRect g_hWndClient,rc
  ScreenToClientRect g_hWndMain,rc
  rc.nRight=rc.nRight-rc.nLeft 'nRight转成存宽度
  rc.nBottom=rc.nBottom-rc.nTop'nBottom转成存高度
  rc.nTop=IIF(rc.nTop<0,0,rc.nTop)
  rc.nLeft=IIF(rc.nLeft<0,0,rc.nLeft)
  tmpStr=RcToStr(rc)
  IniWrite EXE.PATH$ & "config.ini","main","clientrc",tmpStr
END FUNCTION
' 从配置文件获取主窗口客户区实体窗口矩形：位置宽高，相对于主窗口
FUNCTION GetMainClientRc(rc AS RECT) AS LONG
  LOCAL tmpStr AS STRING
  tmpStr=IniRead(EXE.PATH$ & "config.ini","main","clientrc","")
  FUNCTION=StrToRc(tmpStr,rc)
END FUNCTION
'保存Dock信息到配置文件
FUNCTION SaveDockInfo()AS LONG
  LOCAL fname AS STRING
  LOCAL i AS LONG
  fname=EXE.PATH$ & "config.ini"
  IniWrite fname,"dock","count",FORMAT$(UBOUND(gdi())+1) '数量
  IF UBOUND(gdi())<0 THEN
    REDIM gdi()
    lA    = 0 : tA     = 0 : rA   = 0 : bA  = 0
    RESET lARc : RESET tARc : RESET rARc : RESET bARc
    FUNCTION=0
    EXIT FUNCTION
  END IF
  '保存泊坞区
  IniWrite fname,"dock","leftA"   ,FORMAT$(lA)
  IniWrite fname,"dock","topA"    ,FORMAT$(tA)
  IniWrite fname,"dock","rightA"  ,FORMAT$(rA)
  IniWrite fname,"dock","bottomA" ,FORMAT$(bA)
  IniWrite fname,"dock","leftrc"  ,RcToStr(lARc)
  IniWrite fname,"dock","toprc"   ,RcToStr(tARc)
  IniWrite fname,"dock","rightrc" ,RcToStr(rARc)
  IniWrite fname,"dock","bottomrc",RcToStr(bARc)
  '保存DockInfo结构数组
  FOR i=0 TO UBOUND(gdi())
    IniWrite fname,"dock" & FORMAT$(i),"id",FORMAT$(gdi(i).ID)
    IniWrite fname,"dock" & FORMAT$(i),"tid",FORMAT$(gdi(i).tid)
    IniWrite fname,"dock" & FORMAT$(i),"ml",gdi(i).ml
    IniWrite fname,"dock" & FORMAT$(i),"l",gdi(i).l
    IniWrite fname,"dock" & FORMAT$(i),"drc",RcToStr(gdi(i).drc)
    GetWindowRect gdi(i).hWndF,gdi(i).frc
    gdi(i).frc.nRight -=gdi(i).frc.nLeft
    gdi(i).frc.nBottom -=gdi(i).frc.nTop
    IniWrite fname,"dock" & FORMAT$(i),"frc",RcToStr(gdi(i).frc)
    IniWrite fname,"dock" & FORMAT$(i),"minW",FORMAT$(gdi(i).minW)
    IniWrite fname,"dock" & FORMAT$(i),"minH",FORMAT$(gdi(i).minH)
    IniWrite fname,"dock" & FORMAT$(i),"byside",FORMAT$(gdi(i).byside)
  NEXT i
  FUNCTION=1
END FUNCTION
'从配置文件中获取Dock信息，保存到gdi()
FUNCTION GetDockInfo()AS LONG
  LOCAL fname AS STRING
  LOCAL i AS LONG
  LOCAL dockUbd AS LONG
  fname=EXE.PATH$ & "config.ini"
  dockUbd=VAL(IniRead(fname,"dock","count","-1"))-1
  IF dockUbd<0 THEN
    REDIM gdi()
    lA    = 0 : tA     = 0  : rA   = 0  : bA  = 0
    RESET lARc : RESET tARc : RESET rARc :RESET bARc
    FUNCTION=0
    EXIT FUNCTION
  END IF
  lA=VAL(IniRead(fname,"dock","leftA","0"))
  tA=VAL(IniRead(fname,"dock","topA","0"))
  rA=VAL(IniRead(fname,"dock","rightA","0"))
  bA=VAL(IniRead(fname,"dock","bottomA","0"))
  StrToRc(IniRead(fname,"dock","leftrc",""),lARc)
  StrToRc(IniRead(fname,"dock","toprc",""),tARc)
  StrToRc(IniRead(fname,"dock","rightrc",""),rARc)
  StrToRc(IniRead(fname,"dock","bottomrc",""),bARc)
  REDIM gdi(dockUbd)
  FOR i=0 TO UBOUND(gdi())
    gdi(i).ID=VAL(IniRead(fname,"dock" & FORMAT$(i),"id","0"))
    gdi(i).tid=VAL(IniRead(fname,"dock" & FORMAT$(i),"tid","0"))
    gdi(i).ml=IniRead(fname,"dock" & FORMAT$(i),"ml","F")
    gdi(i).l=IniRead(fname,"dock" & FORMAT$(i),"l","F")
    StrToRc(IniRead(fname,"dock" & FORMAT$(i),"drc",""),gdi(i).drc)
    StrToRc(IniRead(fname,"dock" & FORMAT$(i),"frc",""),gdi(i).frc)
    gdi(i).minW=VAL(IniRead(fname,"dock" & FORMAT$(i),"minW","50"))
    gdi(i).minH=VAL(IniRead(fname,"dock" & FORMAT$(i),"minH","150"))
    gdi(i).byside=VAL(IniRead(fname,"dock" & FORMAT$(i),"byside","0"))
    IF gdi(i).ml="F" THEN gdi(i).l="F" : gdi(i).byside=0
    IF gdi(i).l="F" THEN gdi(i).ml="F" : gdi(i).byside=0
  NEXT i
  FUNCTION=1
END FUNCTION
' 将矩形值从窗口相对值转为屏幕值
FUNCTION ClientToScreenRECT(BYVAL temphWnd AS DWORD,BYREF rc AS RECT)AS LONG
  LOCAL tmpPt AS POINTAPI
  tmpPt.x=rc.nLeft
  tmpPt.y=rc.nTop
  ClientToScreen temphWnd,tmpPt
  rc.nLeft=tmpPt.x
  rc.nTop=tmpPt.y
  tmpPt.x=rc.nRight
  tmpPt.y=rc.nBottom
  ClientToScreen temphWnd,tmpPt
  rc.nRight=tmpPt.x
  rc.nBottom=tmpPt.y
END FUNCTION
' 将矩形值从屏幕值转为窗口相对值
FUNCTION ScreenToClientRect(BYVAL temphWnd AS DWORD,BYREF rc AS RECT)AS LONG
  LOCAL tmpPt AS POINTAPI
  tmpPt.x=rc.nLeft
  tmpPt.y=rc.nTop
  ScreenToClient temphWnd,tmpPt
  rc.nLeft=tmpPt.x
  rc.nTop=tmpPt.y

  tmpPt.x=rc.nRight
  tmpPt.y=rc.nBottom
  ScreenToClient temphWnd,tmpPt
  rc.nRight=tmpPt.x
  rc.nBottom=tmpPt.y
END FUNCTION
' 移动窗口，对MoveWindow进行缩减
FUNCTION MoveWin(BYVAL hWnd AS DWORD,BYVAL rc AS RECT,BYVAL rpflag AS BYTE)AS LONG
  IF hWnd<=0 OR ISFALSE IsWindow(hWnd) THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  MoveWindow hWnd,rc.nLeft,rc.nTop,rc.nRight,rc.nBottom,rpflag
  FUNCTION=1
END FUNCTION
FUNCTION SED_MsgBox( BYVAL hWnd AS DWORD, BYVAL strMessage AS STRING, OPTIONAL BYVAL dwStyle AS DWORD, BYVAL strCaption AS STRING ) AS LONG
  LOCAL mbp AS MSGBOXPARAMS
  LOCAL szCaption AS ASCIIZ * 255
  szCaption = " SED Editor"
  IF LEN( strCaption ) THEN szCaption = strCaption
  IF dwStyle = 0 THEN dwStyle = %MB_OK
  ' Initializes MSGBOXPARAMS 初始化MSGBOXPARAMS
  mbp.cbSize = SIZEOF( mbp )      ' Size of the structure结构大小
  mbp.hwndOwner = hWnd        ' Handle of main window 主窗口句柄
  mbp.hInstance = GETMODULEHANDLE( "" )       ' Instance of application应用程序实例
  mbp.lpszText = STRPTR( strMessage )         ' Text of the message消息的文本
  mbp.lpszCaption = VARPTR( szCaption )       ' Caption 标题
  mbp.dwStyle = dwStyle OR %MB_USERICON       ' Style 类型
  mbp.lpszIcon = 100      ' Icon identifier in the resource file 在资源文件中的图标标识
  FUNCTION = MESSAGEBOXINDIRECT( mbp )
END FUNCTION
FUNCTION Ini_Get( BYVAL iniFileName AS STRING, BYVAL iniSect AS STRING, _
          inis AS settings ) AS LONG
  LOCAL sResult AS STRING
  LOCAL iniSection AS STRING
  LOCAL iniFN AS STRING
  LOCAL IniFile AS STRING
  LOCAL localdir AS STRING
  localdir=EXE.PATH$ & "winapi"
  IniFile=EXE.PATH$ & "config.ini"
  iniSection = TRIM$( iniSect )
  iniFN = TRIM$( iniFileName )
  IF iniSection = "" THEN iniSection = "编辑器"
  IF iniFN = "" THEN iniFN = IniFile
  inis.showGrid = INT( VAL( IniRead( iniFN, iniSection, "showGrid", "1" )))
  inis.snapGrid = INT( VAL( IniRead( iniFN, iniSection, "snapGrid", "0" )))
  inis.gridSize = INT( VAL( IniRead( iniFN, iniSection, "gridSize", "8" )))
  inis.stickFrames = INT( VAL( IniRead( iniFN, iniSection, "stickFrames", "0" )))
  inis.autoSizeImg = INT( VAL( IniRead( iniFN, iniSection, "autoSizeImg", "0" )))
  inis.handles = INT( VAL( IniRead( iniFN, iniSection, "handles", FORMAT$( %use_HANDLEMEDIUM ))))
  inis.backgrndType = INT( VAL( IniRead( iniFN, iniSection, "backgrndType", FORMAT$( %use_backColor ))))
  inis.backColor = INT( VAL( IniRead( iniFN, iniSection, "backColor", "0" )))
  inis.backBrush = INT( VAL( IniRead( iniFN, iniSection, "backBrush", "0" )))
  inis.Include = IniRead( iniFile, iniSection, "IncludePath", localdir & "\win32api.inc" )
END FUNCTION
FUNCTION Ini_Set( BYVAL iniFileName AS STRING, BYVAL iniSect AS STRING, inis AS settings ) AS LONG
  LOCAL sResult AS STRING
  LOCAL iniSection AS STRING
  LOCAL iniFN AS STRING
  LOCAL IniFile AS STRING
  IniFile=EXE.PATH$ & "config.ini"
  iniSection = TRIM$( iniSect )
  iniFN = TRIM$( iniFileName )
  IF iniSection = "" THEN iniSection = "编辑器"
  IF iniFN = "" THEN iniFN = IniFile
  IniWrite iniFN, iniSection, "showGrid", FORMAT$( inis.showGrid )
  IniWrite iniFN, iniSection, "snapGrid", FORMAT$( inis.snapGrid )
  IniWrite iniFN, iniSection, "gridSize", FORMAT$( inis.gridSize )
  IniWrite iniFN, iniSection, "stickFrames", FORMAT$( inis.stickFrames )
  IniWrite iniFN, iniSection, "autoSizeImg", FORMAT$( inis.autoSizeImg )
  IniWrite iniFN, iniSection, "handles", FORMAT$( inis.handles )
  IniWrite iniFN, iniSection, "backgrndType", FORMAT$( inis.backgrndType )
  IniWrite iniFN, iniSection, "backColor", FORMAT$( inis.backColor )
  IniWrite iniFN, iniSection, "backBrush", FORMAT$( inis.backBrush )
  IniWrite iniFN, iniSection, "IncludePath", inis.Include
END FUNCTION
' 显示消息提示窗口
FUNCTION MsgInfoBox(BYVAL msgStr AS STRING)AS LONG
  MSGBOX msgStr,%MB_OK OR %MB_ICONINFORMATION,"提示"
END FUNCTION
' 显示错误提示窗口
FUNCTION MsgErrBox(BYVAL errStr AS STRING)AS LONG
  MSGBOX errStr,%MB_OK OR %MB_ICONERROR,"错误"
END FUNCTION
' 显示询问提示窗口
FUNCTION MsgQueBox(BYVAL msgStr AS STRING)AS LONG
  FUNCTION=MSGBOX(msgStr,%MB_YESNO OR %MB_ICONQUESTION,"提示")
END FUNCTION
' *********************************************************************************************
' TITLE: StatusBar_SetPartsBySize
' DESC: Sets n parts of an existing status bar
' SYNTAX: StatusBar_SetPartsBySize(hStatus,sSizes)
' Parameters:
'   hStatus   - The handle of the Status Bar control.
'   sSizes    - Comma delited string of desired sizes
'               e.g.:  sSizes = "90,34,42,184"  'in pixels
'                  creates 4 parts of 90, 34, 42 and 184 pixels
'               e.g.:  sSizes = "90,34,42,184,-1"
'                  creates 5 parts, the last one extending until the
'                  right edge of the window.
' Return Values:
'   If the operation succeeds, the return value is TRUE.
'   If the operation fails, the return value is FALSE.
' *********************************************************************************************
FUNCTION StatusBar_SetPartsBySize (BYVAL hStatus  AS DWORD, BYVAL sSizes AS STRING) AS LONG
  DIM PARTS AS LOCAL LONG
  DIM X AS LOCAL LONG
  DIM Y AS LOCAL LONG

  PARTS = MAX(PARSECOUNT(sSizes),1)
  IF PARTS < 2 THEN
   FUNCTION = SendMessage(hStatus, %SB_SIMPLE, (PARTS = 1), 0)
   EXIT FUNCTION
  END IF

  DIM Part(1:PARTS) AS LOCAL LONG

  FOR X = 1 TO PARTS
    Y = VAL(PARSE$(sSizes, ",", X))
    IF X = 1 THEN
      Part(X) = Y
    ELSE
      IF Y < 1 THEN
        Part(X) = -1
      ELSE
        Part(X) = Part(X-1) + Y
      END IF
    END IF
  NEXT

  IF SendMessage (hStatus, %SB_SETPARTS, PARTS, VARPTR(Part(1))) <> 0 THEN
    FUNCTION = SendMessage(hStatus, %SB_SIMPLE, (PARTS = 1), 0)
 END IF
END FUNCTION
'-----------------------------------------------
' 返回用于文件名的当前日期字符串，例:20131012
'-----------------------------------------------
FUNCTION DateTimeForFileName()AS STRING
  LOCAL st AS SYSTEMTIME
  GetLocalTime st
  FUNCTION = FORMAT$(st.wYear,"0000") & FORMAT$(st.wMonth,"00") & FORMAT$(st.wDay,"00") & _
              FORMAT$(st.wHour,"00") & FORMAT$(st.wMinute,"00") & FORMAT$(st.wSecond,"00") & _
              FORMAT$(st.wMilliseconds,"000")
END FUNCTION
