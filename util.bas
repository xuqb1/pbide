'util.bas
'����ʵ��һЩ���õĺ���

'DECLARE FUNCTION FileExist (BYVAL sfName AS STRING) AS LONG     '�ļ��Ƿ�����ж�
'DECLARE FUNCTION FileNam   (BYVAL Src AS STRING) AS STRING      '��ȡ�ļ���
'DECLARE FUNCTION FilePath  (BYVAL Src AS STRING) AS STRING      '��ȡ�ļ�·��
'DECLARE FUNCTION IsWin95() AS LONG                              '�Ƿ���Win95
'
TYPE BITMAPINFOTYPE
  bmiHeader    AS BITMAPINFOHEADER
  bmiColors(1) AS RGBQUAD
END TYPE
'------------------------------------------------------------------------------
' UDT (�û��Զ������ͼ��ṹ) UNIONS.
'------------------------------------------------------------------------------
' ����ͼ��س���Ĳ��� - ���Ӵ�UDT������Ҫ��������ص�λ�á�
' ����APPPROPS.INC�Ķ�д����!
' ҲҪ���Ǳ�ģ���е� AppInitialize() �� AppTerminate() ����!
'------------------------------------------------------------------------------
TYPE AppParametersTYPE
  sAppPath     AS STRING * %MAX_PATH      ' ����·��.
  sAppIniPath  AS STRING * %MAX_PATH      ' ����INI�ļ�·��.
  sAppDbPath   AS STRING * %MAX_PATH      ' �������ݿ�·��.
  sAppIdxPath  AS STRING * %MAX_PATH      ' �������ݿ�Ŀ¼·��.
  lSplashTime  AS LONG                    ' ����������ʱ.
  sbDateFmt    AS LONG                    ' ״̬�����ڸ�ʽ.
  sbTimeFmt    AS LONG                    ' ״̬��ʱ���ʽ.
  rbRowCount   AS LONG                    ' Rebar ����.
  rbBandCount  AS LONG                    ' Rebar ����.
  rbBand0      AS LONG                    ' Rebar 0���Ƿ��ƶ�.
  rbBand1      AS LONG                    ' Rebar 0���Ƿ��ƶ�.
  sbTimePart   AS LONG                    ' ״̬��ʱ������.
  sbDatePart   AS LONG                    ' ״̬����������.
END TYPE
GLOBAL udtAp AS AppParametersTYPE         ' ȫ�ֳ�������UDT����.

'------------------------------------------------------------------------------
' ���� : IniRead () '
' ���� : �ӳ���INI�ļ���ȡ����. '
' �÷� : sText = IniRead ("INIFILE", "SECTION", "KEY", "DEFAULT")
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
' ���� : IniWrite ()
'
' ���� : �������ݵ�����INI�ļ�.
'
' �÷� : lResult = IniWrite ("INIFILE", "SECTION", "KEY", "VALUE")
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
' ���� : IniDelSection () '
' ���� : ɾ��INI�ļ��еĶ�.  '
' �÷� : lResult = IniDelSection ("INIFILE", "SECTION")
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
' ���� : IniDelKey ()   '
' ���� : ɾ��INI�ļ��еļ�ֵ. '
' �÷� : lResult = IniDelKey ("INIFLE", "SECTION", "KEY")
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
' ���� : AppPath ()   '
' ���� : ���س���·��.   '
' �÷� : sPath = AppPath ()
'------------------------------------------------------------------------------
FUNCTION AppPath () AS STRING
  LOCAL szTmpAsciiz AS ASCIIZ * %MAX_PATH
  GetModuleFileName GetModuleHandle(BYVAL %NULL), szTmpAsciiz, %MAX_PATH
  FUNCTION = LEFT$(szTmpAsciiz, INSTR(-1, szTmpAsciiz, ANY "\/:"))
END FUNCTION
'------------------------------------------------------------------------------
' ���� : GetProperty ()  '
' ���� : ��config.INI�ļ���ȡ����ֵ  '
' �÷� : sText = GetProperty ("KEY", "DEFAULT")
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
' ���� : SetProperty ()  '
' ���� : �������ݵ�����INI�ļ�.  '
' �÷�  : lResult = SetProperty ("KEY", "VALUE")
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
' ���� : AppLoadParams ()   '
' ���� : ��INI�ļ���ȡ���еĳ������.
'------------------------------------------------------------------------------
FUNCTION AppLoadParams () AS LONG
  ' INI �� [APPLICATION].
  udtAp.sAppPath    = IniRead(udtAp.sAppIniPath, "APPLICATION PATHS", "APPPATH",    "")
  udtAp.sAppIniPath = IniRead(udtAp.sAppIniPath, "APPLICATION PATHS", "APPINIPATH", "")
  udtAp.sAppDbPath  = IniRead(udtAp.sAppIniPath, "APPLICATION PATHS", "APPDBPATH",  "")
  udtAp.sAppIdxPath = IniRead(udtAp.sAppIniPath, "APPLICATION PATHS", "APPIDXPATH", "")
  ' INI �� [SPLASH TIMMING].
  udtAp.lSplashTime = VAL(IniRead(udtAp.sAppIniPath, "SPLASH TIMMING", "SPLASHTIME", ""))
  ' INI �� [DATETIMEFORMAT].
  udtAp.sbDateFmt  = VAL(IniRead(udtAp.sAppIniPath, "DATE TIME FORMAT", "SBDATEFMT", ""))
  udtAp.sbTimeFmt  = VAL(IniRead(udtAp.sAppIniPath, "DATE TIME FORMAT", "SBTIMEFMT", ""))
  udtAp.sbDatePart = VAL(IniRead(udtAp.sAppIniPath, "DATE TIME FORMAT", "SBDATEPART", ""))
  udtAp.sbTimePart = VAL(IniRead(udtAp.sAppIniPath, "DATE TIME FORMAT", "SBTIMEPART", ""))
  ' INI �� [REBAR POSITIONS].
  udtAp.rbRowCount  = VAL(IniRead(udtAp.sAppIniPath, "REBAR POSITIONS", "RBROWS",  ""))
  udtAp.rbBandCount = VAL(IniRead(udtAp.sAppIniPath, "REBAR POSITIONS", "RBBANDS", ""))
  udtAp.rbBand0     = VAL(IniRead(udtAp.sAppIniPath, "REBAR POSITIONS", "RBBAND0", ""))
  udtAp.rbBand1     = VAL(IniRead(udtAp.sAppIniPath, "REBAR POSITIONS", "RBBAND1", ""))
  FUNCTION = %FALSE
END FUNCTION
'------------------------------------------------------------------------------
' ���� : AppSaveParams ()'
' ���� : ����������в�����INI�ļ�.
'------------------------------------------------------------------------------
FUNCTION AppSaveParams () AS LONG
  ' INI �� [Application].
  IniWrite udtAp.sAppIniPath, "APPLICATION PATHS", "APPPATH",    udtAp.sAppPath
  IniWrite udtAp.sAppIniPath, "APPLICATION PATHS", "APPINIPATH", udtAp.sAppIniPath
  IniWrite udtAp.sAppIniPath, "APPLICATION PATHS", "APPDBPATH",  udtAp.sAppDbPath
  IniWrite udtAp.sAppIniPath, "APPLICATION PATHS", "APPIDXPATH", udtAp.sAppIdxPath
  ' INI �� [SPLASH TIMMING].
  IniWrite udtAp.sAppIniPath, "SPLASH TIMMING", "SPLASHTIME", STR$(udtAp.lSplashTime)
  ' INI �� [DATETIMEFORMAT].
  IniWrite udtAp.sAppIniPath, "DATE TIME FORMAT", "SBDATEFMT",  STR$(udtAp.sbDateFmt)
  IniWrite udtAp.sAppIniPath, "DATE TIME FORMAT", "SBTIMEFMT",  STR$(udtAp.sbTimeFmt)
  IniWrite udtAp.sAppIniPath, "DATE TIME FORMAT", "SBDATEPART", STR$(udtAp.sbDatePart)
  IniWrite udtAp.sAppIniPath, "DATE TIME FORMAT", "SBTIMEPART", STR$(udtAp.sbTimePart)
  ' INI �� [REBAR POSITIONS].
  IniWrite udtAp.sAppIniPath, "REBAR POSITIONS", "RBROWS",  STR$(udtAp.rbRowCount)
  IniWrite udtAp.sAppIniPath, "REBAR POSITIONS", "RBBANDS", STR$(udtAp.rbBandCount)
  IniWrite udtAp.sAppIniPath, "REBAR POSITIONS", "RBBAND0", STR$(udtAp.rbBand0)
  IniWrite udtAp.sAppIniPath, "REBAR POSITIONS", "RBBAND1", STR$(udtAp.rbBand1)
  FUNCTION = %FALSE
END FUNCTION
'===============================================================================
FUNCTION OnlyOneInstance(szName AS ASCIIZ) AS LONG
'------------------------------------------------------------------------------
  ' OnlyOneInstance - ֻ������һ������ʵ������
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
  ' FileExist - ȷ��ĳ���ļ����ļ��Ƿ����
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
  ' ��ȡ����·���е��ļ�������
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
  ' ��ȡ����·���е�·������
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
  ' �Ƿ��� Windows 95
  '----------------------------------------------------------------------------
  LOCAL vi AS OSVERSIONINFO
  vi.dwOsVersionInfoSize = SIZEOF(vi)
  GetVersionEx vi
  FUNCTION = ((vi.dwPlatformId = %VER_PLATFORM_WIN32_WINDOWS) AND (vi.dwMinorVersion = 0))
END FUNCTION
'------------------------------------------------------------------------------
' ���� : DrawBitMap () '
' ���� : ������������λͼ.  '
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
' ���� : CreateDibPalette ()  '
' ���� : ʹ��ģ�崴��DIBλͼ.  '
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
' ���� : LoadBitmapRes ()  '
' ���� : ����Դ�ļ�������������λͼ.  '
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
' ���� : CenterWindow ()'
' ���� : ������hWnd������κδ�������.
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
' ���� : CenterWindowRel () '
' ���� : ������hWnd����й�ϵ�����д�������.
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
  ' ʹ���ڹ�ϵ���丸����.
  hWndParent = GetParent(hWnd)
  GetWindowRect(hWnd, rc)
  GetWindowRect(hWndParent, rcp)
  nWidth  = rc.nRight  - rc.nLeft
  nHeight = rc.nBottom - rc.nTop
  X = ((rcp.nRight  - rcp.nLeft) - nWidth)  / 2 + rcp.nLeft
  Y = ((rcp.nBottom - rcp.nTop)  - nHeight) / 2 + rcp.nTop
  ScrWidth  = GetSystemMetrics(%SM_CXSCREEN)
  ScrHeight = GetSystemMetrics(%SM_CYSCREEN)
  ' ȷʹ�Ի��򲻻��ƶ�����Ļ����.
  IF (X < 0) THEN X = 0
  IF (Y < 0) THEN Y = 0
  IF (X + nWidth  > ScrWidth)  THEN X = ScrWidth  - nWidth
  IF (Y + nHeight > ScrHeight) THEN Y = ScrHeight - nHeight
  MoveWindow(hWnd, X, Y, nWidth, nHeight, %FALSE)
  FUNCTION = %TRUE
END FUNCTION
'------------------------------------------------------------------------------
' ���� : MakeFont ()
' ���� : ���������壬������������.
'------------------------------------------------------------------------------
FUNCTION MakeFont (BYVAL sFntName    AS STRING, _
           BYVAL fFntSize    AS LONG, _
           BYVAL lWeight     AS LONG, _
           BYVAL lUnderlined AS LONG, _
           BYVAL lItalic     AS LONG, _
           BYVAL lStrike     AS LONG, _
           BYVAL lCharSet    AS LONG)   AS DWORD
'------------------------------------------------------------------------------
' ʹ��ʾ��:   '
'    hFont = MakeFont ("MS Sans Serif", 8, %FW_NORMAL, %FALSE, %FALSE, %FALSE,
'                      %DEFAULT_CHARSET)  '
'    hFont = MakeFont ("Courier New", 10, %FW_BOLD, %FALSE, %FALSE, %FALSE,
'                      %DEFAULT_CHARSET)
'    hFont = MakeFont ("Marlett", 8, %FW_NORMAL, %FALSE, %FALSE, %FALSE,
'                      %SYMBOL_CHARSET)  '
' ע��: �κ���MakeFont�����������ڲ�����ҪʱӦ��ʹ��DeleteObject�����������Է�ֹ�ڴ�й©��
'       ����:    '
'        Case %WM_DESTROY         OnDestroy () ���.
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
' ���� : ExecuteURL ()
'
' ���� : ִ�д��ڳ���MAIL Web������ȡ�
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
' ���� : ExeFile () '
' ���� : ִ���ļ�.
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
    lResult = MSGBOX(sErr, %MB_ICONERROR OR %MB_OKCANCEL, "�������")
    ERRCLEAR
    FUNCTION = %FALSE
    EXIT FUNCTION
  END IF
  FUNCTION = %TRUE
END FUNCTION
'------------------------------------------------------------------------------
' ���� : PbDoEvents () '
' ���� : Ϊ�������¼���Ӧ���塣Do Events for Application.
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
' ���� : �����ڴ�.
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
' �ж��ļ��Ƿ����
'-----------------------------------------------------------------------------------------------------------
FUNCTION Exist(sFilename AS STRING) AS LONG
  ' Checks to see if a given disk file exists
  LOCAL    dummy              AS LONG
  dummy = GETATTR(sFilename)
  FUNCTION = (ERRCLEAR = 0)
END FUNCTION
'�������ݿ⣬�������еı�
FUNCTION MakeDB(BYVAL dbName AS STRING)AS LONG '�������ݿ�
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  errStr=""
  IF dbName="" OR RIGHT$(dbName,3)<>".db" THEN
    errStr="���ݿ��ļ���Ϊ�ջ򲻱�׼,"
    GOTO abort
  END IF
  IF ISTRUE sqlOpen(dbName,hDB) THEN
    errStr="�޷��򿪻򴴽����ݿ�,"
    GOTO abort
  END IF
  '�������ñ�
  sSql="CREATE TABLE "+$CONFIGTABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,name text,value text,remark text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$CONFIGTABLENAME+","
    GOTO abort
  END IF
  '����Ĭ�����ñ�
  sSql="CREATE TABLE "+$DEFAULTCONFIGTABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,name text,value text,remark text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$DEFAULTCONFIGTABLENAME+","
    GOTO abort
  END IF
  '�������ݼ���
  sSql="CREATE TABLE "+$DATASETTABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,type text,value text,remark text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$DATASETTABLENAME+","
    GOTO abort
  END IF
  '�����ַ������Ա�
  sSql="CREATE TABLE "+$STRINGLANGTABLENAME+" (Id integer NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,name text,english text,chinese text)"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for Create "+$STRINGLANGTABLENAME+","
    GOTO abort
  END IF
  '�����ؼ���ʽ��
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
' ��ʼ����Ĭ�����ñ����ñ����ݼ�������ԣ��ַ������Ա�
FUNCTION InitConfig() AS LONG
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="�޷��򿪻򴴽����ݿ�,"
    GOTO abort
  END IF
  ' ��ʼ��Ĭ�����ñ�
  sSql="BEGIN TRANSACTION;"+$CRLF
  '����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showflashwindow','1');"+$CRLF          '����ʱ��ʾ��ӭ��Ļ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showguiderwindow','1');"+$CRLF         '��������ʾ�򵼴���
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowposopt','0');"+$CRLF         '������λ������:Ĭ�ϣ��̶���ʹ����һ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowpos','300,200');"+$CRLF      '������ָ��λ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowsize','640,480');"+$CRLF     '������ָ���ߴ�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowlastpos','300,200');"+$CRLF  '�������ϴ�λ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('mainwindowlastsize','640,480');"+$CRLF '�������ϴγߴ�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('language','English');"+$CRLF           '����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('clientrect','0,0,0,0');"+$CRLF         '���ͻ������Σ����ϣ�����
  'rebar��band
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarrowcount','2');"+$CRLF            'Rebar����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarbandcount','3');"+$CRLF           'Rebar��band��:�������������༭�����������ڹ�������(�����б�����)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband0','0,200');"+$CRLF           'Rebar�ĵ�һ��(����ֵ)band��Ϣ��wID,cx(��Ӧ��������ID,���)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband1','0,200');"+$CRLF           'Rebar�ĵڶ���(����ֵ)band��Ϣ��wID,cx(��Ӧ��������ID,���)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband2','0,200');"+$CRLF           'Rebar�ĵ�����(����ֵ)band��Ϣ��wID,cx(��Ӧ��������ID,���)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband3','0,200');"+$CRLF                 'Rebar�ĵ��ĸ�(����ֵ)band��Ϣ��wID,cx(��Ӧ��������ID,���)(��ռλ)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband4','0,200');"+$CRLF                 'Rebar�ĵ����(����ֵ)band��Ϣ��wID,cx(��Ӧ��������ID,���)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarband5','0,200');"+$CRLF                 'Rebar�ĵ�����(����ֵ)band��Ϣ��wID,cx(��Ӧ��������ID,���)
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rebarlocked','0');"+$CRLF              'rebar�Ƿ�����
  '�ļ�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('sourcefiletype','.bas');"+$CRLF              'Դ�ļ�������չ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('includefiletype','.inc');"+$CRLF             'ͷ�ļ�������չ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('resourcefiletype','.rc');"+$CRLF             '��Դ�ļ�������չ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rcheadfiletype','.h');"+$CRLF                '��Դͷ�ļ�������չ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('projectfiletype','.pbj');"+$CRLF             '��Ŀ�ļ�������չ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('registryfiletype','1');"+$CRLF               'ע���ļ�������չ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autoreloadfile','0');"+$CRLF                 '�Զ���������༭���ļ�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfilenumber','5');"+$CRLF               '����ļ��б��������
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile1','');"+$CRLF                     '����ļ�1
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile2','');"+$CRLF                     '����ļ�2
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile3','');"+$CRLF                     '����ļ�3
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile4','');"+$CRLF                     '����ļ�4
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile5','');"+$CRLF                     '����ļ�5
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile6','');"+$CRLF                     '����ļ�6
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile7','');"+$CRLF                     '����ļ�7
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile8','');"+$CRLF                     '����ļ�8
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile9','');"+$CRLF                     '����ļ�9
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('recentfile10','');"+$CRLF                    '����ļ�10
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('startfolder','0');"+$CRLF                    '��ʼ�ļ��У�0,��Ĭ���ļ��п�ʼ��1,�����ʹ�õ��ļ��п�ʼ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('lastfolder','');"+$CRLF                      '���ʹ�õ��ļ���
  '����༭
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autoindent','1');"+$CRLF                     '�Զ�����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('undoaftersave','1');"+$CRLF                  '������Կɳ����༭
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('spacepertab','2');"+$CRLF                    'ÿ��Tab�൱�ڵĿո���
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autosave','1');"+$CRLF                       '�Զ�����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autosaveminute','5');"+$CRLF                 'ÿ#�����Զ�����һ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showlinenumber','1');"+$CRLF                 '��ʾ�к�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('autocompletetip','1');"+$CRLF                '�Զ������ʾ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('keywordcase','1');"+$CRLF                    '�ؼ��ִ�Сд��0,����ԭ���Ĵ�Сд;1,��д;2,������ĸ��д;3,Сд
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('font','Arial');"+$CRLF                       '����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('fontsize','12');"+$CRLF                      '�����С
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('folderlevel','2');"+$CRLF                    '��չ����0,�޾�չ;1,�ؼ��ּ�;2,Sub/function��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('foldericon','2');"+$CRLF                     '��չͼ�꣺0,��ͷ;1,+/-;2,Բ��;3,����
  '�﷨��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('usecolorineditor','1');"+$CRLF               '�ڱ༭����ʹ���﷨��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('usecolorwhenprint','1');"+$CRLF              '�ڴ�ӡʱʹ���﷨��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('assembleforecolor','192,0,0');"+$CRLF        '������ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('assemblebackcolor','255,255,255');"+$CRLF    '�����䱳��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('commentforecolor','0,127,0');"+$CRLF         'ע�����ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('commentbackcolor','255,255,255');"+$CRLF     'ע����䱳��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('keywordforecolor','0,0,192');"+$CRLF         '�ؼ���ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('keywordbackcolor','255,255,255');"+$CRLF     '�ؼ��ֱ���ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('pbformforecolor','192,100,0');"+$CRLF        '����ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('pbformbackcolor','255,255,255');"+$CRLF      '���屳��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('stringforecolor','192,32,192');"+$CRLF       '�ַ���ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('stringbackcolor','255,255,255');"+$CRLF      '�ַ�������ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('textforecolor','0,0,0');"+$CRLF              '�ı�ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('textbackcolor','255,255,255');"+$CRLF        '�ı�����ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('selectedforecolor','255,255,255');"+$CRLF    'ѡ�����ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('selectedbackcolor','49,106,197');"+$CRLF     'ѡ����䱳��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('breakpointforecolor','255,255,255');"+$CRLF  '�ϵ����ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('breakpointbackcolor','255,0,0');"+$CRLF      '�ϵ���䱳��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('bookmarkforecolor','0,0,255');"+$CRLF        '��ǩ���ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('bookmarkbackcolor','0,255,255');"+$CRLF      '��ǩ��䱳��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('execpointforecolor','0,0,0');"+$CRLF         'ִ�е����ǰ��ɫ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('execpointbackcolor','255,255,0');"+$CRLF     'ִ�е���䱳��ɫ
  '���ڱ༭
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showgrid','1');"+$CRLF                       '��ʾ����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('snaptogrid','1');"+$CRLF                     '���뵽����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('gridsize','8x8');"+$CRLF                     '�����С
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('snapbetweencontrol','1');"+$CRLF             '�ؼ�֮���Զ�����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('adjusthandlesize','1');"+$CRLF               '��������ߴ磺0,С;1,��;2,��
  '����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('includefilepath','D:\Program Files\PBWin\win32api');"+$CRLF        'PBͷ�ļ�·��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('resourcefilepath','D:\Program Files\PBWin\win32api');"+$CRLF       'RCͷ�ļ�·��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('compilefilepath','D:\Program Files\PBWin\bin\pbwin.exe');"+$CRLF   '�����ļ�·��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('win32helpfilepath','D:\Program Files\PBWin\bin\win32.hlp');"+$CRLF 'win32.hlp�ļ�·��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('beepwhencomplete','0');"+$CRLF               '������ɺ�ʹ����ʾ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('showcompileresult','1');"+$CRLF              '��ʾ������
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('createcompilelog','1');"+$CRLF               '����������־
  '��ӡ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('pagenumberonfooter','1');"+$CRLF             'ҳ�Ŵ���ӡҳ��[#/#]
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('filenamedatatimeonheader','1');"+$CRLF       'ҳü�������ļ��������ں�ʱ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('topmargin','10');"+$CRLF                     '�ϱ߾�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('bottommargin','10');"+$CRLF                  '�±߾�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('leftmargin','10');"+$CRLF                    '��߾�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('rightmargin','10');"+$CRLF                   '�ұ߾�
  '�ȼ�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Commands','New,Open,Save,Find,Find Next,Replace,Goto Line,Cut," & _
                                                    "Copy,Paste,Delete,Select All,Comment,UnComment,Tab Indent,Tab Outdent,Space Indent,Space Outdent');"+$CRLF '�����б�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('New','Ctrl + N');"+$CRLF                    '�½�
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Open','Ctrl + O');"+$CRLF                    '��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Save','Ctrl + S');"+$CRLF                    '����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Find','Ctrl + F');"+$CRLF                    '����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Find Next','F3');"+$CRLF                     '������һ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Replace','Ctrl + H');"+$CRLF                 '�滻
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Goto Line','Ctrl + G');"+$CRLF               'ת����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Cut','Ctrl + X');"+$CRLF                     '����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Copy','Ctrl + C');"+$CRLF                    '����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Paste','Ctrl + V');"+$CRLF                   'ճ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Delete','Del');"+$CRLF                       'ɾ��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Select All','Ctrl + A');"+$CRLF              'ȫѡ
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Comment','Ctrl + Q');"+$CRLF                 'ע��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('UnComment','Ctrl + Shift + Q');"+$CRLF       '��ע��
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Tab Indent','Tab');"+$CRLF                   'Tab����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Tab Outdent','Shift + Tab');"+$CRLF          'Tab������
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Space Indent','Space');"+$CRLF               '�ո�����
  sSql=sSql+"insert into "+$DEFAULTCONFIGTABLENAME+" (name,value)values('Space Outdent','Shift + Space');"+$CRLF      '�ո�����
  sSql=sSql+"COMMIT;"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for insert "+$DEFAULTCONFIGTABLENAME+","
    GOTO abort
  END IF
  ' ��ʼ�����ñ�
  REPLACE $DEFAULTCONFIGTABLENAME WITH $CONFIGTABLENAME IN sSql
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for insert "+$CONFIGTABLENAME+","
    GOTO abort
  END IF
  ' ���ݼ�
  sSql="BEGIN TRANSACTION;"+$CRLF
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('language','English');"+$CRLF    'Ӣ��
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('language','Chinese');"+$CRLF    '����
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','2x2');"+$CRLF    '�����С
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','4x4');"+$CRLF    '�����С
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','6x6');"+$CRLF    '�����С
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','8x8');"+$CRLF    '�����С
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','10x10');"+$CRLF    '�����С
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('gridsize','12x12');"+$CRLF    '�����С
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','2');"+$CRLF    'ˮƽ�Ʊ���൱�Ŀո���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','3');"+$CRLF    'ˮƽ�Ʊ���൱�Ŀո���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','4');"+$CRLF    'ˮƽ�Ʊ���൱�Ŀո���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','6');"+$CRLF    'ˮƽ�Ʊ���൱�Ŀո���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('tabsize','8');"+$CRLF    'ˮƽ�Ʊ���൱�Ŀո���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','1');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','2');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','3');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','4');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','5');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','6');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','7');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','8');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','9');"+$CRLF'����ļ���
  sSql=sSql+"insert into "+$DATASETTABLENAME+" (type,value)values('recentfilenumber','10');"+$CRLF'����ļ���
  sSql=sSql+"COMMIT;"
  IF sqlExe(hDB,sSql)<0 THEN
    errStr=sqlErrMsg(hDB) & " for insert "+$DATASETTABLENAME+","
    GOTO abort
  END IF
  ' �ַ������Ա�
  sSql="BEGIN TRANSACTION;"+$CRLF
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('language','language','����');"+$CRLF    '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('English','English','Ӣ��');"+$CRLF      'Ӣ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Chinese','Chinese','����');"+$CRLF      '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('The program is executing!','The program is executing!','������������!');"+$CRLF '����������ʾ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Visual PowerBASIC','Visual PowerBASIC','Visual PowerBASIC');"+$CRLF '������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New Project','New Project','����Ŀ');"+$CRLF        '�򵼴��ڱ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Don''t show this dialog again','Don''t show this dialog again','����ʱ������ʾ���Ի���');"+$CRLF        '�򵼴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New','New','�½�');"+$CRLF        '�򵼴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Exist','Exist','����');"+$CRLF        '�򵼴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Recent','Recent','���');"+$CRLF        '�򵼴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Search scope:','Search scope:','������Χ:');"+$CRLF        '�򵼴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('File name:','File name:','�ļ���:');"+$CRLF        '�򵼴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('File type:','File type:','�ļ�����:');"+$CRLF        '�򵼴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Normal Window','Normal Window','һ�㴰��');"+$CRLF  '�������ڲ˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('MDI Window','MDI Window','MDI����');"+$CRLF         '�������ڲ˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Top Align','Top Align','�϶���');"+$CRLF            '�ؼ�����˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Horizontal Center Align','HorizontalCenter Align','ˮƽ�ж���');"+$CRLF  '�ؼ�����˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Bottom Align','Bottom Align','�¶���');"+$CRLF      '�ؼ�����˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Left Align','Left Align','�����');"+$CRLF          '�ؼ�����˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Vertical Center Align','Vertical Center Align','��ֱ�ж���');"+$CRLF  '�ؼ�����˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Right Align','Right Align','�Ҷ���');"+$CRLF        '�ؼ�����˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('DDT Code','DDT Code','DDT����');"+$CRLF             '�����ʽ�˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('SDK Code','SDK Code','SDK����');"+$CRLF             '�����ʽ�˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Wait for develop...','Wait for develop...','������');"+$CRLF  '��������ʾ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('procedure','procedure','����');"+$CRLF              'SUB/FUNCTION�б�����ǰ��ʾ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Text File','Text File','�ı��ļ�');"+$CRLF              '���뱣��Ի����õ����ļ���չ�������ַ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('All File','All File','�����ļ�');"+$CRLF                '���뱣��Ի����õ����ļ���չ�������ַ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Source Code File','Source Code File','Դ�����ļ�');"+$CRLF '���뱣��Ի����õ����ļ���չ�������ַ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Include File','Include File','�����ļ�');"+$CRLF        '���뱣��Ի����õ����ļ���չ�������ַ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Resource File','Resource File','��Դ�ļ�');"+$CRLF      '���뱣��Ի����õ����ļ���չ�������ַ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Project File','Project File','��Ŀ�ļ�');"+$CRLF        '���뱣��Ի����õ����ļ���չ�������ַ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('MDI text editor, Visualize form editor, Project manager'," + _
                            "'MDI text editor, Visualize form editor, Project manager','MDI�ı��༭�������ӻ�����༭������Ŀ������');"+$CRLF        '���ڹ��ڶԻ����еĹ���˵��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('ToolBar','ToolBar','������');"+$CRLF
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Message','Message','��Ϣ');"+$CRLF
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show As','Show As','��ʾΪ');"+$CRLF
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Orgi String','Orgi String','ԭ�ַ���');"+$CRLF
  '�ȼ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Hotkey','Hotkey','�ȼ�');"+$CRLF                          '�ȼ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Defined Hotkey:','Defined Hotkey:','�Ѷ����ȼ�:');"+$CRLF '�Ѷ����ȼ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Delete','Delete','ɾ��');"+$CRLF                          'ɾ���ȼ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Command:','Command:','����');"+$CRLF                      '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Hotkey:','Hotkey:','�ȼ�');"+$CRLF                        '�ȼ���ǩ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Update','Update','����');"+$CRLF                          '����


  'menu
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('File','File','�ļ�');"+$CRLF                      '�ļ��˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New','New','�½�');"+$CRLF                        '�½��˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Open','Open','��');"+$CRLF                      '�򿪲˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Reopen...','Reopen...','�ٴδ�...');"+$CRLF     '�ٴδ򿪲˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Open Project','Open Project','����Ŀ');"+$CRLF  '����Ŀ�˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save','Save','����');"+$CRLF                      '����˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save As','Save As','���Ϊ');"+$CRLF              '���Ϊ�˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save All','Save All','����ȫ��');"+$CRLF          '����ȫ���˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save Project','Save Project','������Ŀ');"+$CRLF  '������Ŀ�˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print Setting','Print Setting','��ӡ����');"+$CRLF'��ӡ����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print Preview','Print Preview','��ӡԤ��');"+$CRLF'��ӡԤ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print...','Print...','��ӡ...');"+$CRLF           '��ӡ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Close','Close','�ر�');"+$CRLF                    '�ر�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Close All','Close All','�ر�ȫ��');"+$CRLF        '�ر�ȫ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Open Command','Open Command','��������');"+$CRLF'��������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Exit','Exit','�˳�');"+$CRLF                      '�˳�

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Edit','Edit','�༭');"+$CRLF                      '�༭
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Undo','Undo','����');"+$CRLF                      '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Redo','Redo','����');"+$CRLF                      '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Cut','Cut','����');"+$CRLF                        '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Copy','Copy','����');"+$CRLF                      '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Paste','Paste','ճ��');"+$CRLF                    'ճ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Select all','Select all','ȫѡ');"+$CRLF          'ȫѡ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Block','Block','�����');"+$CRLF                  '�����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Comment','Comment','ע��');"+$CRLF                'ע��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('UnComment','UnComment','��ע��');"+$CRLF          '��ע��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tab Indent','Tab Indent','Tab����');"+$CRLF       'Tab����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tab Outdent','Tab Outdent','Tab������');"+$CRLF   'Tab������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Space Indent','Space Indent','�ո�����');"+$CRLF  '�ո�����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Space Outdent','Space Outdent','Space������');"+$CRLF         'Space������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Insert GUID','Insert GUID','����GUID');"+$CRLF    '����GUID
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Find...','Find...','����...');"+$CRLF             '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Find Next','Find Next','������һ��');"+$CRLF      '������һ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Replace...','Replace...','�滻...');"+$CRLF       '�滻
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Goto Line...','Goto Line...','ת����...');"+$CRLF 'ת����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Goto Bookmark...','Goto Bookmark...','ת����ǩ...');"+$CRLF   'ת����ǩ

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Run','Run','����');"+$CRLF                        '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compile','Compile','����');"+$CRLF                '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compile and Run','Compile and Run','���벢����');"+$CRLF      '���벢����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Set Primary File','Set Primary File','�������ļ�');"+$CRLF    '�������ļ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Set Command Parameter','Set Command Parameter','���������в���');"+$CRLF   '���������в���

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('View','View','��ͼ');"+$CRLF                      '��ͼ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Control ToolBox','Control ToolBox','�ؼ���');"+$CRLF          '�ؼ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Project ToolBox','Project ToolBox','��Ŀ����');"+$CRLF        '��Ŀ����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Property ToolBox','Property ToolBox','���Դ���');"+$CRLF      '���Դ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Find Result Window','Find Result Window','���ҽ������');"+$CRLF      '���ҽ������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compile Info Window','Compile Info Window','������Ϣ����');"+$CRLF    '������Ϣ����

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tool','Tool','����');"+$CRLF                  '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Make Menu','Make Menu','�����˵�');"+$CRLF        '�����˵�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Make Toolbar','Make Toolbar','����������');"+$CRLF            '����������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Find Multi Dir','Find Multi Dir','��Ŀ¼����');"+$CRLF        '��Ŀ¼����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Replace Multi Dir','Replace Multi Dir','��Ŀ¼�滻');"+$CRLF  '��Ŀ¼�滻
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Format code','Format code','�����ʽ��');"+$CRLF              '�����ʽ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('WinSpy','WinSpy','WinSpy');"+$CRLF                'WinSpy
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Code Counter','Code Counter','����ͳ��');"+$CRLF  '����ͳ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Code Manager','Code Manager','�������');"+$CRLF  '�������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Custom Tool Setting','Custom Tool Setting','�Զ��幤������');"+$CRLF   '�Զ��幤������

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Project','Project','��Ŀ');"+$CRLF                '��Ŀ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New Project...','New Project...','����Ŀ');"+$CRLF                '��Ŀ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save Project','Save Project','��Ŀ');"+$CRLF      '������Ŀ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Save as Project...','Save as Project...','�����Ŀ...');"+$CRLF     '�����Ŀ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Achive...','Achive...','�鵵');"+$CRLF            '�鵵
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New Vision','New Vision','�½��汾');"+$CRLF      '�½��汾
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Set Vision for files','Set Vision for files','����汾');"+$CRLF    '����汾
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Package Vision...','Package Vision...','����汾');"+$CRLF          '����汾

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Cooperation','Cooperation','Эͬ');"+$CRLF        'Эͬ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Coop Setting...','Coop Setting...','Эͬ����');"+$CRLF   'Эͬ����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Send Invite','Send Invite','��������');"+$CRLF    '��������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Send Visit','Send Visit','���ͷ���');"+$CRLF      '���ͷ���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Download Files','Download Files','�����ļ�');"+$CRLF'�����ļ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Upload Files','Upload Files','�����ļ�');"+$CRLF    '�����ļ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Chat...','Chat...','�Ự...');"+$CRLF             '�Ự

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Window','Window','����');"+$CRLF               '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Cascade','Cascade','�������');"+$CRLF            '�������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tile horizontally','Tile horizontally','ˮƽ�������д���');"+$CRLF   'ˮƽ�������д���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tile vertically','Tile vertically','�����������д���');"+$CRLF   '�����������д���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Arrange Icons','Arrange Icons','����ͼ��');"+$CRLF   '����ͼ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Close All','Close All','ȫ���ر�');"+$CRLF        'ȫ���ر�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Maximize Window','Maximize Window','��󻯴���');"+$CRLF   '��󻯴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Switch to Window','Switch to Window','�л�����');"+$CRLF   '�л�����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Option...','Option...','ѡ��');"+$CRLF            'ѡ��

  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Help','Help','����');"+$CRLF                      '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('IDE Help','IDE Help','IDE ����');"+$CRLF          'IDE����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB/WIN Content','PB/WIN Content','PB/WIN Ŀ¼');"+$CRLF   'PB/WIN Ŀ¼
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB/WIN Index','PB/WIN Index','PB/WIN ����');"+$CRLF   'PB/WIN ����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB/WIN Search','PB/WIN Search','PB/WIN ����');"+$CRLF   'PB/WIN ����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Windows SDK Help','Windows SDK Help','Windows SDK ����');"+$CRLF   'Windows SDK ����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('About','About','����');"+$CRLF      '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Update','Update','����');"+$CRLF    '����
  '�򵼴���
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('New Project','New Project','����Ŀ');"+$CRLF    '����
  'ѡ����е��ַ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Option','Option','ѡ��');"+$CRLF            'ѡ��
  '��ť
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('OK','OK','ȷ��');"+$CRLF    'ȷ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Cancel','Cancel','ȡ��');"+$CRLF    'ȡ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Apply','Apply','Ӧ��');"+$CRLF    'Ӧ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Default','Default','Ĭ��');"+$CRLF    'Ĭ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('General','General','һ��');"+$CRLF    'һ��
  'ѡ���б�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('File','File','�ļ�');"+$CRLF    '�ļ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Code Editor','Code Editor','����༭��');"+$CRLF    '����༭��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Color','Color','��ɫ');"+$CRLF    '��ɫ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Window Editor','Window Editor','����༭��');"+$CRLF    '����༭��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compile','Compile','����');"+$CRLF    '����
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print','Print','��ӡ');"+$CRLF    '��ӡ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Hotkey','Hotkey','�ȼ�');"+$CRLF    '�ȼ�
  '����ѡ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show welcome window when startup','Show welcome window when startup','����ʱ��ʾ��ӭ����');"+$CRLF    '%IDC_SHOWWELCOM
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show guilder window after startup','Show guilder window after startup','��������ʾ�򵼴���');"+$CRLF    '%IDC_SHOWGUILD
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Main window pos && size','Main window pos && size','������λ�ü���С');"+$CRLF    '%IDC_MAINWINDOWPOS
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use default position and size','Use default position and size','ʹ��Ĭ��λ�ü���С');"+$CRLF    '%IDC_DEFAULTPOSSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use fix position and size','Use fix position and size','ʹ�ù̶�λ�ü���С');"+$CRLF    '%IDC_FIXPOSSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use the position and size for last','Use the position and size for last','ʹ���ϴε�λ�ü���С');"+$CRLF    '%IDC_USELASTPOSSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Position','Position','λ��');"+$CRLF    '%IDC_POSITION
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Center','Center','����');"+$CRLF    '%IDC_SETCENTER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Special','Special','ָ��');"+$CRLF    '%IDC_SETPOSITION
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Size','Size','��С');"+$CRLF    '%IDC_MAINWINDOWSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Width:','Width:','���:');"+$CRLF    '%IDC_LABELWIDTH
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Height:','Height:','�߶�:');"+$CRLF    '%IDC_LABELHEIGHT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Please select language for this application:','Please select language for this application:','ѡ������:');"+$CRLF    '%IDC_LANGUAGE
  '�ļ�ѡ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Type','Type','����');"+$CRLF    '%IDC_FILETYPE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Source File:','PB Source File:','PBԴ�ļ�:');"+$CRLF    '%IDC_LABELSOURCEFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Head File:','PB Head File:','PBͷ�ļ�:');"+$CRLF    '%IDC_LABELINCLUDEFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('RC File:','RC File:','RC�ļ�:');"+$CRLF    '%IDC_LABELRCFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('RC Head File:','RC Head File:','RCͷ�ļ�:');"+$CRLF    '%IDC_LABELRCHEADFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Project File:','PB Project File:','PB��Ŀ�ļ�:');"+$CRLF    '%IDC_LABELPROJECTFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Register these file extend in windows','Register these file extend in windows','��Windows��ע����Щ�ļ���չ��');"+$CRLF    '%IDC_SETFILEEXT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Reload previous file set at startup','Reload previous file set at startup','����ʱ����֮ǰ���ļ���');"+$CRLF    '%IDC_LOADRECENTFILES
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Maximum Recent Files:','Maximum Recent Files:','��¼����ļ�������:');"+$CRLF    '%IDC_LABELMAXFILE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Initial startup folder','Initial startup folder','��ʼ�����ļ���');"+$CRLF    '%IDC_STARTUPFOLDER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Start in Default folder','Start in Default folder','Ĭ���ļ���');"+$CRLF    '%IDC_DEFAULTFOLDER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Start in Last folder','Start in Last folder','���ʹ�õ��ļ���');"+$CRLF    '%IDC_LASTFOLDER
  '����༭��ѡ��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Auto indent','Auto indent','�Զ�����');"+$CRLF '%IDC_AUTOINDENT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Keep Undo after save','Keep Undo after save','������Կɳ���');"+$CRLF '%IDC_KEEPUNDO
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Tab size:','Tab size:','Tab��С:');"+$CRLF '%IDC_LABELTABSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Auto save when','Auto save when','�Զ�����ÿ��');"+$CRLF '%IDC_AUTOSAVE1
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('minutes','minutes','����');"+$CRLF '%IDC_AUTOSAVE2
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show line number','Show line number','��ʾ����');"+$CRLF '%IDC_SHOWLINENUMBER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show auto complete tip','Show auto complete tip','��ʾ�Զ������ʾ');"+$CRLF '%IDC_SHOWAUTOCOMPLETE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Keyword Case','Keyword Case','�ؼ��ִ�Сд');"+$CRLF '%IDC_KEYWORDCASE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('No Case Change','No Case Change','���ı�');"+$CRLF '%IDC_NOCASECHANGE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Upper Case','Upper Case','��д');"+$CRLF '%IDC_UPPERCASE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Mixed Case','Mixed Case','����ĸ��д');"+$CRLF '%IDC_MIXEDCASE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Lower Case','Lower Case','Сд');"+$CRLF '%IDC_LOWERCASE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Editor Font','Editor Font','�༭������');"+$CRLF '%IDC_EDITORFONT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Font:','Font:','����:');"+$CRLF '%IDC_LABELFONT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Size:','Size:','��С:');"+$CRLF '%IDC_LABELFONTSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Sample text','Sample text','ʾ���ı�');"+$CRLF '%IDC_LABELSAMPLETEXT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Collapse && Expand','Collapse && Expand','���� && չ��');"+$CRLF '%IDC_COLLAPSEEXPAND
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Foldding Level','Foldding Level','��չ�㼶');"+$CRLF '%IDC_COLLAPSELEVEL
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('No Collapse','No Collapse','����չ');"+$CRLF '%IDC_NOCOLLAPSE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Keyword level','Keyword level','�ؼ��ּ�');"+$CRLF '%IDC_KEYWORDLEVEL
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Sub/Function level','Sub/Function level','Sub/Function ��');"+$CRLF '%IDC_SUBFUNLEVEL
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Folder icon','Folder icon','��չͼ��');"+$CRLF '%IDC_FOLDERICON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Arrow','Arrow','��ͷ');"+$CRLF '%IDC_ARROWICON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Circle','Circle','Բ��');"+$CRLF '%IDC_CIRCLEICON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Square','Square','����');"+$CRLF '%IDC_SQUAREICON
  '��ɫ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use Syntax Color in Editor','Use Syntax Color in Editor','�ڴ���༭����ʹ���﷨��ɫ');"+$CRLF '%IDC_USECOLORINEDITOR
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use Syntax Color when Print','Use Syntax Color when Print','��ӡʱʹ���﷨��ɫ');"+$CRLF '%IDC_USECOLORINPRINTER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Load &Defaults','Load &Defaults','����Ĭ����ɫ');"+$CRLF '%IDC_LOADDEFAULTCOLOR
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Assembler','Assembler','���');"+$CRLF '%IDC_ASSEMBLERBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Comments','Comments','ע��');"+$CRLF '%IDC_COMMENTBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Keywords','Keywords','�ؼ���');"+$CRLF '%IDC_KEYWORDBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Forms','PB Forms','PB����');"+$CRLF '%IDC_PBFORMSBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Strings','Strings','�ַ���');"+$CRLF '%IDC_STRINGBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Text','Text','�ı�');"+$CRLF '%IDC_TEXTBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Selection','Selection','ѡ��');"+$CRLF '%IDC_SELECTIONBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Breakpoint','Breakpoint','�ϵ�');"+$CRLF '%IDC_BREAKPOINTBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Bookmark','Bookmark','��ǩ');"+$CRLF '%IDC_BOOKMARKBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Exec point','Exec point','ִ�е�');"+$CRLF '%IDC_EXECPOINTBUTTON
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Reserved','Reserved','����');"+$CRLF '%IDC_RESERVED1BUTTON %IDC_RESERVED2BUTTON
  '���ڱ༭��
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Show grid','Show grid','��ʾ����');"+$CRLF '%IDC_SHOWGRID
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Snap to grid','Snap to grid','���뵽����');"+$CRLF '%IDC_GRIDSNAP
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Grid size:','Grid size:','�����С:');"+$CRLF '%IDC_GRIDSIZE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Snap between controls','Snap between controls','�ؼ������');"+$CRLF '%IDC_SNAPCONTROL
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Adjust handle size','Adjust handle size','�������С');"+$CRLF '%IDC_ADJUSTHANDLE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Small','Small','С');"+$CRLF '%IDC_SMALLHANDLE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Medium','Medium','��');"+$CRLF '%IDC_MEDIUMHANDLE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Large','Large','��');"+$CRLF '%IDC_LARGEHANDLE
  '������
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Paths','Paths','·��');"+$CRLF '%IDC_PATH
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('PB Include:','PB Include:','PB����:');"+$CRLF '%IDC_LABELPBINCLUDE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('RC Include:','RC Include:','RC����:');"+$CRLF '%IDC_LABELRCINCLUDE
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compiler File:','Compiler File:','�������ļ�:');"+$CRLF '%IDC_LABELCOMPILER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Win32.hlp File:','Win32.hlp File:','Win32.hlp�ļ�:');"+$CRLF '%IDC_LABELWIN32HELP
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Compiler Options','Compiler Options','����ѡ��');"+$CRLF '%IDC_COMPILEOPTION
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Beep on compiletion','Beep on compiletion','�������ʱ��һ��');"+$CRLF '%IDC_BEEPCOMPILETION
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Display Results','Display Results','��ʾ������');"+$CRLF '%IDC_DISPLAYRESULT
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Create Log file','Create Log file','������־�ļ�');"+$CRLF '%IDC_CREATELOGFILE
  '��ӡ
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Use the color for code in editor when print','Use the color for code in editor when print','��ӡʱʹ�ô����ڱ༭���е���ɫ');"+$CRLF '%IDC_USERCODECOLOR
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print page number on footer with the format [#/#]','Print page number on footer with the format [#/#]','��ҳ���Ը�ʽ[#/#]��ӡҳ��');"+$CRLF '%IDC_PRINTPAGENUMBER
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Print filename, time and date of the file on header','Print filename, time and date of the file on header','��ҳü��ӡ�ļ�����ʱ�������');"+$CRLF '%IDC_PRINTFILENAMETIME
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Margin','Margin','ҳ�߾�');"+$CRLF '%IDC_PAGEMARGIN
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Top:','Top:','�ϱ�:');"+$CRLF '%IDC_LABELTOPMARGIN
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Bottom:','Bottom:','�±�:');"+$CRLF '%IDC_LABELBOTTOMMARGIN
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Left:','Left:','���:');"+$CRLF '%IDC_LABELLEFTMARGIN
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Right:','Right:','�ұ�:');"+$CRLF '%IDC_LABELRIGHTMARGIN
  '�ȼ�
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Defined Hotkey:','Defined Hotkey:','�Ѷ�����ȼ�:');"+$CRLF '%IDC_LABELHADHOTKEY
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Delete','Delete','ɾ��');"+$CRLF '%IDC_DELETEHOTKEY
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Command:','Command:','����:');"+$CRLF '%IDC_LABELHOTKEYCOMMAND
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Hotkey:','Hotkey:','�ȼ�:');"+$CRLF '%IDC_LABELHOTKEY
  sSql=sSql+"insert into "+$STRINGLANGTABLENAME+" (name,english,chinese)values('Update','Update','����');"+$CRLF '%IDC_UPDATEHOTKEY

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
' д��������
' ��������������ֵ��ȫ�ֲ���������t_config
FUNCTION WriteConfig(BYVAL ConfigName AS STRING,BYVAL value AS STRING)AS LONG
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  IF Exist($DATABASENAME)=0 THEN '������ݿ��ļ������ڣ��򴴽����ݿ�
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "�����쳣 0001"
      FUNCTION=0
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="�޷��򿪻򴴽����ݿ�,"
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
'��ȡ����
'�������Ϊ������
FUNCTION ReadConfig(BYVAL ConfigName AS STRING)AS STRING
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL rs() AS STRING
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '������ݿ��ļ������ڣ��򴴽����ݿ�
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "�����쳣 0001"
      FUNCTION=""
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="�޷��򿪻򴴽����ݿ�,"
    GOTO abort
  END IF
  sSql="select value from "+$CONFIGTABLENAME+" where name='"+ConfigName+"'"
  IF sqlQuery(hDB,sSql,rs())<0 THEN
    errStr="���ݿ���󣬲�ѯ" & $CONFIGTABLENAME & ","
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
'��ȡĬ������
FUNCTION ReadDefaultConfig(BYVAL ConfigName AS STRING)AS STRING
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL rs() AS STRING
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '������ݿ��ļ������ڣ��򴴽����ݿ�
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "�����쳣 ReadDefaultConfig"
      FUNCTION=""
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="�޷��򿪻򴴽����ݿ�,"
    GOTO abort
  END IF
  sSql="select value from "+$DEFAULTCONFIGTABLENAME+" where name='"+ConfigName+"'"
  IF sqlQuery(hDB,sSql,rs())<0 THEN
    errStr="���ݿ���󣬲�ѯ" & $DEFAULTCONFIGTABLENAME & ","
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
'��ȡ���ݼ�
FUNCTION ReadDataset(BYVAL datasetName AS STRING)AS STRING
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL rs() AS STRING
  LOCAL i AS INTEGER
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '������ݿ��ļ������ڣ��򴴽����ݿ�
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "�����쳣 ReadDataset"
      FUNCTION=""
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="�޷��򿪻򴴽����ݿ�,"
    GOTO abort
  END IF
  sSql="select value from "+$DATASETTABLENAME+" where type='"+datasetName+"'"
  IF sqlQuery(hDB,sSql,rs())<0 THEN
    errStr="���ݿ���󣬲�ѯ" & $DATASETTABLENAME & ","
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
'��ȡĳ��ԭʼ�ַ�������Ӧ�����Զ�Ӧ���ַ���
FUNCTION ReadLang(BYVAL OriString AS STRING,BYVAL selLang AS STRING)AS STRING
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL rs() AS STRING
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '������ݿ��ļ������ڣ��򴴽����ݿ�
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "�����쳣 ReadLang"
      FUNCTION=""
      EXIT FUNCTION
    ELSE
      InitConfig
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="�޷��򿪻򴴽����ݿ�,"
    GOTO abort
  END IF
  sSql="select "+selLang+" from "+$STRINGLANGTABLENAME+" where name='"+OriString+"'"
  IF sqlQuery(hDB,sSql,rs())<0 THEN
    errStr="���ݿ���󣬲�ѯ" & $STRINGLANGTABLENAME & ","
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
'д�����ӻ��޸ģ�ĳ��ԭʼ�ַ�������Ӧ�����Զ�Ӧ���ַ���
FUNCTION WriteLang(BYVAL OriString AS STRING,BYVAL ShowString AS STRING)AS LONG
  LOCAL hDB AS LONG
  LOCAL errStr AS STRING
  LOCAL sSql AS STRING
  LOCAL tmpStr AS STRING
  LOCAL tmpLng AS LONG
  LOCAL rs()   AS STRING
  REDIM rs()
  IF Exist($DATABASENAME)=0 THEN '������ݿ��ļ������ڣ��򴴽����ݿ�
    IF MakeDB($DATABASENAME)=0 THEN
      MSGBOX "�����쳣 WriteLang"
      FUNCTION=0
      EXIT FUNCTION
    ELSE
      InitConfig
      MSGBOX "init..."
    END IF
  END IF
  errStr=""
  IF ISTRUE sqlOpen($DATABASENAME,hDB) THEN
    errStr="�޷��򿪻򴴽����ݿ�,"
    GOTO abort
  END IF
  sSql="select id from "+$STRINGLANGTABLENAME+" where name='"+OriString+"'"
  tmpLng=sqlQuery(hDB,sSql,rs())
  IF tmpLng<0 THEN
    errStr="���ݿ���󣬲�ѯ" & sSql & ","
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
  MSGBOX "�ѳɹ��ؽ� " & $CRLF & OriString & $CRLF & " ��" & language & "���Զ���Ϊ " & $CRLF & GetLang(OriString),_
          %MB_ICONINFORMATION,GetLang("Message")
  RunLog $CRLF & sSql
  FUNCTION=1
  EXIT FUNCTION
abort:
  sqlClose hDB
  MSGBOX errStr
  FUNCTION=0
END FUNCTION
'����ԭʼ�ַ������ص�ǰ�������Ӧ���ַ���
'����ʧ��ʱ������ԭʼ�ַ���
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
' ��¼��־�����ڵ���
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
'������׼����:
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
' �����ֵת�ַ���
FUNCTION PtToStr(BYVAL pt AS POINTAPI)AS STRING
  FUNCTION=FORMAT$(pt.x) & "," & FORMAT$(pt.y)
END FUNCTION
' �ַ���ת�����
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
' ����ֵת�ַ���
FUNCTION RcToStr(BYVAL rc AS RECT)AS STRING
  FUNCTION=FORMAT$(rc.nLeft) & "," & FORMAT$(rc.nTop) & "," _
          & FORMAT$(rc.nRight) & "," & FORMAT$(rc.nBottom)
END FUNCTION
' �ַ���ת����
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
' ����������λ�ÿ��
FUNCTION SaveMainRc()AS LONG
  LOCAL rc AS RECT
  LOCAL tmpStr AS STRING
  GetWindowRect g_hWndMain,rc
  rc.nRight=rc.nRight-rc.nLeft 'nRightת�ɴ���
  rc.nBottom=rc.nBottom-rc.nTop'nBottomת�ɴ�߶�
  rc.nTop=IIF(rc.nTop<0,0,rc.nTop)
  rc.nLeft=IIF(rc.nLeft<0,0,rc.nLeft)
  tmpStr=RcToStr(rc)
  IniWrite EXE.PATH$ & "config.ini","main","rc",tmpStr
END FUNCTION
' �������ļ���ȡ�����ھ��Σ�λ�ÿ��
FUNCTION GetMainRc(rc AS RECT) AS LONG
  LOCAL tmpStr AS STRING
  tmpStr=IniRead(EXE.PATH$ & "config.ini","main","rc","")
  FUNCTION=StrToRc(tmpStr,rc)
END FUNCTION
' ����������λ�ü��ߴ�
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
  'ȷ�����ڿ�Ȳ�������Ļ���
  IF WinPla.rcNormalPosition.nRight - WinPla.rcNormalPosition.nLeft > rcDesktop.nRight THEN
     WinPla.rcNormalPosition.nLeft  = 0
     WinPla.rcNormalPosition.nRight = rcDesktop.nRight
   END IF
  'ȷ�����ڸ߶Ȳ�������Ļ�߶�
  IF WinPla.rcNormalPosition.nBottom - WinPla.rcNormalPosition.nTop > rcDesktop.nBottom THEN
     WinPla.rcNormalPosition.nTop    = 0
     WinPla.rcNormalPosition.nBottom = rcDesktop.nBottom
   END IF
  'ȷ���������߿ɼ�
  IF WinPla.rcNormalPosition.nLeft < 0 THEN
     WinPla.rcNormalPosition.nRight = WinPla.rcNormalPosition.nRight - WinPla.rcNormalPosition.nLeft
     WinPla.rcNormalPosition.nLeft = 0
  END IF
  'ȷ�������Ҳ�߿ɼ�
  IF WinPla.rcNormalPosition.nRight > rcDesktop.nRight THEN
     WinPla.rcNormalPosition.nLeft = WinPla.rcNormalPosition.nLeft - (WinPla.rcNormalPosition.nRight - rcDesktop.nRight)
     WinPla.rcNormalPosition.nRight = rcDesktop.nRight
  END IF
  'ȷ�������ϱ߿ɼ�
  IF WinPla.rcNormalPosition.nTop < 0 THEN
     WinPla.rcNormalPosition.nBottom = WinPla.rcNormalPosition.nBottom - WinPla.rcNormalPosition.nTop
     WinPla.rcNormalPosition.nTop = 0
  END IF
  'ȷ�������±߿ɼ�
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
' ���������ڿͻ���ʵ�崰�ھ��Σ�λ�ÿ�ߣ������������
FUNCTION SaveMainClientRc()AS LONG
  LOCAL rc AS RECT
  LOCAL tmpStr AS STRING
  GetWindowRect g_hWndClient,rc
  ScreenToClientRect g_hWndMain,rc
  rc.nRight=rc.nRight-rc.nLeft 'nRightת�ɴ���
  rc.nBottom=rc.nBottom-rc.nTop'nBottomת�ɴ�߶�
  rc.nTop=IIF(rc.nTop<0,0,rc.nTop)
  rc.nLeft=IIF(rc.nLeft<0,0,rc.nLeft)
  tmpStr=RcToStr(rc)
  IniWrite EXE.PATH$ & "config.ini","main","clientrc",tmpStr
END FUNCTION
' �������ļ���ȡ�����ڿͻ���ʵ�崰�ھ��Σ�λ�ÿ�ߣ������������
FUNCTION GetMainClientRc(rc AS RECT) AS LONG
  LOCAL tmpStr AS STRING
  tmpStr=IniRead(EXE.PATH$ & "config.ini","main","clientrc","")
  FUNCTION=StrToRc(tmpStr,rc)
END FUNCTION
'����Dock��Ϣ�������ļ�
FUNCTION SaveDockInfo()AS LONG
  LOCAL fname AS STRING
  LOCAL i AS LONG
  fname=EXE.PATH$ & "config.ini"
  IniWrite fname,"dock","count",FORMAT$(UBOUND(gdi())+1) '����
  IF UBOUND(gdi())<0 THEN
    REDIM gdi()
    lA    = 0 : tA     = 0 : rA   = 0 : bA  = 0
    RESET lARc : RESET tARc : RESET rARc : RESET bARc
    FUNCTION=0
    EXIT FUNCTION
  END IF
  '���沴����
  IniWrite fname,"dock","leftA"   ,FORMAT$(lA)
  IniWrite fname,"dock","topA"    ,FORMAT$(tA)
  IniWrite fname,"dock","rightA"  ,FORMAT$(rA)
  IniWrite fname,"dock","bottomA" ,FORMAT$(bA)
  IniWrite fname,"dock","leftrc"  ,RcToStr(lARc)
  IniWrite fname,"dock","toprc"   ,RcToStr(tARc)
  IniWrite fname,"dock","rightrc" ,RcToStr(rARc)
  IniWrite fname,"dock","bottomrc",RcToStr(bARc)
  '����DockInfo�ṹ����
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
'�������ļ��л�ȡDock��Ϣ�����浽gdi()
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
' ������ֵ�Ӵ������ֵתΪ��Ļֵ
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
' ������ֵ����ĻֵתΪ�������ֵ
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
' �ƶ����ڣ���MoveWindow��������
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
  ' Initializes MSGBOXPARAMS ��ʼ��MSGBOXPARAMS
  mbp.cbSize = SIZEOF( mbp )      ' Size of the structure�ṹ��С
  mbp.hwndOwner = hWnd        ' Handle of main window �����ھ��
  mbp.hInstance = GETMODULEHANDLE( "" )       ' Instance of applicationӦ�ó���ʵ��
  mbp.lpszText = STRPTR( strMessage )         ' Text of the message��Ϣ���ı�
  mbp.lpszCaption = VARPTR( szCaption )       ' Caption ����
  mbp.dwStyle = dwStyle OR %MB_USERICON       ' Style ����
  mbp.lpszIcon = 100      ' Icon identifier in the resource file ����Դ�ļ��е�ͼ���ʶ
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
  IF iniSection = "" THEN iniSection = "�༭��"
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
  IF iniSection = "" THEN iniSection = "�༭��"
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
' ��ʾ��Ϣ��ʾ����
FUNCTION MsgInfoBox(BYVAL msgStr AS STRING)AS LONG
  MSGBOX msgStr,%MB_OK OR %MB_ICONINFORMATION,"��ʾ"
END FUNCTION
' ��ʾ������ʾ����
FUNCTION MsgErrBox(BYVAL errStr AS STRING)AS LONG
  MSGBOX errStr,%MB_OK OR %MB_ICONERROR,"����"
END FUNCTION
' ��ʾѯ����ʾ����
FUNCTION MsgQueBox(BYVAL msgStr AS STRING)AS LONG
  FUNCTION=MSGBOX(msgStr,%MB_YESNO OR %MB_ICONQUESTION,"��ʾ")
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
' ���������ļ����ĵ�ǰ�����ַ�������:20131012
'-----------------------------------------------
FUNCTION DateTimeForFileName()AS STRING
  LOCAL st AS SYSTEMTIME
  GetLocalTime st
  FUNCTION = FORMAT$(st.wYear,"0000") & FORMAT$(st.wMonth,"00") & FORMAT$(st.wDay,"00") & _
              FORMAT$(st.wHour,"00") & FORMAT$(st.wMinute,"00") & FORMAT$(st.wSecond,"00") & _
              FORMAT$(st.wMilliseconds,"000")
END FUNCTION
