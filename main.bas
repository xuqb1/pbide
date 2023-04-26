'==============================================================================
'
'   �ļ���main.bas
'   ʵ�ֿ��ӻ�PowerBASIC��̵��������ļ�
'
'   MDI �ı��༭�� - ���� PBNote ���ӣ����޸��˲˵����ַ�����ͼ��ٴ��룬��
'   ���������������Ĵ�ӡ�ĵ�����
'
'   ע�⣺win95���ڱ༭���"�µ�"Ϊ16λGDI����ˣ�win95�б༭������Ϊ64k
'         WinNT�еı༭�ؼ��нϸߵ����ƣ�������������Ϊ1M
'==============================================================================
'%USEMACROS         = 1  ' ����ͷ�ļ��ж���ĺ�
'------------------------------------------------------------------------------
#COMPILER PBWIN 9
#COMPILE EXE "PBIDE.exe"
#DIM ALL

#RESOURCE "resource.PBR"
'------------------------------------------------------------------------------
#INCLUDE "contants.inc"   ' ��������
#INCLUDE "globalvar.inc"  'ȫ�ֱ�������
#INCLUDE "include.inc"

%ENABLE_DEBUG = 1
MACRO DP(st)
  #IF %ENABLE_DEBUG
    OutputDebugString BYCOPY st
  #ENDIF
END MACRO
'==============================================================================
FUNCTION WINMAIN (BYVAL hInstance     AS DWORD, _               '���������
                  BYVAL hPrevInstance AS DWORD, _
                  BYVAL lpCmdLine     AS ASCIIZ PTR, _
                  BYVAL iCmdShow      AS LONG) AS LONG
  '----------------------------------------------------------------------------
  ' �������
  '----------------------------------------------------------------------------
  LOCAL hAccel        AS DWORD
  LOCAL Msg           AS TAGMSG
  LOCAL wce           AS WNDCLASSEX
  LOCAL szClassName   AS ASCIIZ * 80
  LOCAL szName        AS ASCIIZ * 10
  LOCAL szText        AS ASCIIZ * 255
  LOCAL x             AS LONG
  LOCAL y             AS LONG
  LOCAL cx            AS LONG
  LOCAL cy            AS LONG
  LOCAL tmpStr        AS STRING
  LOCAL tmpLng        AS DWORD
  LOCAL rc            AS RECT
  LOCAL rs            AS STRING
  tmpStr=@lpCmdLine   '�����в���
  '���ó���ʵ����
  szName = "VPB_xqb"
  '����ȫ�������ļ�����������·��
  g_zIni = EXE.PATH$ & $INIFILENAME
  'Ψһʵ������
  IF VAL(IniRead(g_zIni,"Editor options","AllowMultipleInstances","0"))<>1 THEN
    IF ISFALSE OnlyOneInstance(szName) THEN
      tmpStr="���������У�ȷ������������г���!"
      MsgInfoBox tmpStr
      tmpLng=FindWindow($PROGRAMCLASSNAME,BYVAL 0)
      ShowWindow tmpLng,%SW_SHOW
      SetForegroundWindow tmpLng
      EXIT FUNCTION
    END IF
  END IF
  g_hInst = hInstance '�洢ʵ�������ȫ�ֱ����У�������õ�
  '��ʼ��PB�ؼ���
  InitializePbKeywords
  '�ڶ��û�ģʽ�У���Tsunami SciEditor������ʾ�ļ���
  CodeTipsFile = EXE.PATH$ & $SCI_CODETIPSDB
  'MsgInfoBox "������ʾ�ļ�:" & CodeTipsFile
  hCodetipsFile = trm_Open( BYCOPY CodeTipsFile, %TRUE )
  IF hCodetipsFile < 1 THEN hCodetipsFile = 0
  '�򿪴�����������ݿ�
  hCodeKeepFile = trm_Open( EXE.PATH$ & $SCI_CODEKEEPDB, 1 )
  IF hCodeKeepFile < 1 THEN hCodeKeepFile = 0
  '��������
  hFont = GETSTOCKOBJECT( 17 )        '%ANSI_VAR_FONT)
  '----------------------------------------------------------------------------
  ' ע����������
  szClassName       = $PROGRAMCLASSNAME
  wce.cbSize        = SIZEOF(wce)
  wce.STYLE         = %CS_HREDRAW OR %CS_VREDRAW
  wce.lpfnWndProc   = CODEPTR(WndProc)
  wce.cbClsExtra    = 0
  wce.cbWndExtra    = 0
  wce.hInstance     = g_hInst
  wce.hIcon         = LoadIcon(g_hInst, "APPICON")
  IF wce.hIcon = 0 THEN '���û����Դͼ�꣬��ʹ��ϵͳ��Ӧ�ó���ͼ��
    wce.hIcon = LoadIcon(0, BYVAL %IDI_APPLICATION)
  END IF
  wce.hCursor       = LoadCursor(%NULL, BYVAL %IDC_ARROW)
  wce.hbrBackground = %NULL
  wce.lpszMenuName  = %NULL
  wce.lpszClassName = VARPTR(szClassName)
  wce.hIconSm       = LoadIcon(g_hInst, BYVAL %IDI_APPLICATION)
  IF RegisterClassEx(wce) = 0 THEN
    RegisterClass BYVAL (VARPTR(wce) + 4)
  END IF
  szClassName         = $MainEdit
  wce.hbrBackground   = GETSTOCKOBJECT( %LTGRAY_BRUSH )  '�༭����
  wce.lpfnWndProc     = CODEPTR( WndClientProc )
  wce.style           = %CS_HREDRAW OR %CS_VREDRAW OR %CS_BYTEALIGNCLIENT _
                        OR %CS_DBLCLKS
  wce.hIcon           = LOADICON( %NULL, BYVAL %IDI_APPLICATION )
  wce.hIconSm         = LOADICON( %NULL, BYVAL %IDI_APPLICATION )
  REGISTERCLASSEX wce
  szClassName         = $MainMDIEdit
  wce.hbrBackground   = GETSTOCKOBJECT( %LTGRAY_BRUSH )  '�༭MDI����
  wce.lpfnWndProc     = CODEPTR( WndMDIClientProc )
  REGISTERCLASSEX wce
  '----------------------------------------------------------------------------
  ' ע����봰����
  CALL initEditClass()  'edit.bas 2017/11/23 21:10
  ' ** Register (Print Preview) Display Window Class **  ע�ᣨ��ӡԤ������ʾ������
  szClassName         = "NPREVIEW32"
  wce.cbSize          = SIZEOF(wce)
  wce.style           = %CS_HREDRAW OR %CS_VREDRAW
  wce.lpfnWndProc     = CODEPTR(PreviewProc)
  wce.cbClsExtra      = 0
  wce.cbWndExtra      = 4
  wce.hInstance       = hInstance
  wce.hIcon           = LoadIcon(hInstance, BYVAL 108)
  wce.hCursor         = %NULL
  wce.hbrBackground   = %NULL
  wce.lpszMenuName    = %NULL
  wce.lpszClassName   = VARPTR(szClassName)
  wce.hIconSm         = LoadIcon(hInstance, BYVAL 108)
  IF ISFALSE(RegisterClassEx(wce)) THEN RegisterClass BYVAL (VARPTR(wce) + 4)
  '----------------------------------------------------------------------------
  '���طָ����ϵĹ��
  hCurSplit = LOADIMAGE( GETMODULEHANDLE( szName ), "HRZCURSOR", _
                          %IMAGE_CURSOR, GETSYSTEMMETRICS( %SM_CXCURSOR ), _
                          GETSYSTEMMETRICS( %SM_CYCURSOR ), %LR_DEFAULTSIZE )
  hCurSplitV = LOADIMAGE( GETMODULEHANDLE( BYVAL %NULL ), "VRTCURSOR", _
                          %IMAGE_CURSOR, GETSYSTEMMETRICS( %SM_CXCURSOR ), _
                          GETSYSTEMMETRICS( %SM_CYCURSOR ), %LR_DEFAULTSIZE )
  hCurNorm = LOADCURSOR ( %NULL, BYVAL %IDC_ARROW )
  x=(GetSystemMetrics(%SM_CXSCREEN) - 640) / 2      '�������ã����ô���������ʱ
  y=(GetSystemMetrics(%SM_CYSCREEN) - 480) / 2      '   �Ĵ�С��λ��
  cx=640
  cy=480
  tmpStr=ReadConfig("mainwindowposopt")       '��ȡ������λ������ѡ��
  IF tmpStr="1" THEN                          '����Ϊ��ʹ�ù̶�λ�úʹ�С
    tmpStr=ReadConfig("mainwindowpos")
    x=VAL(MID$(tmpStr,1,INSTR(tmpStr,",")-1))
    y=VAL(MID$(tmpStr,INSTR(tmpStr,",")+1))
    tmpStr=ReadConfig("mainwindowsize")
    cx=VAL(MID$(tmpStr,1,INSTR(tmpStr,",")-1))
    cy=VAL(MID$(tmpStr,INSTR(tmpStr,",")+1))
  ELSEIF tmpStr="2" THEN                     '����Ϊ:ʹ�����һ�εĴ�С��λ��
    IF GetMainRc(rc)=0 THEN
      rc.nLeft    = 100
      rc.nTop     = 50
      rc.nRight   = 800
      rc.nBottom  = 600
    END IF
    x=rc.nLeft
    y=rc.nTop
    cx=rc.nRight
    cy=rc.nBottom
  END IF
  IF x<=0 THEN                               '�ų�λ�úʹ�С�����쳣
    x=(GetSystemMetrics(%SM_CXSCREEN) - 640) / 2
  END IF
  IF y<=0 THEN
    y=(GetSystemMetrics(%SM_CYSCREEN) - 480) / 2
  END IF
  IF cx<640 THEN
    cx=640
  END IF
  IF cy<480 THEN
    cy=480
  END IF
  RegisterDockClass
  '----------------------------------------------------------------------------
  ' ��ע������ഴ��������
  g_hWndMain = CreateWindow($PROGRAMCLASSNAME, _    ' ����������
                            "PBIDE v1.0", _         ' �����ڳ�ʼ����
                            %WS_OVERLAPPEDWINDOW, _ ' ��������
                            x, _                    ' ��ʼ x λ��
                            y, _                    ' ��ʼ y λ�� ��Ļ������
                            cx, _                   ' ��ʼ���
                            cy, _                   ' ��ʼ�߶�
                            %NULL, _                ' �����ھ��
                            g_hMenu, _              ' ���ڲ˵����
                            g_hInst, _              ' ����ʵ�����
                            BYVAL %NULL)            ' ��������

  '----------------------------------------------------------------------------
  IF LEN(COMMAND$) THEN
    PostMessage g_hWndMain,%WM_USER + 1000,0,0  '���ʹ�����������Ϣ
  END IF

  'GetRecentFiles                              '�������ã����������ļ��б�
  'HideButtons %TRUE                           '���ز��ֲ˵�

  '�õ��༭��ѡ������ڲ˵�����֮������
  GetEditorOptions EdOpt
  IF ISTRUE EdOpt.StartInLastFolder THEN
    IF LEN( Edopt.LastFolder ) THEN CHDIR EdOpt.LastFolder
  END IF
  '�õ�ȫ���ַ����е����͡�
  SED_GetTypes
  ' �õ�������ѡ��
  GetCompilerOptions CpOpt
  ' ��״̬������ʾĬ�ϵı�����
  IF CpOpt.DefaultCompiler = 1 THEN
    szText = " PBWIN"
  ELSEIF CpOpt.DefaultCompiler = 2 THEN
    szText = " PBCC"
  END IF
  SENDMESSAGE g_hStatus, %SB_SETTEXT, 3, VARPTR( szText )
  '����/���ù�������Ĳ��ְ�ť����ز˵���
  ChangeButtonsState
  '�õ���ɫ������ѡ��
  GetColorOptions SciColorsAndFonts
  '�����Ӳ˵��ı�������ˢ��
  hMenuTextBkBrush = CREATESOLIDBRUSH( SciColorsAndFonts.SubmenuTextBackColor )
  '�����Ӳ˵������ı�������ˢ��
  hMenuHiBrush = CREATESOLIDBRUSH( SciColorsAndFonts.SubmenuHiTextBackColor )
  'hMenuHiBrush = CreateSolidBrush(RGB(64, 134, 191))
  ' �����˵�ͼ�걳��ˢ��
  hMenuIconsBrush = CREATESOLIDBRUSH( RGB( 193, 211, 239 ))
  'LOCAL g_zIni AS STRING
  g_zIni=EXE.PATH$ & "config.ini"
  rs = IniRead( g_zIni, "Scintilla Print Options", "WrapMode", "" )
  IF rs = "" THEN IniWrite g_zIni, "Scintilla Print Options", "WrapMode", FORMAT$( %SC_WRAP_WORD )
  ' ��ק�ļ�����
  DRAGACCEPTFILES g_hWndMain, %TRUE
  ' ����û�ѡ���˴�ѡ��,����ʱ���MDI����
  IF EdOpt.MaximizeEditWindows = %BST_CHECKED THEN fMaximize = %TRUE
  IF EdOpt.MaximizeMainWindow = %BST_CHECKED THEN
    ' ���������
    SHOWWINDOW g_hWndMain, %SW_MAXIMIZE
  ELSE
    ' ������������ʾ״̬
    SHOWWINDOW g_hWndMain, iCmdShow
  END IF
  '  ���������ڵĿͻ���
  UpdateWindow g_hWndMain
  ' �������õ���һ���ļ�
  IF ISTRUE EdOpt.ReloadFilesAtStartup AND LEN( COMMAND$ ) = 0 THEN
    SED_LoadPreviousFileSet
  END IF
  'tmpStr=ReadConfig("showflashwindow")
  tmpStr=IniRead(g_zIni,"startshow","flashwindow","1")
  IF tmpStr="1" THEN
    CreateSplash g_hInst
  END IF
  'tmpStr=ReadConfig("showguiderwindow")
  tmpStr=IniRead(g_zIni,"startshow","guiderwindow","1")
  IF tmpStr="1" THEN
    g_hGuiderWin=CreateGuiderWindow(g_hWndMain)
  END IF
  '----------------------------------------------------------------------------
  ' ��Ϣ����ѭ��
  WHILE GetMessage(Msg, BYVAL %NULL, 0, 0)
    IF ISFALSE TranslateMDISysAccel(g_hWndClient, Msg) THEN
      IF ISFALSE TranslateAccelerator(g_hWndMain, hAccel, Msg) THEN
        IF ISFALSE IsDialogMessage(hFind,Msg) THEN
          IF ISFALSE IsDialogMessage(g_hWndMain,Msg) THEN
            TranslateMessage Msg
            DispatchMessage Msg
          END IF
        END IF
      END IF
    END IF
  WEND
  FUNCTION = msg.wParam
END FUNCTION  ' WinMain ����
' *********************************************************************************************
' ��ȡ�����в�����������
' *********************************************************************************************
SUB ProcessCommandLine (BYVAL hWnd AS DWORD)
  LOCAL lsFileName   AS STRING           ' �ļ���
  LOCAL lsLineNumber AS STRING           ' ����
  LOCAL DataToSend   AS COPYDATASTRUCT   ' ���͵����ݽṹ
  IF ISFALSE hWnd THEN EXIT SUB
  IF IsIconic(hWnd) <> 0 OR IsZoomed(hWnd) <> 0 THEN ShowWindow hWnd, %SW_RESTORE
  SetForegroundWindow hWnd
  lsLineNumber = RIGHT$(COMMAND$, LEN(COMMAND$) - INSTR(-1, COMMAND$, " "))
  IF LEN(TRIM$(lsLineNumber, ANY "0123456789")) THEN
    lsFileName = TRIM$(COMMAND$, ANY $DQ + $SPC)
  ELSE
    lsFileName = TRIM$(LEFT$(COMMAND$, LEN(COMMAND$) - LEN(lsLineNumber)), ANY $DQ + $SPC)
  END IF
  DataToSend.lpData  = STRPTR(lsFileName)
  DataToSend.cbdata  = LEN(lsFileName) + 1
  DataToSend.dwData  = VAL(lsLineNumber) - 1
  IF RIGHT$(UCASE$(lsFileName), 4) = ".PBP" THEN
    ' ��ȡ������
    lsFileName = DIR$(COMMAND$)
    ' ����ļ�����, ������Ŀ
    IF LEN(lsFileName) THEN ProjectLoadTree(COMMAND$)
  ELSE
    SendMessage hWnd, %WM_COPYDATA, LEN(DataToSend), VARPTR(DataToSend)
    IF GetEdit THEN SendMessage GetEdit, %SCI_GOTOLINE, VAL(lsLineNumber) - 1, 0
  END IF
END SUB
' *********************************************************************************************
' Save the path of the loaded files
' *********************************************************************************************
SUB SaveLoadedFilesPaths
  LOCAL hWndActive AS DWORD                ' // Handle of the active child window
  LOCAL szPath     AS ASCIIZ * %MAX_PATH   ' // File path
  LOCAL p          AS LONG                 ' // Number of child windows
  LOCAL p1         AS LONG                 ' // Window handle
  LOCAL fNumber    AS LONG                 ' // File number
  LOCAL curPos     AS LONG                 ' // Caret position
  LOCAL i          AS LONG                 ' // Counter
  LOCAL buffer     AS STRING               ' // Buffer for bookmark lines
  LOCAL bk         AS LONG                 ' // Bookmark counter
  LOCAL fMark      AS LONG                 ' // Bookmark flag variable
  LOCAL nLine      AS LONG                 ' // Number of line
  LOCAL strLine    AS STRING               ' // Line of text to write to the file
  DIM rgFiles()    AS STRING               ' // Array of filenames and paths

  ON ERROR GOTO ErrHandler

  ' Count the number of child windows
  p = 0
  p1 = GetWindow(g_hWndClient, %GW_CHILD)
  DO WHILE p1 <> 0
    p1 = GetWindow(p1, %GW_HWNDNEXT)
    INCR p
  LOOP

  'Save the paths, positions and bookmarks
  FOR i = 1 TO p
    GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
    curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)
    IF (INSTR(szPath, ANY ":\/") <> 0) THEN
      bk = 0 : buffer = ""
      fMark = 0 : BIT SET fMark, 0
      nLine = SendMessage(GetEdit, %SCI_MARKERNEXT, 0, fMark)
      DO WHILE nLine <> -1
        INCR bk
        buffer = buffer + "|" + FORMAT$(nLine)
        fMark = 0 : BIT SET fMark, 0
        nLine = SendMessage(GetEdit, %SCI_MARKERNEXT, nLine + 1, fMark)
      LOOP
      strLine = szPath & "|" & FORMAT$(curPos) & "|" & FORMAT$(bk) & buffer
      REDIM PRESERVE rgFiles(UBOUND(rgFiles) + 1)
      rgFiles(UBOUND(rgFiles)) = strLine
    END IF
    MdiNext g_hWndClient, hWndActive, 0      ' Set focus to next child window
  NEXT

  ' Save the project name
  IF LEN(sProjectName) THEN
    IF UCASE$(sProjectName) <> "NONAME.SPF" THEN
      REDIM PRESERVE rgFiles(UBOUND(rgFiles) + 1)
      rgFiles(UBOUND(rgFiles)) = sProjectName
    'ELSEIF gdi.l <> "I" THEN
    '  REDIM PRESERVE rgFiles(UBOUND(rgFiles) + 1)
    '  rgFiles(UBOUND(rgFiles)) = sProjectName
    END IF
  END IF

  fNumber = FREEFILE
  OPEN EXE.PATH$ & "SED.LFS" FOR OUTPUT AS fNumber

  IF UBOUND(rgFiles) > - 1 THEN
    ARRAY SORT rgFiles()
    FOR i = LBOUND(rgFiles) TO UBOUND(rgFiles)
      PRINT #fNumber, rgFiles(i)
    NEXT
  END IF

ErrHandler:

  CLOSE fNumber

END SUB
'==============================================================================
FUNCTION WndProc (BYVAL hWnd AS DWORD, BYVAL wMsg AS DWORD, _
                  BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG
  '------------------------------------------------------------------------------
  ' ������
  '----------------------------------------------------------------------------
  LOCAL  hMdi             AS DWORD
  LOCAL  hWndFirst        AS DWORD                ' ��һ���Ӵ��ھ��
  LOCAL  hWndActive       AS DWORD                ' ������Ӵ��ھ��
  LOCAL  hWndPrev         AS DWORD                ' ��һ���Ӵ��ھ��
  LOCAL  PID              AS LONG                 ' ���̱�ʶ
  LOCAL  hr               AS DWORD                ' �������
  LOCAL  hSci             AS DWORD                ' SCI�ؼ����
  LOCAL  curSel           AS LONG                 ' ��ǰѡ��
  LOCAL  szText           AS ASCIIZ * 256         ' ������ʱ����
  LOCAL  fMark            AS LONG                 ' ��ǩ��ʶ����
  LOCAL  nLine            AS LONG                 ' ����
  LOCAL  szPath           AS ASCIIZ * %MAX_PATH   ' ·��
  LOCAL  szFilePath       AS ASCIIZ * %MAX_PATH   ' �ļ�·��
  LOCAL  strFileName      AS STRING               ' �ļ���
  LOCAL  hCalc            AS DWORD                ' ���������
  LOCAL  nTab             AS LONG                 ' ��ǩ��
  LOCAL  i                AS LONG                 ' ѭ��������
  LOCAL  tPROC            AS PROC                 ' PROC �ṹ
  STATIC hFocus           AS DWORD                ' ��ý���Ŀؼ����
  DIM    DroppedFiles(0)  AS STRING               ' �����϶��ļ���·�����ļ���������

  LOCAL  wi               AS WININFOSTRUC PTR     ' WININFOSTRUC �ṹָ��
  LOCAL  pSciData         AS DWORD                ' sci�ؼ�ֱ��ָ��
  LOCAL  fFlag            AS LONG                 ' ��ʶ
  LOCAL  nWidth           AS LONG                 ' ���
  LOCAL  bitnum           AS LONG                 ' λ��

  ' ���� / �滻 (�����Ǿ�̬��)
  STATIC fr               AS FINDREPLACE          ' FINDREPLACE �ṹ
  STATIC szFindText       AS ASCIIZ * 256         ' Ҫ���ҵ��ı�
  STATIC szReplText       AS ASCIIZ * 256         ' �滻Ϊ���ı�
  STATIC szLastFind       AS ASCIIZ * 256         ' ����ҵ��ĵ���
  STATIC dwFindMsg        AS DWORD                ' ע����Ϣ���
  STATIC startPos         AS LONG                 ' ��ʼλ��
  STATIC endPos           AS LONG                 ' ����λ��
  STATIC curPos           AS LONG                 ' ��ǰλ��
  STATIC findFlags        AS LONG                 ' ���ұ�ʶ
  STATIC updown           AS LONG                 ' ֱ�Ӳ���
  LOCAL  numitems         AS LONG                 ' �滻����
  LOCAL  startSelPos      AS LONG                 ' ѡ��ʼλ��
  LOCAL  endSelPos        AS LONG                 ' ѡ�����λ��
  LOCAL  fInSelection     AS LONG                 ' ����ѡ���ı��в���
  LOCAL  buffer           AS STRING               ' һ����;����
  LOCAL  strCommand       AS STRING               ' ������
  LOCAL  strTxt           AS STRING               ' һ����;�ַ�������
  LOCAL  x                AS LONG                 ' ���ʿ�ʼλ��
  LOCAL  y                AS LONG                 ' ���ʽ���λ��

  ' ��������(֧�� Lynx)
  LOCAL  pDataToGet       AS COPYDATASTRUCT PTR   ' CopyData �ṹָ��
  LOCAL  pzBuffer         AS ASCIIZ PTR           ' ����ָ��
  LOCAL  tmpLng           AS LONG
  LOCAL  ToolHeight       AS LONG
  LOCAL  StatHeight       AS LONG
  LOCAL  dwProc           AS DWORD
  LOCAL  hLib             AS DWORD

  LOCAL  dwStyle          AS DWORD
  LOCAL  f                AS STRING               '����������ļ��б���ȡ�ļ���
  LOCAL  PATH             AS STRING
  LOCAL  pt               AS POINTAPI
  LOCAL  tRect            AS RECT
  LOCAL  ClientRect       AS RECT
  LOCAL  mii              AS MENUITEMINFO         '�˵�����Ϣ������������ļ��б�˵�
  LOCAL  tbn              AS TBNOTIFY PTR
  LOCAL  lpToolTip        AS TOOLTIPTEXT PTR
  LOCAL  cc               AS CLIENTCREATESTRUCT
  LOCAL  iccex            AS INIT_COMMON_CONTROLSEX
  LOCAL  zText            AS ASCIIZ * %MAX_PATH
  LOCAL  hPopupMenu       AS DWORD                'Rebar���Ҽ������˵�
  LOCAL  MenuState        AS DWORD                '�Ӳ˵�״̬����
  LOCAL  rc               AS RECT
  LOCAL  tmpWnd           AS DWORD                '��������Ҽ���������ȡ�������������ʱ����
  'LOCAL  i                AS INTEGER
  LOCAL  tmpDockInfo      AS DOCKINFO PTR
  LOCAL  tmpRC            AS RECT
  LOCAL  tmpStr           AS STRING
  LOCAL  tmpArr()         AS STRING
  LOCAL  clientRc         AS RECT
  STATIC OldRC            AS RECT
  LOCAL  hDC              AS LONG
  LOCAL  ps               AS PAINTSTRUCT
  LOCAL  lpNmh            AS NMHDR PTR        ' NMHDR�ṹ��ָ��
  LOCAL  strTabTxt        AS STRING
  'LOCAL  nTab             AS LONG
  LOCAL  ht               AS TC_HITTESTINFO
  'LOCAL  pInfo            AS TC_HITTESTINFO
  STATIC tabIndexClose    AS LONG
  LOCAL  hAccel           AS DWORD
  LOCAL  szFmt            AS ASCIIZ * 15
  LOCAL  szDate           AS ASCIIZ * %MAX_PATH

  SELECT CASE AS LONG wMsg
    CASE %WM_NCACTIVATE
      '�����н���Ŀؼ����
      IF ISFALSE wParam THEN hFocus = GetFocus
      ClipCursor BYVAL %NULL
    CASE %WM_SETFOCUS
      IF hFocus THEN
        PostMessage hWnd,%WM_USER+999,hFocus,0
        hFocus=0
      END IF
    CASE %WM_USER+999
      IF ISTRUE wParam THEN SetFocus wParam
      ShowLinCol
      EXIT FUNCTION
    CASE %WM_USER + 1000
      ' ����������
      ProcessCommandLine hWnd
      EXIT FUNCTION
    CASE %WM_USER + 1001
      ' ���ô����������������
      ResetCodefinder
      ' ���༭�ؼ�����
      SetFocus GetEdit
      EXIT FUNCTION
    CASE %WM_CREATE
      tabIndexClose=-1
      '------------------------------------------------------------------------
      ' ĳЩCommCtrl.dll�Ŀؼ�����Ϣ����ʽ�������ڱ�׼Win95A�汾�У������Ҫ��ʼ����dll
      '------------------------------------------------------------------------
      g_NewComCtl = InitComctl32 (%ICC_BAR_CLASSES)
      g_hWndMain=hWnd
      AppLoadBitmaps()
      g_hRebar=CreateRebar(hWnd)
      IF g_hRebar>0 THEN '�޸Ĳ˵�����״̬
        CheckRebarMenu ghRebarMenu
      END IF
      MakeMenu(g_hWndMain)
      'SetMenuLang                                       '����ѡ����������ò˵��ı�
      'hAccel = MakeAccelerators()                 '�������ã��������ٱ�
      hAccel = CreateAccelTable()
      '------------------------------------------------------------------------
      ' ����״̬��
      g_hStatus = CreateStatusWindow (%WS_CHILD OR %WS_VISIBLE OR %SBS_SIZEGRIP, _
                                      "", hWnd, %ID_STATUSBAR)
      Statusbar_SetPartsBySize(g_hStatus, "80, 90, 230, 60, -1")
      szFmt = "yyyy/MM/dd"
      GetDateFormat %LOCALE_USER_DEFAULT, 0, BYVAL %NULL, szFmt, szDate, SIZEOF(szDate)
      SendMessage g_hStatus, %SB_SETTEXT, 0, VARPTR(szDate)
'      '����״̬��Ϊ����
'      DIM prts(4) AS LONG
'      prts(0) = 250 : prts(1) = 100 : prts(2)= 80 : prts(3)= 50 : prts(4)=-1
'      Sendmessage g_hStatus, %SB_SETPARTS, 5, VARPTR(prts(0))
'      SendMessage g_hStatus, %SB_SETTEXT, 0, VARPTR(zText)
      ' ���� MDI �ͻ�����
      cc.idFirstChild = 1
      cc.hWindowMenu  = g_hMenuWindow  '���ڴ��ڲ˵��е��ļ��б�,��"����"�˵���ͻ��������
      GetClientRect hWnd,tmpRc
      IF GetMainClientRc(clientRc)=0 THEN
        clientRc=tmpRc
      END IF
      IF clientRc.nLeft<0 THEN clientRc.nLeft=0
      IF clientRc.nTop<0  THEN clientRc.nTop=0
      IF clientRc.nRight>tmpRc.nRight-tmpRc.nTop THEN
        clientRc.nRight=tmpRc.nRight-clientRc.nLeft
      END IF
      IF clientRc.nBottom>tmpRc.nBottom-tmpRc.nBottom THEN
        clientRc.nBottom=tmpRc.nbottom-clientRc.nTop
      END IF
      g_hWndClient = CreateWindowEx(%WS_EX_CLIENTEDGE, "MDICLIENT", BYVAL %NULL, _
                                    %WS_CHILD OR %WS_CLIPCHILDREN OR _
                                    %WS_VISIBLE OR %WS_VSCROLL OR %WS_HSCROLL, _
                                    clientRc.nLeft, clientRc.nTop, _
                                    clientRc.nRight, clientRc.nBottom, _
                                    hWnd, %ID_CLIENTWINDOW, g_hInst, cc)
      OldRc=clientRc
      tmpStr=ReadConfig("mainwindowposopt")
      GetDockInfo : CreateDockWin
      CreateTabMdiCtl
      GetRecentFiles
      GetRecentProjects
      FUNCTION = 0
      EXIT FUNCTION
    CASE %WM_SIZE
      IF wParam <> %SIZE_MINIMIZED THEN
        SizeDockWin lParam,wParam
      END IF
      FUNCTION=0 : EXIT FUNCTION
    CASE %WM_PAINT '������Ϣ
      hDC = BeginPaint( hWnd, ps ) '���ڱ䶯ʱ����ˢ��
      EndPaint hWnd, ps
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_DROPFILES
      ' ����϶����ļ�
      REDIM DroppedFiles(0)
      ' ÿ���ļ�����һ���µı༭����
      IF ISTRUE GetDroppedFiles(wParam, DroppedFiles()) THEN
        FOR i = LBOUND(DroppedFiles) TO UBOUND(DroppedFiles)
          IF LEN(DroppedFiles(i)) THEN OpenThisFile(DroppedFiles(i))
        NEXT
      END IF
      EXIT FUNCTION
    CASE %WM_MOUSEACTIVATE
      IF HIWRD(lParam)= %WM_RBUTTONDOWN THEN '���������Ҽ��˵�չʾ
        GetCursorPos pt
        tmpWnd = WindowFromPoint(pt.x,pt.y)
        IF tmpWnd=ghToolBar OR tmpWnd=ghComboBox OR tmpWnd=ghWindowbar OR _
                tmpWnd=ghDebugbar OR tmpWnd=ghWinArragBar OR tmpWnd=ghCodeEditBar THEN
          IF ghRebarMenu<=0 THEN
            ghRebarMenu=MakeRebarMenu()
          END IF
          TrackPopupMenu ghRebarMenu,%TPM_LEFTALIGN OR %TPM_RIGHTBUTTON, _
                    pt.x,pt.y,0,hWnd,BYVAL 0
        END IF
      END IF
    CASE %WM_NOTIFY '֪ͨ��Ϣ
      tbn = lParam
      IF @tbn.hdr.code = %TTN_NEEDTEXT THEN         '������ʾ
        lpToolTip = lParam
        zText = GetStringTable(@lpToolTip.hdr.idFrom)
        @lpToolTip.lpszText = VARPTR(zText)
        ' Tab ��ʾ
        GetCursorPos ht.pt
        ScreentoClient g_hTabMdi,ht.pt
        i = TabCtrl_HitTest(g_hTabMdi,ht)
        IF i>=0 THEN
          zText=""
          IF i<=UBOUND(gTabFilePaths()) THEN
            zText=gTabFilePaths(i)
            @lpToolTip.lpszText = VARPTR(zText)
          END IF
        END IF
      ELSEIF g_NewComCtl >= 4.7 AND @tbn.hdr.code = %TBN_DROPDOWN THEN   '�򿪰�ť�������˵�
        GetClientRect hWnd, rc
        SELECT CASE AS LONG @tbn.iItem
          CASE %IDM_NEWBAS
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %IDM_NEWPROJECT, GetLang("New Project")
            'APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %IDM_NEW, GetLang("New BAS File")
            pbtempfile = GetTemplate()
            LOCAL tempArr() AS STRING
            REDIM tempArr(PARSECOUNT(pbtempfile,$CRLF)-1)
            PARSE pbtempfile,tempArr(),$CRLF
            FOR i=0 TO UBOUND(tempArr())
              APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %IDM_PBTEMPLATES + i +1,MID$(tempArr(i),1,INSTR(tempArr(i),",")-1)
            NEXT i
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
          CASE %IDM_OPEN
            SendMessage @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR(tRect)
            MapWindowPoints @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR(tRect), 2
            TrackPopupMenu g_hMenuReopen, 0, tRect.nLeft, tRect.nBottom, 0, hWnd, BYVAL %NULL
          CASE %IDM_WIN  '�½����ڰ�ť
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %IDM_WINTEMPLATES + 1, GetLang("Normal Window") '"һ�㴰��" 'BYVAL %NULL
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %IDM_WINTEMPLATES + 2, GetLang("MDI Window")'"MDI����" 'BYVAL %NULL
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
          CASE %ID_SETV_CENTER '��ֱ���а�ť
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_TOP, "�϶���"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_CENTER, "��ֱ����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_BOTTOM, "�¶���"
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
          CASE %ID_SETV_TOP '�϶��밴ť
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_TOP, "�϶���"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_CENTER, "��ֱ����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_BOTTOM, "�¶���"
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
          CASE %ID_SETV_BOTTOM '�¶��밴ť
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_TOP, "�϶���"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_CENTER, "��ֱ����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETV_BOTTOM, "�¶���"
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
          CASE %ID_SETH_CENTER  'ˮƽ���а�ť
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_LEFT, "�����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_CENTER, "ˮƽ����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_RIGHT, "�Ҷ���"
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
          CASE %ID_SETH_LEFT  '����밴ť
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_LEFT, "�����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_CENTER, "ˮƽ����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_RIGHT, "�Ҷ���"
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
          CASE %ID_SETH_RIGHT  '�Ҷ��밴ť
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_LEFT, "�����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_CENTER, "ˮƽ����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %ID_SETH_RIGHT, "�Ҷ���"
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
          CASE %ID_TOOL_COMPILE
            SENDMESSAGE( @tbn.hdr.hwndFrom, %TB_GETRECT, @tbn.iItem, VARPTR( rc ))
            MAPWINDOWPOINTS( @tbn.hdr.hwndFrom, %HWND_DESKTOP, BYVAL VARPTR( rc ), 2 )
            hPopupMenu = CREATEPOPUPMENU
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %MN_F_DDTCODE, "DDT����"
            APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %MN_F_SDKCODE, "SDK����"
            TRACKPOPUPMENU( hPopupMenu, 0, rc.nLeft, rc.nBottom, 0, hWnd, BYVAL %NULL )
            DESTROYMENU hPopupMenu
        END SELECT
      END IF
      IF @tbn.hdr.hwndFrom=g_hRebar AND @tbn.hdr.code=%RBN_ENDDRAG THEN '�������϶�����
        SendMessage hWnd,%WM_SIZE,wParam, lParam
      END IF
      ' Tab notifications Tab֪ͨ
      lpNmh = lParam      ' Init NMH pointer��ʼ��NMHָ��
      IF @lpNmh.idFrom = %IDC_TABMDI THEN
        SELECT CASE @lpNmh.Code         ' Examine .Code member ���.Code��Ա
'        CASE %TCN_LAST TO %TCN_FIRST        ' Tab control notifications Tab�ؼ�֪ͨ
'          SELECT CASE @lpNmh.idFrom
'            CASE %IDC_TABMDI
'              SELECT CASE @lpNmh.Code
          CASE %TCN_SELCHANGE         ' Identify which tab ʶ�����ĸ�Tab
            nTab = SENDMESSAGE( g_hTabMdi, %TCM_GETCURSEL, 0, 0 )
            IF UBOUND( gTabFilePaths ) = > nTab THEN
              strTabTxt = gTabFilePaths( nTab )
              IF RIGHT$(strTabTxt,1)="*" THEN
                strTabTxt=MID$(strTabTxt,1,LEN(strTabTxt)-2)
              END IF
              SED_ActivateMdiWindow strTabTxt
              IF FileIsInProject( strTabTxt ) THEN
                TreeView_FindItem( ghTreeView, GetFileName( strTabTxt ), %FALSE, %NULL, %TRUE, %TRUE )
              END IF
            END IF
          CASE %NM_RCLICK
            GetCursorPos ht.pt
            ScreenToClient g_hTabMdi, ht.pt
            ht.flags = %TCHT_ONITEM
            i = SendMessage(g_hTabMdi, %TCM_HITTEST, 0, VARPTR(ht))
            IF i <> -1 THEN
              tabIndexClose=i
              ClientToScreen g_hTabMdi,ht.pt
              hPopupMenu = CREATEPOPUPMENU
              APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %IDM_CLOSE, "�ر�"
              APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %IDM_CLOSEOTHERS, "�ر�����"
              APPENDMENU hPopupMenu, %MF_ENABLED OR %MF_OWNERDRAW, %IDM_CLOSEALL, "�ر�ȫ��"
              TRACKPOPUPMENU( hPopupMenu, 0, ht.pt.x, ht.pt.y, 0, hWnd, BYVAL %NULL )
              DESTROYMENU hPopupMenu
            END IF
        END SELECT
      END IF
'          END SELECT
'      END SELECT
      EXIT FUNCTION
    CASE %WM_MENUSELECT                           'ѡ���˲˵�����״̬��������ʾ��Ϣ
      zText = GetStringTable(LOWRD(wParam))
      SendMessage g_hStatus, %SB_SETTEXT, 2, VARPTR(zText)
      EXIT FUNCTION
    CASE %WM_DRAWITEM
      ' �Ի�˵���Ҫ�ػ�
      IF wParam = 0 THEN      ' �����ʶ����0����ʾ��Ϣ�ɲ˵�����
        DrawMenu lParam         ' Draw the menu���Ʋ˵�
        FUNCTION = %TRUE
        EXIT FUNCTION
      END IF
    CASE %WM_MEASUREITEM        ' Get menu item size�õ��˵���ߴ�
      IF wParam = 0 THEN      ' A menu is calling�˵�������
        MeasureMenu hWnd, lParam        ' Do all work in separate Sub�ڶ�����Sub�д���
        FUNCTION = %TRUE
        EXIT FUNCTION
      END IF
    CASE %WM_COPYDATA
      ' For Lynx support
      pDataToGet = lParam
      pzBuffer   = @pDataToGet.lpData
      strTxt = TRIM$(@pzBuffer)
      OpenThisFile strTxt
      endPos = SendMessage (GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
      SendMessage GetEdit, %SCI_GOTOLINE, endPos, 0
      SendMessage GetEdit, %SCI_GOTOLINE, @pDataToGet.dwData, 0
    CASE %WM_NCPAINT '��������˱�ʶ���򷵻����Ա���˵�����
      IF g_FreezeMenu THEN EXIT FUNCTION
    ' �ɲ���/�滻�Ի����͵���Ϣ
    CASE dwFindMsg
      IF (fr.flags AND %FR_DIALOGTERM) = %FR_DIALOGTERM THEN
        IF szLastFind <> szFindText THEN szLastFind = szFindText
        fr.lCustData = 0       ' ���
      ELSE
        IF (fr.flags AND %FR_FINDNEXT) = %FR_FINDNEXT THEN
          ' ������һ��
          ' ��ȡ�Ի�����ѡ���ļ��״̬
          findFlags = 0
          IF (fr.Flags AND %FR_MATCHCASE) = %FR_MATCHCASE THEN findFlags = findFlags OR %SCFIND_MATCHCASE
          IF (fr.Flags AND %FR_WHOLEWORD) = %FR_WHOLEWORD THEN findFlags = findFlags OR %SCFIND_WHOLEWORD
          ' If FR_DOWN = FALSE ��������
          updown = 0
          IF (fr.Flags AND %FR_DOWN) = %FR_DOWN THEN updown = %FR_DOWN
          ' �ӵ�ǰλ�ÿ�ʼ����
          curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)
          IF startPos <> curPos THEN
            IF updown = %FR_DOWN THEN
              startPos = curPos
            ELSE
              IF curPos < startPos THEN startPos = curPos
              IF startPos = 0 THEN startPos = curPos
            END IF
          END IF
          ' ��������ʱ�����һ�б���С�ڿ�ʼλ��
          endPos = SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
          IF updown <> %FR_DOWN THEN
            endPos = 0    ' ��������
            SendMessage(GetEdit, %SCI_SETTARGETSTART, startPos - 1, 0)
          ELSE
            IF startPos = 0 THEN
              SendMessage(GetEdit, %SCI_SETTARGETSTART, startPos, 0)
            ELSE
              SendMessage(GetEdit, %SCI_SETTARGETSTART, startPos + 1, 0)
            END IF
          END IF
          ' �������λ�úͲ��ұ�ʶ
          SendMessage(GetEdit, %SCI_SETTARGETEND, endPos, 0)
          SendMessage(GetEdit, %SCI_SETSEARCHFLAGS, findFlags, 0)
          ' ����Ҫ���ҵ��ı�
          hr = SendMessage(GetEdit, %SCI_SEARCHINTARGET, LEN(szFindText), VARPTR(szFindText))
          IF CINT(hr) = -1 THEN
            MessageBox(BYVAL hFind, "δ�ҵ�   ", " ����", _
                   %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
          ELSE
            ' ������ַ�λ�ü�ѡ����ı���λ��
            SendMessage GetEdit, %SCI_SETCURRENTPOS, hr, 0
            SendMessage GetEdit, %SCI_GOTOPOS, hr, 0
            SendMessage GetEdit, %SCI_SETSELECTIONSTART, hr, 0
            SendMessage GetEdit, %SCI_SETSELECTIONEND, hr + LEN(szFindText), 0
            ' ���Ӻ���Ϊ��ʼλ��
            startPos = hr
          END IF
        ELSEIF (fr.flags AND %FR_REPLACE) = %FR_REPLACE THEN
          findFlags = 0
          ' ��ȡ�Ի�����ѡ�������
          IF (fr.Flags AND %FR_MATCHCASE) = %FR_MATCHCASE THEN findFlags = findFlags OR %SCFIND_MATCHCASE
          IF (fr.Flags AND %FR_WHOLEWORD) = %FR_WHOLEWORD THEN findFlags = findFlags OR %SCFIND_WHOLEWORD
          ' �ӵ�ǰλ�ÿ�ʼ����
          startPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)
          ' ����λ�� = �ĵ�����
          endPos = SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
          ' See if there is text selected
          x = SendMessage(GetEdit, %SCI_GETSELECTIONSTART, 0, 0)
          y = SendMessage(GetEdit, %SCI_GETSELECTIONEND, 0, 0)
          IF y > x THEN startPos = x
          ' ���ÿ�ʼλ��
          SendMessage(GetEdit, %SCI_SETTARGETSTART, startPos, 0)
          ' ���ý���λ��
          SendMessage(GetEdit, %SCI_SETTARGETEND, endPos, 0)
          ' ����������ʶ
          SendMessage(GetEdit, %SCI_SETSEARCHFLAGS, findFlags, 0)
          ' ����Ҫ�滻���ı�
          hr = SendMessage(GetEdit, %SCI_SEARCHINTARGET, LEN(szFindText), VARPTR(szFindText))
          IF CINT(hr) = -1 THEN
            MessageBox(BYVAL hFind, "δ�ҵ�   ", " �滻", _
                    %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
          ELSE
            ' ���ַ�λ�ü�ѡ����ı�
            SendMessage GetEdit, %SCI_SETCURRENTPOS, hr, 0
            SendMessage GetEdit, %SCI_GOTOPOS, hr, 0
            SendMessage GetEdit, %SCI_SETSELECTIONSTART, hr, 0
            SendMessage GetEdit, %SCI_SETSELECTIONEND, hr + LEN(szFindText), 0
            ' �滻ѡ������
            SendMessage GetEdit, %SCI_REPLACESEL, 0, VARPTR(szReplText)
            ' λ����������Ϊ��ʼλ��
            startPos = hr
          END IF
          ' Ϊ������һλ�÷�����Ϣ
          fr.Flags = fr.Flags OR %FR_FINDNEXT
          SendMessage hWnd, dwFindMsg, 0, 0
        ELSEIF (fr.flags AND %FR_REPLACEALL) = %FR_REPLACEALL THEN
          ' ==================================================================
          ' ���δѡ��飬���ʾ�ڲ�ѡ���ı���������������ĵ���ִ���滻
          ' ==================================================================
          startPos = SendMessage(GetEdit, %SCI_GETSELECTIONSTART, 0, 0)
          endPos = SendMessage(GetEdit, %SCI_GETSELECTIONEND, 0, 0)
          IF endPos > startPos THEN
            buffer = SPACE$(endPos - startPos + 1)
            SendMessage(GetEdit, %SCI_GETSELTEXT, 0, STRPTR(buffer))
            IF LEN(buffer) THEN
              IF INSTR(buffer, CHR$(13, 10)) = 0 THEN
                curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)
                SendMessage(GetEdit, %SCI_SETSELECTIONSTART, curPos, 0)
                SendMessage(GetEdit, %SCI_SETSELECTIONEND, curPos, 0)
              END IF
            END IF
          END IF
          ' ==================================================================
          ' �滻����
          findFlags = 0
          hSci = GetEdit   ' ��ȡ�༭�ؼ����
          IF (fr.Flags AND %FR_MATCHCASE) = %FR_MATCHCASE THEN findFlags = findFlags OR %SCFIND_MATCHCASE
          IF (fr.Flags AND %FR_WHOLEWORD) = %FR_WHOLEWORD THEN findFlags = findFlags OR %SCFIND_WHOLEWORD
          ' ��ͷ��ʼ����
          endPos = SendMessage(hSci, %SCI_GETTEXTLENGTH, 0, 0)
          ' ���ü���������ʼλ��
          numItems = 0
          startPos = 0
          fInSelection = 0
          ' ���startSelPos ���� endSelPos ���ʾû��ѡ������
          ' ���򣬽���ѡ�е��ı���ִ��
          startSelPos = SendMessage(GetEdit, %SCI_GETSELECTIONSTART, 0, 0)
          endSelPos = SendMessage(GetEdit, %SCI_GETSELECTIONEND, 0, 0)
          IF endSelPos > startSelPos THEN
            startPos = startSelPos
            endPos = endSelPos
            fInSelection = %TRUE
          END IF
          DO
            ' ���ò��ұ�ʶ����ʼ������λ��
            SendMessage(hSci, %SCI_SETSEARCHFLAGS, findFlags, 0)
            SendMessage(hSci, %SCI_SETTARGETSTART, startPos, 0)
            SendMessage(hSci, %SCI_SETTARGETEND, endPos, 0)
            ' ����Ҫ�滻���ı�
            hr = SendMessage(hSci, %SCI_SEARCHINTARGET, LEN(szFindText), VARPTR(szFindText))
            ' �洢λ��
            curPos = hr
            ' If hr = -1 ��û����Ҫ�滻���ı���
            IF CINT(hr) = -1 THEN
              MessageBox(BYVAL hFind, FORMAT$(numItems) & " ���滻  ", " Replace all", _
                      %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
              EXIT DO
            ELSE
              ' �滻�ı�
              SendMessage hSci, %SCI_REPLACETARGET, -1, VARPTR(szReplText)
              ' ����������
              INCR numItems
            END IF
            ' �����µĿ�ʼλ��
            startPos = curPos + LEN(szReplText)
            ' �����µĽ���λ��(�ı������п��ܻ��б仯)
            IF ISTRUE fInSelection THEN
              endPos = SendMessage(GetEdit, %SCI_GETSELECTIONEND, 0, 0)
            ELSE
              endPos = SendMessage(hSci, %SCI_GETTEXTLENGTH, 0, 0)
            END IF
            IF endPos <= startPos THEN
              MessageBox(BYVAL hFind, FORMAT$(numItems) & " ���滻  ", " Replace all", _
                      %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
              EXIT DO
            END IF
          LOOP
        END IF
      END IF
      EXIT FUNCTION
    CASE %WM_COMMAND
      SELECT CASE AS LONG LO(WORD, wParam)
        ' �ļ��˵���ز˵���
        'CASE %IDM_NEWPROJECT  '�½���Ŀ
        CASE %IDM_NEWBAS      '�½�BAS�ļ�
          g_FileType=0
          IF GetEdit THEN UPDATEWINDOW GetEdit
          hMdi = CreateMdiChild( $EDITCLASSNAME, g_hWndClient, "", IIF( fMaximize, %WS_MAXIMIZE, 0 ))
          ResetCodeFinder         ' Reset the code finde combobox content���ô������combobox������
          SHOWWINDOW hMdi, %SW_SHOW
          EXIT FUNCTION
        CASE %IDM_NEWINC      '�½�INC�ļ�
          g_FileType=1
          IF GetEdit THEN UPDATEWINDOW GetEdit
          hMdi = CreateMdiChild( $EDITCLASSNAME, g_hWndClient, "", IIF( fMaximize, %WS_MAXIMIZE, 0 ))
          ResetCodeFinder         ' Reset the code finde combobox content���ô������combobox������
          SHOWWINDOW hMdi, %SW_SHOW
          EXIT FUNCTION
        CASE %IDM_NEWRC      '�½�RC�ļ�
          g_FileType=2
          IF GetEdit THEN UPDATEWINDOW GetEdit
          hMdi = CreateMdiChild( $EDITCLASSNAME, g_hWndClient, "", IIF( fMaximize, %WS_MAXIMIZE, 0 ))
          ResetCodeFinder         ' Reset the code finde combobox content���ô������combobox������
          SHOWWINDOW hMdi, %SW_SHOW
          EXIT FUNCTION
        CASE %IDM_PBTEMPLATES+1 TO %IDM_PBTEMPLATES+PARSECOUNT(pbtempfile,$CRLF) '����ģ���ļ��½�
          LOCAL tmpArr() AS STRING
          REDIM tmpArr(PARSECOUNT(pbtempfile,$CRLF))
          PARSE pbtempfile,tmpArr(),$CRLF
          LOCAL tmpfilestr AS STRING
          LOCAL fn AS LONG
          fn=FREEFILE
          OPEN PARSE$(tmpArr(LO(WORD, wParam)-%IDM_PBTEMPLATES-1),2) FOR INPUT AS #fn
          LINE INPUT #fn,tmpStr
          LINE INPUT #fn,tmpStr
          IF UCASE$(tmpStr)=".RC" THEN
            g_FileType=2
          ELSEIF UCASE$(tmpStr)=".TXT" THEN
            g_FileType=3
          END IF
          LINE INPUT #fn,tmpStr
          WHILE ISFALSE EOF(#fn)
            LINE INPUT #fn,tmpStr
            tmpfilestr=tmpfilestr & tmpStr & $CRLF
          WEND
          CLOSE #fn
          IF GetEdit THEN UPDATEWINDOW GetEdit
          hMdi = CreateMdiChild( $EDITCLASSNAME, g_hWndClient,"", IIF( fMaximize, %WS_MAXIMIZE, 0 ))
          SHOWWINDOW hMdi, %SW_SHOW
          'ResetCodeFinder         ' Reset the code finde combobox content���ô������combobox������
          tmpLng=INSTR(tmpfilestr,"|")-1
          REPLACE "|" WITH "" IN tmpfilestr
          wi=GetWindowLong(hMdi,%GWL_USERDATA)
          SENDMESSAGE @wi.hFocus, %SCI_INSERTTEXT, 0, BYVAL STRPTR( tmpfilestr )
          sendmessage @wi.hFocus, %SCI_SETCURRENTPOS,tmpLng,0
          sendmessage @wi.hFocus, %SCI_SETCURRENTPOS,tmpLng,0
          sendmessage @wi.hFocus, %SCI_SETSELECTIONSTART,tmpLng,0
          sendmessage @wi.hFocus, %SCI_SETSELECTIONEND,tmpLng,0
          sendmessage @wi.hFocus, %SCI_SCROLLCARET,tmpLng,0
          ' Empty the undo buffer (it al sets the state of the document as unmodified)
          SENDMESSAGE @wi.hFocus, %SCI_EMPTYUNDOBUFFER, 0, 0
        CASE %IDM_OPEN        '��
          IF GetEdit THEN UpdateWindow GetEdit
          PATH  = CURDIR$
          f     = ""
          dwStyle = %OFN_FILEMUSTEXIST OR %OFN_LONGNAMES
          IF OpenFileDialog(g_hWndMain, "", f, PATH, _
                  "PB�����ļ�(*.BAS)|*.BAS|PBͷ�ļ�(*.INC)|*.inc|��Դ�ļ�(*.RC)|*.RC|�ı��ļ�(*.TXT)|*.TXT|�����ļ�(*.*)|*.*", "BAS", dwStyle) THEN
            InvalidateRect g_hToolbar, BYVAL 0, 0
            UpdateWindow g_hToolbar
            hMdi=OpenThisFile(f)
            IF (dwStyle AND %OFN_READONLY) = %OFN_READONLY THEN
              SendMessage(GetEdit, %SCI_SETREADONLY, %TRUE, 0)
            ELSE
              EnableMenuItem g_hMenuFile, %IDM_INSERTFILE, %MF_ENABLED
            END IF
            EXIT FUNCTION
          END IF
        CASE %IDM_INSERTFILE  '�����ļ�
          IF GetEdit THEN UpdateWindow GetEdit
          InsertFile hWnd
          EXIT FUNCTION
        CASE %IDM_RECENT1 TO %IDM_RECENT8      '����ļ��б� MRU �˵���(�ļ���)
          IF GetEdit THEN UpdateWindow GetEdit
          ' Get the name of the file from the registry
          IF LOWRD(wParam) - %IDM_RECENTFILES => LBOUND(RecentFiles) AND _
                     LOWRD(wParam) - %IDM_RECENTFILES <= UBOUND(RecentFiles) THEN _
                     f = RecentFiles(LOWRD(wParam) - %IDM_RECENTFILES)
          IF ISFALSE FileExist(f) THEN
            MessageBox(g_hWndMain, "�ļ� " & f & $CRLF & " �Ѳ�����   ", _
                       FUNCNAME$, %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
            GetRecentFiles
            EXIT FUNCTION
          END IF
          OpenThisFile f
          EXIT FUNCTION
        'CASE %IDM_OPENPROJECT  '����Ŀ
        '  msgbox "test"
        CASE %IDM_SAVE        '����
          IF GetEdit THEN UpdateWindow GetEdit
          SAVEFILEs hWnd, %FALSE
          ChangeButtonsState
          EXIT FUNCTION
        CASE %IDM_SAVEAS      '���Ϊ
          IF GetEdit THEN UpdateWindow GetEdit
          SAVEFILEs hWnd, %TRUE
          ChangeButtonsState
          EXIT FUNCTION
        CASE %IDM_SAVEALL
          'msgbox "ȫ����"
          'tmpEdit = GetEdit
          LOCAL fOptions  AS STRING
          LOCAL fPos      AS LONG
          LOCAL fExt      AS STRING
          LOCAL fBak      AS STRING
          LOCAL tmpStr1   AS STRING
          'local dwStyle as dword
          LOCAL ttc_item  AS TC_ITEM
          LOCAL nFile     AS LONG
          LOCAL j         AS LONG
          LOCAL hEdit     AS DWORD
          LOCAL nLen      AS DWORD
          LOCAL p         AS DWORD
          LOCAL tmpAsc    AS ASCIIZ * 99
          LOCAL curtime   AS STRING
          LOCAL fTime     AS DWORD
          'local path      as string
          'local f         as string
          fOptions = fOptions & "PB �����ļ� (*.BAS)|*.BAS|"
          fOptions = fOptions & "PB ͷ�ļ� (*.INC)|*.INC|"
          fOptions = fOptions & "PB ģ���ļ� (*.PBTPL)|*.PBTPL|"
          fOptions = fOptions & "��Դ�ļ� (*.RC)|*.RC|"
          fOptions = fOptions & "��ҳ�ļ� (*.HTML)|*.HTML|"
          fOptions = fOptions & "��ҳ�ļ� (*.HTM)|*.HTM|"
          fOptions = fOptions & "�ı��ļ� (*.TXT)|*.TXT|"
          fOptions = fOptions & "�����ļ� (*.*)|*.*"
          dwStyle = %OFN_EXPLORER OR %OFN_FILEMUSTEXIST OR %OFN_HIDEREADONLY OR _
                  %OFN_OVERWRITEPROMPT
          tmpLng=TabCtrl_GetItemCount(g_hTabMdi)
          'msgbox str$(tmpLng)
          ttc_item.mask=%TCIF_TEXT 'or %TCIF_IMAGE OR %TCIF_PARAM OR %TCIF_RTLREADING
          ttc_item.cchTextMax=99
          ttc_item.pszText=VARPTR(tmpAsc)
          curtime=DateTimeForFileName
          FOR i=0 TO tmpLng-1
            'msgbox "test" & str$(i) & str$(TabCtrl_GetItem(g_hTabMdi,i,ttc_item))
            IF ISFALSE(TabCtrl_GetItem(g_hTabMdi,i,ttc_item)) THEN
              ITERATE FOR
            END IF
            'msgbox "test1 " & tmpAsc
            tmpStr=TRIM$(tmpAsc)
            'msgbox tmpStr
            IF RIGHT$(tmpStr,2)<>" *" THEN
              ITERATE FOR
            END IF
            'tmpStr=MID$(tmpStr,1,LEN(tmpStr)-2)
            tmpStr=gTabFilePaths(i)
            'msgbox tmpStr
            PATH = GetFilePath( tmpStr )
            f = GetFileName( tmpStr )
            'msgbox "test1.1"
            IF PATH="" THEN
              IF ISFALSE( SaveFileDialog( g_hWndMain, "", f, PATH, _
                      fOptions, "BAS", dwStyle )) THEN ITERATE FOR
            ELSE
              f=tmpStr
            END IF
            'msgbox "test1.2"
            'canLog=1
            'msgbox g_zIni & $crlf & IniRead( g_zIni, "Editor options", "BackupEditorFiles", "" )
            IF ISTRUE VAL( IniRead( g_zIni, "Editor options", "BackupEditorFiles", "" )) THEN
              tmpStr=MID$(f,1,INSTR(-1,f,ANY "\/")) & curtime & "bak"
              IF DIR$(tmpStr)="" THEN
                MKDIR tmpStr
                'msgbox tmpStr
              END IF
              fBak=tmpStr & MID$(f,INSTR(-1,f,ANY "\/"))
              'msgbox fBak
              IF FileExist( f ) THEN FILECOPY f, fBak
            END IF
            'canLog=0
            'msgbox "test1"
            TRY
              nFile = FREEFILE
              OPEN f FOR BINARY AS nFile
            CATCH
              MESSAGEBOX( hWnd, "����:" & STR$( ERR ) & " �����ļ���������,������.   ", _
                      " �����ļ�", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL )
              ITERATE FOR
            END TRY
            'hEdit=GetDlgItem(MdiGetActive(g_hWndClient), %IDC_EDIT)
            'MdiChildCount
            hEdit=0
            tmpStr=gTabFilePaths(i)
            FOR j=1 TO PARSECOUNT(ghEditStr,",")
              IF TRIM$(PARSE$(ghEditStr,",",j))="" THEN
                ITERATE FOR
              END IF
              DIALOG GET TEXT GetParent(VAL(PARSE$(ghEditStr,",",j))) TO tmpStr1
              'MSGBOX tmpStr1 & $CRLF & tmpStr
              IF tmpStr1=tmpStr THEN
                hEdit=VAL(PARSE$(ghEditStr,",",j))
                EXIT FOR
              END IF
            NEXT j
            IF hEdit=0 THEN
              ITERATE FOR
            END IF
            'msgbox "hEdit=" & str$(hEdit)
            nLen = SENDMESSAGE( hEdit, %SCI_GETTEXTLENGTH, 0, 0 )'ȡ�ñ����洰�����ı�����
            Buffer = SPACE$( nLen + 1 )                 '������
            SENDMESSAGE hEdit, %SCI_GETTEXT, BYVAL LEN( Buffer ), BYVAL STRPTR( Buffer )'ȡ���ı���������
            ' Remove trailing spaces and tabs   ɾ��β���Ŀո��Ʊ��
            IF TrimTrailingBlanks THEN
              DO
                p = LEN( buffer )
                REPLACE " " & $CR WITH $CR IN buffer
                REPLACE $TAB & $CR WITH $CR IN buffer
              LOOP UNTIL p = LEN( buffer )
            END IF
            TRY
              PUT$ nFile, LEFT$( Buffer, LEN( Buffer ) - 1 )
              SETEOF nFile
            CATCH
              MESSAGEBOX( hWnd, "����" & STR$( ERR ) & " [" & ERROR$ & "] �����ļ���������,������.   ", _
                      " �����ļ�", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL )
            FINALLY
              CLOSE nFile
            END TRY
            ' ����scintilla���ĵ��ĵ�ǰ״̬Ϊδ�޸�
            SendMessage hEdit, %SCI_SETSAVEPOINT, 0, 0
            tmpAsc=MID$(f,INSTR(-1,f,ANY "\/")+1)
            tmpStr=f
            SetWindowText GetParent(hEdit),BYVAL STRPTR(f)
            TabCtrl_SetItem(g_hTabMdi,i,ttc_item)
            gTabFilePaths(i)=tmpStr
            DIALOG SET TEXT GetParent(hEdit),tmpStr
            fTime = SED_GetFileTime( tmpStr )
            SETPROP GetParent(hEdit), "FTIME", fTime
          NEXT i
        'case %IDM_SAVEPROJECT
        '  MSGBOX "������Ŀ"
        CASE %IDM_PRINTSETTING
          IF VAL(IniRead(g_zIni, "Editor options", "DdocPrinting", "")) = %BST_CHECKED THEN
           ShowPrinterSetupDialog hWnd
          ELSE
           IF g_psDlg.lStructSize = 0 THEN InitPageSetup
           IF PageSetupDlg(g_psDlg) THEN
              IniWrite g_zIni, "Printer options", "MarginLeft",   FORMAT$(g_psDlg.rtMargin.nLeft   / 1000)
              IniWrite g_zIni, "Printer options", "MarginTop",    FORMAT$(g_psDlg.rtMargin.nTop    / 1000)
              IniWrite g_zIni, "Printer options", "MarginRight",  FORMAT$(g_psDlg.rtMargin.nRight  / 1000)
              IniWrite g_zIni, "Printer options", "MarginBottom", FORMAT$(g_psDlg.rtMargin.nBottom / 1000)
           END IF
          END IF
          EXIT FUNCTION
        CASE %IDM_PRINTPREVIEW
          IF VAL(IniRead(g_zIni, "Editor options", "DdocPrinting", "")) = %BST_CHECKED THEN
            ' Get the path from the window caption
            GetWindowText MdiGetActive(g_hWndClient), szText, SIZEOF(szText)
            SED_PrintDoc(szText, %TRUE)
          ELSE
            'LOCAL dwStyle AS DWORD
            GetClassName MdiGetActive(g_hWndClient), szText, SIZEOF(szText)
            IF szText <> "PBEDIT" THEN
              GeneralErrorMsg "��������." & $CRLF & "���ȼ���һ���༭��."
              EXIT FUNCTION
            END IF
            dwStyle = %WS_VISIBLE OR %WS_OVERLAPPED OR %WS_VSCROLL OR %WS_CLIPSIBLINGS OR _
                 %WS_SYSMENU OR %WS_MINIMIZEBOX 'or %WS_THICKFRAME
            CreateWindow "NPREVIEW32", "��ӡԤ��", dwStyle, %CW_USEDEFAULT, _
                     %CW_USEDEFAULT, %CW_USEDEFAULT, %CW_USEDEFAULT, BYVAL %NULL, BYVAL %NULL, g_hInst, BYVAL %NULL
          END IF
          EXIT FUNCTION
        CASE %IDM_PRINT ', %IDT_PRINT
          UpdateWindow GetEdit
          GetClassName MdiGetActive(g_hWndClient), szText, SIZEOF(szText)
          IF szText <> "PBEDIT" THEN
            GeneralErrorMsg "��������." & $CRLF & "���ȼ���һ���༭��."
            EXIT FUNCTION
          END IF
          IF VAL(IniRead(g_zIni, "Editor options", "DdocPrinting", "")) = %BST_CHECKED THEN
            IF LOWRD(wParam) = %IDM_PRINT THEN
              ' Get the path from the window caption
              GetWindowText MdiGetActive(g_hWndClient), szText, SIZEOF(szText)
              SED_PrintDoc(szText, %FALSE)
'            ELSEIF LOWRD(wParam) = %IDT_PRINT THEN
'              GetWindowText MdiGetActive(g_hWndClient), szText, SIZEOF(szText)
'              SED_PrintDoc(szText, -1)
'              EXIT FUNCTION
            END IF
          ELSE
            LOCAL rslt AS LONG
            SetCapture g_HwndMain
            SetCursor LoadCursor(%NULL, BYVAL %IDC_WAIT)
            ' Run all printing routines in a separate thread
            THREAD CREATE PrintText(BYVAL IIF&(LOWRD(wParam) = %IDM_PRINT, %FALSE, %TRUE)) TO rslt
            THREAD CLOSE rslt TO rslt
          END IF
          EXIT FUNCTION
'        CASE %IDM_CLOSE 'FILE
'          IF GetEdit THEN UpdateWindow GetEdit
'          SendMessage(MdiGetActive(g_hWndClient), %WM_CLOSE, 0, 0)
'          ChangeButtonsState
'          ' Enable open just in case it has been disabled
'          EnableMenuItem g_hMenuFile, %IDM_OPEN, %MF_ENABLED
'          SendMessage g_hToolBar, %TB_ENABLEBUTTON, %IDM_OPEN, %TRUE
'          EXIT FUNCTION
        CASE %IDM_CLOSE       '�ر�
          'MSGBOX STR$(tabIndexClose)
          IF tabIndexClose=-1 THEN
            SendMessage(MdiGetActive(g_hWndClient), %WM_CLOSE, 0, 0)
            ChangeButtonsState
            EnableMenuItem g_hMenuFile, %IDM_OPEN, %MF_ENABLED
            SendMessage g_hToolBar, %TB_ENABLEBUTTON, %IDM_OPEN, %TRUE
          ELSE
            nTab = SENDMESSAGE( g_hTabMdi, %TCM_GETCURSEL, 0, 0 )
            IF nTab=tabIndexClose THEN
              nTab=-1
            END IF
            IF UBOUND( gTabFilePaths ) = > tabIndexClose THEN
              IF nTab>tabIndexClose THEN
                nTab-=1
              END IF
              strTabTxt = gTabFilePaths( tabIndexClose )
              IF RIGHT$(strTabTxt,1)="*" THEN
                strTabTxt=MID$(strTabTxt,1,LEN(strTabTxt)-2)
              END IF
              SED_ActivateMdiWindow strTabTxt
              SendMessage(MdiGetActive(g_hWndClient), %WM_CLOSE, 0, 0)
              IF FileIsInProject( strTabTxt ) THEN
                TreeView_FindItem( ghProjectTV, GetFileName( strTabTxt ), %FALSE, %NULL, %TRUE, %TRUE )
              END IF
            END IF
            IF nTab=-1 THEN
              nTab=0
            END IF
            IF UBOUND(gTabFilePaths)=>nTab THEN
              strTabTxt = gTabFilePaths( nTab )
              IF RIGHT$(strTabTxt,1)="*" THEN
                strTabTxt=MID$(strTabTxt,1,LEN(strTabTxt)-2)
              END IF
              SED_ActivateMdiWindow strTabTxt
            END IF
            tabIndexClose=-1
          END IF
        CASE %IDM_CLOSEALL    '�ر�����
           WHILE MdiGetActive(g_hWndClient)
             CALL SendMessage(MdiGetActive(g_hWndClient), %WM_CLOSE, 0, 0)
             IF fClosed = 0 THEN EXIT FUNCTION
           WEND
        CASE %IDM_OPENCMD     '�������� %IDM_DOS
          PID = SHELL(ENVIRON$("COMSPEC"))
          EXIT FUNCTION
        CASE %IDM_EXIT   : SendMessage hWnd, %WM_CLOSE, wParam, lParam

        ' �༭�˵��Ĳ˵���
        CASE %IDM_UNDO        '����
          SendMessage GetEdit, %SCI_UNDO, 0, 0
          ChangeButtonsState
          EXIT FUNCTION
        CASE %IDM_REDO        '����
          SendMessage GetEdit, %SCI_UNDO, 0, 0
          ChangeButtonsState
          EXIT FUNCTION

        CASE %IDM_CUT         '����
          SendMessage GetEdit, %SCI_CUT, 0, 0
          EXIT FUNCTION
        CASE %IDM_COPY        '����
          SendMessage GetEdit, %SCI_COPY, 0, 0
          EXIT FUNCTION
        CASE %IDM_PASTE       'ճ��
          SendMessage GetEdit, %SCI_PASTE, 0, 0
          EXIT FUNCTION
        CASE %IDM_PASTEIE     'ճ����ҳ����
          ' Paste Internet Explorer Html
          PasteIE
          EXIT FUNCTION
        CASE %IDM_SELALL      'ȫѡ
          SendMessage GetEdit, %SCI_SELECTALL, 0, 0
          EXIT FUNCTION
        CASE %IDM_DELETE      'ɾ��
          SendMessage GetEdit, %SCI_CLEAR, 0, 0
          EXIT FUNCTION
        CASE %IDM_LINEDELETE  'ɾ����
          SendMessage GetEdit, %SCI_LINEDELETE,0,0
          EXIT FUNCTION
        CASE %IDM_INITASNEW   '��ʼ��
          wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
          IF wi THEN
            SendMessage @wi.hFocus, %SCI_SETDOCPOINTER, %NULL, BYVAL %NULL
            ' Flag this control as 'Initialized'
            'BIT SET @wi.bInit, GetDlgCtrlID(@wi.hFocus) - %IDC_EDIT2
            ' -- Jos� Roca: Added checking for safety
            bitnum = GetDlgCtrlID(@wi.hFocus) - %IDC_EDIT2
            IF bitnum => 0 AND bitnum < 16 THEN
              BIT SET @wi.bInit, bitnum
            END IF

            ' Force new document to have the standard options
            pSciData = SendMessage(@wi.hFocus, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
            IF pSciData THEN Scintilla_SetOptions pSciData, "Untitled.bas"
            SendMessage @wi.hFocus, %SCI_EMPTYUNDOBUFFER, %NULL, %NULL

            ' Get 'Save' and 'Open'menu items disabled
            EnableMenuItem g_hMenuFile, %IDM_SAVE, %MF_GRAYED
            EnableMenuItem g_hMenuFile, %IDM_OPEN, %MF_GRAYED
            SendMessage ghToolBar, %TB_ENABLEBUTTON, %IDM_OPEN, %FALSE
            SendMessage ghToolBar, %TB_ENABLEBUTTON, %IDM_SAVE, %FALSE
          END IF
          EXIT FUNCTION
        CASE %IDM_CLEARALL    'ȫɾ(���)
          SendMessage GetEdit,%SCI_CLEARALL,0,0
          EXIT FUNCTION
        CASE %IDM_COMMENT, %IDK_COMMENT       'ע��
          BlockComment
          EXIT FUNCTION

        CASE %IDM_UNCOMMENT, %IDK_UNCOMMENT   '��ע��
          BlockUncomment
          EXIT FUNCTION
        CASE %IDM_INDENT      '����
          SendMessage GetEdit, %SCI_TAB, 0, 0
          EXIT FUNCTION

        CASE %IDM_OUTDENT     '������
          SendMessage GetEdit, %SCI_BACKTAB, 0, 0
          EXIT FUNCTION
        CASE %IDM_FORMATREGION  '��ʽ��ѡ��
          SED_FormatRegion
          EXIT FUNCTION
        CASE %IDM_TABULATEREGION '���ѡ��
          SED_TabulateRegion
          EXIT FUNCTION
        CASE %IDM_SELTOUPPERCASE  'ѡ��ת��д
          ChangeSelectedTextCase 1
          EXIT FUNCTION

        ' Convert selected text to lower case
        CASE %IDM_SELTOLOWERCASE  'ѡ��תСд
          ChangeSelectedTextCase 2
          EXIT FUNCTION

        ' Convert selected text to mixed case
        CASE %IDM_SELTOMIXEDCASE  'ѡ��ת�շ�ģʽ
          ChangeSelectedTextCase 3
          EXIT FUNCTION
        CASE %IDM_TEMPLATES       'ģ�����
          ' Activate the Template code popup dialog
          IF GetEdit THEN TemplateCodePopupDialog hWnd
        CASE %IDM_HTMLCODE        'תΪ��ҳ
          IF GetEdit THEN UpdateWindow GetEdit
          CvBasToHtml
          EXIT FUNCTION
        CASE %IDM_GUID        '����GUID

        ' ��ͼ�˵��Ĳ˵���
        CASE %IDM_PROPERTY          '���Դ�����ʾ/����
          On_CommandMenuDockWin %ID_PROPERTY,LO(WORD,wParam)  '���Ʋ��봰����ʾ/���ز˵�������Ӧ
        CASE %IDM_CONTROLS          '�ؼ���������ʾ/����
          On_CommandMenuDockWin %ID_CONTROLS,LO(WORD,wParam)  '���Ʋ��봰����ʾ/���ز˵�������Ӧ
        CASE %IDM_PROJECT           '��Ŀ��������ʾ/����
          On_CommandMenuDockWin %ID_PROJECT,LO(WORD,wParam)   '���Ʋ��봰����ʾ/���ز˵�������Ӧ
        CASE %IDM_COMPILERS         '������������ʾ/����
          On_CommandMenuDockWin %ID_COMPILERS,LO(WORD,wParam) '���Ʋ��봰����ʾ/���ز˵�������Ӧ
        CASE %IDM_FINDRS            '���ҽ��������ʾ/����
          On_CommandMenuDockWin %ID_FINDRS,LO(WORD,wParam)    '���Ʋ��봰����ʾ/���ز˵�������Ӧ
        CASE %IDM_STDTOOLBAR        '������������ʾ/����
          MENU GET STATE  ghRebarMenu,1 TO MenuState
          IF ghToolBar<=0 THEN
            ghToolBar=CreateToolbar(hWnd)
          END IF
          IF MenuState =%MF_CHECKED THEN
            IF ISFALSE(ShowBand(ghToolbar,%ID_TOOLBAR,0)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,1,%MF_ENABLED OR %MF_UNCHECKED
            END IF
          ELSE
            IF ISFALSE(ShowBand(ghToolbar,%ID_TOOLBAR,1)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,1,%MF_ENABLED OR %MF_CHECKED
            END IF
          END IF
          SendMessage hWnd,%WM_SIZE,wParam, lParam
        CASE %IDM_WINDOWBAR       '���ڱ༭��������ʾ/����
          MENU GET STATE  ghRebarMenu,2 TO MenuState
          IF ghWindowBar<=0 THEN
            ghWindowBar=CreateWindowbar(hWnd)
          END IF
          IF MenuState =%MF_CHECKED THEN
            'ShowBand ghWindowbar,0,""
            IF ISFALSE(ShowBand(ghWindowBar,%ID_WINDOWBAR,0)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,2,%MF_ENABLED OR %MF_UNCHECKED
            END IF
          ELSE
            'ShowBand ghWindowbar,1,""
            IF ISFALSE(ShowBand(ghWindowBar,%ID_WINDOWBAR,1)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,2,%MF_ENABLED OR %MF_CHECKED
            END IF
          END IF
          SendMessage hWnd,%WM_SIZE,wParam, lParam
        CASE %IDM_COMBOBAR        '���̹�������ʾ/����
          MENU GET STATE  ghRebarMenu,3 TO MenuState
          IF ghComboBox<=0 THEN
            ghComboBox=CreateComboBox(hWnd)
          END IF
          IF MenuState =%MF_CHECKED THEN
            'ShowBand ghCombobox,0,"����"
            IF ISFALSE(ShowBand(ghCombobox,%ID_COMBOBOX,0)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,3,%MF_ENABLED OR %MF_UNCHECKED
            END IF
          ELSE
            'ShowBand ghCombobox,1,"����"
            IF ISFALSE(ShowBand(ghCombobox,%ID_COMBOBOX,1)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,3,%MF_ENABLED OR %MF_CHECKED
            END IF
          END IF
          SendMessage hWnd,%WM_SIZE,wParam, lParam
        CASE %IDM_DEBUGBAR        '���룬���ԣ����й�������ʾ/����
          IF ghDebugBar<=0 THEN
            ghDebugBar=CreateDebugBar(hWnd)
          END IF
          MENU GET STATE  ghRebarMenu,4 TO MenuState
          IF MenuState =%MF_CHECKED THEN
            'ShowBand ghDebugBar,0,""
            IF ISFALSE(ShowBand(ghDebugBar,%ID_DEBUGBAR,0)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,4,%MF_ENABLED OR %MF_UNCHECKED
            END IF
          ELSE
            'ShowBand ghDebugBar,1,""
            IF ISFALSE(ShowBand(ghDebugbar,%ID_DEBUGBAR,1)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,4,%MF_ENABLED OR %MF_CHECKED
            END IF
          END IF
          SendMessage hWnd,%WM_SIZE,wParam, lParam
        CASE %IDM_CODEEDIT        '����༭��������ʾ/����
          IF ghCodeEditBar<=0 THEN
            ghCodeEditBar=CreateCodeEditBar(hWnd)
          END IF
          MENU GET STATE  ghRebarMenu,5 TO MenuState
          IF MenuState =%MF_CHECKED THEN
            'ShowBand ghDebugBar,0,""
            IF ISFALSE(ShowBand(ghCodeEditBar,%ID_CODEEDITBAR,0)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,5,%MF_ENABLED OR %MF_UNCHECKED
            END IF
          ELSE
            'ShowBand ghDebugBar,1,""
            IF ISFALSE(ShowBand(ghCodeEditBar,%ID_CODEEDITBAR,1)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,5,%MF_ENABLED OR %MF_CHECKED
            END IF
          END IF
          SendMessage hWnd,%WM_SIZE,wParam, lParam
        CASE %IDM_WINARRAG        '�������й�������ʾ/����
          IF ghWinArragBar<=0 THEN
            ghWinArragBar=CreateWinArragBar(hWnd)
          END IF
          MENU GET STATE  ghRebarMenu,6 TO MenuState
          IF MenuState =%MF_CHECKED THEN
            'ShowBand ghDebugBar,0,""
            IF ISFALSE(ShowBand(ghWinArragBar,%ID_WINARRAGBAR,0)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,6,%MF_ENABLED OR %MF_UNCHECKED
            END IF
          ELSE
            'ShowBand ghDebugBar,1,""
            IF ISFALSE(ShowBand(ghWinArragBar,%ID_WINARRAGBAR,1)) THEN
              MSGBOX "����ʧ�ܣ��������"
            ELSE
              MENU SET STATE ghRebarMenu,6,%MF_ENABLED OR %MF_CHECKED
            END IF
          END IF
          SendMessage hWnd,%WM_SIZE,wParam, lParam
        CASE %ID_LOCKBAR         '����/����������
          MENU GET STATE  ghRebarMenu,8 TO MenuState
          IF MenuState =%MF_CHECKED THEN
            IF LockBar(0) THEN
              MENU SET STATE ghRebarMenu,8,%MF_ENABLED OR %MF_UNCHECKED
            END IF
          ELSE
            IF LockBar(1) THEN
              MENU SET STATE ghRebarMenu,8,%MF_ENABLED OR %MF_CHECKED
            END IF
          END IF
          SendMessage hWnd,%WM_SIZE,wParam, lParam
        CASE %IDM_TOGGLE
          ToggleFolding GetCurrentLine
          EXIT FUNCTION
        CASE %IDM_TOGGLEALL
          ToggleAllFoldersBelow GetCurrentLine
          EXIT FUNCTION
        CASE %IDM_FOLDALL
          FoldAllProcedures
          EXIT FUNCTION
        CASE %IDM_EXPANDALL
          ExpandAllProcedures
          EXIT FUNCTION
        CASE %IDM_ZOOMIN
          SendMessage GetEdit, %SCI_ZOOMIN, 0, 0
          EXIT FUNCTION
        CASE %IDM_ZOOMOUT
          SendMessage GetEdit, %SCI_ZOOMOUT, 0, 0
          EXIT FUNCTION
        CASE %IDM_USETABS
          ' Toggle the Use Tabs option
          hr = SendMessage(GetEdit, %SCI_GETUSETABS, 0, 0)
          IF hr THEN
            CheckMenuItem hMenu, %IDM_USETABS, %MF_UNCHECKED
            IniWrite g_zIni, "Editor options", "UseTabs", FORMAT$(%BST_UNCHECKED)
            fFlag = %FALSE
          ELSE
            CheckMenuItem hMenu, %IDM_USETABS, %MF_CHECKED
            SendMessage GetEdit, %SCI_SETUSETABS, %TRUE, 0
            IniWrite g_zIni, "Editor options", "UseTabs", FORMAT$(%BST_CHECKED)
            fFlag = %TRUE
          END IF
          ' Change the option in the four windows
          wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
          IF wi THEN
            IF @wi.hTopL THEN
              pSciData = SendMessage(@wi.hTopL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETUSETABS, fFlag, 0
              END IF
            END IF
            IF @wi.hTopR THEN
              pSciData = SendMessage(@wi.hTopR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETUSETABS, fFlag, 0
              END IF
            END IF
            IF @wi.hBotL THEN
              pSciData = SendMessage(@wi.hBotL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETUSETABS, fFlag, 0
              END IF
            END IF
            IF @wi.hBotR THEN
              pSciData = SendMessage(@wi.hBotR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETUSETABS, fFlag, 0
              END IF
            END IF
          ELSE
            SendMessage GetEdit, %SCI_SETUSETABS, fFlag, 0
          END IF
          EXIT FUNCTION

        CASE %IDM_AUTOINDENT
          ' Toggle the Autoindent option
          IF VAL(IniRead(g_zIni, "Editor options", "AutoIndent", "")) = %BST_CHECKED THEN
            CheckMenuItem hMenu, %IDM_AUTOINDENT, %MF_UNCHECKED
            IniWrite g_zIni, "Editor options", "AutoIndent", FORMAT$(%BST_UNCHECKED)
          ELSE
            CheckMenuItem hMenu, %IDM_AUTOINDENT, %MF_CHECKED
            IniWrite g_zIni, "Editor options", "AutoIndent", FORMAT$(%BST_CHECKED)
          END IF
          EXIT FUNCTION

        CASE %IDM_SHOWLINENUM
          ' Toggle the Line Numbers option
          nWidth = 0
          hr = SendMessage(GetEdit, %SCI_GETMARGINWIDTHN, 0, 0)
          IF hr THEN
            CheckMenuItem hMenu, %IDM_SHOWLINENUM, %MF_UNCHECKED
            IniWrite g_zIni, "Editor options", "LineNumbers", FORMAT$(%BST_UNCHECKED)
          ELSE
            CheckMenuItem hMenu, %IDM_SHOWLINENUM, %MF_CHECKED
            nWidth = VAL(IniRead(g_zIni, "Editor options", "LineNumbersWidth", ""))
            IF nWidth = 0 THEN nWidth = 50
            IniWrite g_zIni, "Editor options", "LineNumbers", FORMAT$(%BST_CHECKED)
          END IF
          ' Change the option in the four windows
          wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
          IF wi THEN
            IF @wi.hTopL THEN
              pSciData = SendMessage(@wi.hTopL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETMARGINTYPEN, 0, %SC_MARGIN_NUMBER
                SciMsg pSciData, %SCI_SETMARGINWIDTHN, 0, nWidth
              END IF
            END IF
            IF @wi.hTopR THEN
              pSciData = SendMessage(@wi.hTopR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETMARGINTYPEN, 0, %SC_MARGIN_NUMBER
                SciMsg pSciData, %SCI_SETMARGINWIDTHN, 0, nWidth
              END IF
            END IF
            IF @wi.hBotL THEN
              pSciData = SendMessage(@wi.hBotL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETMARGINTYPEN, 0, %SC_MARGIN_NUMBER
                SciMsg pSciData, %SCI_SETMARGINWIDTHN, 0, nWidth
              END IF
            END IF
            IF @wi.hBotR THEN
              pSciData = SendMessage(@wi.hBotR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETMARGINTYPEN, 0, %SC_MARGIN_NUMBER
                SciMsg pSciData, %SCI_SETMARGINWIDTHN, 0, nWidth
              END IF
            END IF
          ELSE
            SendMessage GetEdit, %SCI_SETMARGINTYPEN, 0, %SC_MARGIN_NUMBER
            SendMessage GetEdit, %SCI_SETMARGINWIDTHN, 0, nWidth
          END IF
          EXIT FUNCTION

        CASE %IDM_SHOWMARGIN
          ' Toggle the Margin option
          nWidth = 0
          hr = SendMessage(GetEdit, %SCI_GETMARGINWIDTHN, 2, 0)
          IF hr THEN
            CheckMenuItem hMenu, %IDM_SHOWMARGIN, %MF_UNCHECKED
            IniWrite g_zIni, "Editor options", "Margin", FORMAT$(%BST_UNCHECKED)
          ELSE
            CheckMenuItem hMenu, %IDM_SHOWMARGIN, %MF_CHECKED
            nWidth = VAL(IniRead(g_zIni, "Editor options", "MarginWidth", ""))
            IF nWidth=0 THEN nWidth=20
            IniWrite g_zIni, "Editor options", "Margin", FORMAT$(%BST_CHECKED)
          END IF
          ' Change the option in the four windows
          wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
          IF wi THEN
            IF @wi.hTopL THEN
              pSciData = SendMessage(@wi.hTopL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETMARGINTYPEN, 2, %SC_MARGIN_SYMBOL
                SciMsg pSciData, %SCI_SETMARGINWIDTHN, 2, nWidth
              END IF
            END IF
            IF @wi.hTopR THEN
              pSciData = SendMessage(@wi.hTopR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETMARGINTYPEN, 2, %SC_MARGIN_SYMBOL
                SciMsg pSciData, %SCI_SETMARGINWIDTHN, 2, nWidth
              END IF
            END IF
            IF @wi.hBotL THEN
              pSciData = SendMessage(@wi.hBotL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETMARGINTYPEN, 2, %SC_MARGIN_SYMBOL
                SciMsg pSciData, %SCI_SETMARGINWIDTHN, 2, nWidth
              END IF
            END IF
            IF @wi.hBotR THEN
              pSciData = SendMessage(@wi.hBotR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETMARGINTYPEN, 2, %SC_MARGIN_SYMBOL
                SciMsg pSciData, %SCI_SETMARGINWIDTHN, 2, nWidth
              END IF
            END IF
          ELSE
            SendMessage GetEdit, %SCI_SETMARGINTYPEN, 2, %SC_MARGIN_SYMBOL
            SendMessage GetEdit, %SCI_SETMARGINWIDTHN, 2, nWidth
          END IF
          EXIT FUNCTION

        CASE %IDM_SHOWINDENT
          ' Toggle the Use Tabs option
          hr = SendMessage(GetEdit, %SCI_GETINDENTATIONGUIDES, 0, 0)
          IF hr THEN
            fFlag = %FALSE
            CheckMenuItem hMenu, %IDM_SHOWINDENT, %MF_UNCHECKED
            IniWrite g_zIni, "Editor options", "IndentGuides", FORMAT$(%BST_UNCHECKED)
          ELSE
            fFlag = %TRUE
            CheckMenuItem hMenu, %IDM_SHOWINDENT, %MF_CHECKED
            IniWrite g_zIni, "Editor options", "IndentGuides", FORMAT$(%BST_CHECKED)
          END IF
          ' Change the option in the four windows
          wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
          IF wi THEN
            IF @wi.hTopL THEN
              pSciData = SendMessage(@wi.hTopL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETINDENTATIONGUIDES, fFlag, 0
              END IF
            END IF
            IF @wi.hTopR THEN
              pSciData = SendMessage(@wi.hTopR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETINDENTATIONGUIDES, fFlag, 0
              END IF
            END IF
            IF @wi.hBotL THEN
              pSciData = SendMessage(@wi.hBotL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETINDENTATIONGUIDES, fFlag, 0
              END IF
            END IF
            IF @wi.hBotR THEN
              pSciData = SendMessage(@wi.hBotR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
                SciMsg pSciData, %SCI_SETINDENTATIONGUIDES, fFlag, 0
              END IF
            END IF
          ELSE
            SendMessage GetEdit, %SCI_SETINDENTATIONGUIDES, fFlag, 0
          END IF
          EXIT FUNCTION

        CASE %IDM_SHOWEOL
          ' Toggle the End of Line option
          hr = SendMessage(GetEdit, %SCI_GETVIEWEOL, 0, 0)
          IF hr THEN
           fFlag = %FALSE
           CheckMenuItem hMenu, %IDM_SHOWEOL, %MF_UNCHECKED
           IniWrite g_zIni, "Editor options", "EndOfLine", FORMAT$(%BST_UNCHECKED)
          ELSE
           fFlag = %TRUE
           CheckMenuItem hMenu, %IDM_SHOWEOL, %MF_CHECKED
           IniWrite g_zIni, "Editor options", "EndOfLine", FORMAT$(%BST_CHECKED)
          END IF
          ' Change the option in the four windows
          wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
          IF wi THEN
           IF @wi.hTopL THEN
              pSciData = SendMessage(@wi.hTopL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
               SciMsg pSciData, %SCI_SETVIEWEOL, fFlag, 0
              END IF
           END IF
           IF @wi.hTopR THEN
              pSciData = SendMessage(@wi.hTopR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
               SciMsg pSciData, %SCI_SETVIEWEOL, fFlag, 0
              END IF
           END IF
           IF @wi.hBotL THEN
              pSciData = SendMessage(@wi.hBotL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
               SciMsg pSciData, %SCI_SETVIEWEOL, fFlag, 0
              END IF
           END IF
           IF @wi.hBotR THEN
              pSciData = SendMessage(@wi.hBotR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
              IF pSciData THEN
               SciMsg pSciData, %SCI_SETVIEWEOL, fFlag, 0
              END IF
           END IF
          ELSE
           SendMessage GetEdit, %SCI_SETVIEWEOL, fFlag, 0
          END IF
          EXIT FUNCTION

         CASE %IDM_SHOWSPACES
           ' Toggle the WhiteSpace option
           hr = SendMessage(GetEdit, %SCI_GETVIEWWS, 0, 0)
           IF hr THEN
             fFlag = %SCWS_INVISIBLE
             CheckMenuItem hMenu, %IDM_SHOWSPACES, %MF_UNCHECKED
             IniWrite g_zIni, "Editor options", "WhiteSpace", FORMAT$(%BST_UNCHECKED)
           ELSE
             fFlag = %SCWS_VISIBLEALWAYS
             CheckMenuItem hMenu, %IDM_SHOWSPACES, %MF_CHECKED
             IniWrite g_zIni, "Editor options", "WhiteSpace", FORMAT$(%BST_CHECKED)
           END IF
           ' Change the option in the four windows
           wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
           IF wi THEN
             IF @wi.hTopL THEN
               pSciData = SendMessage(@wi.hTopL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
               IF pSciData THEN
                 SciMsg pSciData, %SCI_SETVIEWWS, fFlag, 0
               END IF
             END IF
             IF @wi.hTopR THEN
               pSciData = SendMessage(@wi.hTopR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
               IF pSciData THEN
                 SciMsg pSciData, %SCI_SETVIEWWS, fFlag, 0
               END IF
             END IF
             IF @wi.hBotL THEN
               pSciData = SendMessage(@wi.hBotL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
               IF pSciData THEN
                 SciMsg pSciData, %SCI_SETVIEWWS, fFlag, 0
               END IF
             END IF
             IF @wi.hBotR THEN
               pSciData = SendMessage(@wi.hBotR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
               IF pSciData THEN
                 SciMsg pSciData, %SCI_SETVIEWWS, fFlag, 0
               END IF
             END IF
           ELSE
             SendMessage GetEdit, %SCI_SETVIEWWS, fFlag, 0
           END IF
           EXIT FUNCTION

         CASE %IDM_SHOWEDGE
           ' Toggle the Edge option
           hr = SendMessage(GetEdit, %SCI_GETEDGEMODE, 0, 0)
           IF hr THEN
             fFlag = %EDGE_NONE
             CheckMenuItem hMenu, %IDM_SHOWEDGE, %MF_UNCHECKED
             IniWrite g_zIni, "Editor options", "EdgeColumn", FORMAT$(%BST_UNCHECKED)
           ELSE
             fFlag = %EDGE_LINE
             CheckMenuItem hMenu, %IDM_SHOWEDGE, %MF_CHECKED
             IniWrite g_zIni, "Editor options", "EdgeColumn", FORMAT$(%BST_CHECKED)
           END IF
           ' Change the option in the four windows
           wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
           IF wi THEN
             IF @wi.hTopL THEN
               pSciData = SendMessage(@wi.hTopL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
               IF pSciData THEN
                 SciMsg pSciData, %SCI_SETEDGEMODE, fFlag, 0
               END IF
             END IF
             IF @wi.hTopR THEN
               pSciData = SendMessage(@wi.hTopR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
               IF pSciData THEN
                 SciMsg pSciData, %SCI_SETEDGEMODE, fFlag, 0
               END IF
             END IF
             IF @wi.hBotL THEN
               pSciData = SendMessage(@wi.hBotL, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
               IF pSciData THEN
                 SciMsg pSciData, %SCI_SETEDGEMODE, fFlag, 0
               END IF
             END IF
             IF @wi.hBotR THEN
               pSciData = SendMessage(@wi.hBotR, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
               IF pSciData THEN
                 SciMsg pSciData, %SCI_SETEDGEMODE, fFlag, 0
               END IF
             END IF
           ELSE
             SendMessage GetEdit, %SCI_SETEDGEMODE, fFlag, 0
           END IF
           EXIT FUNCTION

         CASE %IDM_SHOWPROCNAME
           ' Toggle the Show Procedure Name option
           IF VAL(IniRead(g_zIni, "Editor options", "ShowProcedureName", "")) = %BST_CHECKED THEN
             CheckMenuItem hMenu, %IDM_SHOWPROCNAME, %MF_UNCHECKED
             ShowProcedureName = %FALSE
             IniWrite g_zIni, "Editor options", "ShowProcedureName", FORMAT$(%BST_UNCHECKED)
       '            SED_CodeFinder
           ELSE
             CheckMenuItem hMenu, %IDM_SHOWPROCNAME, %MF_CHECKED
             ShowProcedureName = %TRUE
             IniWrite g_zIni, "Editor options", "ShowProcedureName", FORMAT$(%BST_CHECKED)
       '            SED_CodeFinder
       '            SED_WithinProc(tPROC)
       '            buffer = tPROC.ProcName
       '            SendMessage hCodeFinder, %CB_SELECTSTRING , -1, BYVAL STRPTR(buffer)
             SED_WithinProc(tPROC)
             szText = tPROC.ProcName
             SendMessage g_hStatus, %SB_SETTEXT, 4, VARPTR(szText)
           END IF
           EXIT FUNCTION
        CASE %IDM_CVEOLTOCRLF
          SendMessage GetEdit, %SCI_CONVERTEOLS, %SC_EOL_CRLF, 0
          EXIT FUNCTION

        CASE %IDM_CVEOLTOCR
          SendMessage GetEdit, %SCI_CONVERTEOLS, %SC_EOL_CR, 0
          EXIT FUNCTION

        CASE %IDM_CVEOLTOLF
          SendMessage GetEdit, %SCI_CONVERTEOLS, %SC_EOL_LF, 0
          EXIT FUNCTION
        CASE %IDM_REPLSPCWITHTABS
          ReplaceSpacesWithTabs
          EXIT FUNCTION

        CASE %IDM_REPLTABSWITHSPC
          ReplaceTabsWithSpaces
          EXIT FUNCTION
        CASE %IDM_FILEPROPERTIES
          GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
          ShowFileProperties hWnd, szPath
          EXIT FUNCTION

        CASE %IDM_SYSINFO
          SED_ShowSysInfo hWnd
          EXIT FUNCTION

        CASE %IDM_NEWPROJECT
          IF g_hGuiderWin>0 AND IsWindow(g_hGuiderWin) THEN
            ShowWindow g_hGuiderWin,%SW_SHOW
          ELSE
            g_hGuiderWin=CreateGuiderWindow(g_hWndMain)
          END IF
        CASE %IDM_WIN
          hWndEdit = newEditWindow( hWnd )
          EXIT FUNCTION
        CASE %IDM_WINTEMPLATES + 1      '����һ�㴰�ڲ˵�
          hWndEdit = newEditWindow( hWnd )
          EXIT FUNCTION
        CASE %IDM_WINTEMPLATES + 2      '����MDI���ڲ˵�
          'msgbox "bb"
          hWndMDIEdit = NewEditMDIWindow( hWnd )
          EXIT FUNCTION







        CASE %IDM_OPTION   'ѡ��˵�����ʾѡ���
          ShowOptionWin hWnd
        CASE %IDM_CASCADE : MdiCascade (g_hWndClient)
        CASE %IDM_TILEH   : MdiTile (g_hWndClient, %MDITILE_VERTICAL )
        CASE %IDM_TILEV   : MdiTile (g_hWndClient, %MDITILE_HORIZONTAL)
        CASE %IDM_ARRANGE : MdiIconArrange (g_hWndClient)


        CASE %IDHELP
            WinHelp hWnd, $HELPFILE, %HELP_INDEX, 0&  '<- to show a help file
        CASE %IDM_ABOUT  ' ��ʾ�����ʾWindows�ڽ��Ĺ��ڶԻ���
          ShellAbout hWnd, "Visual PowerBASIC 1.0#", _
                     "PowerBASIC IDE - MDI text editor, Visualize form editor, Project manager", _
                     LoadIcon(g_hInst, "APPICON")
        CASE %IDM_GOTOLINE
          IF ISTRUE GetEdit THEN ShowGotoLinePopupDialog hWnd
          EXIT FUNCTION
        CASE %IDM_GOTOBEGINPROC
          curPos = GotoBeginProc
          EXIT FUNCTION

        CASE %IDM_GOTOENDPROC
          curPos = GotoEndProc
          EXIT FUNCTION

        CASE %IDM_GOTOBEGINTHISPROC
          curPos = GotoBeginThisProc
          EXIT FUNCTION

        CASE %IDM_GOTOENDTHISPROC
          curPos = GotoEndThisProc
          EXIT FUNCTION
        CASE %ID_SETV_TOP
          MSGBOX "top"
          '���ȶ�ѡ�пؼ������϶��룬Ȼ���滻��ǰ��������ťΪ�϶��밴ť
        CASE %ID_SETV_CENTER
          MSGBOX "v-center"
        CASE %ID_SETV_BOTTOM
          MSGBOX "bottom"
        CASE %ID_SETH_LEFT
          MSGBOX "left"
        CASE %ID_SETH_CENTER
          MSGBOX "h-center"
        CASE %ID_SETH_RIGHT
          MSGBOX "right"
      END SELECT
    CASE %WM_CLOSE
      'GetLang("Are you sure to Exit Program?") GetLang("prompt")
      IF MsgQueBox("��ȷ��Ҫ�˳���")=%IDNO THEN
        EXIT FUNCTION
      END IF
      SaveLoadedFilesPaths
      WHILE MdiGetActive(g_hWndClient)
        CALL SendMessage(MdiGetActive(g_hWndClient), %WM_CLOSE, 0, 0)
        IF fClosed = 0 THEN
          EXIT FUNCTION
        END IF
      WEND
      CALL SendMessage( hWnd, %WM_DESTROY, wParam, lParam )
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_DESTROY
'      SaveDockInfo "dock.cfg"
'      SaveSizeInfo "size.cfg"
      ShowWindow hWnd,%SW_HIDE
      SaveMainRc
      SaveMainClientRc
      SaveDockInfo
      IniWrite g_zIni,"Editor options", "LastFolder",lastfolder
      IniWrite g_zIni,"startshow","flashwindow","1"
      IniWrite g_zIni,"startshow","guiderwindow","0"
      WriteRecentFiles ""
      IF hCodetipsFile THEN trm_Close(hCodetipsFile)            ' Close the Tsunami codetips database file
      IF hCodeKeepFile THEN trm_Close(hCodeKeepFile)
      'WriteConfig "lastfolder",lastfolder
'      WriteConfig "clientrect", FORMAT$(OldRC.nLeft) & "," & FORMAT$(OldRC.nTop) & "," _
'                  & FORMAT$(OldRC.nRight) & "," & FORMAT$(OldRC.nBottom)
'      GetWindowRect g_hWndMain, tmpRC
'      WriteConfig "mainwindowlastpos",FORMAT$(tmpRC.nLeft) & "," & FORMAT$(tmpRC.nTop)
'      WriteConfig "mainwindowlastsize",FORMAT$(tmpRC.nRight-tmpRC.nLeft) & "," & FORMAT$(tmpRC.nBottom-tmpRC.nTop)
      SaveRebarBand
      DragAcceptFiles hWnd, %FALSE
      'PostMessage hWnd,%WM_QUIT,0,0
      CALL PostQuitMessage(0) 'byval %NULL '0
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_HELP
      WinHelp hWnd, $HELPFILE, %HELP_INDEX, 0&  '��ʾһ�������ļ�
    CASE %WM_SYSCOLORCHANGE
      ' ����û��ı���ϵͳ��ɫ���õĻ�������������Ϣ��ͨ�ÿؼ�������
      ' ���Ǿ�����ȷ�ظ����ˡ�
      SendMessage g_hToolbar,  %WM_SYSCOLORCHANGE, wParam, lParam
      SendMessage g_hStatus,  %WM_SYSCOLORCHANGE, wParam, lParam
  END SELECT
  FUNCTION = DefFrameProc(hWnd, g_hWndClient, wMsg, wParam, lParam)
END FUNCTION


'''==============================================================================
'FUNCTION GetEdit() AS LONG
''------------------------------------------------------------------------------
'  ' �õ� MDI �е�ǰ������ӱ༭�ؼ�������еĻ�
'  '----------------------------------------------------------------------------
'  FUNCTION = GetDlgItem(MdiGetActive(g_hWndClient), %IDC_EDIT)
'END FUNCTION

SUB GetRecentFiles
  LOCAL Bc AS LONG
  LOCAL fileNumber AS INTEGER
  LOCAL fileInDb AS STRING
  LOCAL InString AS ASCIIZ * 300
  LOCAL i AS INTEGER

  REDIM RecentFiles(1 TO 8) AS STRING
  FOR Bc=1 TO GetMenuItemCount(g_hMenuReopen)
    MENU DELETE g_hMenuReopen, 1
  NEXT Bc
  FOR i=1 TO 8
    RecentFiles(i)=GetRecentFileName(i)
    IF RecentFiles(i)<>"" THEN
      AppendMenu g_hMenuReopen,%MF_OWNERDRAW OR %MF_ENABLED, _
              %IDM_RECENTFILES + i, BYVAL %NULL
    END IF
  NEXT i
END SUB
FUNCTION GetRecentFileName(BYVAL idx AS LONG) AS STRING
   LOCAL szSection AS ASCIIZ * 30
   LOCAL szKey     AS ASCIIZ * 30
   LOCAL szDefault AS ASCIIZ * 30
   LOCAL lRes      AS LONG
   LOCAL InString  AS ASCIIZ * 300
   szSection = "Reopen files"
   szKey = "File " & FORMAT$(idx)
   InString=IniRead(g_zIni,szSection,szKey,"")
   'msgbox "test:" & Instring
   'lRes = GetPrivateProfileString(szSection, szKey, szDefault, InString, %MAX_PATH, g_zIni)
   FUNCTION = InString 'IniRead(g_zIni,szSection,szKey,"")
   'IF lRes THEN FUNCTION = TRIM$(EXTRACT$(InString, CHR$(0)))

END FUNCTION
' *********************************************************************************************
' GetDroppedFiles - Retrieves the paths of the dropped files
' *********************************************************************************************
FUNCTION GetDroppedFiles (BYVAL hfInfo AS LONG, DroppedFiles() AS STRING) AS LONG

  LOCAL COUNT AS LONG                 ' // Number of dropped filed
  LOCAL fName AS ASCIIZ * %MAX_PATH   ' // Filename and path
  LOCAL ln    AS LONG                 ' // Length of the path + filename
  LOCAL tmp   AS STRING               ' // Temporary variable
  LOCAL i     AS LONG                 ' // Loop counter

  COUNT = DragQueryFile(hfInfo, &HFFFFFFFF&, BYVAL %NULL, 0) ' Get the number of dropped files

  IF COUNT THEN                                              ' If we got something...
    REDIM DroppedFiles(COUNT)
    FOR i = 0 TO COUNT - 1
      ln = DragQueryFile(hfInfo, i, fName, %MAX_PATH)      ' Retrieve the path and get his length
      IF ln THEN
        tmp = TRIM$(LEFT$(fName, ln))
        IF LEN(tmp) AND (GETATTR(tmp) AND 16) = 0 THEN    ' Make sure it's a file, not a folder
          DroppedFiles(i) = tmp
        END IF
      END IF
    NEXT
  END IF

  DragFinish hfInfo
  FUNCTION = %TRUE
END FUNCTION
'==============================================================================
FUNCTION GetStringTable(BYVAL ID AS LONG) AS STRING
  '------------------------------------------------------------------------------
  ' ����ID��ص��ַ������ı�
  '----------------------------------------------------------------------------
  SELECT CASE AS LONG ID
    CASE %IDM_NEWBAS   : FUNCTION = " �½�BAS�ļ� "
    CASE %IDM_OPEN     : FUNCTION = " ���Ѵ��ڵ��ļ� "
    CASE %IDM_SAVE     : FUNCTION = " ֱ�ӱ����ļ� "
    CASE %IDM_SAVEAS   : FUNCTION = " �򿪱����ļ��Ի���ȷ���󱣴� "
    CASE %IDM_PRINT    : FUNCTION = " ��ӡ��ǰ�ļ� "
    CASE %IDM_EXIT     : FUNCTION = " �رձ����� "
    CASE %IDM_UNDO     : FUNCTION = " �������һ�β��� "
    CASE %IDM_REDO     : FUNCTION = " �������һ�β��� "
    CASE %IDM_CUT      : FUNCTION = " ���е�ճ���� "
    CASE %IDM_COPY     : FUNCTION = " ���Ƶ�ճ���� "
    CASE %IDM_PASTE    : FUNCTION = " ��ճ����ճ�� "
    CASE %IDM_DELETE   : FUNCTION = " ɾ��ѡ����ı� "
    CASE %IDM_SELALL   : FUNCTION = " ѡ���������� "
    CASE %IDM_CASCADE  : FUNCTION = " ���ʽ���� "
    CASE %IDM_TILEH    : FUNCTION = " ������������ "
    CASE %IDM_TILEV    : FUNCTION = " ������������ "
    CASE %IDM_ARRANGE  : FUNCTION = " ������С���Ĵ��� "
    CASE %IDM_CLOSE    : FUNCTION = " �ر��ļ� "
    CASE %IDHELP       : FUNCTION = " ��ʾ�����ļ� "
    CASE %IDM_ABOUT    : FUNCTION = " ���ڱ����� "
    CASE %IDM_FIND          : FUNCTION = " ���� "
    CASE %IDM_REPLACE       : FUNCTION = " �滻 "
    CASE %IDM_COMPILE       : FUNCTION = " ���� "
    CASE %IDM_COMPILERUN : FUNCTION = " ���벢ִ�� "
    CASE %IDM_EXECUTE       : FUNCTION = " ������ֱ������ "
    CASE %IDM_PRINTPREVIEW  : FUNCTION = " ��ӡԤ�� "
  END SELECT
END FUNCTION
'==============================================================================
SUB HideButtons (BYVAL HideState AS LONG)
'------------------------------------------------------------------------------
  ' ��ʾ/���ذ�ť������/�����ò˵�
  '----------------------------------------------------------------------------
  LOCAL ITEM        AS DWORD
  LOCAL mFlag       AS DWORD
  LOCAL tmpDword    AS DWORD
  tmpDword=g_hToolbar
  IF ghToolBar>0 THEN
    tmpDword=ghToolBar
  END IF
  IF GetSubMenu(g_hMenu, 0) = g_hMenuFile THEN ITEM = 1 ELSE ITEM = 2
  IF HideState THEN mFlag = %MF_GRAYED ELSE mFlag = %MF_ENABLED
  'msgbox str$(item)
  EnableMenuItem g_hMenu, ITEM, %MF_BYPOSITION OR mFlag '����/��ʾ�༭�˵�
  'EnableMenuItem g_hMenu, ITEM + 6, %MF_BYPOSITION OR mFlag
  EnableMenuItem g_hMenuWindow, 0, %MF_BYPOSITION OR mFlag
  EnableMenuItem g_hMenuWindow, 1, %MF_BYPOSITION OR mFlag
  EnableMenuItem g_hMenuWindow, 2, %MF_BYPOSITION OR mFlag
  EnableMenuItem g_hMenuWindow, 3, %MF_BYPOSITION OR mFlag
  EnableMenuItem g_hMenuWindow, 4, %MF_BYPOSITION OR mFlag
  EnableMenuItem g_hMenuWindow, 5, %MF_BYPOSITION OR mFlag
  EnableMenuItem g_hMenuWindow, 6, %MF_BYPOSITION OR mFlag
  EnableMenuItem g_hMenuWindow, 7, %MF_BYPOSITION OR mFlag

  EnableMenuItem g_hMenuFile, %IDM_SAVE,   %MF_BYCOMMAND OR mFlag
  EnableMenuItem g_hMenuFile, %IDM_SAVEAS, %MF_BYCOMMAND OR mFlag
  EnableMenuItem g_hMenuFile, %IDM_SAVEALL, %MF_BYCOMMAND OR mFlag
  EnableMenuItem g_hMenuFile, %IDM_SAVEPROJECT, %MF_BYCOMMAND OR mFlag
  EnableMenuItem g_hMenuFile, %IDM_PRINTSETTING, %MF_BYCOMMAND OR mFlag
  EnableMenuItem g_hMenuFile, %IDM_PRINTPREVIEW, %MF_BYCOMMAND OR mFlag
  EnableMenuItem g_hMenuFile, %IDM_PRINT,  %MF_BYCOMMAND OR mFlag
  DrawMenuBar g_hWndMain
  'msgbox str$(HideState)
  'HideState=1-HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, %IDM_UNDO,  ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, %IDM_REDO,  ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, 703,        ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, %IDM_PASTE, ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, %IDM_COPY,  ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, %IDM_CUT,   ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, 702,        ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, %IDM_PRINT, ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, 701,        ISFALSE HideState
  SendMessage tmpDword, %TB_ENABLEBUTTON, %IDM_SAVE,  ISFALSE HideState
  UpdateWindow tmpDword
END SUB
' *********************************************************************************************
' �ı�˵����������ť��״̬
' *********************************************************************************************
SUB ChangeButtonsState
  LOCAL fModify     AS LONG    ' �޸ı�ʶModify flag
  LOCAL hSci        AS DWORD   ' �༭���ؼ����
  LOCAL startSelPos AS LONG    ' ��ѡ���ı��Ŀ�ʼλ��
  LOCAL endSelPos   AS LONG    ' ��ѡ���ı��Ľ���λ��
  LOCAL numLines    AS LONG    ' ����
  LOCAL bitnum      AS LONG    ' λ��
  LOCAL wi          AS WININFOSTRUC PTR

  ' �õ��༭���ڵľ��
  hSci = GetEdit
  ' ���û���κ��ļ����༭�򹤾�����ť�����á�
  IF ISFALSE hSci THEN
    DisableToolbarButtons
    EXIT SUB
  END IF
  ' �õ���ͼ���ڵľ��
  wi = GetWindowLong(MdiGetActive(g_hWndClient), %GWL_USERDATA)
  ' If it is a window "initialized as new" set the modified    �������һ���ճ�ʼ�����µĴ��ڣ������޸�״̬Ϊδ�ı䣬
  ' status to unchanged to disable the Save button.            ʹ���水ť������
  IF wi THEN
    IF @wi.binit THEN
      bitnum = GetDlgCtrlID(@wi.hFocus) - %IDC_EDIT2
      IF bitnum => 0  AND bitnum < 16 THEN
        IF BIT(@wi.bInit, bitnum) THEN
          SendMessage @wi.hFocus, %SCI_SETSAVEPOINT, 0, 0
        END IF
      END IF
    END IF
  END IF
  IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_SAVE, 0) THEN _
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_SAVE, %TRUE

  IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_SAVEAS, 0) THEN _
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_SAVEAS, %TRUE

  IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_REFRESH, 0) THEN _
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_REFRESH, %TRUE

  IF ISFALSE SendMessage(ghCodeEditBar, %TB_ISBUTTONENABLED, %IDM_FIND, 0) THEN _
    SendMessage ghCodeEditBar, %TB_ENABLEBUTTON, %IDM_FIND, %TRUE

  IF ISFALSE SendMessage(ghCodeEditBar, %TB_ISBUTTONENABLED, %IDM_REPLACE, 0) THEN _
    SendMessage ghCodeEditBar, %TB_ENABLEBUTTON, %IDM_REPLACE, %TRUE

  IF ISFALSE SendMessage(ghDebugBar, %TB_ISBUTTONENABLED, %IDM_COMPILE, 0) THEN _
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_COMPILE, %TRUE

  IF ISFALSE SendMessage(ghDebugBar, %TB_ISBUTTONENABLED, %IDM_COMPILERUN, 0) THEN _
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_COMPILERUN, %TRUE

  IF ISFALSE SendMessage(ghDebugBar, %TB_ISBUTTONENABLED, %IDM_EXECUTE, 0) THEN _
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_EXECUTE, %TRUE

  IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_PRINT, 0) THEN _
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_PRINT, %TRUE

  'IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_PRINTPREVIEW, 0) THEN _
  ' SendMessage hToolbar, %TB_ENABLEBUTTON, %IDM_PRINTPREVIEW, %TRUE

  IF ISFALSE SendMessage(ghWindowBar, %TB_ISBUTTONENABLED, %IDM_CLOSE, 0) THEN _
    SendMessage ghWindowBar, %TB_ENABLEBUTTON, %IDM_CLOSE, %TRUE

  IF ISFALSE SendMessage(GetEdit, %SCI_CANPASTE, 0, 0) THEN
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_PASTE, %FALSE
    EnableMenuItem g_hMenuEdit, %IDM_PASTE, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_PASTEIE, %MF_GRAYED
  ELSE
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_PASTE, %TRUE
    EnableMenuItem g_hMenuEdit, %IDM_PASTE, %MF_ENABLED
    EnableMenuItem g_hMenuEdit, %IDM_PASTEIE, %MF_ENABLED
  END IF

  IF SendMessage(GetEdit, %SCI_GETSELECTIONSTART, 0, 0) = SendMessage(GetEdit, %SCI_GETSELECTIONEND, 0, 0) THEN
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_CUT, %FALSE
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_COPY, %FALSE
    EnableMenuItem g_hMenuEdit, %IDM_CUT, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_COPY, %MF_GRAYED
  ELSE
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_CUT, %TRUE
    SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_COPY, %TRUE
    EnableMenuItem g_hMenuEdit, %IDM_CUT, %MF_ENABLED
    EnableMenuItem g_hMenuEdit, %IDM_COPY, %MF_ENABLED
  END IF

  ' Retrieves if the file has been modified since the last saving
  ' and disable or enable the Save button and menu
  fModify = SendMessage(hSci, %SCI_GETMODIFY, 0, 0)
  IF ISFALSE fModify THEN
    IF ISTRUE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_SAVE, 0) THEN _
      SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_SAVE, %FALSE
    EnableMenuItem g_hMenuFile, %IDM_SAVE, %MF_GRAYED
  ELSE
    IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_SAVE, 0) THEN _
      SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_SAVE, %TRUE
    EnableMenuItem g_hMenuFile, %IDM_SAVE, %MF_ENABLED
  END IF

  ' Disable or enable the undo button and menu
  IF ISFALSE SendMessage(hSci, %SCI_CANUNDO, 0, 0) THEN
    IF ISTRUE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_UNDO, 0) THEN _
      SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_UNDO, %FALSE
    EnableMenuItem g_hMenuEdit, %IDM_UNDO, %MF_GRAYED
  ELSE
    IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_UNDO, 0) THEN _
      SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_UNDO, %TRUE
    EnableMenuItem g_hMenuEdit, %IDM_UNDO, %MF_ENABLED
  END IF

  ' Disable or enable the redo button and menu
  IF ISFALSE SendMessage(hSci, %SCI_CANREDO, 0, 0) THEN
    IF ISTRUE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_REDO, 0) THEN _
      SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_REDO, %FALSE
    EnableMenuItem g_hMenuEdit, %IDM_REDO, %MF_GRAYED
  ELSE
    IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_REDO, 0) THEN _
      SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_REDO, %TRUE
    EnableMenuItem g_hMenuEdit, %IDM_REDO, %MF_ENABLED
  END IF

  ' Disable or enable the paste button
  IF ISFALSE SendMessage(hSci, %SCI_CANPASTE, 0, 0) THEN
    IF ISTRUE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_PASTE, 0) THEN _
      SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_PASTE, %FALSE
  ELSE
    IF ISFALSE SendMessage(ghToolbar, %TB_ISBUTTONENABLED, %IDM_PASTE, 0) THEN _
      SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_PASTE, %TRUE
  END IF

  ' Enable Edit and Run submenus
  EnableSubmenus

  ' File menu
  EnableMenuItem g_hMenuFile, %IDM_INSERTFILE, %MF_ENABLED
  EnableMenuItem g_hMenuFile, %IDM_SAVE, %MF_ENABLED
  EnableMenuItem g_hMenuFile, %IDM_SAVEAS, %MF_ENABLED
  EnableMenuItem g_hMenuFile, %IDM_CLOSE, %MF_ENABLED
  EnableMenuItem g_hMenuFile, %IDM_CLOSEALL, %MF_ENABLED
  EnableMenuItem g_hMenuFile, %IDM_PRINT, %MF_ENABLED
  EnableMenuItem g_hMenuFile, %IDM_PRINTPREVIEW, %MF_ENABLED

  ' Edit menu
  EnableMenuItem g_hMenuEdit, %IDM_LINEDELETE, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_CLEAR, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_CLEARALL, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_SELALL, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_INDENT, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_OUTDENT, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_COMMENT, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_UNCOMMENT, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_FORMATREGION, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_TABULATEREGION, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_SELTOUPPERCASE, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_SELTOLOWERCASE, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_SELTOMIXEDCASE, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_TEMPLATES, %MF_ENABLED
  EnableMenuItem g_hMenuEdit, %IDM_HTMLCODE, %MF_ENABLED

  ' Search menu
  EnableMenuItem g_hMenuSearch, %IDM_FIND, %MF_ENABLED
  EnableMenuItem g_hMenuSearch, %IDM_FINDNEXT, %MF_ENABLED
  EnableMenuItem g_hMenuSearch, %IDM_FINDBACKWARDS, %MF_ENABLED
  EnableMenuItem g_hMenuSearch, %IDM_REPLACE, %MF_ENABLED
  EnableMenuItem g_hMenuSearch, %IDM_GOTOLINE, %MF_ENABLED
  EnableMenuItem g_hMenuSearch, %IDM_TOGGLEBOOKMARK, %MF_ENABLED
  EnableMenuItem g_hMenuSearch, %IDM_NEXTBOOKMARK, %MF_ENABLED
  EnableMenuItem g_hMenuSearch, %IDM_PREVIOUSBOOKMARK, %MF_ENABLED
  EnableMenuItem g_hMenuSearch, %IDM_DELETEBOOKMARKS, %MF_ENABLED

  ' Run menu
  EnableMenuItem g_hMenuRun, %IDM_COMPILE, %MF_ENABLED
  EnableMenuItem g_hMenuRun, %IDM_COMPILERUN, %MF_ENABLED
  EnableMenuItem g_hMenuRun, %IDM_COMPILEDEBUG, %MF_ENABLED
  EnableMenuItem g_hMenuRun, %IDM_EXECUTE, %MF_ENABLED
  EnableMenuItem g_hMenuRun, %IDM_SETPRIMARY, %MF_ENABLED
  EnableMenuItem g_hMenuRun, %IDM_DEBUGTOOL, %MF_ENABLED

  ' View menu
  EnableMenuItem g_hMenuView, %IDM_TOGGLE, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_TOGGLEALL, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_FOLDALL, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_EXPANDALL, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_ZOOMIN, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_ZOOMOUT, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_CVEOLTOCRLF, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_CVEOLTOCR, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_CVEOLTOLF, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_REPLSPCWITHTABS, %MF_ENABLED
  EnableMenuItem g_hMenuView, %IDM_REPLTABSWITHSPC, %MF_ENABLED

  ' Window menu
  EnableMenuItem g_hMenuWindow, %IDM_CASCADE, %MF_ENABLED
  EnableMenuItem g_hMenuWindow, %IDM_TILEH, %MF_ENABLED
  EnableMenuItem g_hMenuWindow, %IDM_TILEV, %MF_ENABLED
  EnableMenuItem g_hMenuWindow, %IDM_SWITCHWINDOW, %MF_ENABLED
  EnableMenuItem g_hMenuWindow, %IDM_ARRANGE, %MF_ENABLED
  EnableMenuItem g_hMenuWindow, %IDM_CLOSE, %MF_ENABLED

  ' Project window
  IF LEN(sProjectName) = 0 THEN
    EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECT, %MF_GRAYED
    EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECTAS, %MF_GRAYED
    EnableMenuItem g_hMenuProject, %IDM_CLOSEPROJECT, %MF_GRAYED
    'EnableMenuItem g_hMenuProject, %IDM_HIDEPROJECTWINDOW, %MF_GRAYED
    'EnableMenuItem g_hMenuProject, %IDM_RESTOREPROJECTWINDOW, %MF_GRAYED
    'EnableMenuItem g_hMenuProject, %IDM_TOGGLEPROJWINDOW, %MF_GRAYED
  ELSE
    IF UCASE$(sProjectName) = "NONAME.PBP" THEN
      EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECT, %MF_GRAYED
      EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECTAS, %MF_GRAYED
      EnableMenuItem g_hMenuProject, %IDM_CLOSEPROJECT, %MF_GRAYED
    ELSE
      EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECT, %MF_ENABLED
      EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECTAS, %MF_ENABLED
      EnableMenuItem g_hMenuProject, %IDM_CLOSEPROJECT, %MF_ENABLED
    END IF
  END IF

  ' Tools menu
  EnableMenuItem g_hMenuTool, %IDM_CODEC, %MF_ENABLED

  ' If startSelPos and endSelPos are the same there is not selection,
  startSelPos = SendMessage(GetEdit, %SCI_GETSELECTIONSTART, 0, 0)
  endSelPos = SendMessage(GetEdit, %SCI_GETSELECTIONEND, 0, 0)
  IF startSelPos = endSelPos THEN
    EnableMenuItem g_hMenuEdit, %IDM_INDENT, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_OUTDENT, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_COMMENT, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_UNCOMMENT, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_FORMATREGION, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_TABULATEREGION, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_SELTOUPPERCASE, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_SELTOLOWERCASE, %MF_GRAYED
    EnableMenuItem g_hMenuEdit, %IDM_SELTOMIXEDCASE, %MF_GRAYED
  END IF

  ' If the file is empty, disable find and bookmark options
  numLines = SendMessage(GetEdit, %SCI_GETLINECOUNT, 0, 0) - 1
  IF ISFALSE numLines THEN
    EnableMenuItem g_hMenuSearch, %IDM_FIND, %MF_GRAYED
    EnableMenuItem g_hMenuSearch, %IDM_FINDNEXT, %MF_GRAYED
    EnableMenuItem g_hMenuSearch, %IDM_FINDBACKWARDS, %MF_GRAYED
    EnableMenuItem g_hMenuSearch, %IDM_REPLACE, %MF_GRAYED
    EnableMenuItem g_hMenuSearch, %IDM_GOTOLINE, %MF_GRAYED
    EnableMenuItem g_hMenuSearch, %IDM_TOGGLEBOOKMARK, %MF_GRAYED
    EnableMenuItem g_hMenuSearch, %IDM_NEXTBOOKMARK, %MF_GRAYED
    EnableMenuItem g_hMenuSearch, %IDM_PREVIOUSBOOKMARK, %MF_GRAYED
    EnableMenuItem g_hMenuSearch, %IDM_DELETEBOOKMARKS, %MF_GRAYED
  END IF
END SUB
' *********************************************************************************************
' Disable toolbar buttons and menu items
' *********************************************************************************************
SUB DisableToolbarButtons

   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_SAVE, %FALSE
   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_SAVEAS, %FALSE
   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_REFRESH, %FALSE
   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_UNDO, %FALSE
   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_REDO, %FALSE
   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_CUT, %FALSE
   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_COPY, %FALSE
   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_PASTE, %FALSE
   SendMessage ghCodeEditBar, %TB_ENABLEBUTTON, %IDM_FIND, %FALSE
   SendMessage ghCodeEditBar, %TB_ENABLEBUTTON, %IDM_REPLACE, %FALSE
   SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_COMPILEDEBUG, %FALSE
   SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_PRINT, %FALSE
   'SendMessage ghToolbar, %TB_ENABLEBUTTON, %IDM_PRINTPREVIEW, %FALSE
   SendMessage ghWindowBar, %TB_ENABLEBUTTON, %IDM_CLOSE, %FALSE

   IF LEN(sProjectPrimary) THEN
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_COMPILE, %TRUE
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_COMPILERUN, %TRUE
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_EXECUTE, %TRUE
   ELSE
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_COMPILE, %FALSE
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_COMPILERUN, %FALSE
    SendMessage ghDebugBar, %TB_ENABLEBUTTON, %IDM_EXECUTE, %FALSE
   END IF

   ' Disable Edit and Run submenus
   DisableSubmenus

   ' File menu
   EnableMenuItem g_hMenuFile, %IDM_INSERTFILE, %MF_GRAYED
   EnableMenuItem g_hMenuFile, %IDM_SAVE, %MF_GRAYED
   EnableMenuItem g_hMenuFile, %IDM_SAVEAS, %MF_GRAYED
   EnableMenuItem g_hMenuFile, %IDM_CLOSE, %MF_GRAYED
   EnableMenuItem g_hMenuFile, %IDM_CLOSEALL, %MF_GRAYED
   EnableMenuItem g_hMenuFile, %IDM_PRINT, %MF_GRAYED
   EnableMenuItem g_hMenuFile, %IDM_PRINTPREVIEW, %MF_GRAYED

   ' Edit menu
'   EnableMenuItem g_hMenuEdit, %IDM_UNDO, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_REDO, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_LINEDELETE, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_INITASNEW, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_CLEAR, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_CLEARALL, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_CUT, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_COPY, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_PASTE, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_PASTEIE, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_SELECTALL, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_BLOCKINDENT, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_BLOCKUNINDENT, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_COMMENT, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_UNCOMMENT, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_FORMATREGION, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_TABULATEREGION, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_SELTOUPPERCASE, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_SELTOLOWERCASE, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_SELTOMIXEDCASE, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_TEMPLATES, %MF_GRAYED
'   EnableMenuItem g_hMenuEdit, %IDM_HTMLCODE, %MF_GRAYED

   ' Search menu
   EnableMenuItem g_hMenuSearch, %IDM_FIND, %MF_GRAYED
   EnableMenuItem g_hMenuSearch, %IDM_FINDNEXT, %MF_GRAYED
   EnableMenuItem g_hMenuSearch, %IDM_FINDBACKWARDS, %MF_GRAYED
   EnableMenuItem g_hMenuSearch, %IDM_REPLACE, %MF_GRAYED
   EnableMenuItem g_hMenuSearch, %IDM_GOTOLINE, %MF_GRAYED
   EnableMenuItem g_hMenuSearch, %IDM_TOGGLEBOOKMARK, %MF_GRAYED
   EnableMenuItem g_hMenuSearch, %IDM_NEXTBOOKMARK, %MF_GRAYED
   EnableMenuItem g_hMenuSearch, %IDM_PREVIOUSBOOKMARK, %MF_GRAYED
   EnableMenuItem g_hMenuSearch, %IDM_DELETEBOOKMARKS, %MF_GRAYED

   ' Run menu
'   EnableMenuItem g_hMenuRun, %IDM_COMPILEANDDEBUG, %MF_GRAYED
'   IF LEN(sProjectPrimary) THEN
'      EnableMenuItem g_hMenuRun, %IDM_COMPILE, %MF_ENABLED
'      EnableMenuItem g_hMenuRun, %IDM_COMPILEANDRUN, %MF_ENABLED
'      EnableMenuItem g_hMenuRun, %IDM_EXECUTE, %MF_ENABLED
'      EnableMenuItem g_hMenuRun, %IDM_PRIMARYSOURCE, %MF_ENABLED
'      EnableMenuItem g_hMenuRun, %IDM_DEBUGTOOL, %MF_ENABLED
'   else
'      EnableMenuItem g_hMenuRun, %IDM_COMPILE, %MF_GRAYED
'      EnableMenuItem g_hMenuRun, %IDM_COMPILEANDRUN, %MF_GRAYED
'      EnableMenuItem g_hMenuRun, %IDM_EXECUTE, %MF_GRAYED
'      EnableMenuItem g_hMenuRun, %IDM_PRIMARYSOURCE, %MF_GRAYED
'      EnableMenuItem g_hMenuRun, %IDM_DEBUGTOOL, %MF_GRAYED
'   end if

   ' View menu
   EnableMenuItem g_hMenuView, %IDM_TOGGLE, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_TOGGLEALL, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_FOLDALL, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_EXPANDALL, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_ZOOMIN, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_ZOOMOUT, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_CVEOLTOCRLF, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_CVEOLTOCR, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_CVEOLTOLF, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_REPLSPCWITHTABS, %MF_GRAYED
   EnableMenuItem g_hMenuView, %IDM_REPLTABSWITHSPC, %MF_GRAYED

   ' Window menu
   EnableMenuItem g_hMenuWindow, %IDM_CASCADE, %MF_GRAYED
   EnableMenuItem g_hMenuWindow, %IDM_TILEH, %MF_GRAYED
   EnableMenuItem g_hMenuWindow, %IDM_TILEV, %MF_GRAYED
   EnableMenuItem g_hMenuWindow, %IDM_SWITCHWINDOW, %MF_GRAYED
   EnableMenuItem g_hMenuWindow, %IDM_ARRANGE, %MF_GRAYED
   EnableMenuItem g_hMenuWindow, %IDM_CLOSE, %MF_GRAYED

   ' Project window
   IF LEN(sProjectName) = 0 THEN
    EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECT, %MF_GRAYED
    EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECTAS, %MF_GRAYED
    EnableMenuItem g_hMenuProject, %IDM_CLOSEPROJECT, %MF_GRAYED
    'EnableMenuItem g_hMenuProject, %IDM_HIDEPROJECTWINDOW, %MF_GRAYED
    'EnableMenuItem g_hMenuProject, %IDM_RESTOREPROJECTWINDOW, %MF_GRAYED
    'EnableMenuItem g_hMenuProject, %IDM_TOGGLEPROJWINDOW, %MF_GRAYED
   ELSE
    IF UCASE$(sProjectName) = "NONAME.SPF" THEN
     EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECT, %MF_GRAYED
     EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECTAS, %MF_GRAYED
     EnableMenuItem g_hMenuProject, %IDM_CLOSEPROJECT, %MF_GRAYED
    ELSE
     EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECT, %MF_ENABLED
     EnableMenuItem g_hMenuProject, %IDM_SAVEPROJECTAS, %MF_ENABLED
     EnableMenuItem g_hMenuProject, %IDM_CLOSEPROJECT, %MF_ENABLED
    END IF
    'EnableMenuItem g_hMenuProject, %IDM_HIDEPROJECTWINDOW, %MF_ENABLED
    'EnableMenuItem g_hMenuProject, %IDM_RESTOREPROJECTWINDOW, %MF_ENABLED
    'EnableMenuItem g_hMenuProject, %IDM_TOGGLEPROJWINDOW, %MF_ENABLED
   END IF

'   IF ISTRUE IsWindowVisible(ghDockC) THEN
'    EnableMenuItem g_hMenuProject, %IDM_HIDEPROJECTWINDOW, %MF_ENABLED
'    EnableMenuItem g_hMenuProject, %IDM_RESTOREPROJECTWINDOW, %MF_GRAYED
'    IF gdi.l <> "F" THEN
'     EnableMenuItem g_hMenuProject, %IDM_TOGGLEPROJWINDOW, %MF_ENABLED
'    ELSE
'     EnableMenuItem g_hMenuProject, %IDM_TOGGLEPROJWINDOW, %MF_GRAYED
'    END IF
'   ELSE
'    EnableMenuItem g_hMenuProject, %IDM_HIDEPROJECTWINDOW, %MF_GRAYED
'    EnableMenuItem g_hMenuProject, %IDM_RESTOREPROJECTWINDOW, %MF_ENABLED
'    EnableMenuItem g_hMenuProject, %IDM_TOGGLEPROJWINDOW, %MF_GRAYED
'   END IF
   ' Tools menu
   EnableMenuItem g_hMenuTool, %IDM_CODEC, %MF_GRAYED
END SUB
' *********************************************************************************************
' ���ñ༭�����в˵����Ӳ˵���Լ�������ٲ���������
' If there are MDI windows opened the position changes because MDI adds a system menu
' *********************************************************************************************
SUB DisableSubmenus
  LOCAL lMenuItem AS LONG
  LOCAL hMenuHandle AS DWORD
'   IF GetSubMenu(hMenu, 0) = hMenuFile THEN lMenuItem = 1 ELSE lMenuItem = 2
'   EnableMenuItem hMenu, lMenuItem, %MF_BYPOSITION OR %MF_GRAYED
'   EnableMenuItem hMenu, lMenuItem + 2, %MF_BYPOSITION OR %MF_GRAYED
''   EnableMenuItem hMenu, lMenuItem + 4, %MF_BYPOSITION OR %MF_GRAYED
  FOR lMenuItem = 0 TO 255
    hMenuHandle = GetSubMenu(g_hMenu, lMenuItem)
    IF hMenuHandle = 0 THEN EXIT FOR
    IF hMenuHandle = g_hMenuEdit OR hMenuHandle = g_hMenuRun THEN
      EnableMenuItem g_hMenu, lMenuItem, %MF_BYPOSITION OR %MF_GRAYED
    END IF
  NEXT
  DrawMenuBar g_hWndMain
END SUB
' *********************************************************************************************
' Enable Edit and Run submenus, and the CodeFinder combobox
' If there are MDI windows opened the position changes because MDI adds a system menu
' *********************************************************************************************
SUB EnableSubmenus

   LOCAL lMenuItem AS LONG
   LOCAL hMenuHandle AS DWORD
'   IF GetSubMenu(hMenu, 0) = hMenuFile THEN lMenuItem = 1 ELSE lMenuItem = 2
'   EnableMenuItem hMenu, lMenuItem, %MF_BYPOSITION OR %MF_ENABLED
'   EnableMenuItem hMenu, lMenuItem + 2, %MF_BYPOSITION OR %MF_ENABLED
''   EnableMenuItem hMenu, lMenuItem + 4, %MF_BYPOSITION OR %MF_ENABLED
   FOR lMenuItem = 0 TO 255
    hMenuHandle = GetSubMenu(g_hMenu, lMenuItem)
    IF hMenuHandle = 0 THEN EXIT FOR
    IF hMenuHandle = g_hMenuEdit OR hMenuHandle = g_hMenuRun THEN
     EnableMenuItem g_hMenu, lMenuItem, %MF_BYPOSITION OR %MF_ENABLED
    END IF
   NEXT
   DrawMenuBar g_hWndMain

END SUB
' *********************************************************************************************
' Creates a table of keyboard accelerators
' *********************************************************************************************
FUNCTION CreateAccelTable () AS DWORD

   DIM ac(98) AS ACCELAPI

   ' // Alt+X - Quit the application
   ac(0).fvirt = %FVIRTKEY OR %FALT
   ac(0).key   = ASC("X")
   ac(0).cmd   = %IDM_EXIT

   ' // Ctrl+O - Open
   ac(1).fvirt = %FVIRTKEY OR %FCONTROL
   ac(1).key   = ASC("O")
   ac(1).cmd   = %IDM_OPEN

   ' // Ctrl+N - New
   ac(2).fvirt = %FVIRTKEY OR %FCONTROL
   ac(2).key   = ASC("N")
   ac(2).cmd   = %IDM_NEW

   ' // Ctrl+S - Save As...
   ac(3).fvirt = %FVIRTKEY OR %FCONTROL
   ac(3).key   = ASC("S")
   ac(3).cmd   = %IDM_SAVE

   ' // Ctr+F4 - Close file
   ac(4).fvirt = %FVIRTKEY OR %FCONTROL
   ac(4).key   = %VK_F4
   ac(4).cmd   = %IDM_CLOSE 'FILE

   ' // Alt+F4 - Exit the editor
   ac(5).fvirt = %FVIRTKEY OR %FALT
   ac(5).key   = %VK_F4
   ac(5).cmd   = %IDM_EXIT

   ' // F3 - Find next
   ac(6).fvirt = %FVIRTKEY
   ac(6).key   = %VK_F3
   ac(6).cmd   = %IDM_FINDNEXT

   ' // F8 - Toggle current sub/function
   ac(7).fvirt = %FVIRTKEY
   ac(7).key   = %VK_F8
   ac(7).cmd   = %IDM_TOGGLE

   ' // Ctrl+F8 - Toggle current and all below
   ac(8).fvirt = %FVIRTKEY OR %FCONTROL
   ac(8).key   = %VK_F8
   ac(8).cmd   = %IDM_TOGGLEALL

   ' // Alt+F8 - Fold all subs/functions
   ac(9).fvirt = %FVIRTKEY OR %FALT
   ac(9).key   = %VK_F8
   ac(9).cmd   = %IDM_FOLDALL

   ' // Shift+F8 - Expand all subs/functions
   ac(10).fvirt = %FVIRTKEY OR %FSHIFT
   ac(10).key   = %VK_F8
   ac(10).cmd   = %IDM_EXPANDALL

   ' // Ctrl+F - Find
   ac(11).fvirt = %FVIRTKEY OR %FCONTROL
   ac(11).key   = ASC("F")
   ac(11).cmd   = %IDM_FIND

   ' // Ctrl+R - Find and replace
   ac(12).fvirt = %FVIRTKEY OR %FCONTROL
   ac(12).key   = ASC("R")
   ac(12).cmd   = %IDM_REPLACE

   ' // Ctrl+W - Switch windows
   ac(13).fvirt = %FVIRTKEY OR %FCONTROL
   ac(13).key   = ASC("W")
   ac(13).cmd   = %IDM_SWITCHWINDOW

   ' // Ctrl+G - Go to line...
   ac(14).fvirt = %FVIRTKEY OR %FCONTROL
   ac(14).key   = ASC("G")
   ac(14).cmd   = %IDM_GOTOLINE

   ' // Ctrl+Q - Block comment
   ac(15).fvirt = %FVIRTKEY OR %FCONTROL
   ac(15).key   = ASC("Q")
   ac(15).cmd   = %IDK_COMMENT

   ' // Ctrl+Shift+Q - Block uncomment
   ac(16).fvirt = %FVIRTKEY OR %FSHIFT OR %FCONTROL
   ac(16).key   = ASC("Q")
   ac(16).cmd   = %IDK_UNCOMMENT

   ' // Ctrl+F2 - Togle bookmark
   ac(17).fvirt = %FVIRTKEY OR %FCONTROL
   ac(17).key   = %VK_F2
   ac(17).cmd   = %IDM_TOGGLEBOOKMARK

   ' // F2 - Next bookmark
   ac(18).fvirt = %FVIRTKEY
   ac(18).key   = %VK_F2
   ac(18).cmd   = %IDM_NEXTBOOKMARK

   ' // Shift+F2 - Previous bookmark
   ac(19).fvirt = %FVIRTKEY OR %FSHIFT
   ac(19).key   = %VK_F2
   ac(19).cmd   = %IDM_PREVIOUSBOOKMARK

   ' // Ctrl+Shift+F2 - Delete bookmarks
   ac(20).fvirt = %FVIRTKEY OR %FSHIFT OR %FCONTROL
   ac(20).key   = %VK_F2
   ac(20).cmd   = %IDM_DELETEBOOKMARKS

   ' // Ctrl+Shift+V - Paste Internet Explorer html text
   ac(21).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(21).key   = ASC("V")
   ac(21).cmd   = %IDM_PASTEIE

   ' // Ctrl+Shift+L - Show/hide line numbers
   ac(22).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(22).key   = ASC("L")
   ac(22).cmd   = %IDM_SHOWLINENUM

   ' // Ctrl+Shift+M - Show/hide margin
   ac(23).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(23).key   = ASC("M")
   ac(23).cmd   = %IDM_SHOWMARGIN

   ' // Ctrl+Shift+I - Show/hide indentation lines
   ac(24).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(24).key   = ASC("I")
   ac(24).cmd   = %IDM_SHOWINDENT

   ' // Ctrl+Shift+W - Show/hide white spaces
   ac(25).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(25).key   = ASC("W")
   ac(25).cmd   = %IDM_SHOWSPACES

   ' // Ctrl+Shift+D - Show/hide end of lines
   ac(26).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(26).key   = ASC("D")
   ac(26).cmd   = %IDM_SHOWEOL

   ' // F1 - Help
   ac(27).fvirt =  %FVIRTKEY
   ac(27).key   = %VK_F1
   ac(27).cmd   = %IDM_HELP

   ' // Ctrl+Shift+G - Show/hide edge
   ac(28).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(28).key   = ASC("G")
   ac(28).cmd   = %IDM_SHOWEDGE

   ' // Ctrl+Shift+A - Auto indentation on/off
   ac(29).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(29).key   = ASC("A")
   ac(29).cmd   = %IDM_AUTOINDENT

   ' // Ctrl+Shift+T - Use tabs on/off
   ac(30).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(30).key   = ASC("T")
   ac(30).cmd   = %IDM_USETABS

   ' // Ctrl+E - Compile and execute
   ac(31).fvirt = %FVIRTKEY OR %FCONTROL
   ac(31).key   = ASC("E")
   ac(31).cmd   = %IDM_COMPILERUN

   ' // Ctrl+Y - Delete line
   ac(32).fvirt = %FVIRTKEY OR %FCONTROL
   ac(32).key   = ASC("Y")
   ac(32).cmd   = %IDM_LINEDELETE

   ' // Ctrl+X - Redo
   ac(33).fvirt = %FVIRTKEY OR %FCONTROL
   ac(33).key   = ASC("K")
   ac(33).cmd   = %IDM_REDO

   ' // Ctr+F5 - Compile
   ac(34).fvirt = %FVIRTKEY OR %FCONTROL
   ac(34).key   = %VK_F5
   ac(34).cmd   = %IDM_COMPILE

   ' // F5 - Compile and execute
   ac(35).fvirt = %FVIRTKEY
   ac(35).key   = %VK_F5
   ac(35).cmd   = %IDM_COMPILERUN

   ' // Ctrl+Enter - Autocomplete
   ac(36).fvirt = %FVIRTKEY OR %FCONTROL
   ac(36).key   = %VK_RETURN
   ac(36).cmd   = %IDM_AUTOCOMPLETE

   ' // Ctrl+Shift+C - Templates
   ac(37).fvirt = %FVIRTKEY OR %FSHIFT OR %FCONTROL
   ac(37).key   = ASC("C")
   ac(37).cmd   = %IDM_TEMPLATES

   ' // Alt+Shift+F - FileFind
   ac(38).fvirt = %FVIRTKEY OR %FSHIFT OR %FALT
   ac(38).key   = ASC("F")
   ac(38).cmd   = %IDM_FILEFIND

   ' // Ctrl+Alt+U - Convert selected text to upper case
   ac(39).fvirt = %FVIRTKEY OR %FCONTROL OR %FALT
   ac(39).key   = ASC("U")
   ac(39).cmd   = %IDM_SELTOUPPERCASE

   ' // Ctrl+Alt+L - Convert selected text to lower case
   ac(40).fvirt = %FVIRTKEY OR %FCONTROL OR %FALT
   ac(40).key   = ASC("L")
   ac(40).cmd   = %IDM_SELTOLOWERCASE

   ' // Ctrl+Alt+M - Convert selected text to mixed case
   ac(41).fvirt = %FVIRTKEY OR %FCONTROL OR %FALT
   ac(41).key   = ASC("M")
   ac(41).cmd   = %IDM_SELTOMIXEDCASE

   ' // Ctrl+P - Print
   ac(42).fvirt = %FVIRTKEY OR %FCONTROL
   ac(42).key   = ASC("P")
   ac(42).cmd   = %IDM_PRINT

   ' // Ctrl+Shift+P - Print preview
   ac(43).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(43).key   = ASC("P")
   ac(43).cmd   = %IDM_PRINTPREVIEW

   ' // F7 - PBFORMS
   ac(44).fvirt = %FVIRTKEY
   ac(44).key   = %VK_F7
   ac(44).cmd   = %IDM_PBFORMS

   ' // F9 - Code keeper
   ac(45).fvirt = %FVIRTKEY
   ac(45).key   = %VK_F9
   ac(45).cmd   = %IDM_CODEKEEPER

   ' // F6 - Compile and debug
   ac(46).fvirt = %FVIRTKEY
   ac(46).key   = %VK_F6
   ac(46).cmd   = %IDM_COMPILEDEBUG

   ' // Ctrl+B - Formats selected text
   ac(47).fvirt = %FVIRTKEY OR %FCONTROL
   ac(47).key   = ASC("B")
   ac(47).cmd   = %IDM_FORMATREGION

   ' // Ctrl+Shift+B - Tabulates text
   ac(48).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(48).key   = ASC("B")
   ac(48).cmd   = %IDM_TABULATEREGION

   ' // Ctrl+Shift+E - Launches Explorer
   ac(49).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(49).key   = ASC("O")
   ac(49).cmd   = %IDM_EXPLORER

   ' // Ctrl+Shift+K - Charmap
   ac(50).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(50).key   = ASC("K")
   ac(50).cmd   = %IDM_CHARMAP

   ' // Ctrl+Shift+Up arrow - Goto to the begining of the procedure or function
   ac(51).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(51).key   = %VK_UP
   ac(51).cmd   = %IDM_GOTOBEGINPROC

   ' // Ctrl+Shift+Down arrow - Goto to the end of the procedure or function
   ac(52).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(52).key   = %VK_DOWN
   ac(52).cmd   = %IDM_GOTOENDPROC

   ' // Ctrl+Alt+Up arrow - Goto to the begining of the current procedure or function
   ac(53).fvirt = %FVIRTKEY OR %FCONTROL OR %FALT
   ac(53).key   = %VK_UP
   ac(53).cmd   = %IDM_GOTOBEGINTHISPROC

   ' // Ctrl+Alt+Down arrow - Goto to the end of the current procedure or function
   ac(54).fvirt = %FVIRTKEY OR %FCONTROL OR %FALT
   ac(54).key   = %VK_DOWN
   ac(54).cmd   = %IDM_GOTOENDTHISPROC

   ' // Ctrl+Shift+S - Show procedure/function name
   ac(55).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(55).key   = ASC("S")
   ac(55).cmd   = %IDM_SHOWPROCNAME

   ' // Ctrl+Spacebar - Autocomplete (same as Ctrl+Enter)
   ac(56).fvirt = %FVIRTKEY OR %FCONTROL
   ac(56).key   = %VK_SPACE
   ac(56).cmd   = %IDM_AUTOCOMPLETE

   ' // F10 - Message box designer
   ac(57).fvirt = %FVIRTKEY
   ac(57).key   = %VK_F10
   ac(57).cmd   = %IDM_MSGBOXDESIGNER

   ' // Ctrl+Alt+F - Windows Find
   ac(58).fvirt = %FVIRTKEY OR %FALT OR %FCONTROL
   ac(58).key   = ASC("F")
   ac(58).cmd   = %IDM_WINDOWSFIND

   ' // Ctrl+Shift+F - Find in Files
   ac(59).fvirt = %FVIRTKEY OR %FSHIFT OR %FCONTROL
   ac(59).key   = ASC("F")
   ac(59).cmd   = %IDM_FINDINFILES

   ' // F11 - More tools
   ac(60).fvirt = %FVIRTKEY
   ac(60).key   = %VK_F11
   ac(60).cmd   = %IDM_MORETOOLS

   ' // Ctrl+Shift+E - Execute without compiling
   ac(61).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(61).key   = ASC("E")
   ac(61).cmd   = %IDM_EXECUTE

   ' // Shift+F3 - Search backwards
   ac(62).fvirt = %FVIRTKEY OR %FSHIFT
   ac(62).key   = %VK_F3
   ac(62).cmd   = %IDM_FINDBACKWARDS

   ' // Alt+F7 - Goto or load the file with has the same name as the one under the caret
   ac(63).fvirt = %FVIRTKEY OR %FALT
   ac(63).key   = %VK_F7
   ac(63).cmd   = %IDM_GOTOSELFILE

   ' // Ctrl+F7 - Search the function/sub that has the same name as the word under the caret
   ac(64).fvirt = %FVIRTKEY OR %FCONTROL
   ac(64).key   = %VK_F7
   ac(64).cmd   = %IDM_GOTOSELPROC

   ' // Shift+F7 - Returns to the last saved position
   ac(65).fvirt = %FVIRTKEY OR %FSHIFT
   ac(65).key   = %VK_F7
   ac(65).cmd   = %IDM_GOTOLASTPOSITION

   ' // Ctrl+M - Compile
   ac(66).fvirt = %FVIRTKEY OR %FCONTROL
   ac(66).key   = ASC("M")
   ac(66).cmd   = %IDM_COMPILE

'   ' // F12 - Toggle Projects/Messages view
'   ac(67).fvirt = %FVIRTKEY
'   ac(67).key   = %VK_F12
'   ac(67).cmd   = %IDM_TOGGLEPROJWINDOW

'   ' // Shift+F12 - Hide Projects/Messages vindow
'   ac(68).fvirt = %FVIRTKEY OR %FSHIFT
'   ac(68).key   = %VK_F12
'   ac(68).cmd   = %IDM_HIDEPROJECTWINDOW
'
'   ' // Ctrl+F12 - Restore/Show Projects/Messages vindow
'   ac(69).fvirt = %FVIRTKEY OR %FCONTROL
'   ac(69).key   = %VK_F12
'   ac(69).cmd   = %IDM_RESTOREPROJECTWINDOW

   ' // Ctrl+Shift+J - GotCha
   ac(70).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(70).key   = ASC("J")
   ac(70).cmd   = %IDM_GOTCHA

   ' // Ctrl+Shift+Y - Code analyzer
   ac(71).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(71).key   = ASC("Y")
   ac(71).cmd   = %IDM_CODEC

   ' // Ctrl+Shift+U - Calculator
   ac(72).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(72).key   = ASC("U")
   ac(72).cmd   = %IDM_CALCULATOR

   ' // Ctrl+Alt+P - Page setup
   ac(73).fvirt = %FVIRTKEY OR %FCONTROL OR %FALT
   ac(73).key   = ASC("P")
   ac(73).cmd   = %IDM_PRINTSETTING'SETUP

   ' // Ctrl+Alt+O - Command prompt
   ac(74).fvirt = %FVIRTKEY OR %FCONTROL OR %FALT
   ac(74).key   = ASC("O")
   ac(74).cmd   = %IDM_OPENCMD '%IDM_DOS

   ' // Ctrl+Alt+Y - Primary source file
   ac(75).fvirt = %FVIRTKEY OR %FCONTROL OR %FALT
   ac(75).key   = ASC("Y")
   ac(75).cmd   = %IDM_SETPRIMARY '%IDM_PRIMARYSOURCE

'   ' // Ctrl+Alt+V - Visual designer
'   ac(76).fvirt = %FVIRTKEY OR %FCONTROL
'   ac(76).key   = %VK_F10
'   ac(76).cmd   = %IDM_VISDES

   ' // Ctrl+F9 - Poffs
   ac(77).fvirt = %FVIRTKEY OR %FCONTROL
   ac(77).key   = %VK_F9
   ac(77).cmd   = %IDM_POFFS

   ' // Alt+0...Alt+9 / Ctrl+0...Ctrl+9 - Macros
   ac(78).fvirt = %FVIRTKEY OR %FALT
   ac(78).key   = %VK_0
   ac(78).cmd   = %IDM_ALT0

   ac(79).fvirt = %FVIRTKEY OR %FALT
   ac(79).key   = %VK_1
   ac(79).cmd   = %IDM_ALT1

   ac(80).fvirt = %FVIRTKEY OR %FALT
   ac(80).key   = %VK_2
   ac(80).cmd   = %IDM_ALT2

   ac(81).fvirt = %FVIRTKEY OR %FALT
   ac(81).key   = %VK_3
   ac(81).cmd   = %IDM_ALT3

   ac(82).fvirt = %FVIRTKEY OR %FALT
   ac(82).key   = %VK_4
   ac(82).cmd   = %IDM_ALT4

   ac(83).fvirt = %FVIRTKEY OR %FALT
   ac(83).key   = %VK_5
   ac(83).cmd   = %IDM_ALT5

   ac(84).fvirt = %FVIRTKEY OR %FALT
   ac(84).key   = %VK_6
   ac(84).cmd   = %IDM_ALT6

   ac(85).fvirt = %FVIRTKEY OR %FALT
   ac(85).key   = %VK_7
   ac(85).cmd   = %IDM_ALT7

   ac(86).fvirt = %FVIRTKEY OR %FALT
   ac(86).key   = %VK_8
   ac(86).cmd   = %IDM_ALT8

   ac(87).fvirt = %FVIRTKEY OR %FALT
   ac(87).key   = %VK_9
   ac(87).cmd   = %IDM_ALT9

   ac(88).fvirt = %FVIRTKEY OR %FCONTROL
   ac(88).key   = %VK_0
   ac(88).cmd   = %IDM_CTRL0

   ac(89).fvirt = %FVIRTKEY OR %FCONTROL
   ac(89).key   = %VK_1
   ac(89).cmd   = %IDM_CTRL1

   ac(90).fvirt = %FVIRTKEY OR %FCONTROL
   ac(90).key   = %VK_2
   ac(90).cmd   = %IDM_CTRL2

   ac(91).fvirt = %FVIRTKEY OR %FCONTROL
   ac(91).key   = %VK_3
   ac(91).cmd   = %IDM_CTRL3

   ac(92).fvirt = %FVIRTKEY OR %FCONTROL
   ac(92).key   = %VK_4
   ac(92).cmd   = %IDM_CTRL4

   ac(93).fvirt = %FVIRTKEY OR %FCONTROL
   ac(93).key   = %VK_5
   ac(93).cmd   = %IDM_CTRL5

   ac(94).fvirt = %FVIRTKEY OR %FCONTROL
   ac(94).key   = %VK_6
   ac(94).cmd   = %IDM_CTRL6

   ac(95).fvirt = %FVIRTKEY OR %FCONTROL
   ac(95).key   = %VK_7
   ac(95).cmd   = %IDM_CTRL7

   ac(96).fvirt = %FVIRTKEY OR %FCONTROL
   ac(96).key   = %VK_8
   ac(96).cmd   = %IDM_CTRL8

   ac(97).fvirt = %FVIRTKEY OR %FCONTROL
   ac(97).key   = %VK_9
   ac(97).cmd   = %IDM_CTRL9

   ac(98).fvirt = %FVIRTKEY OR %FCONTROL OR %FSHIFT
   ac(98).key   = ASC("N")
   ac(98).cmd   = %IDM_INITASNEW

   FUNCTION = CreateAcceleratorTable(ac(0), UBOUND(ac) - LBOUND(ac) + 1)

END FUNCTION
'==============================================================================
'�����ȼ����ٱ�
'==============================================================================
FUNCTION MakeAccelerators() AS DWORD
  LOCAL tmpStr AS STRING
  LOCAL tmpStr1 AS STRING
  LOCAL tmpStr2 AS STRING
  LOCAL tmpInt  AS INTEGER
  LOCAL tmpCmd AS DWORD
  LOCAL tmpArr() AS STRING
  LOCAL ac() AS ACCELAPI
  LOCAL i AS INTEGER
  tmpStr=ReadConfig("Commands")
  IF tmpStr="" THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  REDIM tmpArr(1 TO PARSECOUNT(tmpStr,","))
  PARSE tmpStr,tmpArr(),","
  REDIM ac(1 TO UBOUND(tmpArr()))
  FOR i=1 TO UBOUND(tmpArr())
    tmpStr=ReadConfig(tmpArr(i))
    SELECT CASE tmpArr(i)
      CASE "New"
        tmpCmd=%IDM_NEWBAS
      CASE "Open"
        tmpCmd=%IDM_OPEN
      CASE "Save"
        tmpCmd=%IDM_SAVE
      CASE "Find"
        tmpCmd=%IDM_FIND
      CASE "Find Next"
        tmpCmd=%IDM_FINDNEXT
      CASE "Replace"
        tmpCmd=%IDM_REPLACE
      CASE "Goto Line"
        tmpCmd=%IDM_GOTOLINE
      CASE "Cut"
        tmpCmd=%IDM_CUT
      CASE "Copy"
        tmpCmd=%IDM_COPY
      CASE "paste"
        tmpCmd=%IDM_PASTE
      CASE "Delete"
        tmpCmd=%IDM_DELETE
      CASE "Select All"
        tmpCmd=%IDM_SELALL
      CASE "Comment"
        tmpCmd=%IDM_COMMENT
      CASE "UnComment"
        tmpCmd=%IDM_UNCOMMENT
      CASE "Indent"
        tmpCmd=%IDM_INDENT
      CASE "Outdent"
        tmpCmd=%IDM_OUTDENT
'      CASE "Space Indent"
'        tmpCmd=%IDM_SPACEINDENT
'      CASE "Space Outdent"
'        tmpCmd=%IDM_SPACEOUTDENT
    END SELECT
    tmpStr=TRIM$(tmpStr)
    IF tmpStr="" THEN
      ITERATE FOR
    END IF
    tmpInt=PARSECOUNT(tmpStr," ")
    IF tmpInt=0 OR tmpInt>3 THEN
      ITERATE FOR
    END IF
    IF tmpInt=1 THEN
      IF tmpStr="Ctrl" OR tmpStr="Alt" OR tmpStr="Shift" THEN
        ITERATE FOR
      ELSE
        ac(i).fvirt=%FVIRTKEY
        SELECT CASE MCASE$(tmpStr)
          CASE "F1"
            ac(i).key=%VK_F1
          CASE "F2"
            ac(i).key=%VK_F2
          CASE "F3"
            ac(i).key=%VK_F3
          CASE "F4"
            ac(i).key=%VK_F4
          CASE "F5"
            ac(i).key=%VK_F5
          CASE "F6"
            ac(i).key=%VK_F6
          CASE "F7"
            ac(i).key=%VK_F7
          CASE "F8"
            ac(i).key=%VK_F8
          CASE "F9"
            ac(i).key=%VK_F9
          CASE "F10"
            ac(i).key=%VK_F10
          CASE "F11"
            ac(i).key=%VK_F11
          CASE "F12"
            ac(i).key=%VK_F12
          CASE "1"
            ac(i).key=%VK_1
          CASE "2"
            ac(i).key=%VK_2
          CASE "3"
            ac(i).key=%VK_3
          CASE "4"
            ac(i).key=%VK_4
          CASE "5"
            ac(i).key=%VK_5
          CASE "6"
            ac(i).key=%VK_6
          CASE "7"
            ac(i).key=%VK_7
          CASE "8"
            ac(i).key=%VK_8
          CASE "9"
            ac(i).key=%VK_9
          CASE "0"
            ac(i).key=%VK_0
          CASE "A"
            ac(i).key=%VK_A
          CASE "B"
            ac(i).key=%VK_B
          CASE "C"
            ac(i).key=%VK_C
          CASE "D"
            ac(i).key=%VK_D
          CASE "E"
            ac(i).key=%VK_E
          CASE "F"
            ac(i).key=%VK_F
          CASE "G"
            ac(i).key=%VK_G
          CASE "H"
            ac(i).key=%VK_H
          CASE "I"
            ac(i).key=%VK_I
          CASE "J"
            ac(i).key=%VK_J
          CASE "K"
            ac(i).key=%VK_K
          CASE "L"
            ac(i).key=%VK_L
          CASE "M"
            ac(i).key=%VK_M
          CASE "N"
            ac(i).key=%VK_N
          CASE "O"
            ac(i).key=%VK_O
          CASE "P"
            ac(i).key=%VK_P
          CASE "Q"
            ac(i).key=%VK_Q
          CASE "R"
            ac(i).key=%VK_R
          CASE "S"
            ac(i).key=%VK_S
          CASE "T"
            ac(i).key=%VK_T
          CASE "U"
            ac(i).key=%VK_U
          CASE "V"
            ac(i).key=%VK_V
          CASE "W"
            ac(i).key=%VK_W
          CASE "X"
            ac(i).key=%VK_X
          CASE "Y"
            ac(i).key=%VK_Y
          CASE "Z"
            ac(i).key=%VK_Z
          CASE "Space"
            ac(i).key=%VK_SPACE
          CASE "Up"
            ac(i).key=%VK_UP
          CASE "Down"
            ac(i).key=%VK_DOWN
          CASE "Left"
            ac(i).key=%VK_LEFT
          CASE "Right"
            ac(i).key=%VK_RIGHT
        END SELECT
      END IF
    END IF
    IF PARSECOUNT(tmpStr," ")=3 THEN
      tmpStr1=PARSE$(tmpStr," ",1)
      tmpStr2=PARSE$(tmpStr," ",3)
      SELECT CASE MCASE$(tmpStr1)
        CASE "Ctrl"
          ac(i).fvirt=%FVIRTKEY OR %FCONTROL
        CASE "Shift"
          ac(i).fvirt=%FVIRTKEY OR %FSHIFT
        CASE "Alt"
          ac(i).fvirt=%FVIRTKEY OR %FALT
      END SELECT
      SELECT CASE MCASE$(tmpStr2)
        CASE "F1"
          ac(i).key=%VK_F1
        CASE "F2"
          ac(i).key=%VK_F2
        CASE "F3"
          ac(i).key=%VK_F3
        CASE "F4"
          ac(i).key=%VK_F4
        CASE "F5"
          ac(i).key=%VK_F5
        CASE "F6"
          ac(i).key=%VK_F6
        CASE "F7"
          ac(i).key=%VK_F7
        CASE "F8"
          ac(i).key=%VK_F8
        CASE "F9"
          ac(i).key=%VK_F9
        CASE "F10"
          ac(i).key=%VK_F10
        CASE "F11"
          ac(i).key=%VK_F11
        CASE "F12"
          ac(i).key=%VK_F12
        CASE "1"
          ac(i).key=%VK_1
        CASE "2"
          ac(i).key=%VK_2
        CASE "3"
          ac(i).key=%VK_3
        CASE "4"
          ac(i).key=%VK_4
        CASE "5"
          ac(i).key=%VK_5
        CASE "6"
          ac(i).key=%VK_6
        CASE "7"
          ac(i).key=%VK_7
        CASE "8"
          ac(i).key=%VK_8
        CASE "9"
          ac(i).key=%VK_9
        CASE "0"
          ac(i).key=%VK_0
        CASE "A"
          ac(i).key=%VK_A
        CASE "B"
          ac(i).key=%VK_B
        CASE "C"
          ac(i).key=%VK_C
        CASE "D"
          ac(i).key=%VK_D
        CASE "E"
          ac(i).key=%VK_E
        CASE "F"
          ac(i).key=%VK_F
        CASE "G"
          ac(i).key=%VK_G
        CASE "H"
          ac(i).key=%VK_H
        CASE "I"
          ac(i).key=%VK_I
        CASE "J"
          ac(i).key=%VK_J
        CASE "K"
          ac(i).key=%VK_K
        CASE "L"
          ac(i).key=%VK_L
        CASE "M"
          ac(i).key=%VK_M
        CASE "N"
          ac(i).key=%VK_N
        CASE "O"
          ac(i).key=%VK_O
        CASE "P"
          ac(i).key=%VK_P
        CASE "Q"
          ac(i).key=%VK_Q
        CASE "R"
          ac(i).key=%VK_R
        CASE "S"
          ac(i).key=%VK_S
        CASE "T"
          ac(i).key=%VK_T
        CASE "U"
          ac(i).key=%VK_U
        CASE "V"
          ac(i).key=%VK_V
        CASE "W"
          ac(i).key=%VK_W
        CASE "X"
          ac(i).key=%VK_X
        CASE "Y"
          ac(i).key=%VK_Y
        CASE "Z"
          ac(i).key=%VK_Z
        CASE "Space"
          ac(i).key=%VK_SPACE
        CASE "Up"
          ac(i).key=%VK_UP
        CASE "Down"
          ac(i).key=%VK_DOWN
        CASE "Left"
          ac(i).key=%VK_LEFT
        CASE "Right"
          ac(i).key=%VK_RIGHT
      END SELECT
    END IF
    IF tmpCmd>0 THEN
      ac(i).cmd   = tmpCmd
    END IF
  NEXT i
  FUNCTION=CreateAcceleratorTable(ac(0), UBOUND(ac))
END FUNCTION
FUNCTION GetTemplate()AS STRING
  LOCAL Listing$()
  LOCAL x&
  LOCAL temp$
  LOCAL rsStr AS STRING
  LOCAL tmpStr AS STRING
  LOCAL lineStr AS STRING
  LOCAL i AS INTEGER
  LOCAL fn AS LONG
  REDIM Listing$(19)
  temp$ = DIR$("*.pbtpl")
  WHILE LEN(temp$) AND x& < 1000 ' max = 1000
    Listing$(x&)=temp$
    INCR x&
    temp$ = DIR$(NEXT)
  WEND
  DIR$ CLOSE
  REDIM PRESERVE Listing$(x&-1)
  FOR i=0 TO UBOUND(Listing$())
    fn=FREEFILE
    OPEN Listing(i) FOR INPUT AS #fn
    LINE INPUT #fn,tmpStr
    LINE INPUT #fn,tmpStr
    LINE INPUT #fn,tmpStr
    CLOSE #fn
    IF RIGHT$(tmpStr,4)="ģ��" THEN
      tmpStr=MID$(tmpStr,1,LEN(tmpStr)-4)
    END IF
    rsStr=rsStr & tmpStr & "," & Listing$(i) & $CRLF
  NEXT i
  IF LEN(rsStr)>=3 THEN
    rsStr=MID$(rsStr,1,LEN(rsStr)-2)
  END IF
  FUNCTION=rsStr
END FUNCTION
'�����˵�
FUNCTION MakeMenu (BYVAL hWnd AS DWORD) AS DWORD
  '------------------------------------------------------------------------------
  ' �����˵�
  '----------------------------------------------------------------------------
  LOCAL hMnu        AS DWORD
  LOCAL hSubMenu    AS DWORD
  LOCAL hSubMenu1   AS DWORD
  LOCAL hSubMenu2   AS DWORD
  LOCAL hRebarMenu  AS DWORD
  LOCAL hBlockMenu  AS DWORD
  LOCAL tempArr()   AS STRING
  LOCAL i           AS LONG
  LOCAL tmpMII      AS MENUITEMINFO
  LOCAL tmpLng      AS LONG
  tmpMII.cbsize=44 'sizeof(tmpMII)
  tmpMII.fMask=%MIIM_ID
  tmpMII.fState=%MFS_DEFAULT
  'tmpMII.fType=%MFT_OWNERDRAW OR %MFT_STRING
  'MENU NEW BAR TO hMnu
  g_hMenu = CreateMenu
  '�ļ��˵�
  g_hMenuFile=CreatePopUpMenu
    AppendMenu g_hMenu,%MF_POPUP OR %MF_ENABLED,g_hMenuFile,"�ļ�"
    'MENU ADD POPUP, g_hMenu, "�ļ�", g_hMenuFile, %MF_enabled
    hSubMenu1=CreatePopUpMenu
    AppendMenu g_hMenuFile,%MF_OWNERDRAW OR %MF_POPUP OR %MF_ENABLED,hSubMenu1,GetMenuTxtBmp(%IDM_NEW,0)
    tmpMII.wID=%IDM_NEW
    tmpLng=SetMenuItemInfo(g_hMenuFile,0,1,tmpMII)
      AppendMenu hSubMenu1, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_NEWPROJECT, GetMenuTxtBmp(%IDM_NEWPROJECT, 0)
      AppendMenu hSubMenu1, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_NEWBAS, GetMenuTxtBmp(%IDM_NEWBAS, 0)
      AppendMenu hSubMenu1, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_NEWINC, GetMenuTxtBmp(%IDM_NEWINC, 0)
      AppendMenu hSubMenu1, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_NEWRC, GetMenuTxtBmp(%IDM_NEWRC, 0)
      pbtempfile = GetTemplate()
      REDIM tempArr(PARSECOUNT(pbtempfile,$CRLF)-1)
      PARSE pbtempfile,tempArr(),$CRLF
      FOR i=0 TO UBOUND(tempArr())
        AppendMenu hSubMenu1,%MF_OWNERDRAW OR %MF_ENABLED,%IDM_PBTEMPLATES+i+1,MID$(tempArr(i),1,INSTR(tempArr(i),",")-1)
      NEXT i
    AppendMenu g_hMenuFile,%MF_OWNERDRAW OR %MF_ENABLED,  %IDM_OPEN, "��" + $TAB + "Ctrl+O"
    AppendMenu g_hMenuFile,%MF_OWNERDRAW OR %MF_ENABLED,  %IDM_INSERTFILE, "�����ļ�..."
    g_hMenuReopen = CreatePopUpMenu
    AppendMenu g_hMenuFile,%MF_OWNERDRAW OR %MF_POPUP OR %MF_ENABLED,g_hMenuReopen,"�ش�"
    tmpMII.wID=%IDM_REOPEN
    tmpLng=SetMenuItemInfo(g_hMenuFile,3,1,tmpMII)
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_ENABLED,     %IDM_OPENPROJECT,"����Ŀ"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_SEPARATOR, 0, ""
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_GRAYED,     %IDM_SAVE,"����" + $TAB + "Ctrl+S"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_GRAYED,     %IDM_SAVEAS,"���Ϊ"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_GRAYED,     %IDM_SAVEALL,"ȫ����"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_GRAYED,     %IDM_SAVEPROJECT,"������Ŀ"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_SEPARATOR, 0, ""
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_GRAYED,     %IDM_PRINTSETTING,"��ӡ����"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_GRAYED,     %IDM_PRINTPREVIEW,"��ӡԤ��"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_GRAYED,     %IDM_PRINT,"��ӡ"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_SEPARATOR, 0, ""
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_ENABLED,     %IDM_CLOSE,"�ر��ļ�"
    'AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_ENABLED,     %IDM_CLOSEOTHERS,"�ر�����"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_ENABLED,     %IDM_CLOSEALL,"ȫ�ر�"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_SEPARATOR, 0, ""
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_ENABLED,     %IDM_OPENCMD,"��������"
    AppendMenu g_hMenuFile, %MF_OWNERDRAW OR %MF_ENABLED,     %IDM_EXIT,"�˳�" + $TAB + "Alt+F4"
    '�༭�˵�
    g_hMenuEdit = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_GRAYED,g_hMenuEdit,"�༭"
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_UNDO, ""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_REDO, ""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_CUT,   ""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_COPY,  ""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_PASTE, ""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_PASTEIE, ""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_SELALL, ""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_DELETE,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_LINEDELETE,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_INITASNEW,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_CLEARALL,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_COMMENT,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_UNCOMMENT,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_INDENT,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_OUTDENT,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_FORMATREGION,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_TABULATEREGION,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_SELTOUPPERCASE,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_SELTOLOWERCASE,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_SELTOMIXEDCASE,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_TEMPLATES,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_HTMLCODE,""
      AppendMenu g_hMenuEdit, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_GUID,""
    '�����˵�
    g_hMenuSearch = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED,g_hMenuSearch,"����"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FIND,"����" + $TAB + "Ctrl+F"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FINDNEXT,"��һ��" + $TAB + "F3"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FINDBACKWARDS,"��һ��" + $TAB + "Shift+F3"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_REPLACE,"�滻" + $TAB + "Ctrl+H"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_GOTOLINE,"ת����" + $TAB + "Ctrl+G"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_TOGGLEBOOKMARK,"���ñ�ǩ" + $TAB + "Ctrl+F2"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_NEXTBOOKMARK,"��һ��ǩ" + $TAB + "F2"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PREVIOUSBOOKMARK,"��һ��ǩ" + $TAB + "Shift+F2"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_DELETEBOOKMARKS,"�Ƴ���ǩ" + $TAB + "Ctrl+Shift+F2"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FINDINFILES,"�ļ����ݲ���" + $TAB + "Ctrl+Shift+F"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_EXPLORER,"��Դ������" + $TAB + "Ctrl+Shift+O"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_WINDOWSFIND,"ϵͳ����" + $TAB + "Ctrl+Alt+F"
      AppendMenu g_hMenuSearch, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FILEFIND,"�����ļ�..." + $TAB + "Alt+Shift+F"
    '���в˵�
    g_hMenuRun = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED,g_hMenuRun,"����"
      AppendMenu g_hMenuRun, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_COMPILE,"����" + $TAB + "Ctrl+F5"
      AppendMenu g_hMenuRun, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_COMPILERUN,"���벢ִ��" + $TAB + "F5"
      AppendMenu g_hMenuRun, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_COMPILEDEBUG,"���벢����" + $TAB + "F6"
      AppendMenu g_hMenuRun, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_EXECUTE,"ִ��" + $TAB +"Ctrl+E"
      AppendMenu g_hMenuRun, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuRun, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SETPRIMARY,"������Դ�ļ�" + $TAB + "Ctrl+Alt+Y"
      AppendMenu g_hMenuRun, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CMDPARAMETER,"���в���"
      AppendMenu g_hMenuRun, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_DEBUGTOOL,"���Թ���"
    '��ͼ�˵�
    g_hMenuView = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED,g_hMenuView,"��ͼ"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PROPERTY,"����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CONTROLS,"�ؼ�"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PROJECT,"��Ŀ"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_COMPILERS,"������"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FINDRS,"���ҽ��"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      hRebarMenu = MakeRebarMenu
      ghRebarMenu=hRebarMenu
      AppendMenu g_hMenuView,%MF_OWNERDRAW OR %MF_POPUP OR %MF_ENABLED,ghRebarMenu,"������"
      tmpMII.wID=%IDM_REBAR
      tmpLng=SetMenuItemInfo(g_hMenuView,6,1,tmpMII)
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_TOGGLE,"��չ����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_TOGGLEALL,"�����չ"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FOLDALL,"ȫ������"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_EXPANDALL,"ȫ��չ��"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_ZOOMIN,"��С����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_ZOOMOUT,"�Ŵ����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_USETABS,"ʹ���Ʊ��"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_AUTOINDENT,"�Զ�����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SHOWLINENUM,"��ʾ����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SHOWMARGIN,"��ʾ�հ�"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SHOWINDENT,"��ʾ����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SHOWSPACES,"��ʾ�ո�"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SHOWEOL,"��ʾ��ĩ"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SHOWEDGE,"��ʾ��Ե"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SHOWPROCNAME,"��ʾ������"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CVEOLTOCRLF,"תΪCRLF����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CVEOLTOCR,"תΪCR����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CVEOLTOLF,"תΪLF����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_REPLSPCWITHTABS,"�ո�תΪTAB"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_REPLTABSWITHSPC,"TABתΪ�ո�"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FILEPROPERTIES,"�ļ�����"
      AppendMenu g_hMenuView, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SYSINFO,"ϵͳ��Ϣ"

    '���ڲ˵�
    g_hMenuWindow = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED, g_hMenuWindow,"����"
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CASCADE,"�������"
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_TILEH,"����"
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_TILEV,"����"
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_RESTOREWSIZE,"���ش��ڳߴ�"
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SWITCHWINDOW,"ת������"
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_ARRANGE,"����ͼ��"
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CLOSEWIN,"�ر�"
      AppendMenu g_hMenuWindow, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_MAXWIN,"���"
      'MENU ADD STRING, hSubMenu, "Option..."                        , %IDM_OPTION         , %MF_ENABLED

    '��Ŀ�˵�
    g_hMenuProject = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED, g_hMenuProject,"��Ŀ"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_NEWPROJECT,"����Ŀ"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_OPENPROJECT,"����Ŀ"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_LOADPROJECT,"������Ŀ"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SAVEPROJECT,"������Ŀ"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SAVEPROJECTAS,"�����Ŀ"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CLOSEPROJECT,"�ر���Ŀ"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_REGPROJECTEXT,"������Ŀ�ļ�"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_ACHIVE,"�鵵��Ŀ"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_NEWVISION,"��Ŀ����"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FILEVISIONSETTING,"�ļ��汾����"
      AppendMenu g_hMenuProject, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_VISIONPACKAGE,"��Ŀ�汾"

    g_hMenuOptions = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED, g_hMenuOptions, "����"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_INTERFACEOPT,"��������"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FILEEXTOPT,"�����ļ�"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_EDITOROPT,"����༭"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_COLORSOPT,"��ɫ����"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FORMOROPT,"����༭"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_COMPILEROPT,"��������"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_TOOLSOPT,"��������"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FOLDINGOPT,"��������"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PRINTOPT,"��ӡ����"
      AppendMenu g_hMenuOptions, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SHORTKEYOPT,"�ȼ�����"

    g_hMenuTool = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED,g_hMenuTool , "����"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_MAKEMENU,"�����˵�"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_MAKETOOLBAR,"����������"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_GOTCHA,"��ȷ����"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FINDMULTIDIR,"��Ŀ¼����"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_REPLACEMULTIDIR,"��Ŀ¼�滻"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_FORMATCODE,"��ʽ������"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CODEC,"��������"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CODECOUNT,"����ͳ��"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CODEMANAGER,"�������"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CODEKEEPER,"���뱣��"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_KBDMACROS,"�����ݼ�"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_INCLEAN,"�������"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_WINSPY,"SPY����"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CHARMAP,"�ַ���"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_MSGBOXDESIGNER,"��Ϣ�����"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PBFORMS,"PB����"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PBCOMBR,"PB COM���"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_TYPELIBBR,"PB LIB���"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_IMGEDITOR,"�༭ͼƬ"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_RCEDITOR,"�༭��Դ"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CODETIPSBUILDER,"��ʾ����"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CODETYPEBUILDER,"���͹���"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_DLGEDITOR,"�Ի���༭"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_SEPARATOR,0,""
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_POFFS,"Poffs"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_COPYCAT,"CopyCat"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CALCULATOR,"������"
      AppendMenu g_hMenuTool, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_MORETOOLS,"����..."

    g_hMenuCoop = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED,g_hMenuCoop , "Эͬ"
      AppendMenu g_hMenuCoop, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_COOPSETTING,"Эͬ����"
      AppendMenu g_hMenuCoop, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SENDINVITE,"����Эͬ"
      AppendMenu g_hMenuCoop, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_SENDVISIT,"�μ�Эͬ"
      AppendMenu g_hMenuCoop, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_DOWNLOADFILE,"�����ļ�"
      AppendMenu g_hMenuCoop, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_UPLOADFILE,"�ϴ��ļ�"
      AppendMenu g_hMenuCoop, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_CHAT,"�Ự"

    g_hMenuHelp = CreatePopUpMenu
    AppendMenu g_hMenu, %MF_POPUP OR %MF_ENABLED, g_hMenuHelp, "����"
      AppendMenu g_hMenuHelp, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_IDEHELP,"IDE ����"
      AppendMenu g_hMenuHelp, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PBWINCONTENT,"PB/WIN ����"
      AppendMenu g_hMenuHelp, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PBWININDEX,"PB/WIN ����"
      AppendMenu g_hMenuHelp, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_PBWINSEARCH,"PB/WIN ��ѯ"
      AppendMenu g_hMenuHelp, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_WINDOWSSDK,"WinSDK ����"
      AppendMenu g_hMenuHelp, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_ABOUT,"����"
      AppendMenu g_hMenuHelp, %MF_OWNERDRAW OR %MF_ENABLED,%IDM_UPDATE,"����"
  SetMenu hWnd, g_hMenu
END FUNCTION
' ���Ʋ��봰����ʾ/���ز˵�������Ӧ
FUNCTION On_CommandMenuDockWin(BYVAL wID AS DWORD,BYVAL menuId AS DWORD)AS LONG ' %ID_PROPERTY,lo(word,wParam)
  LOCAL i AS LONG
  i =FindGdiIndex(wID)
  IF i<0 THEN
    REDIM PRESERVE gdi(UBOUND(gdi())+1)
    i=UBOUND(gdi())
    InitGdiF(i,wID)
    SELECT CASE wID
      CASE %ID_PROPERTY
        CreatePropertyWin i ',hWnd
      CASE %ID_CONTROLS
        CreateToolboxWin i
      CASE %ID_PROJECT
        CreateProjectWin i
      CASE %ID_COMPILERS
        CreateCompileRsWin i
      CASE %ID_FINDRS
        CreateFindRsWin i
    END SELECT
    SendMessage gdi(i).hWnd,%WM_SIZE,0,0
    EXIT FUNCTION
  END IF
  IF IsWindowVisible(gdi(i).hWnd) THEN
    DockWinHide i
    gdi(i).l="I"
    gdi(i).hWnd=gdi(i).hWndF
    MENU SET STATE g_hMenuView,BYCMD menuId,%MF_UNCHECKED OR %MF_ENABLED
  ELSE
    DockWinShow i
    gdi(i).hWnd=gdi(i).hWndF
    gdi(i).ml="F"
    gdi(i).l="F"
    gdi(i).tid=0
    gdi(i).byside=0
    SendMessage gdi(i).hWnd,%WM_SIZE,0,0
    MENU SET STATE g_hMenuView,BYCMD menuId,%MF_CHECKED OR %MF_ENABLED
  END IF
END FUNCTION
'==============================================================================
FUNCTION OpenThisFile(BYVAL fn AS STRING) AS DWORD
  '------------------------------------------------------------------------------
  ' ���ļ�����
  '----------------------------------------------------------------------------
  LOCAL hMdi AS DWORD, zText AS ASCIIZ * %MAX_PATH

  hMdi = GetWindow(g_hWndClient, %GW_CHILD) '���ȼ���Ƿ��Ѿ�����
  WHILE hMdi
      GetWindowText hMdi, zText, %MAX_PATH
      IF UCASE$(zText) = UCASE$(fn) THEN                     '����Ѿ���
          SendMessage g_hWndClient, %WM_MDIACTIVATE, hMdi, 0  '������
          WriteRecentFiles fn                                 '���� MRU �˵����˳�����
          EXIT FUNCTION
      END IF
      hMdi = GetWindow(hMdi, %GW_HWNDNEXT)
  WEND
  '------------------------------------------------------------------------------
  IF MdiGetActive(g_hWndClient) AND _      '������ǵ�һ���ĵ����ĵ�����󻯵ģ����л��˵������ػ�
     IsZoomed(MdiGetActive(g_hWndClient)) THEN g_FreezeMenu = 1

  hMdi = CreateMdiChild($EDITCLASSNAME, g_hWndClient, fn, %WS_MAXIMIZE)

  IF g_FreezeMenu THEN          '����˵������Ѿ��ر���
      g_FreezeMenu = 0          '���ñ�ʶreset flag
      DrawMenuBar g_hWndMain    '���ػ�˵�
  END IF

  FUNCTION = hMdi
END FUNCTION
' *********************************************************************************************
' Save the list of recently opened files
' *********************************************************************************************
SUB WriteRecentFiles (BYVAL OpenFName AS STRING)

   LOCAL Ac        AS LONG
   LOCAL szText    AS ASCIIZ * %MAX_PATH
   LOCAL szSection AS ASCIIZ * 30
   LOCAL szKey     AS ASCIIZ * 30

   IF INSTR(OpenFName, ANY ":\/") = 0 THEN   ' Path not available?
      IF LEFT$(UCASE$(GetFileName(OpenFName)), 8) = "UNTITLED" THEN EXIT SUB
   END IF

   szSection = "Reopen files"

   IF LEN(OpenFName) THEN
      ARRAY SCAN RecentFiles(), COLLATE UCASE, = UCASE$(OpenFName), TO Ac
      IF Ac THEN ARRAY DELETE RecentFiles(Ac)
      ARRAY INSERT RecentFiles(), OpenFName
   END IF

   FOR Ac = 1 TO 8
      szText = RecentFiles(Ac)
      IF ISFALSE FileExist(szText) THEN szText = ""
      szKey   = "File " & FORMAT$(Ac)
      WritePrivateProfileString szSection, szKey, szText, g_zIni
   NEXT

   GetRecentFiles ' update MRU menu

END SUB
'SUB WriteRecentFiles(BYVAL OpenFName AS STRING)
'  LOCAL fileInDB AS STRING
'  LOCAL recentFileNumber AS INTEGER
'  LOCAL i AS INTEGER
'  IF OpenFName="" THEN EXIT SUB
'  recentFileNumber=VAL(ReadConfig("recentfilenumber"))
'  IF recentFileNumber=0 THEN
'    EXIT SUB
'  END IF
'  FOR i=1 TO recentFileNumber
'    fileInDB=ReadConfig("recentfile" & FORMAT$(i))
'    IF fileInDB=OpenFName THEN EXIT SUB
'  NEXT i
'  FOR i=recentFileNumber-1 TO 1  STEP -1
'    fileInDB=ReadConfig("recentfile" & FORMAT$(i))
'    WriteConfig("recentfile" & FORMAT$(i+1),fileInDB)
'  NEXT i
'  WriteConfig("recentfile1",OpenFName )
'  GetRecentFiles
'END SUB
'SUB LoadDefaultLayout(BYVAL hWnd AS DWORD)
'  hToolWin=CreatetoolWin(hWnd)
'END SUB
'SUB LoadSettingLayout(BYVAL hWnd AS DWORD)
'  LOCAL i AS LONG
'  FOR i=0 TO UBOUND(dockArr())
'    SELECT CASE dockArr(i).thisWin.wID
'      CASE %ID_TOOLWINDOW
'        hToolWin = CreateToolWin(hWnd)
'      CASE %ID_PROJECTWINDOW
'
'      CASE %ID_PROPERTYWINDOW
'
'      CASE %ID_COMPILEWINDOW
'
'      CASE %ID_FINDRESULTWINDOW
'
'      CASE %ID_CLIENTWINDOW
'
'    END SELECT
'  NEXT i
'  FOR i=1 TO UBOUND(nSize())
'    IF nSize(i).thisWin.hWnd>0 THEN
'      nSize(i).thisWin.hWnd=CreateDivSize()
'      SetWindowLong nSize(i).thisWin.hWnd, %GWL_USERDATA, VARPTR(nSize(i))
'      IF nSize(i).isVisible=0 THEN
'        showWindow nSize(i).thisWin.hWnd,%SW_HIDE
'      ELSE
'        MoveWindow nSize(i).thisWin.hWnd, nSize(i).rc.nLeft, nSize(i).rc.nTop, _
'                    nSize(i).rc.nRight-nSize(i).rc.nLeft, nSize(i).rc.nBottom-nSize(i).rc.ntop, %TRUE
'      END IF
'    END IF
'  NEXT i
'  Repaint g_hWndMain
'END SUB
FUNCTION CreateGuiderWindow(BYVAL hParent AS DWORD)AS LONG
  LOCAL tmpStr AS STRING
  LOCAL tmpRC  AS RECT
  LOCAL hDlg AS DWORD
  IF hParent<=0 THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  tmpStr=GetLang("New Project")
  DIALOG NEW hParent, tmpStr, , , 303, 195, _
              %WS_POPUP OR %WS_VISIBLE OR %WS_CLIPSIBLINGS OR %WS_CLIPCHILDREN OR _
              %WS_CAPTION OR %WS_SYSMENU OR %DS_3DLOOK OR %DS_SETFONT OR _
              %DS_MODALFRAME OR %DS_CONTEXTHELP, _
              %WS_EX_CONTEXTHELP OR %WS_EX_DLGMODALFRAME OR %WS_EX_CONTROLPARENT OR %WS_EX_WINDOWEDGE TO hDlg
  FUNCTION = hDlg
  'Note: in PB/WIN 7.0 and later, you can also use DIALOG SET ICON hDlg, newicon$
  'For icon from resource, use something like, LoadIcon(hInst, "APP_ICON")
  DIALOG SEND hDlg, %WM_SETICON, %ICON_SMALL, LoadIcon(%NULL, BYVAL %IDI_APPLICATION)
  DIALOG SEND hDlg, %WM_SETICON, %ICON_BIG, LoadIcon(%NULL, BYVAL %IDI_APPLICATION)

  InitCommonControls
  REDIM g_hGuiderTab(2)
  CONTROL ADD TAB, hDlg, %ID_GUIDERTAB, "", 4, 5, 295, 170
  CONTROL ADD CHECKBOX, hDlg, %IDC_GUILDSHOWCHECKBOX, GetLang("Don''t show this dialog again"), 4, 177, 303, 9     '������ʾ����Ի���
  tmpStr=ReadConfig("showguiderwindow")
  IF tmpStr="0" THEN
    CONTROL SET CHECK hDlg,%IDC_GUILDSHOWCHECKBOX,1
  END IF
  TAB INSERT PAGE hDlg, %ID_GUIDERTAB, 1, 0,GetLang("New"), CALL GuiderTabProc TO g_hGuiderTab(0)
  CONTROL ADD LISTVIEW, g_hGuiderTab(0), %IDC_GUILDNEWLIST, "", 6, 3, 279, 95, _
              %WS_CHILD OR %WS_VSCROLL OR _
              %WS_TABSTOP OR %LVS_SINGLESEL OR %LVS_SHOWSELALWAYS OR %LVS_AUTOARRANGE, _
              %WS_EX_CLIENTEDGE
  CONTROL ADD BUTTON, g_hGuiderTab(0), %IDOK, GetLang("Open"), 230, 103, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, g_hGuiderTab(0), %IDCANCEL, GetLang("Cancel"), 230, 120, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, g_hGuiderTab(0), %IDC_GUILDHELPBUTTON, GetLang("Help"), 230, 137, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY

  TAB INSERT PAGE hDlg, %ID_GUIDERTAB, 2, 0,GetLang("Exist"), CALL GuiderTabProc TO g_hGuiderTab(1)
  CONTROL ADD LABEL, g_hGuiderTab(1), %IDC_GUILDEXISTFINDLABEL, GetLang("Search scope:"),6, 7, 48, 9, _      '���ҷ�Χ:
              %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %SS_NOTIFY, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD COMBOBOX, g_hGuiderTab(1), %IDC_GUILDEXISTFINDCOMBOBOX, ,55, 4, 139, 98, _
              %WS_CHILD OR %WS_VISIBLE OR %CBS_DROPDOWNLIST OR %CBS_OWNERDRAWFIXED OR %CBS_HASSTRINGS, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD LISTVIEW, g_hGuiderTab(1), %IDC_GUILDEXISTFINDLIST,"" ,6, 23, 279, 76, _
              %WS_CHILD OR %LBS_NOTIFY OR %LBS_SORT OR %LBS_NOINTEGRALHEIGHT OR %LBS_MULTICOLUMN, _
              %WS_EX_CLIENTEDGE OR %WS_EX_NOPARENTNOTIFY
  CONTROL ADD LABEL, g_hGuiderTab(1), %IDC_GUILDEXISTFILENAMELABEL, GetLang("File name:"),6, 106, 48, 9, _  '�ļ���(&N):
              %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %SS_NOTIFY OR %SS_RIGHT, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD TEXTBOX, g_hGuiderTab(1), %IDC_GUILDEXISTFILENAMETEXTBOX, "",60, 105, 160, 12, _
              %WS_CHILD OR %WS_VISIBLE OR %ES_AUTOHSCROLL,%WS_EX_CLIENTEDGE OR %WS_EX_NOPARENTNOTIFY
  CONTROL ADD LABEL, g_hGuiderTab(1), %IDC_GUILDEXISTFILETYPELABEL, GetLang("File type:"),6, 124, 48, 9, _ '�ļ�����(&T):
              %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %SS_NOTIFY OR %SS_RIGHT,%WS_EX_NOPARENTNOTIFY
  CONTROL ADD COMBOBOX, g_hGuiderTab(1), %IDC_GUILDEXISTFILETYPECOMBOBOX, ,60, 123, 160, 98, _
              %WS_CHILD OR %WS_VISIBLE OR %CBS_DROPDOWNLIST OR %CBS_HASSTRINGS, %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, g_hGuiderTab(1), %IDOK, GetLang("Open"), 230, 103, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, g_hGuiderTab(1), %IDCANCEL, GetLang("Cancel"), 230, 120, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, g_hGuiderTab(1), %IDC_GUILDHELPBUTTON, GetLang("Help"), 230, 137, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY

  TAB INSERT PAGE hDlg, %ID_GUIDERTAB, 3, 0,GetLang("Recent"), CALL GuiderTabProc TO g_hGuiderTab(2)
  CONTROL ADD BUTTON, g_hGuiderTab(2), %IDOK, GetLang("Open"), 230, 103, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD LISTVIEW, g_hGuiderTab(2), %IDC_GUILDRECENTLIST,"" ,6, 3, 279, 95, _
              %WS_CHILD OR %LBS_NOTIFY OR %LBS_SORT OR %LBS_NOINTEGRALHEIGHT OR %LBS_MULTICOLUMN, _
              %WS_EX_CLIENTEDGE OR %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, g_hGuiderTab(2), %IDCANCEL, GetLang("Cancel"), 230, 120, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, g_hGuiderTab(2), %IDC_GUILDHELPBUTTON, GetLang("Help"), 230, 137, 50, 14, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY

  DIALOG SHOW MODELESS hDlg 'call GuiderWinProc
END FUNCTION
'callback function GuiderWinProc
'  SELECT CASE AS LONG CB.MSG
'    case %wm_initdialog
'      msgbox "initdialog"
'    case %wm_create
'      msgbox "create"
'    case %wm_command
'      select case as long cb.ctl
'        case %IDOK
'          msgbox "ok"
'          dialog end cb.hndl,0
'        case %idcancel
'          msgbox "cancel"
'          dialog end cb.hndl,1
'      end select
'  end select
'end function
CALLBACK FUNCTION GuiderTabProc
  SELECT CASE AS LONG CB.MSG
    CASE %WM_COMMAND
      SELECT CASE AS LONG CB.CTL
        CASE %IDOK
          IF CB.CTLMSG = %BN_CLICKED OR CB.CTLMSG = 1 THEN
            SELECT CASE CB.HNDL 'GetParent(GetDlgItem(CB.HNDL,%IDOK))
              CASE g_hGuiderTab(0)
                ? "In Tab 0"
              CASE g_hGuiderTab(1)
                ? "In Tab 1"
              CASE g_hGuiderTab(2)
                ? "In Tab 2"
            END SELECT
          END IF
        CASE %IDCANCEL
          IF CB.CTLMSG = %BN_CLICKED OR CB.CTLMSG = 1 THEN
            'MSGBOX "Submit button on Tab 1 clicked", %MB_TASKMODAL
            DIALOG END g_hGuiderWin,1
          END IF
        CASE %IDC_GUILDHELPBUTTON
          IF CB.CTLMSG = %BN_CLICKED OR CB.CTLMSG = 1 THEN
            'MSGBOX "Submit button on Tab 1 clicked", %MB_TASKMODAL
          END IF
        CASE %IDC_GUILDEXISTFINDCOMBOBOX
          IF CB.CTLMSG = %CBN_SELCHANGE THEN
'            COMBOBOX GET SELECT CB.HNDL, %ID_TAB2_CBCHOICES TO i
'            COMBOBOX GET TEXT   CB.HNDL, %ID_TAB2_CBCHOICES, i TO s
            '? "You selected "+s
          END IF

        CASE %IDC_GUILDEXISTFILETYPECOMBOBOX
          IF CB.CTLMSG = %LBN_SELCHANGE THEN
'            LISTBOX GET SELECT CB.HNDL, %ID_TAB3_LBCHOICES TO i
'            LISTBOX GET TEXT   CB.HNDL, %ID_TAB3_LBCHOICES, i TO s
'            CONTROL SET TEXT   CB.HNDL, %ID_TAB3_TEXTBOX, s
          END IF
      END SELECT
  END SELECT
END FUNCTION
' *********************************************************************************************
' Insert file procedure
' *********************************************************************************************
SUB InsertFile (BYVAL hWnd AS DWORD)

  LOCAL hEdit    AS DWORD
  LOCAL PATH     AS STRING
  LOCAL f        AS STRING
  LOCAL dwStyle  AS DWORD
  LOCAL fOptions AS STRING
  LOCAL nFile    AS DWORD
  LOCAL Buffer   AS STRING

  hEdit = GetEdit
  IF ISFALSE hEdit THEN EXIT SUB

  PATH  = CURDIR$
  fOptions =          "PB �����ļ�(*.BAS)|*.BAS|"
  fOptions = fOptions & "PB ͷ�ļ�(*.INC)|*.INC|"
  fOptions = fOptions & "�����ļ�(*.*)|*.*"
  f     = "*.BAS;*.INC"
  dwStyle = %OFN_FILEMUSTEXIST OR %OFN_HIDEREADONLY OR %OFN_LONGNAMES
  IF ISFALSE OpenFileDialog(g_hWndMain, "", f, PATH, fOptions, "BAS", dwStyle) THEN EXIT SUB

  ' Open the file and read it
  TRY
    nFile = FREEFILE
    OPEN f FOR BINARY AS nFile
  CATCH
    MessageBox(hWnd, "Error" & STR$(ERR) & " inserting the file   ", _
       " InsertFile", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL)
    EXIT SUB
  END TRY
  GET$ nFile, LOF(nFile), Buffer
  CLOSE nFile
  ' Insert the text at the current position
  SendMessage hEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(Buffer)
  ' Change statusbar information
  ShowLinCol
END SUB
