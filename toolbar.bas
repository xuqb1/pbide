'toolbar.bas
'������
'------------------------------------------------------------------------------
' �����ͺ���.
'------------------------------------------------------------------------------
' ʹ��ReBar����ͼƬ.
%USE_RBIMAGE = 0 '%FALSE        ' ��Ϊ%TRUE��ʹ��ReBar����ͼƬ
' ʹ�õĹ�������ť��.
%TOOLBUTTONS = 12 '��׼�������İ�ť��
%MENUBUTTONS = 6
%WM_UPDATESIZEWINDOW  = %WM_USER +1024  'Size/Position window
%APP_REBARS           = %WM_USER + 3072
%APP_TOOLBARS         = %WM_USER + 4096
%APP_STATUSBARS       = %WM_USER + 5120
%APP_CONTROLS         = %WM_USER + 6144
' ReBar �� ToolBar ��ͬ. ����ʼ��3000��������...
%ID_REBAR             = %APP_REBARS
%ID_TOOLBAR           = %APP_TOOLBARS
%ID_WINDOWBAR         = %WM_USER + 4500
%ID_COMBOBOX          = %APP_CONTROLS
%ID_DEBUGBAR          = %WM_USER + 4700
%ID_CODEEDITBAR       = %WM_USER + 4800
%ID_WINARRAGBAR       = %WM_USER + 4900

'%ID_MENUBAR  = %APP_MENUBARS

' ��Щ�ǹ�������ťID.
%ID_NEW           = 101
%ID_OPEN          = 102
%ID_SAVE          = 103
%ID_PRINT         = 104
%ID_PRINTPREVIEW  = 105
%ID_CUT           = 106
%ID_COPY          = 107
%ID_PASTE         = 108
%ID_UNDOS         = 109
%ID_REDOS         = 110
%ID_UNDOC         = 111
%ID_REDOC         = 112
%ID_VIEW          = 113
%ID_EMAIL         = 114
%ID_FOLDERCLOSED  = 115
%ID_FOLDEROPEN    = 116
%ID_FOLDERS       = 117

%IDM_STDTOOLBAR   = %WM_USER + 2100& '��׼��������ʾ/���ز˵�����
%IDM_WINDOWBAR    = %WM_USER + 2101& '���ڹ�������ʾ/���ز˵�����
%IDM_COMBOBAR     = %WM_USER + 2102& '�����б�������ʾ/���ز˵�����
%IDM_DEBUGBAR     = %WM_USER + 2103& '�������/���й�������ʾ/���ز˵�����
%ID_LOCKBAR       = %WM_USER + 2104& '����/����������
%IDM_CODEEDIT     = %WM_USER + 2105& '����༭��������ʾ/���ز˵�����
%IDM_WINARRAG     = %WM_USER + 2016& '�������й�������ʾ/���ز˵�����

%I_IMAGENONE      = -2
%IMG_DIS = 0    ' ������Disabled
%IMG_NOR = 1    ' һ��״̬Normal
%IMG_HOT = 2    ' ѡ��״̬Selected
%OMENU_EXTRAWIDTH = 30    ' Extra width in pixels
%OMENU_CHECKEDICON = 106    ' Identifier of the checked icon
'------------------------------------------------------------------------------
' ȫ�ֱ���.
'------------------------------------------------------------------------------
GLOBAL hMenuTextBkBrush AS DWORD        ' // Submenu text brush          �Ӳ˵��ı�ˢ
GLOBAL hMenuHiBrush     AS DWORD        ' // Submenu highlighted text brush�Ӳ˵������ı�ˢ
GLOBAL hMenuIconsBrush      AS DWORD        ' // Brush for the icon's backgroundͼ�걳��ˢ

'= �˵��͹�����ͼƬ�б���.
' ��׼��������ť
GLOBAL ghImlHot    AS DWORD               ' ��ѡ��ͼƬ�б���.
GLOBAL ghImlDis    AS DWORD               ' ������ͼƬ�б���.
GLOBAL ghImlNor    AS DWORD               ' ����״̬ͼƬ�б���.
' ���ڹ�������ť
GLOBAL hImlHot    AS DWORD               ' ��ѡ��ͼƬ�б���.
GLOBAL hImlDis    AS DWORD               ' ������ͼƬ�б���.
GLOBAL hImlNor    AS DWORD               ' ����״̬ͼƬ�б���.

GLOBAL gBmpSize    AS LONG                ' λͼ�ߴ�.
'= Rebar ����ͼƬλͼ���.
GLOBAL ghRbBack    AS DWORD               ' Rebar �ؼ�����λͼ���.

'------------------------------------------------------------------------------
' ���� : AppLoadBitmaps ()
' ���� : ����Դ�ļ����س���λͼ.
' ע�� : �����Ҫ�ڹ�����ʹ��24x24λͼ����ͱ�����˵�����16x16λͼ��!
'------------------------------------------------------------------------------
FUNCTION AppLoadBitmaps() AS LONG
  LOCAL bm      AS BITMAP
  LOCAL hBmpHot AS DWORD
  LOCAL hBmpDis AS DWORD
  LOCAL hBmpNor AS DWORD
  '----------------------------------------------------------------------------
  ' ���ò���ʼ���˵���������λͼ��ͼƬ�б�.
  ' ��Ҫ : ��Ҫ�˵�λͼ��16x16�Ĵ������ӵ�����!
  '----------------------------------------------------------------------------
  ' ������λͼ.
  hBmpHot = LoadImage(g_hInst, "TBHOT", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  ' ������λͼ.
  hBmpDis = LoadImage(g_hInst, "TBDIS", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  ' ����ʱλͼ.
  hBmpNor = LoadImage(g_hInst, "TBNOR", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  '----------------------------------------------------------------------------
  ' Ϊ�Ժ�ʹ�ã���ȡ������λͼ�ߴ�.
  '----------------------------------------------------------------------------
  GetObject hBmpNor, SIZEOF(bm), bm           ' �õ�λͼ�ߴ�.
  'msgbox "hBmpHot=" & str$(hBmpHot)
  gBmpSize = bm.bmHeight                      ' ����λͼ�ߴ��Ա�����.
  'if gBmpSize=0 then gBmpSize=16
  '----------------------------------------------------------------------------
  ' �����˵��͹�����ͼƬ��ѡ�С������ü�����״̬�б�.
  ' ��Ҫ: ��Ҫ���˵�λͼΪ16x16�Ĵ������ӵ�����!
  '----------------------------------------------------------------------------
  ghImlHot = ImageList_Create(gBmpSize, gBmpSize, %ILC_COLORDDB OR %ILC_MASK, 1, 0)
       ImageList_AddMasked ghImlHot, hBmpHot, RGB(255,0,255)
       ImageList_Add ghImlHot,       hBmpHot, RGB(255,0,255)
  ghImlDis = ImageList_Create(gBmpSize, gBmpSize, %ILC_COLORDDB OR %ILC_MASK, 1, 0)
       ImageList_AddMasked ghImlDis, hBmpDis, RGB(255,0,255)
       ImageList_Add ghImlDis,       hBmpDis, RGB(255,0,255)
  ghImlNor = ImageList_Create(gBmpSize, gBmpSize, %ILC_COLORDDB OR %ILC_MASK, 1, 0)
       ImageList_AddMasked ghImlNor, hBmpNor, RGB(255,0,255)
       ImageList_Add ghImlNor,       hBmpNor, RGB(255,0,255)
  '----------------------------------------------------------------------------
  ' ����Rebar����ͼƬ.
  '----------------------------------------------------------------------------
  ghRbBack = LoadImage(g_hInst, "RBBACK", %IMAGE_BITMAP, 0, 0, _
             %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR _
             %LR_DEFAULTSIZE)
  '----------------------------------------------------------------------------
  ' �����ɾ��������Ҫ��λͼ���.
  '----------------------------------------------------------------------------
  IF hBmpHot THEN DeleteObject(hBmpHot)
  IF hBmpDis THEN DeleteObject(hBmpDis)
  IF hBmpNor THEN DeleteObject(hBmpNor)
  '----------------------------------------------------------------------------
  ' ���ò���ʼ���˵���������λͼ��ͼƬ�б�.
  ' ��Ҫ : ��Ҫ�˵�λͼ��16x16�Ĵ������ӵ�����!
  '----------------------------------------------------------------------------
  ' ������λͼ.
  hBmpHot = LoadImage(g_hInst, "WINTBHOT", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  ' ������λͼ.
  hBmpDis = LoadImage(g_hInst, "WINTBDIS", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  ' ����ʱλͼ.
  hBmpNor = LoadImage(g_hInst, "WINTBNOR", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  '----------------------------------------------------------------------------
  ' Ϊ�Ժ�ʹ�õõ�������λͼ�ߴ�.
  '----------------------------------------------------------------------------
'  GetObject hBmpNor, SIZEOF(bm), bm           ' �õ�λͼ�ߴ�.
'  'msgbox "hBmpHot=" & str$(hBmpHot)
'  gBmpSize = bm.bmHeight                      ' ����λͼ�ߴ��Ա�����.
'  'if gBmpSize=0 then gBmpSize=16
  '----------------------------------------------------------------------------
  ' �����˵��͹�����ͼƬ��ѡ�С������ü�����״̬�б�.
  ' ��Ҫ: ��Ҫ���˵�λͼΪ16x16�Ĵ������ӵ�����!
  '----------------------------------------------------------------------------
  hImlHot = ImageList_Create(gBmpSize, gBmpSize, %ILC_COLORDDB OR %ILC_MASK, 1, 0)
       ImageList_AddMasked hImlHot, hBmpHot, RGB(255,0,255)
       ImageList_Add ghImlHot,       hBmpHot, RGB(255,0,255)
  hImlDis = ImageList_Create(gBmpSize, gBmpSize, %ILC_COLORDDB OR %ILC_MASK, 1, 0)
       ImageList_AddMasked hImlDis, hBmpDis, RGB(255,0,255)
       ImageList_Add ghImlDis,       hBmpDis, RGB(255,0,255)
  hImlNor = ImageList_Create(gBmpSize, gBmpSize, %ILC_COLORDDB OR %ILC_MASK, 1, 0)
       ImageList_AddMasked hImlNor, hBmpNor, RGB(255,0,255)
       ImageList_Add ghImlNor,       hBmpNor, RGB(255,0,255)
  IF hBmpHot THEN DeleteObject(hBmpHot)
  IF hBmpDis THEN DeleteObject(hBmpDis)
  IF hBmpNor THEN DeleteObject(hBmpNor)
  FUNCTION = %TRUE
END FUNCTION
'------------------------------------------------------------------------------
' ���� : CreateToolBar ()
' ���� : �����������ؼ�.
'------------------------------------------------------------------------------
FUNCTION CreateToolBar (BYVAL hWnd AS DWORD) AS DWORD
  ' Ϊ�Ż��ٶ�ʹ��ע�����.
  REGISTER x AS LONG
  LOCAL c AS LONG
  ' ����������.
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP
  DIM Tbb(0 TO %TOOLBUTTONS - 1) AS LOCAL TBBUTTON
  ' �����ȷ��λͼ�ߴ�!
  IF gBmpSize <> 24 THEN
    IF gBmpSize <> 16 THEN
      ' �û�ʹ���˲���ȷ��λͼ�ߴ�.
      MessageBox(hWnd, "������λͼ�ߴ����Ϊ16x16 �� 24x24!" & $CRLF & STR$(gBmpSize), "������ʾ", _
             %MB_OK OR %MB_ICONERROR OR %MB_TOPMOST)
      FUNCTION = 1
      EXIT FUNCTION
    END IF
  END IF
  '----------------------------------------------------------------------------
  ' ��������������.
  '----------------------------------------------------------------------------
  ghToolBar = CreateWindowEx(0, _
               "ToolbarWindow32", _
               "", _
               %WS_CHILD OR %WS_VISIBLE OR %TBSTYLE_TOOLTIPS OR _
               %TBSTYLE_FLAT OR %TBSTYLE_LIST OR %TBSTYLE_TRANSPARENT OR _
               %WS_CLIPCHILDREN  OR %WS_CLIPSIBLINGS OR %CCS_NORESIZE OR _
               %CCS_NODIVIDER, _
               0, 0, 0, 0, _
               hWnd, _
               %ID_TOOLBAR, _
               g_hInst, _
               BYVAL %NULL)
  '----------------------------------------------------------------------------
  ' ��鴴���������ؼ�����ʱ�Ĵ���.
  '----------------------------------------------------------------------------
  IF ISFALSE ghToolBar THEN         ' ����!
    ' ��������������ʧ�ܣ�֪ͨ�û�.
    FUNCTION = %TRUE
    EXIT FUNCTION
  END IF
  '----------------------------------------------------------------------------
  ' ��ʼ����������ť��Tbb()����.
  '----------------------------------------------------------------------------
  ' �ð�ť��Ϣ��� TBBUTTON ����
  tbb(c).iBitmap   = 0'%STD_FILENEW     '�½�
  tbb(c).idCommand = %IDM_NEWBAS
  tbb(c).fsState   = %TBSTATE_ENABLED
  'tbb(c).fsStyle   = %TBSTYLE_BUTTON
  IF g_NewComCtl >= 4.7 THEN
    tbb(c).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb(c).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR c
  tbb(c).iBitmap   = 1'%STD_FILEOPEN    '��
  tbb(c).idCommand = %IDM_OPEN
  tbb(c).fsState   = %TBSTATE_ENABLED
  IF g_NewComCtl >= 4.7 THEN
    tbb(c).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb(c).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR c
  tbb(c).iBitmap   = 2'%STD_FILESAVE    '����
  tbb(c).idCommand = %IDM_SAVE
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
'  INCR c
'  tbb(c).iBitmap   = 108
'  tbb(c).idCommand = %IDM_REFRESH       'ˢ��
'  tbb(c).fsState   = %TBSTATE_ENABLED
'  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 0                  '�ָ���
  tbb(c).idCommand = 701
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c
  tbb(c).iBitmap   = 3'%STD_PRINT       '��ӡ
  tbb(c).idCommand = %IDM_PRINT
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 0                  '�ָ���
  tbb(c).idCommand = 702
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c
  tbb(c).iBitmap   = 5'%STD_CUT         '����
  tbb(c).idCommand = %IDM_CUT
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 6'%STD_COPY        '����
  tbb(c).idCommand = %IDM_COPY
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 7'%STD_PASTE       'ճ��
  tbb(c).idCommand = %IDM_PASTE
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 0                  '�ָ���
  tbb(c).idCommand = 703
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c
  tbb(c).iBitmap   = 9'%STD_UNDO        '����
  tbb(c).idCommand = %IDM_UNDO
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 8'%STD_REDO        ������
  tbb(c).idCommand = %IDM_REDO
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  '----------------------------------------------------------------------------
  ' ��ʼ�����������ڲ����ʹ�����Ϣ.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_TOOLBAR '%IDB_STD_LARGE_COLOR ' ʹ��λͼͼƬ��ID����, %ID_TOOLBAR
  Tabm.hInst = g_hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
                         ' ���ʹ��������Դ����λͼ
  ' ���ù�����λͼ�ߴ�Ϊ16x16.
  IF gBmpSize = 16 THEN
   SendMessage ghToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(16, 16)
   SendMessage ghToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(16, 16)
  END IF
  ' ���ù�����λͼ�ߴ�Ϊ24x24.
  IF gBmpSize = 24 THEN
   SendMessage ghToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(24, 24)
   SendMessage ghToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(24, 24)
  END IF
  ' ����ͼƬ�б�������.
  SendMessage ghToolBar, %TB_SETIMAGELIST,         0, ghImlNor  ' ����
  SendMessage ghToolBar, %TB_SETDISABLEDIMAGELIST, 0, ghImlDis  ' ������
  SendMessage ghToolBar, %TB_SETHOTIMAGELIST,      0, ghImlHot  ' ��ѡ��
  ' ����ͼƬ�б�λͼ����������ť.
  SendMessage ghToolBar, %TB_ADDBITMAP, %TOOLBUTTONS, VARPTR(Tabm)
  ' ���ù�������ť.
  SendMessage ghToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage ghToolBar, %TB_ADDBUTTONS, %TOOLBUTTONS, VARPTR(Tbb(0))
  ' ǿ�ƹ�����Ϊ��ʼ�ߴ�.
  IF ghToolBar THEN
    SendMessage ghToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' ͨ�ÿؼ��汾���£����ʹ�ø����ԣ���ƽ����ʽ��
       SetWindowLong ghToolBar, %GWL_STYLE, _
          GetWindowLong(ghToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage ghToolBar, %TB_SETEXTENDEDSTYLE, 0,%TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF
  FUNCTION = ghToolBar            ' ���ع��������ھ��.
END FUNCTION
' *********************************************************************************************
' ������    : CreateWindowBar
' ����      : �������ڹ������ؼ�
' ����      : ���������
' ����      : hWnd      = �����ھ��
'             lBmpSize  = λͼ��С�� 16x16 �� 24x24
' *********************************************************************************************
FUNCTION CreateWindowBar( BYVAL hWnd AS DWORD) AS DWORD
  ' ʹ�üĴ����������Ż��ٶ�
  REGISTER x AS LONG
  ' ����������
  LOCAL TOOLBUTTONS AS LONG
  LOCAL Tbb( ) AS TBBUTTON
  LOCAL Tabm AS TBADDBITMAP
  LOCAL hToolbar AS DWORD
  LOCAL a$
  LOCAL NewComCtl AS LONG
  TOOLBUTTONS = 12
  DIM Tbb( 0 TO TOOLBUTTONS-1) AS LOCAL TBBUTTON
  '----------------------------------------------------------------------------
  ' ��������������
  '----------------------------------------------------------------------------
  hToolbar = CreateWindowEx(0, _
               "ToolbarWindow32", _
               "", _
               %WS_CHILD OR %WS_VISIBLE OR %TBSTYLE_TOOLTIPS OR _
               %TBSTYLE_FLAT OR %TBSTYLE_LIST OR %TBSTYLE_TRANSPARENT OR _
               %WS_CLIPCHILDREN  OR %WS_CLIPSIBLINGS OR %CCS_NORESIZE OR _
               %CCS_NODIVIDER, _
               0, 0, 0, 0, _
               hWnd, _
               %ID_WINDOWBAR, _
               g_hInst, _
               BYVAL %NULL)
  IF ISFALSE hToolbar THEN EXIT FUNCTION
  ' ��ȡͨ�ÿؼ���汾
  NewComCtl = InitComctl32 (%ICC_BAR_CLASSES)
  IF NewComCtl>4.7 THEN
    SENDMESSAGE hToolbar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
  END IF
  '----------------------------------------------------------------------------
  ' ��ʼ����������ť��Tbb()����
  '----------------------------------------------------------------------------
  FOR x = 0 TO TOOLBUTTONS-1
    ' ����ÿ����ť�ĳ�ʼ״̬
    Tbb( x ).iBitmap = 0
    Tbb( x ).idCommand = 0
    Tbb( x ).fsState = %TBSTATE_ENABLED
    Tbb( x ).fsStyle = %TBSTYLE_BUTTON OR %TBSTYLE_AUTOSIZE
    Tbb( x ).dwData = 0
    Tbb( x ).iString = 0
  NEXT x
  x=0
  Tbb( x ).iBitmap = 75                   '�½�����
  Tbb( x ).idCommand = %IDM_WIN
  IF NewComCtl THEN
    tbb( x ).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb( x ).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR x
  Tbb( x ).iBitmap = 132                  '�����˵�
  Tbb( x ).idCommand = %ID_TOOL_MENU
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 133                  '����������
  Tbb( x ).idCommand = %ID_TOOL_TOOLBAR
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 138                  '�ؼ�����
  Tbb( x ).idCommand = %ID_BARONLY_STICKY
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 139                  '�ؼ���������
  Tbb( x ).idCommand = %ID_BARONLY_SNAP
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 140                  '��ʾ/��������
  Tbb( x ).idCommand = %ID_BARONLY_GRID
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 134                  '��ֱ����ؼ�
  Tbb( x ).idCommand = %ID_SETV_CENTER
  IF NewComCtl THEN
    tbb( x ).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb( x ).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR x
  Tbb( x ).iBitmap = 135                  'ˮƽ����ؼ�
  Tbb( x ).idCommand = %ID_SETH_CENTER
  IF NewComCtl THEN
    tbb( x ).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb( x ).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR x
  Tbb( x ).iBitmap = 136                  '��ֱ�Ⱦ�
  Tbb( x ).idCommand = %ID_SETV_EQUATE
  INCR x
  Tbb( x ).iBitmap = 137                  'ˮƽ�Ⱦ�
  Tbb( x ).idCommand = %ID_SETH_EQUATE
  INCR x
  Tbb( x ).iBitmap = 141                  '���ɴ���
  Tbb( x ).idCommand = %ID_TOOL_COMPILE
  IF NewComCtl THEN
    tbb( x ).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb( x ).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR x
  Tbb( x ).fsStyle = %TBSTYLE_SEP         '�ָ���
  '----------------------------------------------------------------------------
  ' ��ʼ���������������ڷ�����Ϣ
  '----------------------------------------------------------------------------
  Tabm.nID = %ID_WINDOWBAR    ' %IDB_STD_LARGE_COLOR
  Tabm.hInst = g_hInst
  SENDMESSAGE hToolBar, %TB_SETBITMAPSIZE, 0, MAKLNG( 16, 16 )
  ' ������������ͼƬ�б�
  SENDMESSAGE hToolBar, %TB_SETIMAGELIST, 0, hImlNor          ' ����ͼƬ
  SENDMESSAGE hToolBar, %TB_SETDISABLEDIMAGELIST, 0, hImlDis  ' ������ͼƬ
  SENDMESSAGE hToolBar, %TB_SETHOTIMAGELIST, 0, hImlHot       ' ����ͼƬ
  ' ����������ť����ͼƬ�б��ϵ�λͼ
  SENDMESSAGE hToolBar, %TB_ADDBITMAP, TOOLBUTTONS, VARPTR( Tabm )
  ' ���ù�������ť
  SENDMESSAGE hToolBar, %TB_BUTTONSTRUCTSIZE, SIZEOF( Tbb( 0 )), 0
  SENDMESSAGE hToolBar, %TB_ADDBUTTONS, TOOLBUTTONS, VARPTR( Tbb( 0 ))
'   a$ = $NUL & "0,0   " &  $NUL & "0,0   " & $NUL & $NUL
'   Sendmessage hToolBar, %TB_ADDSTRING,0, strptr(a$)
  ' ǿ�ƹ�������ʼ���óߴ�
  SENDMESSAGE hToolBar, %WM_SIZE, 0, 0
  ghWndSize = CreateWindowEx(0,"STATIC",_
                              "",_
                              %WS_CHILD OR %WS_VISIBLE OR _
                              %SS_SUNKEN OR %SS_LEFT, _
                              365, 1, 160, 24,_
                              hToolBar,_
                              BYVAL %NULL,_
                              g_hInst,_
                              BYVAL %NULL)
  '���໯��̬�ؼ��������Ϳ��Ի���λ�úͳߴ�
  LOCAL szOldWndProc AS ASCIIZ*11
  LOCAL lpOldWndProc AS LONG
  lpOldWndProc = SetWindowLong(ghWndSize,%GWL_WNDPROC,BYVAL CODEPTR(WndStatusSizeProc))
  'ͨ�����������ԣ���������ɵĴ��ڴ�����̹���
  szOldWndProc = "OldWndProc"
  CALL SetProp(ghWndSize,szOldWndProc,lpOldWndProc&)
  ' ǿ�ƹ�����Ϊ��ʼ�ߴ�.
  IF hToolBar THEN
    SendMessage hToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' �汾֧��ʱ�����ӱ�ƽ����������ť��ʽ
       SetWindowLong hToolBar, %GWL_STYLE, _
          GetWindowLong(hToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage hToolBar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF
  ' ���ع��������
  FUNCTION = hToolbar
END FUNCTION
'------------------------------------------------------------------------------
' ���� : CreateDebugBar ()
' ���� : �������룬���У����Թ������ؼ�.
'------------------------------------------------------------------------------
FUNCTION CreateDebugBar (BYVAL hWnd AS DWORD) AS DWORD

  ' Ϊ�Ż��ٶ�ʹ��ע�����.
  REGISTER x AS LONG
  LOCAL c AS LONG
  ' ����������.
  LOCAL TOOLBUTTONS AS INTEGER
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP
  LOCAL hToolbar AS DWORD
  TOOLBUTTONS = 9
  DIM Tbb(0 TO TOOLBUTTONS - 1) AS LOCAL TBBUTTON

  ' �����ȷ��λͼ�ߴ�!
  IF gBmpSize <> 24 THEN
   IF gBmpSize <> 16 THEN
    ' �û�ʹ���˲���ȷ��λͼ�ߴ�.
    MessageBox(hWnd, "������λͼ�ߴ����Ϊ16x16 �� 24x24!" & $CRLF & STR$(gBmpSize), "������ʾ", _
           %MB_OK OR %MB_ICONERROR OR %MB_TOPMOST)
    FUNCTION = 1
    EXIT FUNCTION
   END IF
  END IF

  '----------------------------------------------------------------------------
  ' ��������������.
  '----------------------------------------------------------------------------
  hToolBar = CreateWindowEx(0, _
               "ToolbarWindow32", _
               "", _
               %WS_CHILD OR %WS_VISIBLE OR %TBSTYLE_TOOLTIPS OR _
               %TBSTYLE_FLAT OR %TBSTYLE_LIST OR %TBSTYLE_TRANSPARENT OR _
               %WS_CLIPCHILDREN  OR %WS_CLIPSIBLINGS OR %CCS_NORESIZE OR _
               %CCS_NODIVIDER, _
               0, 0, 0, 0, _
               hWnd, _
               %ID_DEBUGBAR, _
               g_hInst, _
               BYVAL %NULL)

  '----------------------------------------------------------------------------
  ' ��鴴���������ؼ�����ʱ�Ĵ���.
  '----------------------------------------------------------------------------
  IF ISFALSE hToolBar THEN         ' ����!
   ' ��������������ʧ�ܣ�֪ͨ�û�.
   FUNCTION = %TRUE
   EXIT FUNCTION
  END IF
  FOR c = 0 TO TOOLBUTTONS-1
    ' Set the initial states for each button.
    Tbb( c ).iBitmap = 0
    Tbb( c ).idCommand = 0
    Tbb( c ).fsState = %TBSTATE_ENABLED
    Tbb( c ).fsStyle = %TBSTYLE_BUTTON OR %TBSTYLE_AUTOSIZE
    Tbb( c ).dwData = 0
    Tbb( c ).iString = 0
  NEXT c
  '----------------------------------------------------------------------------
  ' ��ʼ����������ť��Tbb()����.
  '----------------------------------------------------------------------------

  ' �ð�ť��Ϣ��� TBBUTTON ����
  c=0
  tbb(c).iBitmap   = 37 '���öϵ�
  tbb(c).idCommand = %IDM_SETBRKPO
  INCR c

  tbb(c).iBitmap   = 39 '��һ�ϵ�
  tbb(c).idCommand = %IDM_PRVBRKPO
  INCR c

  tbb(c).iBitmap   = 38 '��һ�ϵ�
  tbb(c).idCommand = %IDM_NEXBRKPO
  INCR c

  tbb(c).iBitmap   = 40 'ɾ���ϵ�
  tbb(c).idCommand = %IDM_DELBRKPO
  INCR c

  tbb(c).iBitmap   = 125 'ת���ϵ�
  tbb(c).idCommand = %IDM_GOBRKPO
  INCR c

  tbb(c).iBitmap   = -1
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c

  tbb(c).iBitmap   = 42
  tbb(c).idCommand = %IDM_COMPILE  '����
  INCR c

  tbb(c).iBitmap   = 43
  tbb(c).idCommand = %IDM_CMPRUN  '���벢����
  INCR c

  tbb(c).iBitmap   = 44 '���벢����
  tbb(c).idCommand = %IDM_CMPDEBUG
  tbb(c).fsState   = %TBSTATE_DISABLED
  INCR c

'  tbb(c).iBitmap   = -1                  '������Ҫ������������չ����ʾ���������룬��������ֹ�����еȰ�ť
'  'tbb(c).idCommand = 701
'  tbb(c).fsState   = %TBSTATE_ENABLED
'  tbb(c).fsStyle   = %TBSTYLE_SEP
'  INCR c
'
'  tbb(c).iBitmap   = -1
'  tbb(c).idCommand = 701
'  tbb(c).fsState   = %TBSTATE_ENABLED
'  tbb(c).fsStyle   = %TBSTYLE_SEP
'  INCR c
  '----------------------------------------------------------------------------
  ' ��ʼ�����������ڲ����ʹ�����Ϣ.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_DEBUGBAR '%IDB_STD_LARGE_COLOR ' ʹ��λͼͼƬ��ID����, %ID_TOOLBAR
  Tabm.hInst =  g_hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
                         ' ���ʹ��������Դ����λͼ

  ' ���ù�����λͼ�ߴ�Ϊ16x16.
  IF gBmpSize = 16 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(16, 16)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(16, 16)
  END IF

  ' ���ù�����λͼ�ߴ�Ϊ24x24.
  IF gBmpSize = 24 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(24, 24)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(24, 24)
  END IF

  ' ����ͼƬ�б�������.
  SendMessage hToolBar, %TB_SETIMAGELIST,         0, hImlNor  ' ����
  SendMessage hToolBar, %TB_SETDISABLEDIMAGELIST, 0, hImlDis  ' ������
  SendMessage hToolBar, %TB_SETHOTIMAGELIST,      0, hImlHot  ' ��ѡ��

  ' ����ͼƬ�б�λͼ����������ť.
  SendMessage hToolBar, %TB_ADDBITMAP, TOOLBUTTONS, VARPTR(Tabm)

  ' ���ù�������ť.
  SendMessage hToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage hToolBar, %TB_ADDBUTTONS, TOOLBUTTONS, VARPTR(Tbb(0))

  ' ǿ�ƹ�����Ϊ��ʼ�ߴ�.

  IF hToolBar THEN
    SendMessage hToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' then we can use more modern features, like flat style, etc.
       SetWindowLong hToolBar, %GWL_STYLE, _
          GetWindowLong(hToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage hToolBar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF

  FUNCTION = hToolBar            ' ���ع��������ھ��.

END FUNCTION
'------------------------------------------------------------------------------
' ������WndStatusSizeProc
' �������Զ���״̬���ڻص�����:
'       ��������ͼƬ�������µ�ǰ���༭���ڵ� x,y,cx,cy λ�ü���С��Ϣ
'------------------------------------------------------------------------------
FUNCTION WndStatusSizeProc(BYVAL hWnd AS LONG, BYVAL wMsg AS LONG, _
                           BYVAL wParam AS LONG,BYVAL lParam AS LONG) AS LONG
  LOCAL  szOldWndProc AS ASCIIZ*11
  LOCAL  ps           AS PAINTSTRUCT
  LOCAL  hDC          AS LONG
  LOCAL  hOldFont     AS LONG
  LOCAL  x            AS LONG
  LOCAL  y            AS LONG
  STATIC hbitmapLT    AS LONG
  STATIC hbitmapRB    AS LONG
  LOCAL  hMemDc       AS LONG
  LOCAL  Bm           AS BITMAP
  LOCAL  rc           AS RECT
  LOCAL  hInst        AS LONG
  STATIC lSizeXY      AS ASCIIZ*40
  STATIC lSizeCXY     AS ASCIIZ*40
  LOCAL  lpRect       AS RECT PTR
  STATIC fBmpLoaded   AS LONG
  LOCAL  lpOldWndProc AS LONG
  LOCAL ghMyFont      AS DWORD
  'hInst = GetModuleHandle(BYVAL %NULL)
  hInst = g_hInst
  IF lpOldWndProc<=0 THEN
    szOldWndProc = "OldWndProc"
    lpOldWndProc = GetProp(hWnd, szOldWndProc)
  END IF
  IF ghMyFont<=0 THEN
    ghMyFont = MakeFont("MS Sans Serif", 8, %FW_NORMAL, %FALSE, %FALSE, %FALSE, _
                      %DEFAULT_CHARSET)
  END IF
  SELECT CASE ( wMsg )
    CASE %WM_UPDATESIZEWINDOW
      'wParam = lpRect
      'lParam = handle, assume a control if it is valid
      lpRect = wParam
      rc = @lpRect
      lSizeXY  = STR$(rc.nLeft) + " ," + STR$(rc.nTop)
      lSizeCXY = STR$(rc.nRight-rc.nLeft) + " x" + STR$(rc.nBottom-rc.nTop)
      '��ˢ��...
      InvalidateRect hWnd, BYVAL %NULL,%TRUE
      UpdateWindow hWnd
      FUNCTION =0 :EXIT FUNCTION
    CASE %WM_ERASEBKGND
      FUNCTION =1 :EXIT FUNCTION
    CASE %WM_PAINT
      hDC = BeginPaint(hWnd, ps)
      '���֮ǰ�ĵ�ɫ...
      GetClientRect hWnd,rc
      FillRect hDC, rc, GetStockObject(%LTGRAY_BRUSH)'GRAY_BRUSH
      ' WM_CREATE ��Ϣ������ͼ���Ȼ��ѭ��
      IF fBmpLoaded = 0 THEN
        fBmpLoaded = 1
        '��ȡ x/y λͼ
        hbitmapLT = LoadBitmap( g_hInst,"XYBM" )
        '��ȡ cx/cy λ��
        hbitmapRB = LoadBitmap( g_hInst,"CXYBM" )
      END IF
      '����λͼ...
      hMemDc = CreateCompatibleDc(hDC)
      GetObject hbitmapLT,LEN(Bm),Bm
      SelectObject hMemDc ,hbitmapLT
      BitBlt hDC ,0 ,4 ,Bm.bmWidth ,Bm.bmHeight ,hMemDc ,0 ,0 ,%SRCCOPY
      GetObject hbitmapRB,LEN(Bm),Bm
      SelectObject hMemDc ,hBitmapRB
      BitBlt hDC ,81 ,4 ,Bm.bmWidth ,Bm.bmHeight ,hMemDc ,0 ,0 ,%SRCCOPY
      DeleteDc hMemDc
      '����ı�...
      hOldFont = SelectObject(hDC,ghMyFont)
      SetTextColor hDC, RGB(0,0,0)
      SetBkMode hDC, %TRANSPARENT
      TextOut hDC, 15 , 4, lSizeXY, LEN(lSizeXY)
      TextOut hDC, 96, 4, lSizeCXY, LEN(lSizeCXY)
      SelectObject hDC,hOldFont
      EndPaint hWnd,ps
      FUNCTION = 0 :EXIT FUNCTION
    CASE %WM_DESTROY
      '���ڴ�...
      IF ISTRUE(hbitmapLT) THEN CALL DeleteObject(hbitmapLT)
      IF ISTRUE(hbitmapRB) THEN CALL DeleteObject(hbitmapRB)
      ' ȡ�����໯�ؼ����ָ���Ĭ�ϴ������
      CALL SetWindowLong(hWnd, %GWL_WNDPROC, GetProp(hWnd,szOldWndProc))
      '�Ƴ���������
      RemoveProp hWnd,szOldWndProc
    END SELECT
    '��ʼ��Ϣ�������ദ�����
    FUNCTION = CallWindowProc( lpOldWndProc, hWnd, wMsg, wParam, lParam )

END FUNCTION
'------------------------------------------------------------------------------
' ���� : CreateCodeEditBar ()
' ���� : ��������༭�������ؼ�.
'------------------------------------------------------------------------------
FUNCTION CreateCodeEditBar (BYVAL hWnd AS DWORD) AS DWORD

  ' Ϊ�Ż��ٶ�ʹ��ע�����.
  REGISTER x AS LONG
  LOCAL c AS LONG
  ' ����������.
  LOCAL TOOLBUTTONS AS INTEGER
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP
  LOCAL hToolbar AS DWORD
  TOOLBUTTONS = 13
  DIM Tbb(0 TO TOOLBUTTONS - 1) AS LOCAL TBBUTTON

  ' �����ȷ��λͼ�ߴ�!
  IF gBmpSize <> 24 THEN
   IF gBmpSize <> 16 THEN
    ' �û�ʹ���˲���ȷ��λͼ�ߴ�.
    MessageBox(hWnd, "������λͼ�ߴ����Ϊ16x16 �� 24x24!" & $CRLF & STR$(gBmpSize), "������ʾ", _
           %MB_OK OR %MB_ICONERROR OR %MB_TOPMOST)
    FUNCTION = 1
    EXIT FUNCTION
   END IF
  END IF

  '----------------------------------------------------------------------------
  ' ��������������.
  '----------------------------------------------------------------------------
  hToolBar = CreateWindowEx(0, _
               "ToolbarWindow32", _
               "", _
               %WS_CHILD OR %WS_VISIBLE OR %TBSTYLE_TOOLTIPS OR _
               %TBSTYLE_FLAT OR %TBSTYLE_LIST OR %TBSTYLE_TRANSPARENT OR _
               %WS_CLIPCHILDREN  OR %WS_CLIPSIBLINGS OR %CCS_NORESIZE OR _
               %CCS_NODIVIDER, _
               0, 0, 0, 0, _
               hWnd, _
               %ID_CODEEDITBAR, _
               g_hInst, _
               BYVAL %NULL)

  '----------------------------------------------------------------------------
  ' ��鴴���������ؼ�����ʱ�Ĵ���.
  '----------------------------------------------------------------------------
  IF ISFALSE hToolBar THEN         ' ����!
   ' ��������������ʧ�ܣ�֪ͨ�û�.
   FUNCTION = %TRUE
   EXIT FUNCTION
  END IF
  FOR c = 0 TO TOOLBUTTONS-1
    ' Set the initial states for each button.
    Tbb( c ).iBitmap = 0
    Tbb( c ).idCommand = 0
    Tbb( c ).fsState = %TBSTATE_ENABLED
    Tbb( c ).fsStyle = %TBSTYLE_BUTTON OR %TBSTYLE_AUTOSIZE
    Tbb( c ).dwData = 0
    Tbb( c ).iString = 0
  NEXT c
  '----------------------------------------------------------------------------
  ' ��ʼ����������ť��Tbb()����.
  '----------------------------------------------------------------------------

  ' �ð�ť��Ϣ��� TBBUTTON ����
  c=0
  tbb(c).iBitmap   = 22 '��ע��
  tbb(c).idCommand = %IDM_BLOCKCOMMIT
  INCR c

  tbb(c).iBitmap   = 23 '����ע��
  tbb(c).idCommand = %IDM_REBLOCKCOMMIT
  INCR c

  tbb(c).iBitmap   = 24 '������
  tbb(c).idCommand = %IDM_BLOCKINDENT
  INCR c

  tbb(c).iBitmap   = 25 '��������
  tbb(c).idCommand = %IDM_REBLOCKINDENT
  INCR c

  tbb(c).iBitmap   = -1
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c

  tbb(c).iBitmap   = 28 'ת��д
  tbb(c).idCommand = %IDM_CAPWORD
  INCR c

  tbb(c).iBitmap   = 29 'תСд
  tbb(c).idCommand = %IDM_SMALLWORD
  INCR c

  tbb(c).iBitmap   = 30 '����ĸ��д
  tbb(c).idCommand = %IDM_FIRSTCAPWORD
  INCR c

  tbb(c).iBitmap   = -1
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c

  tbb(c).iBitmap   = 33 '����
  tbb(c).idCommand = %IDM_FIND
  INCR c

  tbb(c).iBitmap   = 34 '������һ��
  tbb(c).idCommand = %IDM_FINDNEXT
  INCR c

  tbb(c).iBitmap   = 35 '���Ҳ��滻
  tbb(c).idCommand = %IDM_FINDREPLACE
  'tbb(c).fsState   = %TBSTATE_DISABLED
  INCR c

  tbb(c).iBitmap   = 36 'ת����
  tbb(c).idCommand = %IDM_GOLINE
  INCR c

'  tbb(c).iBitmap   = -1                  '������Ҫ������������չ����ʾ���������룬��������ֹ�����еȰ�ť
'  'tbb(c).idCommand = 701
'  tbb(c).fsState   = %TBSTATE_ENABLED
'  tbb(c).fsStyle   = %TBSTYLE_SEP
'  INCR c
'
'  tbb(c).iBitmap   = -1
'  tbb(c).idCommand = 701
'  tbb(c).fsState   = %TBSTATE_ENABLED
'  tbb(c).fsStyle   = %TBSTYLE_SEP
'  INCR c
  '----------------------------------------------------------------------------
  ' ��ʼ�����������ڲ����ʹ�����Ϣ.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_CODEEDITBAR '%IDB_STD_LARGE_COLOR ' ʹ��λͼͼƬ��ID����, %ID_TOOLBAR
  Tabm.hInst =  g_hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
                         ' ���ʹ��������Դ����λͼ

  ' ���ù�����λͼ�ߴ�Ϊ16x16.
  IF gBmpSize = 16 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(16, 16)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(16, 16)
  END IF

  ' ���ù�����λͼ�ߴ�Ϊ24x24.
  IF gBmpSize = 24 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(24, 24)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(24, 24)
  END IF

  ' ����ͼƬ�б�������.
  SendMessage hToolBar, %TB_SETIMAGELIST,         0, hImlNor  ' ����
  SendMessage hToolBar, %TB_SETDISABLEDIMAGELIST, 0, hImlDis  ' ������
  SendMessage hToolBar, %TB_SETHOTIMAGELIST,      0, hImlHot  ' ��ѡ��

  ' ����ͼƬ�б�λͼ����������ť.
  SendMessage hToolBar, %TB_ADDBITMAP, TOOLBUTTONS, VARPTR(Tabm)

  ' ���ù�������ť.
  SendMessage hToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage hToolBar, %TB_ADDBUTTONS, TOOLBUTTONS, VARPTR(Tbb(0))

  ' ǿ�ƹ�����Ϊ��ʼ�ߴ�.

  IF hToolBar THEN
    SendMessage hToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' then we can use more modern features, like flat style, etc.
       SetWindowLong hToolBar, %GWL_STYLE, _
          GetWindowLong(hToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage hToolBar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF

  FUNCTION = hToolBar            ' ���ع��������ھ��.

END FUNCTION
'------------------------------------------------------------------------------
' ���� : CreateWinArragBar ()
'
' ���� : �����������й������ؼ�.
'------------------------------------------------------------------------------
FUNCTION CreateWinArragBar (BYVAL hWnd AS DWORD) AS DWORD

  ' Ϊ�Ż��ٶ�ʹ��ע�����.
  'REGISTER x AS LONG
  LOCAL c AS LONG
  ' ����������.
  LOCAL TOOLBUTTONS AS INTEGER
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP
  LOCAL hToolbar AS DWORD
  TOOLBUTTONS = 7
  DIM Tbb(0 TO TOOLBUTTONS - 1) AS LOCAL TBBUTTON

  ' �����ȷ��λͼ�ߴ�!
  IF gBmpSize <> 24 THEN
   IF gBmpSize <> 16 THEN
    ' �û�ʹ���˲���ȷ��λͼ�ߴ�.
    MessageBox(hWnd, "������λͼ�ߴ����Ϊ16x16 �� 24x24!" & $CRLF & STR$(gBmpSize), "������ʾ", _
           %MB_OK OR %MB_ICONERROR OR %MB_TOPMOST)
    FUNCTION = 1
    EXIT FUNCTION
   END IF
  END IF

  '----------------------------------------------------------------------------
  ' ��������������.
  '----------------------------------------------------------------------------
  hToolBar = CreateWindowEx(0, _
               "ToolbarWindow32", _
               "", _
               %WS_CHILD OR %WS_VISIBLE OR %TBSTYLE_TOOLTIPS OR _
               %TBSTYLE_FLAT OR %TBSTYLE_LIST OR %TBSTYLE_TRANSPARENT OR _
               %WS_CLIPCHILDREN  OR %WS_CLIPSIBLINGS OR %CCS_NORESIZE OR _
               %CCS_NODIVIDER, _
               0, 0, 0, 0, _
               hWnd, _
               %ID_WINARRAGBAR, _
               g_hInst, _
               BYVAL %NULL)

  '----------------------------------------------------------------------------
  ' ��鴴���������ؼ�����ʱ�Ĵ���.
  '----------------------------------------------------------------------------
  IF ISFALSE hToolBar THEN         ' ����!
   ' ��������������ʧ�ܣ�֪ͨ�û�.
   FUNCTION = %TRUE
   EXIT FUNCTION
  END IF
  FOR c = 0 TO TOOLBUTTONS-1
    ' Set the initial states for each button.
    Tbb( c ).iBitmap = 0
    Tbb( c ).idCommand = 0
    Tbb( c ).fsState = %TBSTATE_ENABLED
    Tbb( c ).fsStyle = %TBSTYLE_BUTTON OR %TBSTYLE_AUTOSIZE
    Tbb( c ).dwData = 0
    Tbb( c ).iString = 0
  NEXT c
  '----------------------------------------------------------------------------
  ' ��ʼ����������ť��Tbb()����.
  '----------------------------------------------------------------------------

  ' �ð�ť��Ϣ��� TBBUTTON ����
  c=0
  tbb(c).iBitmap   = 57 '�������
  tbb(c).idCommand = %IDM_CASCADE
  INCR c

  tbb(c).iBitmap   = 58 'ˮƽ����
  tbb(c).idCommand = %IDM_TILEH
  INCR c

  tbb(c).iBitmap   = 59 '��ֱ����
  tbb(c).idCommand = %IDM_TILEV
  INCR c

  tbb(c).iBitmap   = 62 'ͼ������
  tbb(c).idCommand = %IDM_ARRANGE
  INCR c

  tbb(c).iBitmap   = 60 '���
  tbb(c).idCommand = %IDM_MAXWIN
  INCR c

  tbb(c).iBitmap   = 61 'ת������
  tbb(c).idCommand = %IDM_SWITCHWINDOW
  INCR c

  tbb(c).iBitmap   = 63
  tbb(c).idCommand = %IDM_CLOSE  '�رմ���
  'INCR c

  '----------------------------------------------------------------------------
  ' ��ʼ�����������ڲ����ʹ�����Ϣ.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_WINARRAGBAR '%IDB_STD_LARGE_COLOR ' ʹ��λͼͼƬ��ID����, %ID_TOOLBAR
  Tabm.hInst =  g_hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
                         ' ���ʹ��������Դ����λͼ

  ' ���ù�����λͼ�ߴ�Ϊ16x16.
  IF gBmpSize = 16 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(16, 16)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(16, 16)
  END IF

  ' ���ù�����λͼ�ߴ�Ϊ24x24.
  IF gBmpSize = 24 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(24, 24)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(24, 24)
  END IF

  ' ����ͼƬ�б�������.
  SendMessage hToolBar, %TB_SETIMAGELIST,         0, hImlNor  ' ����
  SendMessage hToolBar, %TB_SETDISABLEDIMAGELIST, 0, hImlDis  ' ������
  SendMessage hToolBar, %TB_SETHOTIMAGELIST,      0, hImlHot  ' ��ѡ��

  ' ����ͼƬ�б�λͼ����������ť.
  SendMessage hToolBar, %TB_ADDBITMAP, TOOLBUTTONS, VARPTR(Tabm)

  ' ���ù�������ť.
  SendMessage hToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage hToolBar, %TB_ADDBUTTONS, TOOLBUTTONS, VARPTR(Tbb(0))

  ' ǿ�ƹ�����Ϊ��ʼ�ߴ�.

  IF hToolBar THEN
    SendMessage hToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' then we can use more modern features, like flat style, etc.
       SetWindowLong hToolBar, %GWL_STYLE, _
          GetWindowLong(hToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage hToolBar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF

  FUNCTION = hToolBar            ' ���ع��������ھ��.

END FUNCTION
'------------------------------------------------------------------------------
' ���� : CreateComboBox ()
'
' ���� : ����Combo Box�����ӵ�ReBar�ؼ���.
'        ʹ��SetParent���㼸�����������κοؼ���������!
'------------------------------------------------------------------------------
FUNCTION CreateComboBox (BYVAL hWnd AS DWORD) AS DWORD

  REGISTER i AS LONG

  LOCAL dwBaseUnits AS DWORD, sStr AS STRING
  LOCAL hToolbar AS DWORD
  LOCAL ghMyFont AS DWORD
  LOCAL tmpDword AS DWORD
  dwBaseUnits = GetDialogBaseUnits()

  '----------------------------------------------------------------------------
  ' ����ComboBox�ؼ�����.
  '----------------------------------------------------------------------------
  hToolbar = CreateWindowEx(0, _
                "ComboBox", _
                "", _
                %WS_CHILD OR %WS_VISIBLE OR %CBS_DROPDOWNLIST OR %DS_SETFONT OR _
                %WS_VSCROLL OR %WS_CLIPSIBLINGS OR %WS_CLIPCHILDREN, _
                -100,_ '(6 * LOWRD(dwBaseUnits)) / 4, _
                0,_ '(2 * HIWRD(dwBaseUnits)) / 8, _
                40,_ '(100 * LOWRD(dwBaseUnits)) / 4, _
                100,_ '(50 * HIWRD(dwBaseUnits)) / 8, _
                hWnd, _   'ghRebar,_
                %ID_COMBOBOX, _
                g_hInst, _
                BYVAL %NULL)
'  tmpDword = CreateWindowEx(0,"STATIC",_
'                                    "����",_
'                                    %WS_CHILD OR %WS_VISIBLE OR _
'                                    %SS_SUNKEN OR %SS_LEFT, _
'                                    0, 1, 50, 24,_
'                                    hToolBar,_
'                                    BYVAL %NULL,_
'                                    g_hInst,_
'                                    BYVAL %NULL)
'  control add combobox,hWnd,%ID_COMBOBOX,,10,1,100,100
'  control handle hWnd,%ID_COMBOBOX to hToolBar
  '----------------------------------------------------------------------------
  ' ��鴴��ComboBox����ʱ�Ĵ���.
  '----------------------------------------------------------------------------
  IF ISFALSE hToolbar THEN         ' ����!
    ' ����ComboBox����ʧ�ܣ�֪ͨ�û�.
    FUNCTION = %TRUE
    EXIT FUNCTION
  END IF
  '----------------------------------------------------------------------------
  ' ѭ������ComboBox�ı��ַ���.
  '----------------------------------------------------------------------------
  SendMessage hToolbar,%WM_SETFONT,GetStockObject(17),0 '%DEFAULT_GUI_FONT),0 '%ANSI_FIXED_FONT), 0
  'SendMessage tmpDword,%WM_SETFONT,GetStockObject(%ANSI_FIXED_FONT), 0
  FOR i = 1 TO 20
    sStr = "�� " & FORMAT$(i) & " ��"
'    cbexi.pszText=strptr(sStr)
'    cbexi.cchTextMax= len(sStr)
    'SendMessage hToolbar, %CBEM_INSERTITEM, 0&, varptr(cbexi)'STRPTR(sStr)
    SendMessage hToolbar,%CB_ADDSTRING,0&,STRPTR(sStr)
  NEXT
  COMBOBOX SELECT hWnd,%ID_COMBOBOX,1
  'SendMessage hToolbar, %CBEM_SETCURSEL, 0&, 0&

  FUNCTION = hToolbar

END FUNCTION
'------------------------------------------------------------------------------
' ���� : CreateRebar ()
'
' ���� : ����Rebar�ؼ�.
'------------------------------------------------------------------------------
FUNCTION CreateRebar (BYVAL hWnd AS DWORD) AS DWORD

  LOCAL rbi       AS REBARINFO
  LOCAL rbBand    AS REBARBANDINFO
  LOCAL rc        AS RECT
  LOCAL szCbText  AS ASCIIZ * 255
  LOCAL szTbText  AS ASCIIZ * 6
  LOCAL szRbImage AS ASCIIZ * %MAX_PATH
  LOCAL dwBtnSize AS DWORD
  LOCAL rowCount  AS LONG
  LOCAL bandCount AS LONG
  LOCAL tmpStr    AS STRING
  LOCAL i         AS INTEGER
  LOCAL wID       AS LONG
  LOCAL bandInfo() AS STRING
  LOCAL baseStyle AS DWORD
  LOCAL trect     AS RECT
  LOCAL barStr    AS STRING
  '----------------------------------------------------------------------------
  ' ����Rebar�ؼ�����.
  '----------------------------------------------------------------------------
  ghRebar = CreateWindowEx(0, _
               "ReBarWindow32", _
               "", _
               %WS_BORDER OR %WS_CHILD OR %WS_VISIBLE OR _
               %WS_CLIPSIBLINGS OR %WS_CLIPCHILDREN OR _
               %RBS_VARHEIGHT OR %RBS_BANDBORDERS OR _
               %RBBS_FIXEDSIZE, _
               0, 0, 0, 0, _
               hWnd, _
               %ID_REBAR, _
               g_hInst, _
               BYVAL %NULL)

  '----------------------------------------------------------------------------
  ' ��鴴��Rebar�ؼ����ڵĴ���.
  '----------------------------------------------------------------------------
  IF ISFALSE ghRebar THEN         ' ����!
   '= ����Rebar����ʧ�ܣ�֪ͨ�û�.
   FUNCTION = %TRUE
   EXIT FUNCTION
  END IF
  FUNCTION=ghRebar
  '----------------------------------------------------------------------------
  ' ��ʼ��������REBARINFO�ṹ.
  '----------------------------------------------------------------------------
  rbi.cbSize = SIZEOF(rbi)
  rbi.fMask  = 0
  rbi.himl   = 0

  SendMessage ghRebar, %RB_SETBARINFO, 0, VARPTR(rbi)

  '----------------------------------------------------------------------------
  ' ��ʼ������Rebar�ε�REBARBANDINFO.
  '----------------------------------------------------------------------------
  rbBand.cbSize     = SIZEOF(rbBand)
  rbBand.fMask      = %RBBIM_COLORS    OR _    ' clrFore �� clrBack ����
                      %RBBIM_CHILD     OR _    ' hwndChild ����
                      %RBBIM_CHILDSIZE OR _    ' cxMinChild �� cyMinChild ����
                      %RBBIM_STYLE     OR _    ' fStyle ����
                      %RBBIM_ID        OR _    ' wID ����
                      %RBBIM_SIZE      OR _    ' cx ����
                      %RBBIM_TEXT      OR _    ' lpText ����
                      %RBBIM_BACKGROUND        ' hbmBack ����

  rbBand.clrFore    = GetSysColor(%COLOR_BTNTEXT)
  rbBand.clrBack    = GetSysColor(%COLOR_BTNFACE)
  baseStyle     = %RBBS_NOVERT     OR _    ' ����ʾ��ֱ����
                  %RBBS_CHILDEDGE  OR _
                  %RBBS_FIXEDBMP   OR _
                  %RBBS_GRIPPERALWAYS
  'rbBand.fStyle =

'  '----------------------------------------------------------------------------
'  ' ����û�ѡ����һ��Rebar����ͼƬ���Ǽ�����!
'  '----------------------------------------------------------------------------
'  IF ISTRUE %USE_RBIMAGE THEN
'   ghRbBack   = LoadImage(g_hInst, "RBBACK", %IMAGE_BITMAP, 0, 0, _
'              %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR _
'              %LR_DEFAULTSIZE)
'
'   rbBand.hbmBack    = ghRbBack           ' Rebar����ͼƬ.
'  END IF
  '������
  rowCount=VAL(ReadConfig("rebarrowcount"))
  bandCount=VAL(ReadConfig("rebarbandcount"))
  tmpStr=ReadConfig("rebarband0")
  IF rowCount<=0 OR bandCount<=0 OR MID$(tmpStr,1,1)="0" THEN 'ȡ�õ�ֵ�쳣����Ϊ�յ�����£�����ʾ���湤������Ȼ���˳�������
    'msgbox "rowCount=" & str$(rowCount) & $crlf & "bandCount=" & str$(bandCount)
    'RunLog "�����������쳣��û�п���ʾ�Ĺ�������rowCount=" & STR$(rowCount) & $CRLF & "bandCount=" & STR$(bandCount)
    g_hToolbar=CreateToolBar(hWnd)
    '----------------------------------------------------------------------------
    ' ��ʼ�������ӹ�������Rebar�Ķ�.
    '----------------------------------------------------------------------------
    dwBtnSize         = SendMessage(ghToolBar, %TB_GETBUTTONSIZE, 0, 0)
    'szTbText          = "ToolBar"
    rbBand.lpText     = VARPTR(szTbText)
    rbBand.hwndChild  = ghToolBar
    rbBand.wID        = %ID_TOOLBAR
    rbBand.cxMinChild = 55
    rbBand.cyMinChild = HIWRD(dwBtnSize) + 2
    rbBand.cx         = 250
    SendMessage ghRebar, %RB_INSERTBAND, -1&, VARPTR(rbBand)
    LockBar VAL(ReadConfig("rebarlocked"))
    EXIT FUNCTION
  END IF
'  IF rowCount>=2 THEN
'    rbBand.fStyle = rbBand.fStyle OR %RBBS_BREAK
'  END IF
  FOR i=0 TO bandCount-1
    tmpStr=ReadConfig("rebarband" & FORMAT$(i))
    'RunLog "�����Ĺ�������" & "rebarband" & FORMAT$(i) & ":" & tmpStr
    REDIM bandInfo(3)
    PARSE tmpStr,bandInfo(),","
    wID=VAL(bandInfo(0))
    IF wID<=0 THEN
      EXIT FUNCTION
    END IF
    rbBand.cx         = VAL(bandInfo(1))
    IF rbBand.cx<=0 THEN
      rbBand.cx=200
    END IF
    IF bandInfo(2)="" THEN  'Ĭ�ϲ����أ�����ʾ
      bandInfo(2)="0"
    END IF
    rbBand.fStyle = baseStyle
    IF VAL(bandInfo(2))=1 THEN
      rbBand.fStyle=rbBand.fStyle OR %RBBS_HIDDEN
    END IF
    IF bandInfo(3)="" THEN  'Ĭ�ϻ���
      bandInfo(3)="1"
    END IF
    IF VAL(bandInfo(3))=1 THEN
      rbBand.fStyle=rbBand.fStyle OR %RBBS_BREAK
    END IF
    rbBand.wID        = wID
    SELECT CASE wID
      CASE %ID_TOOLBAR
        'MSGBOX STR$(i) & " toolbar " & STR$(wID)
        barStr="stdtool"
        ghToolbar=CreateToolBar(hWnd)
        dwBtnSize         = SendMessage(ghToolBar, %TB_GETBUTTONSIZE, 0, 0)
        szTbText          = ""
        rbBand.lpText     = VARPTR(szTbText)
        rbBand.hwndChild  = ghToolBar
        rbBand.cxMinChild = 55
        rbBand.cyMinChild = HIWRD(dwBtnSize) + 2
        'rbBand.cx         = 250
      CASE %ID_COMBOBOX
        barStr="combobox"
        'MSGBOX STR$(i) & " combobox " & STR$(wID)
        ghComboBox=CreateComboBox(hWnd)
        'GetWindowRect ghComboBox, trect
        szTbText          = "����:"
        rbBand.lpText     = VARPTR(szTbText)
        rbBand.hwndChild  = ghComboBox
        rbBand.cxMinChild = 50 'trect.nRight '100
        rbBand.cyMinChild = 23 '18 'HIWRD(dwBtnSize) '+ 2
      CASE %ID_WINDOWBAR
        barStr="winEdit"
        'MSGBOX STR$(i) & " windowbar " & STR$(wID)
        ghWindowBar=CreateWindowBar(hWnd)
        dwBtnSize         = SendMessage(ghWindowBar, %TB_GETBUTTONSIZE, 0, 0)
        szTbText          = ""
        rbBand.lpText     = VARPTR(szTbText)
        rbBand.hwndChild  = ghWindowBar
        rbBand.cxMinChild = 60
        'rbBand.cyMinChild = HIWRD(dwBtnSize) + 2
        'rbBand.cx         = 200
      CASE %ID_DEBUGBAR
        barStr="debug"
        'MSGBOX STR$(i) & " debugbar " & STR$(wID)
        ghDebugBar=CreateDebugBar(hWnd)
        dwBtnSize         = SendMessage(ghDebugBar, %TB_GETBUTTONSIZE, 0, 0)
        szTbText          = ""
        rbBand.lpText     = VARPTR(szTbText)
        rbBand.hwndChild  = ghDebugBar
        rbBand.cxMinChild = 55
        'rbBand.cyMinChild = HIWRD(dwBtnSize) + 2
      CASE %ID_WINARRAGBAR
        barStr="winArrag"
        'MSGBOX STR$(i) & " arragbar " & STR$(wID)
        ghWinArragBar=CreateWinArragBar(hWnd)
        dwBtnSize         = SendMessage(ghWinArragBar, %TB_GETBUTTONSIZE, 0, 0)
        szTbText          = ""
        rbBand.lpText     = VARPTR(szTbText)
        rbBand.hwndChild  = ghWinArragBar
        rbBand.cxMinChild = 55
        'rbBand.cyMinChild = HIWRD(dwBtnSize) + 2
      CASE %ID_CODEEDITBAR
        barStr="codeEdit"
        'MSGBOX STR$(i) & " codeeditbar " & STR$(wID)
        ghCodeEditBar=CreateCodeEditBar(hWnd)
        dwBtnSize         = SendMessage(ghCodeEditBar, %TB_GETBUTTONSIZE, 0, 0)
        szTbText          = ""
        rbBand.lpText     = VARPTR(szTbText)
        rbBand.hwndChild  = ghCodeEditBar
        rbBand.cxMinChild = 55
        'rbBand.cx=200
    END SELECT
    'RunLog "�����Ĺ�������" & barStr & " rebarband" & FORMAT$(i) & ":" & tmpStr
    IF wID<>%ID_COMBOBOX THEN
      rbBand.cyMinChild = HIWRD(dwBtnSize) + 2
    END IF
    'MSGBOX tmpStr & $CRLF & RIGHT$(tmpStr,1)
    SendMessage ghRebar, %RB_INSERTBAND, -1&, VARPTR(rbBand)
  NEXT i
  LockBar VAL(ReadConfig("rebarlocked"))
  FUNCTION =ghRebar'%FALSE
END FUNCTION
'------------------------------------------------------------------------------
' ���� : SaveRebarBandPos ()
' ���� : ����Rebar��λ��.
'------------------------------------------------------------------------------
FUNCTION SaveRebarBand() AS LONG
  DIM lResult AS LONG, lBands AS LONG
  DIM rBand   AS REBARBANDINFO
  LOCAL ishidden AS INTEGER
  LOCAL isbreak AS INTEGER
  LOCAL tmpStr AS STRING
  REGISTER lCount AS LONG
  ' �õ�Rebar����.
  lResult = SendMessage(ghRebar, %RB_GETROWCOUNT, 0&, 0&)
  'udtAp.rbRowCount = lResult
  WriteConfig("rebarrowcount",FORMAT$(lResult))
  'RunLog "����Ĺ�������������" & FORMAT$(lResult)
  ' �õ�Rebar����.
  lBands = SendMessage(ghRebar, %RB_GETBANDCOUNT, 0&, 0&)
  'udtAp.rbBandCount = lResult
  WriteConfig("rebarbandcount",FORMAT$(lBands))
  'RunLog "����Ĺ�������������" & FORMAT$(lBands)
  ' ����Rebar��REBARBANDINFO�ṹ.
  rBand.cbSize = SIZEOF(rBand)
  rBand.fMask  = %RBBIM_ID OR %RBBIM_SIZE OR %RBBIM_STYLE
  'MSGBOX "rowcount=" & STR$(lResult) & $CRLF & "bandscount=" & STR$(lBands)
  ' ѭ���õ�ÿ�ε�ID.
  FOR lCount = 0 TO lBands - 1
    'lResult = SendMessage(hRebar, %RB_IDTOINDEX, lCount, VARPTR(rBand))
    lResult = SendMessage(ghRebar, %RB_GETBANDINFO, lCount, VARPTR(rBand))
    IF ISTRUE(rBand.fStyle AND %RBBS_HIDDEN) THEN
      ishidden=1
    ELSE
      ishidden=0
    END IF
    IF ISTRUE(rBand.fStyle AND %RBBS_BREAK) THEN
      isbreak=1
    ELSE
      isbreak=0
    END IF
    tmpStr=FORMAT$(rBand.wID) & "," & FORMAT$(rBand.cx) & "," & FORMAT$(ishidden) & "," & FORMAT$(isbreak)
    'RunLog "����Ĺ�������" & "rebarband" & FORMAT$(lCount) & ":" & tmpStr
    WriteConfig "rebarband" & FORMAT$(lCount),tmpStr
'    IF lCount = 0 THEN udtAp.rbBand0 = rBand.Wid
'    IF lCount = 1 THEN udtAp.rbBand1 = rBand.Wid
  NEXT
  'msgbox str$(istrue(rBand.fStyle OR %RBBS_GRIPPERALWAYS)) & $crlf & str$(istrue(rBand.fStyle AND NOT(%RBBS_NOGRIPPER)))
  IF ISTRUE(rBand.fStyle AND %RBBS_GRIPPERALWAYS) THEN   ' AND ISTRUE(rBand.fStyle AND NOT(%RBBS_NOGRIPPER))
    WriteConfig("rebarlocked","0")
  ELSE
    WriteConfig("rebarlocked","1")
  END IF
END FUNCTION

FUNCTION MakeRebarMenu() AS DWORD
  LOCAL hPopupMenu AS DWORD
  hPopupMenu = CREATEPOPUPMENU
  IF ghToolBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_STDTOOLBAR, "��׼"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_STDTOOLBAR, "��׼"
  END IF
  IF ghWindowBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_WINDOWBAR, "���ڱ༭"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_WINDOWBAR, "���ڱ༭"
  END IF
  IF ghComboBox THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_COMBOBAR, "����"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_COMBOBAR, "����"
  END IF
  IF ghDebugBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_DEBUGBAR, "���������"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_DEBUGBAR, "���������"
  END IF
  IF ghCodeEditBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_CODEEDIT, "����༭"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_CODEEDIT, "����༭"
  END IF
  IF ghWinArragBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_WINARRAG, "��������"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_WINARRAG, "��������"
  END IF
  APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_SEPARATOR  , 0, "-"
  IF ReadConfig("rebarlocked")="0" THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %ID_LOCKBAR, "����������"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %ID_LOCKBAR, "����������"
  END IF
  FUNCTION=hPopupMenu
END FUNCTION
FUNCTION CheckRebarMenu(BYVAL hRebarMenu AS DWORD) AS DWORD
  LOCAL rbBand AS REBARBANDINFO
  LOCAL bandCount AS INTEGER
  LOCAL tmpRC AS RECT
  LOCAL i AS INTEGER
  LOCAL ishidden AS INTEGER
  LOCAL isbreak AS INTEGER
  LOCAL tmpStr AS STRING
  LOCAL menuStyle AS DWORD
  LOCAL menuIndex AS DWORD
  LOCAL islock AS INTEGER
  bandCount=SendMessage(ghRebar,%RB_GETBANDCOUNT,0,0)
  rbBand.cbSize=SIZEOF(rbBand)
  rbBand.fMask = %RBBIM_ID OR %RBBIM_SIZE OR %RBBIM_CHILD OR %RBBIM_STYLE
  FOR i=0 TO bandCount-1
    menuStyle=%MF_ENABLED OR %MF_CHECKED
    SendMessage(ghRebar,%RB_GETBANDINFO,i,VARPTR(rbBand))
    IF ISTRUE(rbBand.fStyle AND %RBBS_HIDDEN) THEN
      menuStyle=%MF_ENABLED OR %MF_UNCHECKED
    END IF
    menuIndex=0
    IF rbBand.hwndChild>0 THEN
      SELECT CASE rbBand.hwndChild
        CASE ghToolBar
          menuIndex=1
        CASE ghWindowBar
          menuIndex=2
        CASE ghComboBox
          menuIndex=3
        CASE ghDebugBar
          menuIndex=4
        CASE ghCodeEditBar
          menuIndex=5
        CASE ghWinArragBar
          menuIndex=6
      END SELECT
      IF menuIndex>0 THEN
        MENU SET STATE hRebarMenu,menuIndex,menuStyle
      END IF
    END IF
  NEXT i
  IF VAL(ReadConfig("rebarlocked"))=1 THEN
    MENU SET STATE hRebarMenu,8,%MF_ENABLED OR %MF_CHECKED
  ELSE
    MENU SET STATE hRebarMenu,8,%MF_ENABLED OR %MF_UNCHECKED
  END IF
END FUNCTION
SUB SaveRebarConfig()
  LOCAL rbBand AS REBARBANDINFO
  LOCAL bandCount AS INTEGER
  LOCAL tmpRC AS RECT
  LOCAL i AS INTEGER
  LOCAL ishidden AS INTEGER
  LOCAL isbreak AS INTEGER
  LOCAL tmpStr AS STRING
  bandCount=SendMessage(ghRebar,%RB_GETBANDCOUNT,0,0)
  WriteConfig "rebarbandcount",FORMAT$(bandCount)
  GetWindowRect g_hRebar, tmpRC
  WriteConfig "rebarrowcount",FORMAT$((tmpRC.nBottom-tmpRC.nTop)\28)
  rbBand.cbSize=SIZEOF(rbBand)
  rbBand.fMask = %RBBIM_ID OR %RBBIM_SIZE
  FOR i=0 TO bandCount-1
    SendMessage(ghRebar,%RB_GETBANDINFO,i,VARPTR(rbBand))
    IF ISTRUE(rbBand.fStyle AND %RBBS_HIDDEN) THEN
      ishidden=1
    END IF
    IF ISTRUE(rbBand.fStyle AND %RBBS_BREAK) THEN
      isbreak=1
    END IF
    tmpStr=FORMAT$(rbBand.wID) & "," & FORMAT$(rbBand.cx) & "," & FORMAT$(ishidden) & "," & FORMAT$(isbreak)
    WriteConfig "rebarband" & FORMAT$(i),tmpStr
  NEXT i
END SUB
FUNCTION ShowBand(BYVAL hBar AS DWORD,BYVAL wID AS DWORD,BYVAL sign AS BYTE) AS LONG ',BYVAL TbText AS STRING)AS LONG
  LOCAL rbBand AS REBARBANDINFO
  LOCAL bandCount AS INTEGER
  LOCAL i AS INTEGER
  LOCAL hasDone AS INTEGER
  LOCAL szTbText AS ASCIIZ * 255
  LOCAL rowCount AS INTEGER
  LOCAL MenuState AS LONG
  'msgbox "ShowBand hBar=" & str$(hBar) & ",wID=" & str$(wID)
  'szTbText=TbText
  hasDone=0
  IF hBar<=0 THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  bandCount=SendMessage(ghRebar,%RB_GETBANDCOUNT,0,0)'varptr(bandCount)
  'msgbox str$(bandCount)
  IF bandCount<1 THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  '----------------------------------------------------------------------------
  ' ��ʼ������Rebar�ε�REBARBANDINFO.
  '----------------------------------------------------------------------------
  rbBand.cbSize     = SIZEOF(rbBand)
  rbBand.fMask      = %RBBIM_COLORS    OR _    ' clrFore �� clrBack ����
                      %RBBIM_CHILD     OR _    ' hwndChild ����
                      %RBBIM_CHILDSIZE OR _    ' cxMinChild �� cyMinChild ����
                      %RBBIM_STYLE     OR _    ' fStyle ����
                      %RBBIM_ID        OR _    ' wID ����
                      %RBBIM_SIZE      OR _    ' cx ����
                      %RBBIM_TEXT      OR _    ' lpText ����
                      %RBBIM_BACKGROUND        ' hbmBack ����

  rbBand.clrFore    = GetSysColor(%COLOR_BTNTEXT)
  rbBand.clrBack    = GetSysColor(%COLOR_BTNFACE)

  rbBand.fStyle = %RBBS_NOVERT     OR _    ' ����ʾ��ֱ����
                  %RBBS_CHILDEDGE  OR _
                  %RBBS_FIXEDBMP   OR _
                  %RBBS_GRIPPERALWAYS
  FOR i=0 TO bandCount-1
    SendMessage(ghRebar,%RB_GETBANDINFO,i,VARPTR(rbBand))
    IF rbBand.hwndChild=hBar THEN
      SendMessage ghRebar,%RB_SHOWBAND,i,sign
      hasDone=1
      EXIT FOR
    END IF
  NEXT i
  IF hasDone=0 AND sign=1 THEN   '���û�У���Ҫ����ʾ�Ļ����������ﴴ��������Band
    IF ghRebarMenu THEN
      MENU GET STATE ghRebarMenu,8 TO MenuState
      IF MenuState=%MF_CHECKED THEN
        rbBand.fStyle= rbBand.fStyle AND NOT(%RBBS_GRIPPERALWAYS)
        rbBand.fStyle= rbBand.fStyle OR %RBBS_NOGRIPPER
      END IF
    END IF
    IF wID=%ID_COMBOBOX THEN
      szTbText          = "����:"
    END IF
    rbBand.lpText     = VARPTR(szTbText)
    rbBand.hwndChild  = hBar
    rbBand.wID        = wID '%ID_TOOLBAR
    rbBand.cxMinChild = 100
    rbBand.cyMinChild = 23 '18 'HIWRD(dwBtnSize) '+ 2
    rbBand.cx         = 280
    IF bandCount>=2 THEN
      rbBand.fStyle = rbBand.fStyle OR %RBBS_BREAK
    END IF
    SendMessage ghRebar, %RB_INSERTBAND, 0&, VARPTR(rbBand)
    rowCount=SendMessage(ghRebar, %RB_GETROWCOUNT, 0,0)
    SendMessage ghRebar, %RB_MOVEBAND, rowCount, 0&
  END IF
  FUNCTION=1
END FUNCTION
' *****************************************************************************
' ���ز˵����ַ�������0��ͷ��λͼ������(-1 = ��λͼ).
' λͼͼƬ�б����Ǵ�0��ʼ�ģ����� BmpNum ��0��ʼ
' *****************************************************************************
FUNCTION GetMenuTxtBmp( BYVAL ItemId AS LONG, BmpNum AS LONG ) AS STRING
  SELECT CASE ItemId
    ' �ļ��˵�
    'CASE %IDM_FILE :        BmpNum = -1 : FUNCTION = "�ļ�"
    CASE %IDM_NEW :           BmpNum = 1 : FUNCTION = "�½�"
    CASE %IDM_NEWPROJECT :    BmpNum = 72 : FUNCTION = "�½���Ŀ"
    CASE %IDM_NEWBAS :        BmpNum = 45 : FUNCTION = "��BAS�ļ�" & $TAB & "Ctrl+N"    '&New
    CASE %IDM_NEWINC :        BmpNum = 45 : FUNCTION = "��INC�ļ�"
    CASE %IDM_NEWRC :         BmpNum = 45 : FUNCTION = "��RC�ļ�"

    CASE %IDM_OPEN :          BmpNum = 2 : FUNCTION = "��..." & $TAB & "Ctrl+O"     '&Open
    CASE %IDM_REOPEN :        BmpNum = 2 : FUNCTION = "�ش�"
    CASE %IDM_INSERTFILE :    BmpNum = 0 : FUNCTION = "�����ļ�..."           '&Insert File
    CASE %IDM_SAVE :          BmpNum = 3 : FUNCTION = "����" & $TAB & "Ctrl+S"    '&Save
    'CASE %IDM_REFRESH :       BmpNum = 108 : FUNCTION = "ˢ��"
    CASE %IDM_SAVEAS :        BmpNum = 4 : FUNCTION = "���Ϊ..."             'Save File &As
    CASE %IDM_SAVEALL :       BmpNum = 96 : FUNCTION = "ȫ����"             'Save File &As
    CASE %IDM_SAVEPROJECT :   BmpNum = 112: FUNCTION = "������Ŀ"
    CASE %IDM_CLOSE :         BmpNum = 63 : FUNCTION = "�ر�" & $TAB & "Ctrl+F4"     'Close
    CASE %IDM_CLOSEOTHERS :   BmpNum = 97 : FUNCTION = "�ر�����"     'Close
    CASE %IDM_CLOSEALL :      BmpNum = 6 : FUNCTION = "ȫ���ر�"            'Close A&ll Files
    CASE %IDM_PRINTSETTING :  BmpNum = 7 : FUNCTION = "ҳ������..." + $TAB + "Ctrl+Alt+P"     'Page Set&up
    CASE %IDM_PRINTPREVIEW :  BmpNum = 8 : FUNCTION = "��ӡԤ��" + $TAB + "Ctrl+Shift+P"
    CASE %IDM_PRINT :         BmpNum = 9 : FUNCTION = "��ӡ����" + $TAB + "Ctrl+P"  '&Print
    'CASE %IDM_FORM_PRINT :      BmpNum = 9 : FUNCTION = "��ӡ����"
    CASE %IDM_OPENCMD :       BmpNum = 10 : FUNCTION = "������" + $TAB + "Ctrl+Alt+O"     'Comman&d Prompt
    CASE %IDM_EXIT :          BmpNum = 11 : FUNCTION = "�˳�" & $TAB & "Alt+F4"     'E&xit
    CASE %IDM_RECENT1 TO %IDM_RECENT8
      BmpNum = 2
      IF ( ItemId - %IDM_RECENTFILES ) = > LBOUND( RecentFiles ) AND _
              ( ItemId - %IDM_RECENTFILES ) < = UBOUND( RecentFiles ) THEN _
              FUNCTION = "&" & FORMAT$( ItemId - %IDM_RECENTFILES ) & " " & RecentFiles( ItemId - %IDM_RECENTFILES )
'    CASE %IDM_PBTEMPLATES + 1 TO %IDM_PBTEMPLATES + %MAXPBTEMPLATES + 1
'      BmpNum = 31
'      FUNCTION = PARSE$( g_PbTemplatesDesc, "|", ItemId - %IDM_PBTEMPLATES )

    ' �༭�˵�
    'CASE %IDM_EDITHEADER :        BmpNum = - 1    : FUNCTION = "�༭"               '&Edit
    CASE %IDM_UNDO :          BmpNum = 12 : FUNCTION = "����" + $TAB + "Ctrl+Z"     'U&ndo
    CASE %IDM_REDO :          BmpNum = 13 : FUNCTION = "�ظ�" + $TAB + "Ctrl+K"     'Red&o
    CASE %IDM_CUT :           BmpNum = 17 : FUNCTION = "����" + $TAB + "Ctrl+X"     'Cu&t
    CASE %IDM_COPY :          BmpNum = 18 : FUNCTION = "����" + $TAB + "Ctrl+C"     '&Copy
    CASE %IDM_PASTE :         BmpNum = 19 : FUNCTION = "ճ��" + $TAB + "Ctrl+V"     '&Paste
    CASE %IDM_PASTEIE :       BmpNum = 20 : FUNCTION = "ճ����ҳ" + $TAB + "Ctrl+Shift+V"     'Paste from Internet Explorer
    CASE %IDM_SELALL :        BmpNum = 21 : FUNCTION = "ȫѡ" + $TAB + "Ctrl+A"
    CASE %IDM_DELETE :        BmpNum = 15 : FUNCTION = "ɾ��" + $TAB + "Delete"
    CASE %IDM_LINEDELETE :    BmpNum = 14 : FUNCTION = "ɾ����" & $TAB & "Ctrl+Y"     '&Delete Line
    CASE %IDM_INITASNEW :     BmpNum = 1  : FUNCTION = "��ʼ��" & $TAB & "Ctrl+Shift+N"     ' *** NTM ***'&Initialize as New
    CASE %IDM_CLEAR :         BmpNum = 15 : FUNCTION = "���"               'Clea&r
    CASE %IDM_CLEARALL :      BmpNum = 16 : FUNCTION = "ȫɾ"             'Cl&ear All
    CASE %IDM_INDENT :        BmpNum = 24 : FUNCTION = "����" + $TAB + "Tab"    'Block Indent
    CASE %IDM_OUTDENT :       BmpNum = 25 : FUNCTION = "������" + $TAB + "Shift+Tab"    'Block Unindent
    CASE %IDM_COMMENT :       BmpNum = 22 : FUNCTION = "ע��" + $TAB + "Ctrl+Q"
    CASE %IDM_UNCOMMENT :     BmpNum = 23 : FUNCTION = "��ע��" + $TAB + "Ctrl+Shift+Q"
    CASE %IDM_FORMATREGION :  BmpNum = 26 : FUNCTION = "��ʽ���ı�" + $TAB + "Ctrl+B"
    CASE %IDM_TABULATEREGION :BmpNum = 103: FUNCTION = "��TAB�����ı�" + $TAB + "Ctrl+Shift+B"
    CASE %IDM_SELTOUPPERCASE :BmpNum = 28 : FUNCTION = "ת��Ϊ��д" + $TAB + "Ctrl+Alt+U"
    CASE %IDM_SELTOLOWERCASE :BmpNum = 29 : FUNCTION = "ת��ΪСд" + $TAB + "Ctrl+Alt+L"
    CASE %IDM_SELTOMIXEDCASE :BmpNum = 30 : FUNCTION = "ת��Ϊ����ĸ��д" + $TAB + "Ctrl+Alt+M"
    CASE %IDM_TEMPLATES :     BmpNum = 31 : FUNCTION = "ģ�����" + $TAB + "Ctrl+Shift+C"
    CASE %IDM_HTMLCODE :      BmpNum = 32 : FUNCTION = "ת��Ϊ��ҳ"
    CASE %IDM_GUID :          BmpNum = 115: FUNCTION = "������GUID"

    ' �����˵�
    'CASE %IDM_SEARCHHEADER :      BmpNum = - 1 : FUNCTION = "����"
    CASE %IDM_FIND :          BmpNum = 33 : FUNCTION = "����..." + $TAB + "Ctrl+F"
    CASE %IDM_FINDNEXT :      BmpNum = 34 : FUNCTION = "��һ��" + $TAB + "F3"
    CASE %IDM_FINDBACKWARDS : BmpNum = 125 : FUNCTION = "��һ��" + $TAB + "Shift+F3"
    CASE %IDM_REPLACE :       BmpNum = 35 : FUNCTION = "�滻..." + $TAB + "Ctrl+R"
    CASE %IDM_GOTOLINE :      BmpNum = 36 : FUNCTION = "ת����..." + $TAB + "Ctrl+G"
    ' SEARCH MENU BOOKMARKS
    CASE %IDM_TOGGLEBOOKMARK :BmpNum = 37 : FUNCTION = "���ñ�ǩ" + $TAB + "Ctrl+F2"
    CASE %IDM_NEXTBOOKMARK :  BmpNum = 38 : FUNCTION = "��һ��ǩ" + $TAB + "F2"
    CASE %IDM_PREVIOUSBOOKMARK : BmpNum = 39 : FUNCTION = "��һ��ǩ" + $TAB + "Shift+F2"
    CASE %IDM_DELETEBOOKMARKS :BmpNum = 40 : FUNCTION = "�Ƴ���ǩ" + $TAB + "Ctrl+Shift+F2"
    ' SEARCH MENU FILEFIND
    CASE %IDM_FINDINFILES :    BmpNum = 123 : FUNCTION = "���ļ��в���..." + $TAB + "Ctrl+Shift+F"
    CASE %IDM_EXPLORER :       BmpNum = 116 : FUNCTION = "��Դ������..." + $TAB + "Ctrl+Shift+O"
    CASE %IDM_WINDOWSFIND :    BmpNum = 122 : FUNCTION = "ϵͳ����..." + $TAB + "Ctrl+Alt+F"
    CASE %IDM_FILEFIND :       BmpNum = 41 : FUNCTION = "�����ļ�..." + $TAB + "Alt+Shift+F"

    ' ���в˵�
    'CASE %IDM_RUNHEADER :       BmpNum = - 1 : FUNCTION = "����"
    CASE %IDM_COMPILE :       BmpNum = 42 : FUNCTION = "����" + $TAB + "Ctrl+F5"
    CASE %IDM_COMPILERUN :    BmpNum = 43 : FUNCTION = "���벢ִ��" + $TAB + "F5"
    CASE %IDM_COMPILEDEBUG :  BmpNum = 44 : FUNCTION = "���벢����" + $TAB + "F6"
    CASE %IDM_EXECUTE :       BmpNum = 91 : FUNCTION = "ִ��" + $TAB + "Ctrl+E"
    CASE %IDM_SETPRIMARY :    BmpNum = 45 : FUNCTION = "������Դ�ļ�" + $TAB + "Ctrl+Alt+Y"
    CASE %IDM_CMDPARAMETER :  BmpNum = 46 : FUNCTION = "���в���"
    CASE %IDM_DEBUGTOOL :     BmpNum = 47 : FUNCTION = "���Թ���"

    ' ��ͼ�˵�
    CASE %IDM_PROPERTY :      BmpNum =-1 : FUNCTION = "����"
    CASE %IDM_CONTROLS :      BmpNum = -1 : FUNCTION = "�ؼ�"
    CASE %IDM_PROJECT :       BmpNum =-1 : FUNCTION = "��Ŀ"
    CASE %IDM_COMPILERS :     BmpNum =-1 : FUNCTION = "������"
    CASE %IDM_FINDRS :        BmpNum =-1 : FUNCTION = "���ҽ��"
    CASE %IDM_REBAR :         BmpNum =133 : FUNCTION = "������"
    CASE %IDM_STDTOOLBAR:     BmpNum = -1 : FUNCTION = "��������"
    CASE %IDM_WINDOWBAR:     BmpNum = -1 : FUNCTION = "���ڱ༭"
    CASE %IDM_COMBOBAR:       BmpNum = -1 : FUNCTION = "�������"
    CASE %IDM_DEBUGBAR:       BmpNum = -1 : FUNCTION = "�������"
    CASE %IDM_CODEEDIT:       BmpNum = -1 : FUNCTION = "����༭"
    CASE %IDM_WINARRAG:       BmpNum = -1 : FUNCTION = "��������"
    CASE %ID_LOCKBAR:         BmpNum = -1 : FUNCTION = "����"
    CASE %IDM_TOGGLE :        BmpNum = 48 : FUNCTION = "չ��/�������" + $TAB + "F8"
    CASE %IDM_TOGGLEALL :     BmpNum = 49 : FUNCTION = "չ��/��������" + $TAB + "Ctrl+F8"
    CASE %IDM_FOLDALL :       BmpNum = 51 : FUNCTION = "��������" + $TAB + "Alt+F8"
    CASE %IDM_EXPANDALL :     BmpNum = 50 : FUNCTION = "չ������" + $TAB + "Shift+F8"
    CASE %IDM_ZOOMIN :        BmpNum = 52 : FUNCTION = "��С�����ı�" + $TAB + "Ctrl++"
    CASE %IDM_ZOOMOUT :       BmpNum = 53 : FUNCTION = "�Ŵ�����ı�" + $TAB + "Ctrl+-"
    CASE %IDM_USETABS :       BmpNum = - 1 : FUNCTION = "TAB����" + $TAB + "Ctrl+Shift+T"
    CASE %IDM_AUTOINDENT :    BmpNum = - 1 : FUNCTION = "�Զ�����" + $TAB + "Ctrl+Shift+A"
    CASE %IDM_SHOWLINENUM :   BmpNum = - 1 : FUNCTION = "����" + $TAB + "Ctrl+Shift+L"
    CASE %IDM_SHOWMARGIN :    BmpNum = - 1 : FUNCTION = "�շ���" + $TAB + "Ctrl+Shift+M"
    CASE %IDM_SHOWINDENT :    BmpNum = - 1 : FUNCTION = "����������" + $TAB + "Ctrl+Shift+I"
    CASE %IDM_SHOWSPACES :    BmpNum = - 1 : FUNCTION = "�հ�" + $TAB + "Ctrl+Shift+W"
    CASE %IDM_SHOWEOL :       BmpNum = - 1 : FUNCTION = "��ĩ" + $TAB + "Ctrl+Shift+D"
    CASE %IDM_SHOWEDGE :      BmpNum = - 1 : FUNCTION = "��Ե" + $TAB + "Ctrl+Shift+G"
    CASE %IDM_SHOWPROCNAME :  BmpNum = - 1 : FUNCTION = "��ʾ��������" + $TAB + "Ctrl+Shift+S"
    CASE %IDM_CVEOLTOCRLF :   BmpNum = 54 : FUNCTION = "ת����ĩ�ַ�Ϊ $CRLF"
    CASE %IDM_CVEOLTOCR :     BmpNum = 55 : FUNCTION = "ת����ĩ�ַ�Ϊ $CR"
    CASE %IDM_CVEOLTOLF :     BmpNum = 56 : FUNCTION = "ת����ĩ�ַ�Ϊ $LF"
    CASE %IDM_REPLSPCWITHTABS:BmpNum = 120 : FUNCTION = "TAB�滻�ո�"
    CASE %IDM_REPLTABSWITHSPC:BmpNum = 119 : FUNCTION = "�ո��滻TAB"
    CASE %IDM_FILEPROPERTIES :BmpNum = 127 : FUNCTION = "�ļ�����"
    CASE %IDM_SYSINFO :       BmpNum = 128 : FUNCTION = "ϵͳ��Ϣ"

    ' ���ڲ˵�
    CASE %IDM_CASCADE :       BmpNum = 57 : FUNCTION = "�������"
    CASE %IDM_TILEH :         BmpNum = 58 : FUNCTION = "ˮƽ����"
    CASE %IDM_TILEV :         BmpNum = 59 : FUNCTION = "��ֱ����"
    CASE %IDM_RESTOREWSIZE :  BmpNum = 60 : FUNCTION = "���������ڳߴ�"
    CASE %IDM_SWITCHWINDOW :  BmpNum = 61 : FUNCTION = "�л�����    " + $TAB + "Ctrl+W"
    CASE %IDM_ARRANGE :       BmpNum = 62 : FUNCTION = "����ͼ��"
    CASE %IDM_CLOSEWIN :      BmpNum = 63 : FUNCTION = "�ر�"
    CASE %IDM_MAXWIN :        BmpNum = 93 : FUNCTION = "���"

    ' ��Ŀ�˵�
    CASE %IDM_NEWPROJECT :    BmpNum = 1 : FUNCTION = "�½���Ŀ"
    CASE %IDM_OPENPROJECT :   BmpNum = 2 : FUNCTION = "����Ŀ"
    CASE %IDM_LOADPROJECT :   BmpNum = 101 : FUNCTION = "������Ŀ"
    CASE %IDM_CLOSEPROJECT :  BmpNum = 5 : FUNCTION = "�ر���Ŀ"
    CASE %IDM_SAVEPROJECT :   BmpNum = 3 : FUNCTION = "������Ŀ"
    CASE %IDM_SAVEPROJECTAS : BmpNum = 4 : FUNCTION = "�����Ŀ..."
    CASE %IDM_REGPROJECTEXT : BmpNum = 46 : FUNCTION = "������Ŀ�ļ�"
    CASE %IDM_ACHIVE :        BmpNum = 83 : FUNCTION = "�鵵��Ŀ"
    CASE %IDM_NEWVISION :     BmpNum = 95 : FUNCTION = "��Ŀ����"
    CASE %IDM_FILEVISIONSETTING : BmpNum = 127 : FUNCTION = "�ļ��汾"
    CASE %IDM_VISIONPACKAGE :     BmpNum = 121 : FUNCTION = "��Ŀ�汾"
    CASE %IDM_INSERTINPROJECT :   BmpNum = 112 : FUNCTION = "�����ļ�����Ŀ"
'    CASE %IDM_RECENTPRJ1 TO %IDM_RECENTPRJ4
'      BmpNum = 2
'      IF ( ItemId - %IDM_RECENTPROJECTS ) = > LBOUND( RecentProjects ) AND _
'              ( ItemId - %IDM_RECENTPROJECTS ) < = UBOUND( RecentProjects ) THEN _
'              FUNCTION = "&" & FORMAT$( ItemId - %IDM_RECENTPROJECTS ) & " " & RecentProjects( ItemId - %IDM_RECENTPROJECTS )

    ' ���ò˵�
    'CASE %IDM_OPTIONSHEADER :     BmpNum = - 1 : FUNCTION = "ѡ��"
    CASE %IDM_INTERFACEOPT :  BmpNum = 129 : FUNCTION = "��������"
    CASE %IDM_FILEEXTOPT :    BmpNum = 102 : FUNCTION = "�����ļ�"
    CASE %IDM_EDITOROPT :     BmpNum = 64 : FUNCTION = "����༭"
    CASE %IDM_COLORSOPT :     BmpNum = 66 : FUNCTION = "������ɫ����"
    CASE %IDM_FORMOROPT :     BmpNum = 78 : FUNCTION = "����༭"
    CASE %IDM_COMPILEROPT :   BmpNum = 65 : FUNCTION = "��������"
    CASE %IDM_TOOLSOPT :      BmpNum = 67 : FUNCTION = "��������"
    CASE %IDM_FOLDINGOPT :    BmpNum = 100 : FUNCTION = "��������"
    CASE %IDM_PRINTOPT :      BmpNum = 9 : FUNCTION = "��ӡ����"
    CASE %IDM_SHORTKEYOPT :   BmpNum = 105 : FUNCTION = "�ȼ�����"

    ' ���߲˵�
    'CASE %IDM_TOOLSHEADER :     BmpNum = - 1 : FUNCTION = "����"
    CASE %IDM_MAKEMENU :      BmpNum = 69 : FUNCTION = "�����˵�"
    CASE %IDM_MAKETOOLBAR :   BmpNum = 69 : FUNCTION = "�����˵�"
    CASE %IDM_GOTCHA :        BmpNum = 123 : FUNCTION = "��ȷ����"
    CASE %IDM_FINDMULTIDIR :  BmpNum = 123 : FUNCTION = "��Ŀ¼����"
    CASE %IDM_REPLACEMULTIDIR : BmpNum = 123 : FUNCTION = "��Ŀ¼�滻"
    CASE %IDM_FORMATCODE :    BmpNum = 123 : FUNCTION = "��ʽ������"
    CASE %IDM_CODEC :         BmpNum = 69 : FUNCTION = "��������"
    CASE %IDM_CODECOUNT :     BmpNum = 69 : FUNCTION = "����ͳ��"
    CASE %IDM_CODEMANAGER :   BmpNum = 69 : FUNCTION = "�������"
    CASE %IDM_CODEKEEPER :    BmpNum = 71 : FUNCTION = "���뱣��"
    CASE %IDM_CODEMANAGER :   BmpNum = 69 : FUNCTION = "�������"
    CASE %IDM_KBDMACROS :     BmpNum = 129 : FUNCTION = "�����ݼ�"
    CASE %IDM_INCLEAN :       BmpNum = 82 : FUNCTION = "�������"
    CASE %IDM_WINSPY :        BmpNum = 82 : FUNCTION = "SPY����"
    CASE %IDM_CHARMAP :       BmpNum = 117 : FUNCTION = "�ַ���"
    CASE %IDM_MSGBOXDESIGNER :BmpNum = 121 : FUNCTION = "��Ϣ�����"
    CASE %IDM_PBFORMS :       BmpNum = 72 : FUNCTION = "PB����"
    CASE %IDM_PBCOMBR :       BmpNum = 73 : FUNCTION = "PB COM ���"
    CASE %IDM_TYPELIBBR :     BmpNum = 126 : FUNCTION = "PB LIB ���"
    CASE %IDM_IMGEDITOR :     BmpNum = 76 : FUNCTION = "�༭ͼƬ"
    CASE %IDM_RCEDITOR :      BmpNum = 77 : FUNCTION = "�༭��Դ"
    CASE %IDM_CODETIPSBUILDER : BmpNum = 70 : FUNCTION = "������ʾ����"
    CASE %IDM_CODETYPEBUILDER : BmpNum = 98 : FUNCTION = "�������͹���"
    CASE %IDM_DLGEDITOR :       BmpNum = 75 : FUNCTION = "�Ի���༭"
    CASE %IDM_POFFS :         BmpNum = 80 : FUNCTION = "Poffs"
    CASE %IDM_COPYCAT :       BmpNum = 83 : FUNCTION = "CopyCat"
    CASE %IDM_CALCULATOR :    BmpNum = 84 : FUNCTION = "������"
    CASE %IDM_MORETOOLS :     BmpNum = 124 : FUNCTION = "����..."

    'Эͬ�˵�
    CASE %IDM_COOPSETTING :   BmpNum = 124 : FUNCTION = "Эͬ����"
    CASE %IDM_SENDINVITE :    BmpNum = 124 : FUNCTION = "����Эͬ"
    CASE %IDM_SENDVISIT :     BmpNum = 124 : FUNCTION = "�μ�Эͬ"
    CASE %IDM_DOWNLOADFILE :  BmpNum = 124 : FUNCTION = "�����ļ�"
    CASE %IDM_UPLOADFILE :    BmpNum = 124 : FUNCTION = "�ϴ��ļ�"
    CASE %IDM_CHAT :          BmpNum = 124 : FUNCTION = "�Ự"

    ' �����˵�
    'CASE %IDM_HELPHEADER :        BmpNum = - 1 : FUNCTION = "����"
    CASE %IDM_IDEHELP :       BmpNum = 85 : FUNCTION = "IDE ����"
    CASE %IDM_PBWINCONTENT :  BmpNum = 85 : FUNCTION = "PB/WIN ����"
    CASE %IDM_PBWININDEX :    BmpNum = 85 : FUNCTION = "PB/WIN ����"
    CASE %IDM_PBWINSEARCH :   BmpNum = 85 : FUNCTION = "PB/WIN ��ѯ"
    CASE %IDM_WINDOWSSDK :    BmpNum = 85 : FUNCTION = "WinSDK ����"
    CASE %IDM_ABOUT :         BmpNum = 86 : FUNCTION = "����"
    CASE %IDM_UPDATE :        BmpNum = 86 : FUNCTION = "����"

    '�������������������Ҽ��˵�
    CASE %IDM_WINTEMPLATES + 1
      BmpNum = 75
      FUNCTION = GetLang("Normal Window")'"һ�㴰��"
    CASE %IDM_WINTEMPLATES + 2
      BmpNum = 75
      FUNCTION = GetLang("MDI Window") '"MDI����"
    CASE %ID_SETH_LEFT
      BmpNum=143
      FUNCTION = GetLang("Left Align")'"�����"
    CASE %ID_SETH_CENTER
      BmpNum=135
      FUNCTION = GetLang("H-Center Align")'"ˮƽ����"
    CASE %ID_SETH_RIGHT
      BmpNum=144
      FUNCTION = GetLang("Right Align")'"�Ҷ���"
    CASE %ID_SETV_TOP
      BmpNum=145
      FUNCTION = GetLang("Top Align")'"�϶���"
    CASE %ID_SETV_CENTER
      BmpNum=134
      FUNCTION = GetLang("V-Center Align")'"��ֱ����"
    CASE %ID_SETV_BOTTOM
      BmpNum=146
      FUNCTION = GetLang("Bottom Align")'"�¶���"
    CASE %MN_F_DDTCODE
      BmpNum=-1
      FUNCTION = GetLang("DDT Code")'"DDT����"
    CASE %MN_F_SDKCODE
      BmpNum=-1
      FUNCTION = GetLang("SDK Code")'"SDK����"
    CASE %IDM_PBTEMPLATES + 1
      BmpNum=-1
      FUNCTION = GetLang("Normal PB Program")
    CASE %IDM_PBTEMPLATES + 2
      BmpNum=-1
      FUNCTION = GetLang("DLL Program")
    CASE %IDM_PBTEMPLATES + 3
      BmpNum=-1
      FUNCTION = GetLang("PB Dialog Program")
    CASE %IDM_PBTEMPLATES + 4
      BmpNum=-1
      FUNCTION = GetLang("PB Window Program")
    CASE %IDM_PBTEMPLATES + 5
      BmpNum=-1
      FUNCTION = GetLang("RC File")
    CASE %IDM_PBTEMPLATES + 6
      BmpNum=-1
      FUNCTION = GetLang("Plain Text File")
    CASE %IDM_GOTOSELFILE     ' Ownerdraw context menu (include files)
      BmpNum = 2 : FUNCTION = strContextMenu
    CASE %IDM_GOTOSELPROC     ' Ownerdraw context menu (Find Subs/Functions)
      BmpNum = 33 : FUNCTION = strContextMenu
    CASE %IDM_GOTOLASTPOSITION    ' Ownerdraw context menu (Goto last position)
      BmpNum = 36 : FUNCTION = "ת�����λ��"
  END SELECT
END FUNCTION
' *********************************************************************************************
' Function    : GetMenuBmpHandle ()
' Description : Get desired Menu Item's bitmap handle.
' *********************************************************************************************
FUNCTION GetMenuBmpHandle( BYVAL BmpNum AS LONG, BYVAL nState AS LONG ) AS LONG
  LOCAL hBmp1 AS DWORD
  LOCAL hBmp2 AS DWORD
  LOCAL hDC AS DWORD
  LOCAL memDC1 AS DWORD
  LOCAL memDC2 AS DWORD
  '----------------------------------------------------------------------------
  ' Load the Bitmap Strips and create the ImageLists.
  '----------------------------------------------------------------------------
  ' Disabled bitmap strip.
  IF nState = %IMG_DIS THEN
    hBmp1 = LOADIMAGE( GETMODULEHANDLE( "" ), "WINTBDIS", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE )
  END IF
  ' Normal bitmap strip.
  IF nState = %IMG_NOR THEN
    hBmp1 = LOADIMAGE( GETMODULEHANDLE( "" ), "WINTBNOR", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE )
  END IF
  ' Hot Selected bitmap strip.
  IF nState = %IMG_HOT THEN
    hBmp1 = LOADIMAGE( GETMODULEHANDLE( "" ), "WINTBHOT", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE )
  END IF
  '----------------------------------------------------------------------------
  ' Create memory DC's and draw the bitmaps.
  '----------------------------------------------------------------------------
  hdc = GETDC( 0 )    ' GET device context FOR screen
  memDC1 = CREATECOMPATIBLEDC( hDC )    ' Create two memory device contexts
  memDC2 = CREATECOMPATIBLEDC( hDC )
  hBmp2 = CREATECOMPATIBLEBITMAP( hDC, 16, 16 )     ' Create compatible 16x16 pixel bmp
  RELEASEDC 0, hdc    ' Release the DC - don't need it anymore.
  SELECTOBJECT memDC1, hBmp1    ' Select the bitmaps into the mem dc's
  SELECTOBJECT memDC2, hBmp2
  BITBLT memDC2, 0, 0, 16, 16, _    ' copy bNum part of hBmp1 strip to hBmp2
          memDC1, BmpNum * 16, 0, %SRCCOPY
  DELETEDC memDC1     ' Cleanup memDC's and delete main bmp strip object
  DELETEDC memDC2
  DELETEOBJECT hBmp1    ' (note: 16x16 pixel hBmp2 is deleted in DrawMenu routine)
  FUNCTION = hBmp2    ' return handle to 16x16 pixel bitmap
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' �Ի�˵��� - �������ڵ�%WM_DRAWITEM ��Ϣ����
' *********************************************************************************************
SUB DrawMenu( BYREF lp AS LONG )
  REGISTER c          AS LONG     ' ���� for/next ����
  LOCAL ShortCutWidth AS LONG     ' ��ݷ�ʽ�ı�����(�����)
  LOCAL ID            AS LONG     ' ��ʱ�洢�˵���ID
  LOCAL BmpNum        AS LONG     ' λͼ����ֵ
  LOCAL hBmp          AS LONG     ' Ҫ���Ƶ�ͼ����
  LOCAL hBrush        AS LONG     ' ѡ��ʱ�Ļ���ˢ
  LOCAL bHighlight    AS LONG     ' ������ʶ
  LOCAL drawTxtFlags  AS DWORD    ' �����ı���ʽ��ʶ
  LOCAL drawBmpFlags  AS DWORD    ' ����ͼ�귽ʽ��ʶ
  LOCAL menuBtnW      AS LONG     ' �˵���ť���ؿ��
  LOCAL sCaption      AS STRING   ' �˵����ı��ַ���
  LOCAL sShortCut     AS STRING   ' ��ݷ�ʽ�ı�
  LOCAL rc            AS RECT     ' ��������
  LOCAL szl           AS SIZEL    ' ���ڼ����ı�����
  LOCAL lpDis AS DRAWITEMSTRUCT PTR ' we fill this from received pointer
  LOCAL bGrayed       AS LONG
  LOCAL bDisabled     AS LONG
  LOCAL nState        AS LONG     ' For ImageList Bitmap Strip ID.
  lpDis = lp    ' �Ӹ���ָ���л�ȡ�ṹ
  menuBtnW = GETSYSTEMMETRICS( %SM_CYMENU ) - 1     ' ���� "��ť" �ߴ�
  '----------------------------------------------------------------------------
  ' �����ݷ�ʽ����λ�á���Ҫ������������в˵����������ȣ�
  ' �������ȼ�Ϊ���п�ݷ�ʽ�ı����λ�õ����
  '----------------------------------------------------------------------------
  FOR c = 0 TO GETMENUITEMCOUNT( @lpDis.hwndItem ) - 1 ' hwndItem �Ǳ������Ĳ˵�����
    ID = GETMENUITEMID( @lpDis.hwndItem, c )
    IF ID THEN
      sCaption = GetMenuTxtBmp( ID, 0 )
      IF INSTR( sCaption, $TAB ) THEN     ' ����п�ݷ�ʽ�ı� (Ctrl+X, etc.)
        sShortCut = TRIM$( PARSE$( sCaption, $TAB, 2 )) 'ȡ��ݷ�ʽ�ı�
      ELSE
        sShortCut = ""
      END IF
      IF LEN( sShortCut ) THEN
        GETTEXTEXTENTPOINT32 @lpDis.hDC, BYVAL STRPTR( sShortCut ), _
                LEN( sShortCut ), BYVAL VARPTR( szl )
        ShortCutWidth = MAX& ( ShortCutWidth, szl.cx )
      END IF
    END IF
  NEXT
  '----------------------------------------------------------------------------
  ' ��ȡ�˵����ַ������ֳ����յ��ı�����ݷ�ʽ
  '----------------------------------------------------------------------------
  sCaption = GetMenuTxtBmp( @lpDis.itemID, BmpNum )
  IF sCaption="" THEN
    'control get text g_hWndMain,@lpDis.itemID to sCaption
    'msgbox sCaption
    sCaption="  "
  END IF
  IF INSTR( sCaption, $TAB ) THEN     ' if it has shortcut (Ctrl+X, etc.)
    '      sShortCut = TRIM$(PARSE$(sCaption, $TAB, 2))
    sShortCut = PARSE$( sCaption, $TAB, 2 )
    sCaption = TRIM$( PARSE$( sCaption, $TAB, 1 ))
  ELSE
    sCaption = TRIM$(sCaption)
    'sCaption = sCaption
    sShortCut = ""
  END IF
  'msgbox sCaption
  '----------------------------------------------------------------------------
  ' ����˵�����ַ����ߴ�
  '----------------------------------------------------------------------------
  GETTEXTEXTENTPOINT32 @lpDis.hDC, BYVAL STRPTR( sCaption ), LEN( sCaption ), szl
  '----------------------------------------------------------------------------
  ' ���ݶ���������ɫ
  '----------------------------------------------------------------------------
  hMenuTextBkBrush = CREATESOLIDBRUSH( GETSYSCOLOR(%COLOR_HIGHLIGHTTEXT))'SciColorsAndFonts.SubmenuTextBackColor )
  ' Creates the brush for the submenu highlighted text background �����Ӳ˵������ı�������ˢ��
  hMenuHiBrush = CREATESOLIDBRUSH( GETSYSCOLOR(%COLOR_MENUHILIGHT))'SciColorsAndFonts.SubmenuHiTextBackColor )
  'hMenuHiBrush = CreateSolidBrush(RGB(64, 134, 191))
  ' Creates the brush for the menu icons background �����˵�ͼ�걳��ˢ��
  hMenuIconsBrush = CREATESOLIDBRUSH( RGB( 193, 211, 239 ))
  IF ( @lpDis.itemState AND %ODS_SELECTED ) THEN    ' �˵��ѡ��
    nState = %IMG_HOT
    hBrush = hMenuHiBrush
    SETBKCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_HIGHLIGHTTEXT ) '%COLOR_HIGHLIGHT )
    SETTEXTCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_HIGHLIGHTTEXT )
    bHighlight = %TRUE    ' ��ʶ�������ں���������ӻ��Ƴ�ͼ����Χ�ı߿�
    ' ��ѡ�У������ǻ�ɫ�Ļ򲻿��õ�
    IF ( @lpDis.itemState AND %ODS_GRAYED ) THEN
      nState = %IMG_DIS
      bGrayed = %TRUE
    ELSEIF ( @lpDis.itemState AND %ODS_DISABLED ) THEN
      nState = %IMG_DIS
      bDisabled = %TRUE
    END IF
  ELSEIF ( @lpDis.itemState AND %ODS_GRAYED ) THEN    ' �˵����ǻҵ�
    nState = %IMG_DIS
    hBrush = GETSYSCOLORBRUSH( %COLOR_MENU )
    SETBKCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENU )
    SETTEXTCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENUTEXT )
    bGrayed = %TRUE
  ELSEIF ( @lpDis.itemState AND %ODS_DISABLED ) THEN    ' �˵������
    nState = %IMG_DIS
    hBrush = GETSYSCOLORBRUSH( %COLOR_MENU )
    SETBKCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENU )
    SETTEXTCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENUTEXT )
    bDisabled = %TRUE
  ELSE    ' δѡ��򲻿���
    nState = %IMG_NOR
    hBrush = GETSYSCOLORBRUSH( %COLOR_MENU )
    SETBKCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENU )
    SETTEXTCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENUTEXT )
  END IF
  drawTxtFlags = %DST_PREFIXTEXT
  drawBmpFlags = %DST_BITMAP
  IF ISTRUE bGrayed OR ISTRUE bDisabled THEN    ' if grayed or disabled item
    IF ( @lpDis.itemState AND %ODS_SELECTED ) THEN    ' if it's selected
      SETTEXTCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_GRAYTEXT )
    ELSE
      drawTxtFlags = drawTxtFlags OR %DSS_DISABLED
    END IF
  END IF
  '----------------------------------------------------------------------------
  ' Calculate rect for highlight area (selected item) and fill rect with proper
  ' color, either COLOR_MENU or COLOR_HIGHLIGHT, depending on selection state.
  '----------------------------------------------------------------------------
  rc = @lpDis.rcItem
  ' Draw the background of the icons
  FILLRECT @lpDis.hDC, rc, hMenuIconsBrush
  IF ISTRUE bGrayed OR ISTRUE bDisabled OR BmpNum = - 1 THEN    ' grayed or no bitmap
    rc.nLeft = 0
  ELSE
    rc.nLeft = menuBtnW + 4     ' enabled, with bitmap
  END IF
  ' For GRAYED With Bitmap.
  IF ISTRUE bGrayed OR ISTRUE bDisabled THEN
    IF BmpNum = > 0 THEN
      rc.nLeft = menuBtnW + 4
      hBrush = GETSYSCOLORBRUSH( %COLOR_MENU )
      drawTxtFlags = drawTxtFlags OR %DSS_DISABLED
    END IF
  END IF
  '   FillRect @lpDis.hDC, rc, hBrush
  IF @lpDis.itemId = 0 THEN 'OR @lpDis.itemId = %IDM_RECENTSEPARATOR OR _
         ' @lpDis.itemId = %IDM_RECENTPROJECTSSEPARATOR THEN     ' Separator
    rc.nLeft = menuBtnW + 4
    FILLRECT @lpDis.hDC, rc, hMenuTextBkBrush
    rc.nLeft = rc.nLeft + 4
    rc.nTop = rc.nTop + ( rc.nBottom - rc.nTop ) \ 2
    DRAWEDGE @lpdis.hDC, rc, %EDGE_ETCHED, %BF_TOP
    EXIT SUB
  END IF
  '----------------------------------------------------------------------------
  ' Draw the disabled bitmap.
  '----------------------------------------------------------------------------
  IF ISTRUE bGrayed OR ISTRUE bDisabled THEN
    ' Draw Disabled Bitmap.
    IF BmpNum = > 0 THEN hBmp = GetMenuBmpHandle( BmpNum, nState )
    IF nState = %IMG_DIS THEN
      IMAGELIST_DRAWEX( hImlDis, BmpNum, @lpDis.hDC,(( menuBtnW + 3 ) \ 2 ) - 8, _
              (( @lpDis.rcItem.nTop + @lpDis.rcItem.nBottom ) \ 2 ) - 8, _
              16, 16, %CLR_NONE, %CLR_NONE, %ILD_TRANSPARENT )
    END IF
  END IF
  '----------------------------------------------------------------------------
  ' Draw the bitmap "button", if item is not grayed out (disabled).
  '----------------------------------------------------------------------------
  ' ImageList Draw Style Constants.
  ' %IMG_DIS = 0 : %IMG_NOR = 1 : %IMG_HOT = 2
  IF ( @lpdis.itemState AND %ODS_CHECKED ) THEN
    ' Make adjustements for checked items
    BmpNum = %OMENU_CHECKEDICON
    hBmp = GetMenuBmpHandle( %OMENU_CHECKEDICON, nState )
    nState = %IMG_NOR
  ELSE
    IF BmpNum = > 0 THEN hBmp = GetMenuBmpHandle( BmpNum, nState )
  END IF
  IF hBmp <> 0 THEN     ' ����λͼ��������"��ť"���м�
    IF nState = %IMG_NOR THEN
      IMAGELIST_DRAWEX( hImlNor, BmpNum, @lpDis.hDC,(( menuBtnW + 3 ) \ 2 ) - 8, _
              (( @lpDis.rcItem.nTop + @lpDis.rcItem.nBottom ) \ 2 ) - 8, _
              16, 16, %CLR_NONE, %CLR_NONE, %ILD_TRANSPARENT )
    ELSEIF nState = %IMG_HOT THEN
      IMAGELIST_DRAWEX( hImlHot, BmpNum, @lpDis.hDC,(( menuBtnW + 3 ) \ 2 ) - 8, _
              (( @lpDis.rcItem.nTop + @lpDis.rcItem.nBottom ) \ 2 ) - 8, _
              16, 16, %CLR_NONE, %CLR_NONE, %ILD_TRANSPARENT )
    ELSEIF nState = %IMG_DIS THEN
      IMAGELIST_DRAWEX( hImlDis, BmpNum, @lpDis.hDC,(( menuBtnW + 3 ) \ 2 ) - 8, _
              (( @lpDis.rcItem.nTop + @lpDis.rcItem.nBottom ) \ 2 ) - 8, _
              16, 16, %CLR_NONE, %CLR_NONE, %ILD_TRANSPARENT )
    END IF
    '----------------------------------------------------------------------
    ' Ϊ3D�߿������Ҫ��RECT
    '----------------------------------------------------------------------
    rc = @lpDis.rcItem
    rc.nLeft = 0    ' Size and pos for "button"..
    rc.nRight = menuBtnW + 4
    rc.nBottom = rc.nTop + GETSYSTEMMETRICS( %SM_CYMENU ) + 1
    '----------------------------------------------------------------------
    ' Draw "button" flat
    '----------------------------------------------------------------------
    'DrawEdge @lpDis.hDC, rc, %BDR_RAISEDINNER, %BF_FLAT OR %BF_RECT
  END IF
  DELETEOBJECT hBmp     ' ���ʱɾ��λͼ�������ڴ�й©
  '----------------------------------------------------------------------------
  ' ���Ʋ˵��ı���λͼ��ť�����������
  '----------------------------------------------------------------------------
  rc = @lpDis.rcItem
  rc.nLeft = menuBtnW + 4 : rc.nRight = @lpDis.rcItem.nRight
  rc.nTop = rc.nTop + (( rc.nBottom - rc.nTop ) - szl.cy ) \ 2
  '   rc.nTop = rc.nTop - 3
  rc.nTop = rc.nTop - 4
  IF ISTRUE bHighlight THEN
    FILLRECT @lpDis.hDC, rc, hMenuHiBrush
    IF ISTRUE bGrayed OR ISTRUE bDisabled THEN
      SETTEXTCOLOR @lpdis.hDC, GETSYSCOLOR ( %COLOR_GRAYTEXT )
    ELSE
      SETTEXTCOLOR @lpdis.hDC, GETSYSCOLOR ( %COLOR_HIGHLIGHTTEXT )'SciColorsAndFonts.SubMenuHiTextForeColor
    END IF
    'Draw frame using custom colored pen
    MOVETOEX @lpdis.hDC, @lpdis.rcItem.nLeft, @lpdis.rcItem.nTop, BYVAL 0
    LINETO @lpdis.hDC, @lpdis.rcItem.nRight - 1, @lpdis.rcItem.nTop
    LINETO @lpdis.hDC, @lpdis.rcItem.nRight - 1, @lpdis.rcItem.nBottom - 1
    LINETO @lpdis.hDC, @lpdis.rcItem.nLeft, @lpdis.rcItem.nBottom - 1
    LINETO @lpdis.hDC, @lpdis.rcItem.nLeft, @lpdis.rcItem.nTop
  ELSEIF ISTRUE bGrayed OR ISTRUE bDisabled THEN
    FILLRECT @lpDis.hDC, rc, hMenuTextBkBrush
    SETTEXTCOLOR @lpdis.hDC, GETSYSCOLOR ( %COLOR_GRAYTEXT )
  ELSE
    FILLRECT @lpDis.hDC, rc, hMenuTextBkBrush
    SETTEXTCOLOR @lpdis.hDC, GETSYSCOLOR(%COLOR_MENUTEXT)'SciColorsAndFonts.SubMenuTextForeColor
  END IF
  SETBKMODE @lpdis.hDC, %TRANSPARENT
  rc.nTop = rc.nTop + 3
  rc.nLeft = menuBtnW + 8
  DRAWSTATE @lpDis.hDC, hMenuTextBkBrush, 0, BYVAL STRPTR( sCaption ), LEN( sCaption ), _
          rc.nLeft, rc.nTop, rc.nRight, szl.cy, drawTxtFlags
  IF ShortCutWidth THEN     ' if there's shortcut text (Ctrl+N, etc)
    DRAWSTATE @lpDis.hDC, 0, 0, BYVAL STRPTR( sShortCut ), LEN( sShortCut ), _
            @lpDis.rcItem.nRight - ShortCutWidth - 8, rc.nTop, _
            rc.nRight, szl.cy, drawTxtFlags
  END IF
'  if sCaption="�½�" then
'    msgbox "test" & str$(id)
'  end if
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Sub Routine : MeasureMenu ()
' Description : ��%WM_MEASUREITEM ��Ϣ����
' *********************************************************************************************
SUB MeasureMenu( BYVAL hWnd AS DWORD, BYVAL lParam AS DWORD )
  LOCAL hDC AS DWORD
  LOCAL hf AS DWORD
  LOCAL txt AS STRING
  LOCAL sl AS SIZEL
  LOCAL lpMis AS MEASUREITEMSTRUCT PTR
  LOCAL ncm AS NONCLIENTMETRICS
  lpMis = lParam    ' lParam ָ�� MEASUREITEMSTRUCT
  txt = GetMenuTxtBmp( @lpMis.itemID, 0 )    ' ��ȡ�˵�����ı�
  hDC = GETDC( hWnd )     ' ��ȡ�Ի��� DC
  ncm.cbSize = SIZEOF( ncm )    ' ��ȡ�˵�����
  SYSTEMPARAMETERSINFO %SPI_GETNONCLIENTMETRICS, SIZEOF( ncm ), BYVAL VARPTR( ncm ), 0
  IF LEN( ncm.lfMenuFont ) THEN
    hf = CREATEFONTINDIRECT( ncm.lfMenuFont )     ' �Բ˵�LogFont���ݴ�������
    IF hf THEN hf = SELECTOBJECT( hDC, hf )     ' ��DC��ѡ������
  END IF
  GETTEXTEXTENTPOINT32 hDC, BYVAL STRPTR( txt ), LEN( txt ), sl     ' ��ȡ�ı�����
  '     @lpMis.itemWidth  = sl.cx                                 ' set width
  @lpMis.itemWidth = sl.cx + %OMENU_EXTRAWIDTH    ' ���ÿ��
  @lpMis.ItemHeight = GETSYSTEMMETRICS( %SM_CYMENU ) + 1    ' ���ø߶�
  IF hf THEN DELETEOBJECT SELECTOBJECT( hDC, hf )     ' ɾ����ʱ�˵��������
  RELEASEDC hWnd, hDC     ' �ͷ�DC
  IF @lpMis.itemId = 0 THEN 'OR @lpMis.itemId = %IDM_RECENTSEPARATOR OR @lpMis.itemId = %IDM_RECENTPROJECTSSEPARATOR THEN     ' Separator
    @lpMis.itemHeight = @lpMis.itemHeight \ 2
  END IF
END SUB
' ����Rebar������
FUNCTION LockBar(BYVAL sign AS INTEGER) AS LONG
  LOCAL BandCount AS INTEGER
  LOCAL i AS INTEGER
  LOCAL rbbi AS REBARBANDINFO
  BandCount=SendMessage(ghRebar,%RB_GETBANDCOUNT,0,0)
  IF BandCount<=0 THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  rbbi.cbSize=SIZEOF(rbbi)
  rbbi.fMask=%RBBIM_STYLE
  IF sign=0 THEN     '0Ϊ����, 1Ϊ����
    FOR i=0 TO BandCount-1
      SendMessage ghRebar,%RB_GETBANDINFO,i,VARPTR(rbbi)
      rbbi.fStyle = rbbi.fStyle OR %RBBS_GRIPPERALWAYS
      rbbi.fStyle = rbbi.fStyle AND NOT(%RBBS_NOGRIPPER)
      SendMessage ghRebar,%RB_SETBANDINFO,i,VARPTR(rbbi)
    NEXT i
  ELSE
    FOR i=0 TO BandCount-1
      SendMessage ghRebar,%RB_GETBANDINFO,i,VARPTR(rbbi)
      rbbi.fStyle = rbbi.fStyle AND NOT(%RBBS_GRIPPERALWAYS)
      rbbi.fStyle = rbbi.fStyle OR %RBBS_NOGRIPPER
      SendMessage ghRebar,%RB_SETBANDINFO,i,VARPTR(rbbi)
    NEXT i
  END IF
  FUNCTION=1
END FUNCTION
'------------------------------------------------------------------------------
' TOOLBAR.INC  ����
'------------------------------------------------------------------------------
