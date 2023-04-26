
'------------------------------------------------------------------------------
' �ļ����� : TOOLBAR.INC
'
' ���� : ����һ��Rebar�����ӹ��ߵ�����...
'
'               �������汾   - PB/WIN 7.02
'               Win32Api�汾 - Last Update: 27 August 2003.
'               ��������     -
'               ����ʱ��     -
'               Windowƽ̨   - Windows 95, 98, ME And XP.
'
'
' ����Դ�����õ�״̬��λͼ:
' http://www.glyfx.com/toolbarimages.html
'
' �Ҵ���λͼ���ĳ�������������ҵ�:
' http://www.gamani.com/
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' �����ͺ���.
'------------------------------------------------------------------------------

' ʹ��ReBar����ͼƬ.
%USE_RBIMAGE = %FALSE        ' ��Ϊ%TRUE��ʹ��ReBar����ͼƬ

' ʹ�õĹ�������ť��.
%TOOLBUTTONS = 12
%MENUBUTTONS = 6

' ReBar �� ToolBar ��ͬ. ����ʼ��3000��������...
%ID_REBAR    = %APP_REBARS
%ID_TOOLBAR  = %APP_TOOLBARS
%ID_COMBOBOX = %APP_CONTROLS
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

%I_IMAGENONE      = -2

'------------------------------------------------------------------------------
' ȫ�ֱ���.
'------------------------------------------------------------------------------
GLOBAL ghRebar    AS DWORD         ' ReBar�ؼ����.
GLOBAL ghToolBar  AS DWORD         ' �������ؼ����.
GLOBAL ghComboBox AS DWORD         ' ComboBox�ؼ����.

'------------------------------------------------------------------------------
' �������ӳ���.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���� : CreateToolBar ()
'
' ���� : �����������ؼ�.
'------------------------------------------------------------------------------
FUNCTION CreateToolBar (BYVAL hWnd AS DWORD) AS DWORD

  ' Ϊ�Ż��ٶ�ʹ��ע�����.
  REGISTER x AS LONG

  ' ����������.
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP

  DIM Tbb(0 TO %TOOLBUTTONS - 1) AS LOCAL TBBUTTON

  ' �����ȷ��λͼ�ߴ�!
  IF gBmpSize <> 24 THEN
	 IF gBmpSize <> 16 THEN
		' �û�ʹ���˲���ȷ��λͼ�ߴ�.
		MessageBox(hWnd, "������λͼ�ߴ����Ϊ16x16 �� 24x24!", "������ʾ", _
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
							 hInst, _
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
  FOR x = 0 TO %TOOLBUTTONS - 1
	' Ϊÿ����ť���ó�ʼ��״̬.
	Tbb(x).iBitmap   = 0
	Tbb(x).idCommand = 0
	Tbb(x).fsState   = %TBSTATE_ENABLED
	Tbb(x).fsStyle   = %TBSTYLE_BUTTON OR %TBSTYLE_AUTOSIZE
	Tbb(x).dwData    = 0
	Tbb(x).iString   = 0

	SELECT CASE AS LONG x
	  CASE 3,6,10
		' ������ť���.
		Tbb(x).fsStyle   = %TBSTYLE_SEP

	  CASE 0&
		Tbb(x).iBitmap   = 0
		Tbb(x).idCommand = %IDM_NEW
		Tbb(x).iString   = 2
		Tbb(x).fsState   = %TBSTATE_ENABLED ' ��������壬�������%TBSTATE_DISABLED

	  CASE 1&
		Tbb(x).iBitmap   = 1
		Tbb(x).idCommand = %IDM_OPEN
		Tbb(x).iString   = 1
		Tbb(x).fsState   = %TBSTATE_ENABLED ' ��������壬�������%TBSTATE_DISABLED

	  CASE 2&
		Tbb(x).iBitmap   = 2
		Tbb(x).idCommand = %IDM_SAVE
		Tbb(x).iString   = 2

	  CASE 4&
		Tbb(x).iBitmap   = 3
		Tbb(x).idCommand = %ID_PRINT
		Tbb(x).iString   = 3
		Tbb(x).fsState   = %TBSTATE_ENABLED

	  CASE 5&
		Tbb(x).iBitmap   = 4
		Tbb(x).idCommand = %ID_PRINTPREVIEW
		Tbb(x).iString   = 3

	  CASE 7&
		Tbb(x).iBitmap   = 5
		Tbb(x).idCommand = %IDM_CUT
		Tbb(x).iString   = 4

	  CASE 8&
		Tbb(x).iBitmap   = 6
		Tbb(x).idCommand = %IDM_COPY
		Tbb(x).iString   = 5

	  CASE 9&
		Tbb(x).iBitmap   = 7
		Tbb(x).idCommand = %IDM_PASTE
		Tbb(x).iString   = 6

	  CASE 11&
		Tbb(x).iBitmap   = 13
		Tbb(x).idCommand = %IDM_EMAIL
		Tbb(x).iString   = 7

	END SELECT
  NEXT x

  '----------------------------------------------------------------------------
  ' ��ʼ�����������ڲ����ʹ�����Ϣ.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_TOOLBAR '%IDB_STD_LARGE_COLOR ' ʹ��λͼͼƬ��ID����, %ID_TOOLBAR
  Tabm.hInst =  hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
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

  ' ����ͼƬ�б���������.
  SendMessage ghToolBar, %TB_SETIMAGELIST,         0, ghImlNor  ' ����
  SendMessage ghToolBar, %TB_SETDISABLEDIMAGELIST, 0, ghImlDis  ' ������
  SendMessage ghToolBar, %TB_SETHOTIMAGELIST,      0, ghImlHot  ' ��ѡ��

  ' ����ͼƬ�б�λͼ����������ť.
  SendMessage ghToolBar, %TB_ADDBITMAP, %TOOLBUTTONS, VARPTR(Tabm)

  ' ���ù�������ť.
  SendMessage ghToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage ghToolBar, %TB_ADDBUTTONS, %TOOLBUTTONS, VARPTR(Tbb(0))

  ' ǿ�ƹ�����Ϊ��ʼ�ߴ�.
  SendMessage ghToolBar, %WM_SIZE, 0, 0

  FUNCTION = ghToolBar            ' ���ع��������ھ��.

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

  dwBaseUnits = GetDialogBaseUnits()

  '----------------------------------------------------------------------------
  ' ����ComboBox�ؼ�����.
  '----------------------------------------------------------------------------
  ghComboBox = CreateWindowEx(0, _
							  "ComboBox", _
							  "", _
							  %WS_CHILD OR %WS_VISIBLE OR %CBS_DROPDOWNLIST OR _
							  %WS_VSCROLL OR %WS_CLIPSIBLINGS OR %WS_CLIPCHILDREN, _
							  (6 * LOWRD(dwBaseUnits)) / 4, _
							  (2 * HIWRD(dwBaseUnits)) / 8, _
							  (100 * LOWRD(dwBaseUnits)) / 4, _
							  (50 * HIWRD(dwBaseUnits)) / 8, _
							  hWnd, _
							  %ID_COMBOBOX, _
							  hInst, _
							  BYVAL %NULL)

  '----------------------------------------------------------------------------
  ' ��鴴��ComboBox����ʱ�Ĵ���.
  '----------------------------------------------------------------------------
  IF ISFALSE ghComboBox THEN         ' ����!
	 ' ����ComboBox����ʧ�ܣ�֪ͨ�û�.
	 FUNCTION = %TRUE
	 EXIT FUNCTION
  END IF

  '----------------------------------------------------------------------------
  ' ѭ������ComboBox�ı��ַ���.
  '----------------------------------------------------------------------------
  FOR i = 1 TO 20
	sStr = "�� " & FORMAT$(i) & " ��"
	SendMessage ghComboBox, %CB_ADDSTRING, 0&, STRPTR(sStr)
  NEXT

  SendMessage ghComboBox, %CB_SETCURSEL, 0&, 0&

  FUNCTION = ghComboBox

END FUNCTION

'------------------------------------------------------------------------------
' ���� : SaveRebarBandPos ()
'
' ���� : ����Rebar��λ��.
'------------------------------------------------------------------------------
FUNCTION SaveRebarBandPos() AS LONG

  DIM lResult AS LONG, lBands AS LONG
  DIM rBand   AS REBARBANDINFO

  REGISTER lCount AS LONG

  ' �õ�Rebar����.
  lResult = SendMessage(ghRebar, %RB_GETROWCOUNT, 0&, 0&)
  udtAp.rbRowCount = lResult

  ' �õ�Rebar����.
  lBands = SendMessage(ghRebar, %RB_GETBANDCOUNT, 0&, 0&)
  udtAp.rbBandCount = lResult

  ' ����Rebar��REBARBANDINFO�ṹ.
  rBand.cbSize = SIZEOF(rBand)
  rBand.fMask  = %RBBIM_ID

  ' ѭ���õ�ÿ�ε�ID.
  FOR lCount = 0 TO lBands - 1
	'lResult = SendMessage(hRebar, %RB_IDTOINDEX, lCount, VARPTR(rBand))
	lResult = SendMessage(ghRebar, %RB_GETBANDINFO, lCount, VARPTR(rBand))
	IF lCount = 0 THEN udtAp.rbBand0 = rBand.Wid
	IF lCount = 1 THEN udtAp.rbBand1 = rBand.Wid
  NEXT

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
  LOCAL szTbText  AS ASCIIZ * 255
  LOCAL szRbImage AS ASCIIZ * %MAX_PATH
  LOCAL dwBtnSize AS DWORD

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
						   hInst, _
						   BYVAL %NULL)

  '----------------------------------------------------------------------------
  ' ��鴴��Rebar�ؼ����ڵĴ���.
  '----------------------------------------------------------------------------
  IF ISFALSE ghRebar THEN         ' ����!
	 '= ����Rebar����ʧ�ܣ�֪ͨ�û�.
	 FUNCTION = %TRUE
	 EXIT FUNCTION
  END IF

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

  rbBand.fStyle = %RBBS_NOVERT     OR _    ' ����ʾ��ֱ����
				  %RBBS_CHILDEDGE  OR _
				  %RBBS_FIXEDBMP   OR _
				  %RBBS_GRIPPERALWAYS

  '----------------------------------------------------------------------------
  ' ����û�ѡ����һ��Rebar����ͼƬ���Ǽ�����!
  '----------------------------------------------------------------------------
  IF ISTRUE %USE_RBIMAGE THEN
	 ghRbBack   = LoadImage(hInst, "RBBACK", %IMAGE_BITMAP, 0, 0, _
							%LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR _
							%LR_DEFAULTSIZE)

	 rbBand.hbmBack    = ghRbBack           ' Rebar����ͼƬ.
  END IF

  '----------------------------------------------------------------------------
  ' ����Rebar��������.
  '----------------------------------------------------------------------------
  IF ISFALSE CreateToolBar(hWnd) THEN         ' ����!
	 ' λͼ�ߴ�Ƿ���֪ͨ�û�.
	 FUNCTION = %TRUE
	 EXIT FUNCTION
  END IF

  '----------------------------------------------------------------------------
  ' ��ʼ�������ӹ�������Rebar�Ķ�.
  '----------------------------------------------------------------------------
  dwBtnSize         = SendMessage(ghToolBar, %TB_GETBUTTONSIZE, 0, 0)
  'szTbText          = "ToolBar"
  rbBand.lpText     = VARPTR(szTbText)
  rbBand.hwndChild  = ghToolBar
  rbBand.wID        = %ID_TOOLBAR
  rbBand.cxMinChild = 230
  rbBand.cyMinChild = HIWRD(dwBtnSize) '+ 2
  rbBand.cx         = 280

  ' ����һ��ʱ�����϶ε�һ������.
  IF udtAp.rbRowCount > 1 THEN
	rbBand.fStyle = rbBand.fStyle OR %RBBS_BREAK
  END IF

  '----------------------------------------------------------------------------
  ' ���벢���ӹ������ε�Rebar�ؼ�����.
  '----------------------------------------------------------------------------
  SendMessage ghRebar, %RB_INSERTBAND, -1&, VARPTR(rbBand)


  '----------------------------------------------------------------------------
  ' ȥ��ע��: ����������ComboBox��Rebar����.
  '----------------------------------------------------------------------------

  ' ����Rebar��ComboBox��.
  CreateComboBox(hWnd)
  szCbText          = "����:"
  rbBand.lpText     = VARPTR(szCbText)
  rbBand.hwndChild  = ghComboBox
  rbBand.wID        = %ID_COMBOBOX
  GetWindowRect ghComboBox, rc
  rbBand.cxMinChild = 150
  rbBand.cyMinChild = rc.nBottom - rc.nTop
  rbBand.cx         = 400

 '  ����ε�Rebar��
  SendMessage ghRebar, %RB_INSERTBAND, 0&, VARPTR(rbBand)

 '  �ƶ�Rebar�ε������λ��!
  'IF udtAp.rbBand0 <>10  THEN           '%ID_MENUBAR
	SendMessage ghRebar, %RB_MOVEBAND, 1&, 0&
	SendMessage ghRebar, %RB_MOVEBAND, 2&, 1&
  'END IF
  '----------------------------------------------------------------------------

  FUNCTION = %FALSE

END FUNCTION

'------------------------------------------------------------------------------
' TOOLBAR.INC  ����
'------------------------------------------------------------------------------