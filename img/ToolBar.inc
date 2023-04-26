
'------------------------------------------------------------------------------
' 文件名称 : TOOLBAR.INC
'
' 描述 : 创建一个Rebar并增加工具到其上...
'
'               编译器版本   - PB/WIN 7.02
'               Win32Api版本 - Last Update: 27 August 2003.
'               创建日期     -
'               创建时间     -
'               Window平台   - Windows 95, 98, ME And XP.
'
'
' 你可以从这里得到状态栏位图:
' http://www.glyfx.com/toolbarimages.html
'
' 我创建位图条的程序，你可在这里找到:
' http://www.gamani.com/
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 常量和恒量.
'------------------------------------------------------------------------------

' 使用ReBar背景图片.
%USE_RBIMAGE = %FALSE        ' 变为%TRUE则使用ReBar背景图片

' 使用的工具栏按钮数.
%TOOLBUTTONS = 12
%MENUBUTTONS = 6

' ReBar 与 ToolBar 相同. 都开始于3000并逐渐增加...
%ID_REBAR    = %APP_REBARS
%ID_TOOLBAR  = %APP_TOOLBARS
%ID_COMBOBOX = %APP_CONTROLS
'%ID_MENUBAR  = %APP_MENUBARS

' 这些是工具栏按钮ID.
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
' 全局变量.
'------------------------------------------------------------------------------
GLOBAL ghRebar    AS DWORD         ' ReBar控件句柄.
GLOBAL ghToolBar  AS DWORD         ' 工具栏控件句柄.
GLOBAL ghComboBox AS DWORD         ' ComboBox控件句柄.

'------------------------------------------------------------------------------
' 函数及子程序.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 函数 : CreateToolBar ()
'
' 描述 : 创建工具栏控件.
'------------------------------------------------------------------------------
FUNCTION CreateToolBar (BYVAL hWnd AS DWORD) AS DWORD

  ' 为优化速度使用注册变量.
  REGISTER x AS LONG

  ' 工具栏变量.
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP

  DIM Tbb(0 TO %TOOLBUTTONS - 1) AS LOCAL TBBUTTON

  ' 检查正确的位图尺寸!
  IF gBmpSize <> 24 THEN
	 IF gBmpSize <> 16 THEN
		' 用户使用了不正确的位图尺寸.
		MessageBox(hWnd, "工具栏位图尺寸必须为16x16 或 24x24!", "错误提示", _
				   %MB_OK OR %MB_ICONERROR OR %MB_TOPMOST)
		FUNCTION = 1
		EXIT FUNCTION
	 END IF
  END IF

  '----------------------------------------------------------------------------
  ' 创建工具栏窗口.
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
  ' 检查创建工具栏控件窗口时的错误.
  '----------------------------------------------------------------------------
  IF ISFALSE ghToolBar THEN         ' 错误!
	 ' 创建工具栏窗口失败，通知用户.
	 FUNCTION = %TRUE
	 EXIT FUNCTION
  END IF

  '----------------------------------------------------------------------------
  ' 初始化工具栏按钮的Tbb()数组.
  '----------------------------------------------------------------------------
  FOR x = 0 TO %TOOLBUTTONS - 1
	' 为每个按钮设置初始化状态.
	Tbb(x).iBitmap   = 0
	Tbb(x).idCommand = 0
	Tbb(x).fsState   = %TBSTATE_ENABLED
	Tbb(x).fsStyle   = %TBSTYLE_BUTTON OR %TBSTYLE_AUTOSIZE
	Tbb(x).dwData    = 0
	Tbb(x).iString   = 0

	SELECT CASE AS LONG x
	  CASE 3,6,10
		' 创建按钮间隔.
		Tbb(x).fsStyle   = %TBSTYLE_SEP

	  CASE 0&
		Tbb(x).iBitmap   = 0
		Tbb(x).idCommand = %IDM_NEW
		Tbb(x).iString   = 2
		Tbb(x).fsState   = %TBSTATE_ENABLED ' 如果被定义，则可能是%TBSTATE_DISABLED

	  CASE 1&
		Tbb(x).iBitmap   = 1
		Tbb(x).idCommand = %IDM_OPEN
		Tbb(x).iString   = 1
		Tbb(x).fsState   = %TBSTATE_ENABLED ' 如果被定义，则可能是%TBSTATE_DISABLED

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
  ' 初始化工具栏窗口并发送窗口信息.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_TOOLBAR '%IDB_STD_LARGE_COLOR ' 使用位图图片的ID，如, %ID_TOOLBAR
  Tabm.hInst =  hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
											   ' 如果使用了链接源代码位图

  ' 设置工具栏位图尺寸为16x16.
  IF gBmpSize = 16 THEN
	 SendMessage ghToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(16, 16)
	 SendMessage ghToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(16, 16)
  END IF

  ' 设置工具栏位图尺寸为24x24.
  IF gBmpSize = 24 THEN
	 SendMessage ghToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(24, 24)
	 SendMessage ghToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(24, 24)
  END IF

  ' 增加图片列表到工具栏.
  SendMessage ghToolBar, %TB_SETIMAGELIST,         0, ghImlNor  ' 正常
  SendMessage ghToolBar, %TB_SETDISABLEDIMAGELIST, 0, ghImlDis  ' 不可用
  SendMessage ghToolBar, %TB_SETHOTIMAGELIST,      0, ghImlHot  ' 被选中

  ' 增加图片列表位图到工具栏按钮.
  SendMessage ghToolBar, %TB_ADDBITMAP, %TOOLBUTTONS, VARPTR(Tabm)

  ' 设置工具栏按钮.
  SendMessage ghToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage ghToolBar, %TB_ADDBUTTONS, %TOOLBUTTONS, VARPTR(Tbb(0))

  ' 强制工具栏为初始尺寸.
  SendMessage ghToolBar, %WM_SIZE, 0, 0

  FUNCTION = ghToolBar            ' 返回工具栏窗口句柄.

END FUNCTION

'------------------------------------------------------------------------------
' 函数 : CreateComboBox ()
'
' 描述 : 创建Combo Box并增加到ReBar控件上.
'        使用SetParent，你几乎可以增加任何控件到工具栏!
'------------------------------------------------------------------------------
FUNCTION CreateComboBox (BYVAL hWnd AS DWORD) AS DWORD

  REGISTER i AS LONG

  LOCAL dwBaseUnits AS DWORD, sStr AS STRING

  dwBaseUnits = GetDialogBaseUnits()

  '----------------------------------------------------------------------------
  ' 创建ComboBox控件窗口.
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
  ' 检查创建ComboBox窗口时的错误.
  '----------------------------------------------------------------------------
  IF ISFALSE ghComboBox THEN         ' 错误!
	 ' 创建ComboBox窗口失败，通知用户.
	 FUNCTION = %TRUE
	 EXIT FUNCTION
  END IF

  '----------------------------------------------------------------------------
  ' 循环增加ComboBox文本字符串.
  '----------------------------------------------------------------------------
  FOR i = 1 TO 20
	sStr = "第 " & FORMAT$(i) & " 项"
	SendMessage ghComboBox, %CB_ADDSTRING, 0&, STRPTR(sStr)
  NEXT

  SendMessage ghComboBox, %CB_SETCURSEL, 0&, 0&

  FUNCTION = ghComboBox

END FUNCTION

'------------------------------------------------------------------------------
' 函数 : SaveRebarBandPos ()
'
' 描述 : 保存Rebar条位置.
'------------------------------------------------------------------------------
FUNCTION SaveRebarBandPos() AS LONG

  DIM lResult AS LONG, lBands AS LONG
  DIM rBand   AS REBARBANDINFO

  REGISTER lCount AS LONG

  ' 得到Rebar行数.
  lResult = SendMessage(ghRebar, %RB_GETROWCOUNT, 0&, 0&)
  udtAp.rbRowCount = lResult

  ' 得到Rebar段数.
  lBands = SendMessage(ghRebar, %RB_GETBANDCOUNT, 0&, 0&)
  udtAp.rbBandCount = lResult

  ' 设置Rebar的REBARBANDINFO结构.
  rBand.cbSize = SIZEOF(rBand)
  rBand.fMask  = %RBBIM_ID

  ' 循环得到每段的ID.
  FOR lCount = 0 TO lBands - 1
	'lResult = SendMessage(hRebar, %RB_IDTOINDEX, lCount, VARPTR(rBand))
	lResult = SendMessage(ghRebar, %RB_GETBANDINFO, lCount, VARPTR(rBand))
	IF lCount = 0 THEN udtAp.rbBand0 = rBand.Wid
	IF lCount = 1 THEN udtAp.rbBand1 = rBand.Wid
  NEXT

END FUNCTION

'------------------------------------------------------------------------------
' 函数 : CreateRebar ()
'
' 描述 : 创建Rebar控件.
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
  ' 创建Rebar控件窗口.
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
  ' 检查创建Rebar控件窗口的错误.
  '----------------------------------------------------------------------------
  IF ISFALSE ghRebar THEN         ' 错误!
	 '= 创建Rebar窗口失败，通知用户.
	 FUNCTION = %TRUE
	 EXIT FUNCTION
  END IF

  '----------------------------------------------------------------------------
  ' 初始化并设置REBARINFO结构.
  '----------------------------------------------------------------------------
  rbi.cbSize = SIZEOF(rbi)
  rbi.fMask  = 0
  rbi.himl   = 0

  SendMessage ghRebar, %RB_SETBARINFO, 0, VARPTR(rbi)

  '----------------------------------------------------------------------------
  ' 初始化所有Rebar段的REBARBANDINFO.
  '----------------------------------------------------------------------------
  rbBand.cbSize     = SIZEOF(rbBand)
  rbBand.fMask      = %RBBIM_COLORS    OR _    ' clrFore 和 clrBack 可用
					  %RBBIM_CHILD     OR _    ' hwndChild 可用
					  %RBBIM_CHILDSIZE OR _    ' cxMinChild 和 cyMinChild 可用
					  %RBBIM_STYLE     OR _    ' fStyle 可用
					  %RBBIM_ID        OR _    ' wID 可用
					  %RBBIM_SIZE      OR _    ' cx 可用
					  %RBBIM_TEXT      OR _    ' lpText 可用
					  %RBBIM_BACKGROUND        ' hbmBack 可用

  rbBand.clrFore    = GetSysColor(%COLOR_BTNTEXT)
  rbBand.clrBack    = GetSysColor(%COLOR_BTNFACE)

  rbBand.fStyle = %RBBS_NOVERT     OR _    ' 不显示垂直方向
				  %RBBS_CHILDEDGE  OR _
				  %RBBS_FIXEDBMP   OR _
				  %RBBS_GRIPPERALWAYS

  '----------------------------------------------------------------------------
  ' 如果用户选择了一个Rebar背景图片，是加载它!
  '----------------------------------------------------------------------------
  IF ISTRUE %USE_RBIMAGE THEN
	 ghRbBack   = LoadImage(hInst, "RBBACK", %IMAGE_BITMAP, 0, 0, _
							%LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR _
							%LR_DEFAULTSIZE)

	 rbBand.hbmBack    = ghRbBack           ' Rebar背景图片.
  END IF

  '----------------------------------------------------------------------------
  ' 创建Rebar工具栏段.
  '----------------------------------------------------------------------------
  IF ISFALSE CreateToolBar(hWnd) THEN         ' 错误!
	 ' 位图尺寸非法，通知用户.
	 FUNCTION = %TRUE
	 EXIT FUNCTION
  END IF

  '----------------------------------------------------------------------------
  ' 初始化并增加工具栏到Rebar的段.
  '----------------------------------------------------------------------------
  dwBtnSize         = SendMessage(ghToolBar, %TB_GETBUTTONSIZE, 0, 0)
  'szTbText          = "ToolBar"
  rbBand.lpText     = VARPTR(szTbText)
  rbBand.hwndChild  = ghToolBar
  rbBand.wID        = %ID_TOOLBAR
  rbBand.cxMinChild = 230
  rbBand.cyMinChild = HIWRD(dwBtnSize) '+ 2
  rbBand.cx         = 280

  ' 多于一个时，则打断段到一个新行.
  IF udtAp.rbRowCount > 1 THEN
	rbBand.fStyle = rbBand.fStyle OR %RBBS_BREAK
  END IF

  '----------------------------------------------------------------------------
  ' 插入并增加工具栏段到Rebar控件窗口.
  '----------------------------------------------------------------------------
  SendMessage ghRebar, %RB_INSERTBAND, -1&, VARPTR(rbBand)


  '----------------------------------------------------------------------------
  ' 去掉注释: 如果你想放置ComboBox到Rebar段中.
  '----------------------------------------------------------------------------

  ' 创建Rebar的ComboBox段.
  CreateComboBox(hWnd)
  szCbText          = "程序:"
  rbBand.lpText     = VARPTR(szCbText)
  rbBand.hwndChild  = ghComboBox
  rbBand.wID        = %ID_COMBOBOX
  GetWindowRect ghComboBox, rc
  rbBand.cxMinChild = 150
  rbBand.cyMinChild = rc.nBottom - rc.nTop
  rbBand.cx         = 400

 '  插入段到Rebar中
  SendMessage ghRebar, %RB_INSERTBAND, 0&, VARPTR(rbBand)

 '  移动Rebar段到保存的位置!
  'IF udtAp.rbBand0 <>10  THEN           '%ID_MENUBAR
	SendMessage ghRebar, %RB_MOVEBAND, 1&, 0&
	SendMessage ghRebar, %RB_MOVEBAND, 2&, 1&
  'END IF
  '----------------------------------------------------------------------------

  FUNCTION = %FALSE

END FUNCTION

'------------------------------------------------------------------------------
' TOOLBAR.INC  结束
'------------------------------------------------------------------------------
