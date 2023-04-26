'toolbar.bas
'工具栏
'------------------------------------------------------------------------------
' 常量和恒量.
'------------------------------------------------------------------------------
' 使用ReBar背景图片.
%USE_RBIMAGE = 0 '%FALSE        ' 变为%TRUE则使用ReBar背景图片
' 使用的工具栏按钮数.
%TOOLBUTTONS = 12 '标准工具栏的按钮数
%MENUBUTTONS = 6
%WM_UPDATESIZEWINDOW  = %WM_USER +1024  'Size/Position window
%APP_REBARS           = %WM_USER + 3072
%APP_TOOLBARS         = %WM_USER + 4096
%APP_STATUSBARS       = %WM_USER + 5120
%APP_CONTROLS         = %WM_USER + 6144
' ReBar 与 ToolBar 相同. 都开始于3000并逐渐增加...
%ID_REBAR             = %APP_REBARS
%ID_TOOLBAR           = %APP_TOOLBARS
%ID_WINDOWBAR         = %WM_USER + 4500
%ID_COMBOBOX          = %APP_CONTROLS
%ID_DEBUGBAR          = %WM_USER + 4700
%ID_CODEEDITBAR       = %WM_USER + 4800
%ID_WINARRAGBAR       = %WM_USER + 4900

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

%IDM_STDTOOLBAR   = %WM_USER + 2100& '标准工具栏显示/隐藏菜单命令
%IDM_WINDOWBAR    = %WM_USER + 2101& '窗口工具栏显示/隐藏菜单命令
%IDM_COMBOBAR     = %WM_USER + 2102& '过程列表工具栏显示/隐藏菜单命令
%IDM_DEBUGBAR     = %WM_USER + 2103& '编译调试/运行工具栏显示/隐藏菜单命令
%ID_LOCKBAR       = %WM_USER + 2104& '锁定/解锁工具栏
%IDM_CODEEDIT     = %WM_USER + 2105& '代码编辑工具栏显示/隐藏菜单命令
%IDM_WINARRAG     = %WM_USER + 2016& '窗口排列工具栏显示/隐藏菜单命令

%I_IMAGENONE      = -2
%IMG_DIS = 0    ' 不可用Disabled
%IMG_NOR = 1    ' 一般状态Normal
%IMG_HOT = 2    ' 选中状态Selected
%OMENU_EXTRAWIDTH = 30    ' Extra width in pixels
%OMENU_CHECKEDICON = 106    ' Identifier of the checked icon
'------------------------------------------------------------------------------
' 全局变量.
'------------------------------------------------------------------------------
GLOBAL hMenuTextBkBrush AS DWORD        ' // Submenu text brush          子菜单文本刷
GLOBAL hMenuHiBrush     AS DWORD        ' // Submenu highlighted text brush子菜单高亮文本刷
GLOBAL hMenuIconsBrush      AS DWORD        ' // Brush for the icon's background图标背景刷

'= 菜单和工具栏图片列表句柄.
' 标准工具栏按钮
GLOBAL ghImlHot    AS DWORD               ' 热选中图片列表句柄.
GLOBAL ghImlDis    AS DWORD               ' 不可用图片列表句柄.
GLOBAL ghImlNor    AS DWORD               ' 正常状态图片列表句柄.
' 窗口工具栏按钮
GLOBAL hImlHot    AS DWORD               ' 热选中图片列表句柄.
GLOBAL hImlDis    AS DWORD               ' 不可用图片列表句柄.
GLOBAL hImlNor    AS DWORD               ' 正常状态图片列表句柄.

GLOBAL gBmpSize    AS LONG                ' 位图尺寸.
'= Rebar 背景图片位图句柄.
GLOBAL ghRbBack    AS DWORD               ' Rebar 控件背景位图句柄.

'------------------------------------------------------------------------------
' 函数 : AppLoadBitmaps ()
' 描述 : 从资源文件加载程序位图.
' 注意 : 如果你要在工具栏使用24x24位图，你就必须给菜单增加16x16位图条!
'------------------------------------------------------------------------------
FUNCTION AppLoadBitmaps() AS LONG
  LOCAL bm      AS BITMAP
  LOCAL hBmpHot AS DWORD
  LOCAL hBmpDis AS DWORD
  LOCAL hBmpNor AS DWORD
  '----------------------------------------------------------------------------
  ' 设置并初始化菜单，工具栏位图和图片列表.
  ' 需要 : 需要菜单位图是16x16的代码增加到这里!
  '----------------------------------------------------------------------------
  ' 悬浮热位图.
  hBmpHot = LoadImage(g_hInst, "TBHOT", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  ' 不可用位图.
  hBmpDis = LoadImage(g_hInst, "TBDIS", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  ' 正常时位图.
  hBmpNor = LoadImage(g_hInst, "TBNOR", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  '----------------------------------------------------------------------------
  ' 为以后使用，获取并保存位图尺寸.
  '----------------------------------------------------------------------------
  GetObject hBmpNor, SIZEOF(bm), bm           ' 得到位图尺寸.
  'msgbox "hBmpHot=" & str$(hBmpHot)
  gBmpSize = bm.bmHeight                      ' 保存位图尺寸以备后用.
  'if gBmpSize=0 then gBmpSize=16
  '----------------------------------------------------------------------------
  ' 创建菜单和工具栏图片热选中、不可用及正常状态列表.
  ' 需要: 需要将菜单位图为16x16的代码增加到这里!
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
  ' 加载Rebar背景图片.
  '----------------------------------------------------------------------------
  ghRbBack = LoadImage(g_hInst, "RBBACK", %IMAGE_BITMAP, 0, 0, _
             %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR _
             %LR_DEFAULTSIZE)
  '----------------------------------------------------------------------------
  ' 清除并删除不再需要的位图句柄.
  '----------------------------------------------------------------------------
  IF hBmpHot THEN DeleteObject(hBmpHot)
  IF hBmpDis THEN DeleteObject(hBmpDis)
  IF hBmpNor THEN DeleteObject(hBmpNor)
  '----------------------------------------------------------------------------
  ' 设置并初始化菜单，工具栏位图和图片列表.
  ' 需要 : 需要菜单位图是16x16的代码增加到这里!
  '----------------------------------------------------------------------------
  ' 悬浮热位图.
  hBmpHot = LoadImage(g_hInst, "WINTBHOT", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  ' 不可用位图.
  hBmpDis = LoadImage(g_hInst, "WINTBDIS", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  ' 正常时位图.
  hBmpNor = LoadImage(g_hInst, "WINTBNOR", %IMAGE_BITMAP, 0, 0, _
            %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR %LR_DEFAULTSIZE)
  '----------------------------------------------------------------------------
  ' 为以后使用得到并保存位图尺寸.
  '----------------------------------------------------------------------------
'  GetObject hBmpNor, SIZEOF(bm), bm           ' 得到位图尺寸.
'  'msgbox "hBmpHot=" & str$(hBmpHot)
'  gBmpSize = bm.bmHeight                      ' 保存位图尺寸以备后用.
'  'if gBmpSize=0 then gBmpSize=16
  '----------------------------------------------------------------------------
  ' 创建菜单和工具栏图片热选中、不可用及正常状态列表.
  ' 需要: 需要将菜单位图为16x16的代码增加到这里!
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
' 函数 : CreateToolBar ()
' 描述 : 创建工具栏控件.
'------------------------------------------------------------------------------
FUNCTION CreateToolBar (BYVAL hWnd AS DWORD) AS DWORD
  ' 为优化速度使用注册变量.
  REGISTER x AS LONG
  LOCAL c AS LONG
  ' 工具栏变量.
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP
  DIM Tbb(0 TO %TOOLBUTTONS - 1) AS LOCAL TBBUTTON
  ' 检查正确的位图尺寸!
  IF gBmpSize <> 24 THEN
    IF gBmpSize <> 16 THEN
      ' 用户使用了不正确的位图尺寸.
      MessageBox(hWnd, "工具栏位图尺寸必须为16x16 或 24x24!" & $CRLF & STR$(gBmpSize), "错误提示", _
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
               g_hInst, _
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
  ' 用按钮信息填充 TBBUTTON 数组
  tbb(c).iBitmap   = 0'%STD_FILENEW     '新建
  tbb(c).idCommand = %IDM_NEWBAS
  tbb(c).fsState   = %TBSTATE_ENABLED
  'tbb(c).fsStyle   = %TBSTYLE_BUTTON
  IF g_NewComCtl >= 4.7 THEN
    tbb(c).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb(c).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR c
  tbb(c).iBitmap   = 1'%STD_FILEOPEN    '打开
  tbb(c).idCommand = %IDM_OPEN
  tbb(c).fsState   = %TBSTATE_ENABLED
  IF g_NewComCtl >= 4.7 THEN
    tbb(c).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb(c).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR c
  tbb(c).iBitmap   = 2'%STD_FILESAVE    '保存
  tbb(c).idCommand = %IDM_SAVE
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
'  INCR c
'  tbb(c).iBitmap   = 108
'  tbb(c).idCommand = %IDM_REFRESH       '刷新
'  tbb(c).fsState   = %TBSTATE_ENABLED
'  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 0                  '分隔条
  tbb(c).idCommand = 701
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c
  tbb(c).iBitmap   = 3'%STD_PRINT       '打印
  tbb(c).idCommand = %IDM_PRINT
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 0                  '分隔条
  tbb(c).idCommand = 702
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c
  tbb(c).iBitmap   = 5'%STD_CUT         '剪切
  tbb(c).idCommand = %IDM_CUT
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 6'%STD_COPY        '复制
  tbb(c).idCommand = %IDM_COPY
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 7'%STD_PASTE       '粘贴
  tbb(c).idCommand = %IDM_PASTE
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 0                  '分隔条
  tbb(c).idCommand = 703
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c
  tbb(c).iBitmap   = 9'%STD_UNDO        '撤销
  tbb(c).idCommand = %IDM_UNDO
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  tbb(c).iBitmap   = 8'%STD_REDO        ‘重做
  tbb(c).idCommand = %IDM_REDO
  tbb(c).fsState   = %TBSTATE_ENABLED
  tbb(c).fsStyle   = %TBSTYLE_BUTTON
  INCR c
  '----------------------------------------------------------------------------
  ' 初始化工具栏窗口并发送窗口信息.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_TOOLBAR '%IDB_STD_LARGE_COLOR ' 使用位图图片的ID，如, %ID_TOOLBAR
  Tabm.hInst = g_hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
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
  IF ghToolBar THEN
    SendMessage ghToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' 通用控件版本较新，则可使用更特性，如平整样式等
       SetWindowLong ghToolBar, %GWL_STYLE, _
          GetWindowLong(ghToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage ghToolBar, %TB_SETEXTENDEDSTYLE, 0,%TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF
  FUNCTION = ghToolBar            ' 返回工具栏窗口句柄.
END FUNCTION
' *********************************************************************************************
' 函数名    : CreateWindowBar
' 描述      : 创建窗口工具栏控件
' 返回      : 工具栏句柄
' 参数      : hWnd      = 主窗口句柄
'             lBmpSize  = 位图大小： 16x16 或 24x24
' *********************************************************************************************
FUNCTION CreateWindowBar( BYVAL hWnd AS DWORD) AS DWORD
  ' 使用寄存器变量以优化速度
  REGISTER x AS LONG
  ' 工具栏变量
  LOCAL TOOLBUTTONS AS LONG
  LOCAL Tbb( ) AS TBBUTTON
  LOCAL Tabm AS TBADDBITMAP
  LOCAL hToolbar AS DWORD
  LOCAL a$
  LOCAL NewComCtl AS LONG
  TOOLBUTTONS = 12
  DIM Tbb( 0 TO TOOLBUTTONS-1) AS LOCAL TBBUTTON
  '----------------------------------------------------------------------------
  ' 创建工具栏窗口
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
  ' 获取通用控件库版本
  NewComCtl = InitComctl32 (%ICC_BAR_CLASSES)
  IF NewComCtl>4.7 THEN
    SENDMESSAGE hToolbar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
  END IF
  '----------------------------------------------------------------------------
  ' 初始化工具栏按钮的Tbb()数组
  '----------------------------------------------------------------------------
  FOR x = 0 TO TOOLBUTTONS-1
    ' 设置每个按钮的初始状态
    Tbb( x ).iBitmap = 0
    Tbb( x ).idCommand = 0
    Tbb( x ).fsState = %TBSTATE_ENABLED
    Tbb( x ).fsStyle = %TBSTYLE_BUTTON OR %TBSTYLE_AUTOSIZE
    Tbb( x ).dwData = 0
    Tbb( x ).iString = 0
  NEXT x
  x=0
  Tbb( x ).iBitmap = 75                   '新建窗口
  Tbb( x ).idCommand = %IDM_WIN
  IF NewComCtl THEN
    tbb( x ).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb( x ).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR x
  Tbb( x ).iBitmap = 132                  '创建菜单
  Tbb( x ).idCommand = %ID_TOOL_MENU
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 133                  '创建工具栏
  Tbb( x ).idCommand = %ID_TOOL_TOOLBAR
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 138                  '控件吸附
  Tbb( x ).idCommand = %ID_BARONLY_STICKY
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 139                  '控件对齐网格
  Tbb( x ).idCommand = %ID_BARONLY_SNAP
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 140                  '显示/隐藏网格
  Tbb( x ).idCommand = %ID_BARONLY_GRID
  Tbb( x ).fsStyle = %TBSTYLE_AUTOSIZE
  INCR x
  Tbb( x ).iBitmap = 134                  '垂直对齐控件
  Tbb( x ).idCommand = %ID_SETV_CENTER
  IF NewComCtl THEN
    tbb( x ).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb( x ).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR x
  Tbb( x ).iBitmap = 135                  '水平对齐控件
  Tbb( x ).idCommand = %ID_SETH_CENTER
  IF NewComCtl THEN
    tbb( x ).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb( x ).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR x
  Tbb( x ).iBitmap = 136                  '垂直等距
  Tbb( x ).idCommand = %ID_SETV_EQUATE
  INCR x
  Tbb( x ).iBitmap = 137                  '水平等距
  Tbb( x ).idCommand = %ID_SETH_EQUATE
  INCR x
  Tbb( x ).iBitmap = 141                  '生成代码
  Tbb( x ).idCommand = %ID_TOOL_COMPILE
  IF NewComCtl THEN
    tbb( x ).fsStyle = %TBSTYLE_DROPDOWN
  ELSE
    tbb( x ).fsStyle = %TBSTYLE_BUTTON
  END IF
  INCR x
  Tbb( x ).fsStyle = %TBSTYLE_SEP         '分隔条
  '----------------------------------------------------------------------------
  ' 初始化工具栏并给窗口发送消息
  '----------------------------------------------------------------------------
  Tabm.nID = %ID_WINDOWBAR    ' %IDB_STD_LARGE_COLOR
  Tabm.hInst = g_hInst
  SENDMESSAGE hToolBar, %TB_SETBITMAPSIZE, 0, MAKLNG( 16, 16 )
  ' 给工具栏增加图片列表
  SENDMESSAGE hToolBar, %TB_SETIMAGELIST, 0, hImlNor          ' 常规图片
  SENDMESSAGE hToolBar, %TB_SETDISABLEDIMAGELIST, 0, hImlDis  ' 不可用图片
  SENDMESSAGE hToolBar, %TB_SETHOTIMAGELIST, 0, hImlHot       ' 悬浮图片
  ' 给工具栏按钮增加图片列表上的位图
  SENDMESSAGE hToolBar, %TB_ADDBITMAP, TOOLBUTTONS, VARPTR( Tabm )
  ' 设置工具栏按钮
  SENDMESSAGE hToolBar, %TB_BUTTONSTRUCTSIZE, SIZEOF( Tbb( 0 )), 0
  SENDMESSAGE hToolBar, %TB_ADDBUTTONS, TOOLBUTTONS, VARPTR( Tbb( 0 ))
'   a$ = $NUL & "0,0   " &  $NUL & "0,0   " & $NUL & $NUL
'   Sendmessage hToolBar, %TB_ADDSTRING,0, strptr(a$)
  ' 强制工具样初始重置尺寸
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
  '子类化静态控件，这样就可以绘制位置和尺寸
  LOCAL szOldWndProc AS ASCIIZ*11
  LOCAL lpOldWndProc AS LONG
  lpOldWndProc = SetWindowLong(ghWndSize,%GWL_WNDPROC,BYVAL CODEPTR(WndStatusSizeProc))
  '通过设置新属性，将该类与旧的窗口处理过程关联
  szOldWndProc = "OldWndProc"
  CALL SetProp(ghWndSize,szOldWndProc,lpOldWndProc&)
  ' 强制工具栏为初始尺寸.
  IF hToolBar THEN
    SendMessage hToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' 版本支持时，增加扁平化和下拉按钮样式
       SetWindowLong hToolBar, %GWL_STYLE, _
          GetWindowLong(hToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage hToolBar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF
  ' 返回工具栏句柄
  FUNCTION = hToolbar
END FUNCTION
'------------------------------------------------------------------------------
' 函数 : CreateDebugBar ()
' 描述 : 创建编译，运行，调试工具栏控件.
'------------------------------------------------------------------------------
FUNCTION CreateDebugBar (BYVAL hWnd AS DWORD) AS DWORD

  ' 为优化速度使用注册变量.
  REGISTER x AS LONG
  LOCAL c AS LONG
  ' 工具栏变量.
  LOCAL TOOLBUTTONS AS INTEGER
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP
  LOCAL hToolbar AS DWORD
  TOOLBUTTONS = 9
  DIM Tbb(0 TO TOOLBUTTONS - 1) AS LOCAL TBBUTTON

  ' 检查正确的位图尺寸!
  IF gBmpSize <> 24 THEN
   IF gBmpSize <> 16 THEN
    ' 用户使用了不正确的位图尺寸.
    MessageBox(hWnd, "工具栏位图尺寸必须为16x16 或 24x24!" & $CRLF & STR$(gBmpSize), "错误提示", _
           %MB_OK OR %MB_ICONERROR OR %MB_TOPMOST)
    FUNCTION = 1
    EXIT FUNCTION
   END IF
  END IF

  '----------------------------------------------------------------------------
  ' 创建工具栏窗口.
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
  ' 检查创建工具栏控件窗口时的错误.
  '----------------------------------------------------------------------------
  IF ISFALSE hToolBar THEN         ' 错误!
   ' 创建工具栏窗口失败，通知用户.
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
  ' 初始化工具栏按钮的Tbb()数组.
  '----------------------------------------------------------------------------

  ' 用按钮信息填充 TBBUTTON 数组
  c=0
  tbb(c).iBitmap   = 37 '设置断点
  tbb(c).idCommand = %IDM_SETBRKPO
  INCR c

  tbb(c).iBitmap   = 39 '上一断点
  tbb(c).idCommand = %IDM_PRVBRKPO
  INCR c

  tbb(c).iBitmap   = 38 '下一断点
  tbb(c).idCommand = %IDM_NEXBRKPO
  INCR c

  tbb(c).iBitmap   = 40 '删除断点
  tbb(c).idCommand = %IDM_DELBRKPO
  INCR c

  tbb(c).iBitmap   = 125 '转到断点
  tbb(c).idCommand = %IDM_GOBRKPO
  INCR c

  tbb(c).iBitmap   = -1
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c

  tbb(c).iBitmap   = 42
  tbb(c).idCommand = %IDM_COMPILE  '编译
  INCR c

  tbb(c).iBitmap   = 43
  tbb(c).idCommand = %IDM_CMPRUN  '编译并运行
  INCR c

  tbb(c).iBitmap   = 44 '编译并调试
  tbb(c).idCommand = %IDM_CMPDEBUG
  tbb(c).fsState   = %TBSTATE_DISABLED
  INCR c

'  tbb(c).iBitmap   = -1                  '这里需要在升级版中扩展：显示跳过，跳入，跳出，终止，运行等按钮
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
  ' 初始化工具栏窗口并发送窗口信息.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_DEBUGBAR '%IDB_STD_LARGE_COLOR ' 使用位图图片的ID，如, %ID_TOOLBAR
  Tabm.hInst =  g_hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
                         ' 如果使用了链接源代码位图

  ' 设置工具栏位图尺寸为16x16.
  IF gBmpSize = 16 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(16, 16)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(16, 16)
  END IF

  ' 设置工具栏位图尺寸为24x24.
  IF gBmpSize = 24 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(24, 24)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(24, 24)
  END IF

  ' 增加图片列表到工具栏.
  SendMessage hToolBar, %TB_SETIMAGELIST,         0, hImlNor  ' 正常
  SendMessage hToolBar, %TB_SETDISABLEDIMAGELIST, 0, hImlDis  ' 不可用
  SendMessage hToolBar, %TB_SETHOTIMAGELIST,      0, hImlHot  ' 被选中

  ' 增加图片列表位图到工具栏按钮.
  SendMessage hToolBar, %TB_ADDBITMAP, TOOLBUTTONS, VARPTR(Tabm)

  ' 设置工具栏按钮.
  SendMessage hToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage hToolBar, %TB_ADDBUTTONS, TOOLBUTTONS, VARPTR(Tbb(0))

  ' 强制工具栏为初始尺寸.

  IF hToolBar THEN
    SendMessage hToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' then we can use more modern features, like flat style, etc.
       SetWindowLong hToolBar, %GWL_STYLE, _
          GetWindowLong(hToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage hToolBar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF

  FUNCTION = hToolBar            ' 返回工具栏窗口句柄.

END FUNCTION
'------------------------------------------------------------------------------
' 函数：WndStatusSizeProc
' 描述：自定义状态窗口回调过程:
'       绘制两个图片，并更新当前被编辑窗口的 x,y,cx,cy 位置及大小信息
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
      '自刷新...
      InvalidateRect hWnd, BYVAL %NULL,%TRUE
      UpdateWindow hWnd
      FUNCTION =0 :EXIT FUNCTION
    CASE %WM_ERASEBKGND
      FUNCTION =1 :EXIT FUNCTION
    CASE %WM_PAINT
      hDC = BeginPaint(hWnd, ps)
      '清除之前的底色...
      GetClientRect hWnd,rc
      FillRect hDC, rc, GetStockObject(%LTGRAY_BRUSH)'GRAY_BRUSH
      ' WM_CREATE 消息发出后就加载然后循环
      IF fBmpLoaded = 0 THEN
        fBmpLoaded = 1
        '获取 x/y 位图
        hbitmapLT = LoadBitmap( g_hInst,"XYBM" )
        '获取 cx/cy 位置
        hbitmapRB = LoadBitmap( g_hInst,"CXYBM" )
      END IF
      '绘制位图...
      hMemDc = CreateCompatibleDc(hDC)
      GetObject hbitmapLT,LEN(Bm),Bm
      SelectObject hMemDc ,hbitmapLT
      BitBlt hDC ,0 ,4 ,Bm.bmWidth ,Bm.bmHeight ,hMemDc ,0 ,0 ,%SRCCOPY
      GetObject hbitmapRB,LEN(Bm),Bm
      SelectObject hMemDc ,hBitmapRB
      BitBlt hDC ,81 ,4 ,Bm.bmWidth ,Bm.bmHeight ,hMemDc ,0 ,0 ,%SRCCOPY
      DeleteDc hMemDc
      '输出文本...
      hOldFont = SelectObject(hDC,ghMyFont)
      SetTextColor hDC, RGB(0,0,0)
      SetBkMode hDC, %TRANSPARENT
      TextOut hDC, 15 , 4, lSizeXY, LEN(lSizeXY)
      TextOut hDC, 96, 4, lSizeCXY, LEN(lSizeCXY)
      SelectObject hDC,hOldFont
      EndPaint hWnd,ps
      FUNCTION = 0 :EXIT FUNCTION
    CASE %WM_DESTROY
      '清内存...
      IF ISTRUE(hbitmapLT) THEN CALL DeleteObject(hbitmapLT)
      IF ISTRUE(hbitmapRB) THEN CALL DeleteObject(hbitmapRB)
      ' 取消子类化控件并恢复其默认处理过程
      CALL SetWindowLong(hWnd, %GWL_WNDPROC, GetProp(hWnd,szOldWndProc))
      '移除关联属性
      RemoveProp hWnd,szOldWndProc
    END SELECT
    '开始消息，调用类处理过程
    FUNCTION = CallWindowProc( lpOldWndProc, hWnd, wMsg, wParam, lParam )

END FUNCTION
'------------------------------------------------------------------------------
' 函数 : CreateCodeEditBar ()
' 描述 : 创建代码编辑工具栏控件.
'------------------------------------------------------------------------------
FUNCTION CreateCodeEditBar (BYVAL hWnd AS DWORD) AS DWORD

  ' 为优化速度使用注册变量.
  REGISTER x AS LONG
  LOCAL c AS LONG
  ' 工具栏变量.
  LOCAL TOOLBUTTONS AS INTEGER
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP
  LOCAL hToolbar AS DWORD
  TOOLBUTTONS = 13
  DIM Tbb(0 TO TOOLBUTTONS - 1) AS LOCAL TBBUTTON

  ' 检查正确的位图尺寸!
  IF gBmpSize <> 24 THEN
   IF gBmpSize <> 16 THEN
    ' 用户使用了不正确的位图尺寸.
    MessageBox(hWnd, "工具栏位图尺寸必须为16x16 或 24x24!" & $CRLF & STR$(gBmpSize), "错误提示", _
           %MB_OK OR %MB_ICONERROR OR %MB_TOPMOST)
    FUNCTION = 1
    EXIT FUNCTION
   END IF
  END IF

  '----------------------------------------------------------------------------
  ' 创建工具栏窗口.
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
  ' 检查创建工具栏控件窗口时的错误.
  '----------------------------------------------------------------------------
  IF ISFALSE hToolBar THEN         ' 错误!
   ' 创建工具栏窗口失败，通知用户.
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
  ' 初始化工具栏按钮的Tbb()数组.
  '----------------------------------------------------------------------------

  ' 用按钮信息填充 TBBUTTON 数组
  c=0
  tbb(c).iBitmap   = 22 '块注释
  tbb(c).idCommand = %IDM_BLOCKCOMMIT
  INCR c

  tbb(c).iBitmap   = 23 '反块注释
  tbb(c).idCommand = %IDM_REBLOCKCOMMIT
  INCR c

  tbb(c).iBitmap   = 24 '块缩进
  tbb(c).idCommand = %IDM_BLOCKINDENT
  INCR c

  tbb(c).iBitmap   = 25 '反块缩进
  tbb(c).idCommand = %IDM_REBLOCKINDENT
  INCR c

  tbb(c).iBitmap   = -1
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c

  tbb(c).iBitmap   = 28 '转大写
  tbb(c).idCommand = %IDM_CAPWORD
  INCR c

  tbb(c).iBitmap   = 29 '转小写
  tbb(c).idCommand = %IDM_SMALLWORD
  INCR c

  tbb(c).iBitmap   = 30 '首字母大写
  tbb(c).idCommand = %IDM_FIRSTCAPWORD
  INCR c

  tbb(c).iBitmap   = -1
  tbb(c).fsStyle   = %TBSTYLE_SEP
  INCR c

  tbb(c).iBitmap   = 33 '查找
  tbb(c).idCommand = %IDM_FIND
  INCR c

  tbb(c).iBitmap   = 34 '查找下一个
  tbb(c).idCommand = %IDM_FINDNEXT
  INCR c

  tbb(c).iBitmap   = 35 '查找并替换
  tbb(c).idCommand = %IDM_FINDREPLACE
  'tbb(c).fsState   = %TBSTATE_DISABLED
  INCR c

  tbb(c).iBitmap   = 36 '转到行
  tbb(c).idCommand = %IDM_GOLINE
  INCR c

'  tbb(c).iBitmap   = -1                  '这里需要在升级版中扩展：显示跳过，跳入，跳出，终止，运行等按钮
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
  ' 初始化工具栏窗口并发送窗口信息.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_CODEEDITBAR '%IDB_STD_LARGE_COLOR ' 使用位图图片的ID，如, %ID_TOOLBAR
  Tabm.hInst =  g_hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
                         ' 如果使用了链接源代码位图

  ' 设置工具栏位图尺寸为16x16.
  IF gBmpSize = 16 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(16, 16)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(16, 16)
  END IF

  ' 设置工具栏位图尺寸为24x24.
  IF gBmpSize = 24 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(24, 24)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(24, 24)
  END IF

  ' 增加图片列表到工具栏.
  SendMessage hToolBar, %TB_SETIMAGELIST,         0, hImlNor  ' 正常
  SendMessage hToolBar, %TB_SETDISABLEDIMAGELIST, 0, hImlDis  ' 不可用
  SendMessage hToolBar, %TB_SETHOTIMAGELIST,      0, hImlHot  ' 被选中

  ' 增加图片列表位图到工具栏按钮.
  SendMessage hToolBar, %TB_ADDBITMAP, TOOLBUTTONS, VARPTR(Tabm)

  ' 设置工具栏按钮.
  SendMessage hToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage hToolBar, %TB_ADDBUTTONS, TOOLBUTTONS, VARPTR(Tbb(0))

  ' 强制工具栏为初始尺寸.

  IF hToolBar THEN
    SendMessage hToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' then we can use more modern features, like flat style, etc.
       SetWindowLong hToolBar, %GWL_STYLE, _
          GetWindowLong(hToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage hToolBar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF

  FUNCTION = hToolBar            ' 返回工具栏窗口句柄.

END FUNCTION
'------------------------------------------------------------------------------
' 函数 : CreateWinArragBar ()
'
' 描述 : 创建窗口排列工具栏控件.
'------------------------------------------------------------------------------
FUNCTION CreateWinArragBar (BYVAL hWnd AS DWORD) AS DWORD

  ' 为优化速度使用注册变量.
  'REGISTER x AS LONG
  LOCAL c AS LONG
  ' 工具栏变量.
  LOCAL TOOLBUTTONS AS INTEGER
  LOCAL Tbb() AS TBBUTTON
  LOCAL Tabm  AS TBADDBITMAP
  LOCAL hToolbar AS DWORD
  TOOLBUTTONS = 7
  DIM Tbb(0 TO TOOLBUTTONS - 1) AS LOCAL TBBUTTON

  ' 检查正确的位图尺寸!
  IF gBmpSize <> 24 THEN
   IF gBmpSize <> 16 THEN
    ' 用户使用了不正确的位图尺寸.
    MessageBox(hWnd, "工具栏位图尺寸必须为16x16 或 24x24!" & $CRLF & STR$(gBmpSize), "错误提示", _
           %MB_OK OR %MB_ICONERROR OR %MB_TOPMOST)
    FUNCTION = 1
    EXIT FUNCTION
   END IF
  END IF

  '----------------------------------------------------------------------------
  ' 创建工具栏窗口.
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
  ' 检查创建工具栏控件窗口时的错误.
  '----------------------------------------------------------------------------
  IF ISFALSE hToolBar THEN         ' 错误!
   ' 创建工具栏窗口失败，通知用户.
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
  ' 初始化工具栏按钮的Tbb()数组.
  '----------------------------------------------------------------------------

  ' 用按钮信息填充 TBBUTTON 数组
  c=0
  tbb(c).iBitmap   = 57 '层叠排列
  tbb(c).idCommand = %IDM_CASCADE
  INCR c

  tbb(c).iBitmap   = 58 '水平排列
  tbb(c).idCommand = %IDM_TILEH
  INCR c

  tbb(c).iBitmap   = 59 '垂直排列
  tbb(c).idCommand = %IDM_TILEV
  INCR c

  tbb(c).iBitmap   = 62 '图标排列
  tbb(c).idCommand = %IDM_ARRANGE
  INCR c

  tbb(c).iBitmap   = 60 '最大化
  tbb(c).idCommand = %IDM_MAXWIN
  INCR c

  tbb(c).iBitmap   = 61 '转到窗口
  tbb(c).idCommand = %IDM_SWITCHWINDOW
  INCR c

  tbb(c).iBitmap   = 63
  tbb(c).idCommand = %IDM_CLOSE  '关闭窗口
  'INCR c

  '----------------------------------------------------------------------------
  ' 初始化工具栏窗口并发送窗口信息.
  '----------------------------------------------------------------------------
  Tabm.nID   = %ID_WINARRAGBAR '%IDB_STD_LARGE_COLOR ' 使用位图图片的ID，如, %ID_TOOLBAR
  Tabm.hInst =  g_hInst  'GetModuleHandle(BYVAL %NULL)   ' %HINST_COMMCTRL
                         ' 如果使用了链接源代码位图

  ' 设置工具栏位图尺寸为16x16.
  IF gBmpSize = 16 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(16, 16)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(16, 16)
  END IF

  ' 设置工具栏位图尺寸为24x24.
  IF gBmpSize = 24 THEN
   SendMessage hToolBar, %TB_SETBITMAPSIZE, 0&, MAKLNG(24, 24)
   SendMessage hToolBar, %TB_SETBUTTONSIZE, 0&, MAKLNG(24, 24)
  END IF

  ' 增加图片列表到工具栏.
  SendMessage hToolBar, %TB_SETIMAGELIST,         0, hImlNor  ' 正常
  SendMessage hToolBar, %TB_SETDISABLEDIMAGELIST, 0, hImlDis  ' 不可用
  SendMessage hToolBar, %TB_SETHOTIMAGELIST,      0, hImlHot  ' 被选中

  ' 增加图片列表位图到工具栏按钮.
  SendMessage hToolBar, %TB_ADDBITMAP, TOOLBUTTONS, VARPTR(Tabm)

  ' 设置工具栏按钮.
  SendMessage hToolBar, %TB_BUTTONSTRUCTSIZE,         SIZEOF(Tbb(0)), 0
  SendMessage hToolBar, %TB_ADDBUTTONS, TOOLBUTTONS, VARPTR(Tbb(0))

  ' 强制工具栏为初始尺寸.

  IF hToolBar THEN
    SendMessage hToolBar, %TB_AUTOSIZE, 0, 0
    IF g_NewComCtl >= 4.7 THEN ' then we can use more modern features, like flat style, etc.
       SetWindowLong hToolBar, %GWL_STYLE, _
          GetWindowLong(hToolbar, %GWL_STYLE) OR %TBSTYLE_FLAT
       SendMessage hToolBar, %TB_SETEXTENDEDSTYLE, 0, %TBSTYLE_EX_DRAWDDARROWS
    END IF
  END IF

  FUNCTION = hToolBar            ' 返回工具栏窗口句柄.

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
  LOCAL hToolbar AS DWORD
  LOCAL ghMyFont AS DWORD
  LOCAL tmpDword AS DWORD
  dwBaseUnits = GetDialogBaseUnits()

  '----------------------------------------------------------------------------
  ' 创建ComboBox控件窗口.
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
'                                    "程序",_
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
  ' 检查创建ComboBox窗口时的错误.
  '----------------------------------------------------------------------------
  IF ISFALSE hToolbar THEN         ' 错误!
    ' 创建ComboBox窗口失败，通知用户.
    FUNCTION = %TRUE
    EXIT FUNCTION
  END IF
  '----------------------------------------------------------------------------
  ' 循环增加ComboBox文本字符串.
  '----------------------------------------------------------------------------
  SendMessage hToolbar,%WM_SETFONT,GetStockObject(17),0 '%DEFAULT_GUI_FONT),0 '%ANSI_FIXED_FONT), 0
  'SendMessage tmpDword,%WM_SETFONT,GetStockObject(%ANSI_FIXED_FONT), 0
  FOR i = 1 TO 20
    sStr = "第 " & FORMAT$(i) & " 项"
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
' 函数 : CreateRebar ()
'
' 描述 : 创建Rebar控件.
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
               g_hInst, _
               BYVAL %NULL)

  '----------------------------------------------------------------------------
  ' 检查创建Rebar控件窗口的错误.
  '----------------------------------------------------------------------------
  IF ISFALSE ghRebar THEN         ' 错误!
   '= 创建Rebar窗口失败，通知用户.
   FUNCTION = %TRUE
   EXIT FUNCTION
  END IF
  FUNCTION=ghRebar
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
  baseStyle     = %RBBS_NOVERT     OR _    ' 不显示垂直方向
                  %RBBS_CHILDEDGE  OR _
                  %RBBS_FIXEDBMP   OR _
                  %RBBS_GRIPPERALWAYS
  'rbBand.fStyle =

'  '----------------------------------------------------------------------------
'  ' 如果用户选择了一个Rebar背景图片，是加载它!
'  '----------------------------------------------------------------------------
'  IF ISTRUE %USE_RBIMAGE THEN
'   ghRbBack   = LoadImage(g_hInst, "RBBACK", %IMAGE_BITMAP, 0, 0, _
'              %LR_LOADTRANSPARENT OR %LR_LOADMAP3DCOLORS OR _
'              %LR_DEFAULTSIZE)
'
'   rbBand.hbmBack    = ghRbBack           ' Rebar背景图片.
'  END IF
  '读配置
  rowCount=VAL(ReadConfig("rebarrowcount"))
  bandCount=VAL(ReadConfig("rebarbandcount"))
  tmpStr=ReadConfig("rebarband0")
  IF rowCount<=0 OR bandCount<=0 OR MID$(tmpStr,1,1)="0" THEN '取得的值异常，或为空的情况下，仅显示常规工具栏，然后退出本函数
    'msgbox "rowCount=" & str$(rowCount) & $crlf & "bandCount=" & str$(bandCount)
    'RunLog "创建工具栏异常或没有可显示的工具栏：rowCount=" & STR$(rowCount) & $CRLF & "bandCount=" & STR$(bandCount)
    g_hToolbar=CreateToolBar(hWnd)
    '----------------------------------------------------------------------------
    ' 初始化并增加工具栏到Rebar的段.
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
    'RunLog "创建的工具栏：" & "rebarband" & FORMAT$(i) & ":" & tmpStr
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
    IF bandInfo(2)="" THEN  '默认不隐藏，即显示
      bandInfo(2)="0"
    END IF
    rbBand.fStyle = baseStyle
    IF VAL(bandInfo(2))=1 THEN
      rbBand.fStyle=rbBand.fStyle OR %RBBS_HIDDEN
    END IF
    IF bandInfo(3)="" THEN  '默认换行
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
        szTbText          = "程序:"
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
    'RunLog "创建的工具栏：" & barStr & " rebarband" & FORMAT$(i) & ":" & tmpStr
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
' 函数 : SaveRebarBandPos ()
' 描述 : 保存Rebar条位置.
'------------------------------------------------------------------------------
FUNCTION SaveRebarBand() AS LONG
  DIM lResult AS LONG, lBands AS LONG
  DIM rBand   AS REBARBANDINFO
  LOCAL ishidden AS INTEGER
  LOCAL isbreak AS INTEGER
  LOCAL tmpStr AS STRING
  REGISTER lCount AS LONG
  ' 得到Rebar行数.
  lResult = SendMessage(ghRebar, %RB_GETROWCOUNT, 0&, 0&)
  'udtAp.rbRowCount = lResult
  WriteConfig("rebarrowcount",FORMAT$(lResult))
  'RunLog "保存的工具栏：行数：" & FORMAT$(lResult)
  ' 得到Rebar段数.
  lBands = SendMessage(ghRebar, %RB_GETBANDCOUNT, 0&, 0&)
  'udtAp.rbBandCount = lResult
  WriteConfig("rebarbandcount",FORMAT$(lBands))
  'RunLog "保存的工具栏：个数：" & FORMAT$(lBands)
  ' 设置Rebar的REBARBANDINFO结构.
  rBand.cbSize = SIZEOF(rBand)
  rBand.fMask  = %RBBIM_ID OR %RBBIM_SIZE OR %RBBIM_STYLE
  'MSGBOX "rowcount=" & STR$(lResult) & $CRLF & "bandscount=" & STR$(lBands)
  ' 循环得到每段的ID.
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
    'RunLog "保存的工具栏：" & "rebarband" & FORMAT$(lCount) & ":" & tmpStr
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
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_STDTOOLBAR, "标准"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_STDTOOLBAR, "标准"
  END IF
  IF ghWindowBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_WINDOWBAR, "窗口编辑"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_WINDOWBAR, "窗口编辑"
  END IF
  IF ghComboBox THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_COMBOBAR, "过程"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_COMBOBAR, "过程"
  END IF
  IF ghDebugBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_DEBUGBAR, "编译与调试"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_DEBUGBAR, "编译与调试"
  END IF
  IF ghCodeEditBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_CODEEDIT, "代码编辑"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_CODEEDIT, "代码编辑"
  END IF
  IF ghWinArragBar THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %IDM_WINARRAG, "窗口排列"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %IDM_WINARRAG, "窗口排列"
  END IF
  APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_SEPARATOR  , 0, "-"
  IF ReadConfig("rebarlocked")="0" THEN
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_UNCHECKED, %ID_LOCKBAR, "锁定工具栏"
  ELSE
    APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED OR %MF_CHECKED, %ID_LOCKBAR, "锁定工具栏"
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
  FOR i=0 TO bandCount-1
    SendMessage(ghRebar,%RB_GETBANDINFO,i,VARPTR(rbBand))
    IF rbBand.hwndChild=hBar THEN
      SendMessage ghRebar,%RB_SHOWBAND,i,sign
      hasDone=1
      EXIT FOR
    END IF
  NEXT i
  IF hasDone=0 AND sign=1 THEN   '如果没有，且要求显示的话，则在这里创建并插入Band
    IF ghRebarMenu THEN
      MENU GET STATE ghRebarMenu,8 TO MenuState
      IF MenuState=%MF_CHECKED THEN
        rbBand.fStyle= rbBand.fStyle AND NOT(%RBBS_GRIPPERALWAYS)
        rbBand.fStyle= rbBand.fStyle OR %RBBS_NOGRIPPER
      END IF
    END IF
    IF wID=%ID_COMBOBOX THEN
      szTbText          = "程序:"
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
' 返回菜单项字符串及从0开头的位图索引数(-1 = 无位图).
' 位图图片列表条是从0开始的，所以 BmpNum 从0开始
' *****************************************************************************
FUNCTION GetMenuTxtBmp( BYVAL ItemId AS LONG, BmpNum AS LONG ) AS STRING
  SELECT CASE ItemId
    ' 文件菜单
    'CASE %IDM_FILE :        BmpNum = -1 : FUNCTION = "文件"
    CASE %IDM_NEW :           BmpNum = 1 : FUNCTION = "新建"
    CASE %IDM_NEWPROJECT :    BmpNum = 72 : FUNCTION = "新建项目"
    CASE %IDM_NEWBAS :        BmpNum = 45 : FUNCTION = "空BAS文件" & $TAB & "Ctrl+N"    '&New
    CASE %IDM_NEWINC :        BmpNum = 45 : FUNCTION = "空INC文件"
    CASE %IDM_NEWRC :         BmpNum = 45 : FUNCTION = "空RC文件"

    CASE %IDM_OPEN :          BmpNum = 2 : FUNCTION = "打开..." & $TAB & "Ctrl+O"     '&Open
    CASE %IDM_REOPEN :        BmpNum = 2 : FUNCTION = "重打开"
    CASE %IDM_INSERTFILE :    BmpNum = 0 : FUNCTION = "插入文件..."           '&Insert File
    CASE %IDM_SAVE :          BmpNum = 3 : FUNCTION = "保存" & $TAB & "Ctrl+S"    '&Save
    'CASE %IDM_REFRESH :       BmpNum = 108 : FUNCTION = "刷新"
    CASE %IDM_SAVEAS :        BmpNum = 4 : FUNCTION = "另存为..."             'Save File &As
    CASE %IDM_SAVEALL :       BmpNum = 96 : FUNCTION = "全保存"             'Save File &As
    CASE %IDM_SAVEPROJECT :   BmpNum = 112: FUNCTION = "保存项目"
    CASE %IDM_CLOSE :         BmpNum = 63 : FUNCTION = "关闭" & $TAB & "Ctrl+F4"     'Close
    CASE %IDM_CLOSEOTHERS :   BmpNum = 97 : FUNCTION = "关闭其他"     'Close
    CASE %IDM_CLOSEALL :      BmpNum = 6 : FUNCTION = "全部关闭"            'Close A&ll Files
    CASE %IDM_PRINTSETTING :  BmpNum = 7 : FUNCTION = "页面设置..." + $TAB + "Ctrl+Alt+P"     'Page Set&up
    CASE %IDM_PRINTPREVIEW :  BmpNum = 8 : FUNCTION = "打印预览" + $TAB + "Ctrl+Shift+P"
    CASE %IDM_PRINT :         BmpNum = 9 : FUNCTION = "打印代码" + $TAB + "Ctrl+P"  '&Print
    'CASE %IDM_FORM_PRINT :      BmpNum = 9 : FUNCTION = "打印窗口"
    CASE %IDM_OPENCMD :       BmpNum = 10 : FUNCTION = "命令行" + $TAB + "Ctrl+Alt+O"     'Comman&d Prompt
    CASE %IDM_EXIT :          BmpNum = 11 : FUNCTION = "退出" & $TAB & "Alt+F4"     'E&xit
    CASE %IDM_RECENT1 TO %IDM_RECENT8
      BmpNum = 2
      IF ( ItemId - %IDM_RECENTFILES ) = > LBOUND( RecentFiles ) AND _
              ( ItemId - %IDM_RECENTFILES ) < = UBOUND( RecentFiles ) THEN _
              FUNCTION = "&" & FORMAT$( ItemId - %IDM_RECENTFILES ) & " " & RecentFiles( ItemId - %IDM_RECENTFILES )
'    CASE %IDM_PBTEMPLATES + 1 TO %IDM_PBTEMPLATES + %MAXPBTEMPLATES + 1
'      BmpNum = 31
'      FUNCTION = PARSE$( g_PbTemplatesDesc, "|", ItemId - %IDM_PBTEMPLATES )

    ' 编辑菜单
    'CASE %IDM_EDITHEADER :        BmpNum = - 1    : FUNCTION = "编辑"               '&Edit
    CASE %IDM_UNDO :          BmpNum = 12 : FUNCTION = "撤销" + $TAB + "Ctrl+Z"     'U&ndo
    CASE %IDM_REDO :          BmpNum = 13 : FUNCTION = "重复" + $TAB + "Ctrl+K"     'Red&o
    CASE %IDM_CUT :           BmpNum = 17 : FUNCTION = "剪切" + $TAB + "Ctrl+X"     'Cu&t
    CASE %IDM_COPY :          BmpNum = 18 : FUNCTION = "复制" + $TAB + "Ctrl+C"     '&Copy
    CASE %IDM_PASTE :         BmpNum = 19 : FUNCTION = "粘贴" + $TAB + "Ctrl+V"     '&Paste
    CASE %IDM_PASTEIE :       BmpNum = 20 : FUNCTION = "粘贴网页" + $TAB + "Ctrl+Shift+V"     'Paste from Internet Explorer
    CASE %IDM_SELALL :        BmpNum = 21 : FUNCTION = "全选" + $TAB + "Ctrl+A"
    CASE %IDM_DELETE :        BmpNum = 15 : FUNCTION = "删除" + $TAB + "Delete"
    CASE %IDM_LINEDELETE :    BmpNum = 14 : FUNCTION = "删除行" & $TAB & "Ctrl+Y"     '&Delete Line
    CASE %IDM_INITASNEW :     BmpNum = 1  : FUNCTION = "初始化" & $TAB & "Ctrl+Shift+N"     ' *** NTM ***'&Initialize as New
    CASE %IDM_CLEAR :         BmpNum = 15 : FUNCTION = "清除"               'Clea&r
    CASE %IDM_CLEARALL :      BmpNum = 16 : FUNCTION = "全删"             'Cl&ear All
    CASE %IDM_INDENT :        BmpNum = 24 : FUNCTION = "缩进" + $TAB + "Tab"    'Block Indent
    CASE %IDM_OUTDENT :       BmpNum = 25 : FUNCTION = "反缩进" + $TAB + "Shift+Tab"    'Block Unindent
    CASE %IDM_COMMENT :       BmpNum = 22 : FUNCTION = "注释" + $TAB + "Ctrl+Q"
    CASE %IDM_UNCOMMENT :     BmpNum = 23 : FUNCTION = "反注释" + $TAB + "Ctrl+Shift+Q"
    CASE %IDM_FORMATREGION :  BmpNum = 26 : FUNCTION = "格式化文本" + $TAB + "Ctrl+B"
    CASE %IDM_TABULATEREGION :BmpNum = 103: FUNCTION = "按TAB对齐文本" + $TAB + "Ctrl+Shift+B"
    CASE %IDM_SELTOUPPERCASE :BmpNum = 28 : FUNCTION = "转换为大写" + $TAB + "Ctrl+Alt+U"
    CASE %IDM_SELTOLOWERCASE :BmpNum = 29 : FUNCTION = "转换为小写" + $TAB + "Ctrl+Alt+L"
    CASE %IDM_SELTOMIXEDCASE :BmpNum = 30 : FUNCTION = "转换为首字母大写" + $TAB + "Ctrl+Alt+M"
    CASE %IDM_TEMPLATES :     BmpNum = 31 : FUNCTION = "模板代码" + $TAB + "Ctrl+Shift+C"
    CASE %IDM_HTMLCODE :      BmpNum = 32 : FUNCTION = "转换为网页"
    CASE %IDM_GUID :          BmpNum = 115: FUNCTION = "插入新GUID"

    ' 搜索菜单
    'CASE %IDM_SEARCHHEADER :      BmpNum = - 1 : FUNCTION = "查找"
    CASE %IDM_FIND :          BmpNum = 33 : FUNCTION = "查找..." + $TAB + "Ctrl+F"
    CASE %IDM_FINDNEXT :      BmpNum = 34 : FUNCTION = "下一个" + $TAB + "F3"
    CASE %IDM_FINDBACKWARDS : BmpNum = 125 : FUNCTION = "上一个" + $TAB + "Shift+F3"
    CASE %IDM_REPLACE :       BmpNum = 35 : FUNCTION = "替换..." + $TAB + "Ctrl+R"
    CASE %IDM_GOTOLINE :      BmpNum = 36 : FUNCTION = "转到行..." + $TAB + "Ctrl+G"
    ' SEARCH MENU BOOKMARKS
    CASE %IDM_TOGGLEBOOKMARK :BmpNum = 37 : FUNCTION = "设置标签" + $TAB + "Ctrl+F2"
    CASE %IDM_NEXTBOOKMARK :  BmpNum = 38 : FUNCTION = "下一标签" + $TAB + "F2"
    CASE %IDM_PREVIOUSBOOKMARK : BmpNum = 39 : FUNCTION = "上一标签" + $TAB + "Shift+F2"
    CASE %IDM_DELETEBOOKMARKS :BmpNum = 40 : FUNCTION = "移除标签" + $TAB + "Ctrl+Shift+F2"
    ' SEARCH MENU FILEFIND
    CASE %IDM_FINDINFILES :    BmpNum = 123 : FUNCTION = "在文件中查找..." + $TAB + "Ctrl+Shift+F"
    CASE %IDM_EXPLORER :       BmpNum = 116 : FUNCTION = "资源管理器..." + $TAB + "Ctrl+Shift+O"
    CASE %IDM_WINDOWSFIND :    BmpNum = 122 : FUNCTION = "系统查找..." + $TAB + "Ctrl+Alt+F"
    CASE %IDM_FILEFIND :       BmpNum = 41 : FUNCTION = "查找文件..." + $TAB + "Alt+Shift+F"

    ' 运行菜单
    'CASE %IDM_RUNHEADER :       BmpNum = - 1 : FUNCTION = "运行"
    CASE %IDM_COMPILE :       BmpNum = 42 : FUNCTION = "编译" + $TAB + "Ctrl+F5"
    CASE %IDM_COMPILERUN :    BmpNum = 43 : FUNCTION = "编译并执行" + $TAB + "F5"
    CASE %IDM_COMPILEDEBUG :  BmpNum = 44 : FUNCTION = "编译并调试" + $TAB + "F6"
    CASE %IDM_EXECUTE :       BmpNum = 91 : FUNCTION = "执行" + $TAB + "Ctrl+E"
    CASE %IDM_SETPRIMARY :    BmpNum = 45 : FUNCTION = "设置主源文件" + $TAB + "Ctrl+Alt+Y"
    CASE %IDM_CMDPARAMETER :  BmpNum = 46 : FUNCTION = "运行参数"
    CASE %IDM_DEBUGTOOL :     BmpNum = 47 : FUNCTION = "调试工具"

    ' 视图菜单
    CASE %IDM_PROPERTY :      BmpNum =-1 : FUNCTION = "属性"
    CASE %IDM_CONTROLS :      BmpNum = -1 : FUNCTION = "控件"
    CASE %IDM_PROJECT :       BmpNum =-1 : FUNCTION = "项目"
    CASE %IDM_COMPILERS :     BmpNum =-1 : FUNCTION = "编译结果"
    CASE %IDM_FINDRS :        BmpNum =-1 : FUNCTION = "查找结果"
    CASE %IDM_REBAR :         BmpNum =133 : FUNCTION = "工具栏"
    CASE %IDM_STDTOOLBAR:     BmpNum = -1 : FUNCTION = "基本工具"
    CASE %IDM_WINDOWBAR:     BmpNum = -1 : FUNCTION = "窗口编辑"
    CASE %IDM_COMBOBAR:       BmpNum = -1 : FUNCTION = "代码查找"
    CASE %IDM_DEBUGBAR:       BmpNum = -1 : FUNCTION = "编译调试"
    CASE %IDM_CODEEDIT:       BmpNum = -1 : FUNCTION = "代码编辑"
    CASE %IDM_WINARRAG:       BmpNum = -1 : FUNCTION = "窗口排列"
    CASE %ID_LOCKBAR:         BmpNum = -1 : FUNCTION = "锁定"
    CASE %IDM_TOGGLE :        BmpNum = 48 : FUNCTION = "展开/收起过程" + $TAB + "F8"
    CASE %IDM_TOGGLEALL :     BmpNum = 49 : FUNCTION = "展开/收起以下" + $TAB + "Ctrl+F8"
    CASE %IDM_FOLDALL :       BmpNum = 51 : FUNCTION = "收起所有" + $TAB + "Alt+F8"
    CASE %IDM_EXPANDALL :     BmpNum = 50 : FUNCTION = "展开所有" + $TAB + "Shift+F8"
    CASE %IDM_ZOOMIN :        BmpNum = 52 : FUNCTION = "缩小代码文本" + $TAB + "Ctrl++"
    CASE %IDM_ZOOMOUT :       BmpNum = 53 : FUNCTION = "放大代码文本" + $TAB + "Ctrl+-"
    CASE %IDM_USETABS :       BmpNum = - 1 : FUNCTION = "TAB可用" + $TAB + "Ctrl+Shift+T"
    CASE %IDM_AUTOINDENT :    BmpNum = - 1 : FUNCTION = "自动缩进" + $TAB + "Ctrl+Shift+A"
    CASE %IDM_SHOWLINENUM :   BmpNum = - 1 : FUNCTION = "行数" + $TAB + "Ctrl+Shift+L"
    CASE %IDM_SHOWMARGIN :    BmpNum = - 1 : FUNCTION = "收放线" + $TAB + "Ctrl+Shift+M"
    CASE %IDM_SHOWINDENT :    BmpNum = - 1 : FUNCTION = "缩进导引线" + $TAB + "Ctrl+Shift+I"
    CASE %IDM_SHOWSPACES :    BmpNum = - 1 : FUNCTION = "空白" + $TAB + "Ctrl+Shift+W"
    CASE %IDM_SHOWEOL :       BmpNum = - 1 : FUNCTION = "行末" + $TAB + "Ctrl+Shift+D"
    CASE %IDM_SHOWEDGE :      BmpNum = - 1 : FUNCTION = "边缘" + $TAB + "Ctrl+Shift+G"
    CASE %IDM_SHOWPROCNAME :  BmpNum = - 1 : FUNCTION = "显示过程名称" + $TAB + "Ctrl+Shift+S"
    CASE %IDM_CVEOLTOCRLF :   BmpNum = 54 : FUNCTION = "转换行末字符为 $CRLF"
    CASE %IDM_CVEOLTOCR :     BmpNum = 55 : FUNCTION = "转换行末字符为 $CR"
    CASE %IDM_CVEOLTOLF :     BmpNum = 56 : FUNCTION = "转换行末字符为 $LF"
    CASE %IDM_REPLSPCWITHTABS:BmpNum = 120 : FUNCTION = "TAB替换空格"
    CASE %IDM_REPLTABSWITHSPC:BmpNum = 119 : FUNCTION = "空格替换TAB"
    CASE %IDM_FILEPROPERTIES :BmpNum = 127 : FUNCTION = "文件属性"
    CASE %IDM_SYSINFO :       BmpNum = 128 : FUNCTION = "系统信息"

    ' 窗口菜单
    CASE %IDM_CASCADE :       BmpNum = 57 : FUNCTION = "层叠窗口"
    CASE %IDM_TILEH :         BmpNum = 58 : FUNCTION = "水平排列"
    CASE %IDM_TILEV :         BmpNum = 59 : FUNCTION = "垂直排列"
    CASE %IDM_RESTOREWSIZE :  BmpNum = 60 : FUNCTION = "重载主窗口尺寸"
    CASE %IDM_SWITCHWINDOW :  BmpNum = 61 : FUNCTION = "切换窗口    " + $TAB + "Ctrl+W"
    CASE %IDM_ARRANGE :       BmpNum = 62 : FUNCTION = "排列图标"
    CASE %IDM_CLOSEWIN :      BmpNum = 63 : FUNCTION = "关闭"
    CASE %IDM_MAXWIN :        BmpNum = 93 : FUNCTION = "最大化"

    ' 项目菜单
    CASE %IDM_NEWPROJECT :    BmpNum = 1 : FUNCTION = "新建项目"
    CASE %IDM_OPENPROJECT :   BmpNum = 2 : FUNCTION = "打开项目"
    CASE %IDM_LOADPROJECT :   BmpNum = 101 : FUNCTION = "加载项目"
    CASE %IDM_CLOSEPROJECT :  BmpNum = 5 : FUNCTION = "关闭项目"
    CASE %IDM_SAVEPROJECT :   BmpNum = 3 : FUNCTION = "保存项目"
    CASE %IDM_SAVEPROJECTAS : BmpNum = 4 : FUNCTION = "另存项目..."
    CASE %IDM_REGPROJECTEXT : BmpNum = 46 : FUNCTION = "关联项目文件"
    CASE %IDM_ACHIVE :        BmpNum = 83 : FUNCTION = "归档项目"
    CASE %IDM_NEWVISION :     BmpNum = 95 : FUNCTION = "项目升版"
    CASE %IDM_FILEVISIONSETTING : BmpNum = 127 : FUNCTION = "文件版本"
    CASE %IDM_VISIONPACKAGE :     BmpNum = 121 : FUNCTION = "项目版本"
    CASE %IDM_INSERTINPROJECT :   BmpNum = 112 : FUNCTION = "增加文件到项目"
'    CASE %IDM_RECENTPRJ1 TO %IDM_RECENTPRJ4
'      BmpNum = 2
'      IF ( ItemId - %IDM_RECENTPROJECTS ) = > LBOUND( RecentProjects ) AND _
'              ( ItemId - %IDM_RECENTPROJECTS ) < = UBOUND( RecentProjects ) THEN _
'              FUNCTION = "&" & FORMAT$( ItemId - %IDM_RECENTPROJECTS ) & " " & RecentProjects( ItemId - %IDM_RECENTPROJECTS )

    ' 设置菜单
    'CASE %IDM_OPTIONSHEADER :     BmpNum = - 1 : FUNCTION = "选项"
    CASE %IDM_INTERFACEOPT :  BmpNum = 129 : FUNCTION = "界面设置"
    CASE %IDM_FILEEXTOPT :    BmpNum = 102 : FUNCTION = "关联文件"
    CASE %IDM_EDITOROPT :     BmpNum = 64 : FUNCTION = "代码编辑"
    CASE %IDM_COLORSOPT :     BmpNum = 66 : FUNCTION = "代码颜色字体"
    CASE %IDM_FORMOROPT :     BmpNum = 78 : FUNCTION = "窗体编辑"
    CASE %IDM_COMPILEROPT :   BmpNum = 65 : FUNCTION = "编译设置"
    CASE %IDM_TOOLSOPT :      BmpNum = 67 : FUNCTION = "工具设置"
    CASE %IDM_FOLDINGOPT :    BmpNum = 100 : FUNCTION = "卷起设置"
    CASE %IDM_PRINTOPT :      BmpNum = 9 : FUNCTION = "打印设置"
    CASE %IDM_SHORTKEYOPT :   BmpNum = 105 : FUNCTION = "热键设置"

    ' 工具菜单
    'CASE %IDM_TOOLSHEADER :     BmpNum = - 1 : FUNCTION = "工具"
    CASE %IDM_MAKEMENU :      BmpNum = 69 : FUNCTION = "创建菜单"
    CASE %IDM_MAKETOOLBAR :   BmpNum = 69 : FUNCTION = "创建菜单"
    CASE %IDM_GOTCHA :        BmpNum = 123 : FUNCTION = "精确查找"
    CASE %IDM_FINDMULTIDIR :  BmpNum = 123 : FUNCTION = "多目录查找"
    CASE %IDM_REPLACEMULTIDIR : BmpNum = 123 : FUNCTION = "多目录替换"
    CASE %IDM_FORMATCODE :    BmpNum = 123 : FUNCTION = "格式化代码"
    CASE %IDM_CODEC :         BmpNum = 69 : FUNCTION = "分析代码"
    CASE %IDM_CODECOUNT :     BmpNum = 69 : FUNCTION = "代码统计"
    CASE %IDM_CODEMANAGER :   BmpNum = 69 : FUNCTION = "代码管理"
    CASE %IDM_CODEKEEPER :    BmpNum = 71 : FUNCTION = "代码保持"
    CASE %IDM_CODEMANAGER :   BmpNum = 69 : FUNCTION = "代码管理"
    CASE %IDM_KBDMACROS :     BmpNum = 129 : FUNCTION = "代码快捷键"
    CASE %IDM_INCLEAN :       BmpNum = 82 : FUNCTION = "代码清洁"
    CASE %IDM_WINSPY :        BmpNum = 82 : FUNCTION = "SPY代码"
    CASE %IDM_CHARMAP :       BmpNum = 117 : FUNCTION = "字符表"
    CASE %IDM_MSGBOXDESIGNER :BmpNum = 121 : FUNCTION = "消息框设计"
    CASE %IDM_PBFORMS :       BmpNum = 72 : FUNCTION = "PB窗体"
    CASE %IDM_PBCOMBR :       BmpNum = 73 : FUNCTION = "PB COM 浏览"
    CASE %IDM_TYPELIBBR :     BmpNum = 126 : FUNCTION = "PB LIB 浏览"
    CASE %IDM_IMGEDITOR :     BmpNum = 76 : FUNCTION = "编辑图片"
    CASE %IDM_RCEDITOR :      BmpNum = 77 : FUNCTION = "编辑资源"
    CASE %IDM_CODETIPSBUILDER : BmpNum = 70 : FUNCTION = "代码提示构建"
    CASE %IDM_CODETYPEBUILDER : BmpNum = 98 : FUNCTION = "代码类型构建"
    CASE %IDM_DLGEDITOR :       BmpNum = 75 : FUNCTION = "对话框编辑"
    CASE %IDM_POFFS :         BmpNum = 80 : FUNCTION = "Poffs"
    CASE %IDM_COPYCAT :       BmpNum = 83 : FUNCTION = "CopyCat"
    CASE %IDM_CALCULATOR :    BmpNum = 84 : FUNCTION = "计算器"
    CASE %IDM_MORETOOLS :     BmpNum = 124 : FUNCTION = "更多..."

    '协同菜单
    CASE %IDM_COOPSETTING :   BmpNum = 124 : FUNCTION = "协同设置"
    CASE %IDM_SENDINVITE :    BmpNum = 124 : FUNCTION = "邀请协同"
    CASE %IDM_SENDVISIT :     BmpNum = 124 : FUNCTION = "参加协同"
    CASE %IDM_DOWNLOADFILE :  BmpNum = 124 : FUNCTION = "下载文件"
    CASE %IDM_UPLOADFILE :    BmpNum = 124 : FUNCTION = "上传文件"
    CASE %IDM_CHAT :          BmpNum = 124 : FUNCTION = "会话"

    ' 帮助菜单
    'CASE %IDM_HELPHEADER :        BmpNum = - 1 : FUNCTION = "帮助"
    CASE %IDM_IDEHELP :       BmpNum = 85 : FUNCTION = "IDE 帮助"
    CASE %IDM_PBWINCONTENT :  BmpNum = 85 : FUNCTION = "PB/WIN 主题"
    CASE %IDM_PBWININDEX :    BmpNum = 85 : FUNCTION = "PB/WIN 索引"
    CASE %IDM_PBWINSEARCH :   BmpNum = 85 : FUNCTION = "PB/WIN 查询"
    CASE %IDM_WINDOWSSDK :    BmpNum = 85 : FUNCTION = "WinSDK 帮助"
    CASE %IDM_ABOUT :         BmpNum = 86 : FUNCTION = "关于"
    CASE %IDM_UPDATE :        BmpNum = 86 : FUNCTION = "升级"

    '工具栏下拉，及部分右键菜单
    CASE %IDM_WINTEMPLATES + 1
      BmpNum = 75
      FUNCTION = GetLang("Normal Window")'"一般窗口"
    CASE %IDM_WINTEMPLATES + 2
      BmpNum = 75
      FUNCTION = GetLang("MDI Window") '"MDI窗口"
    CASE %ID_SETH_LEFT
      BmpNum=143
      FUNCTION = GetLang("Left Align")'"左对齐"
    CASE %ID_SETH_CENTER
      BmpNum=135
      FUNCTION = GetLang("H-Center Align")'"水平居中"
    CASE %ID_SETH_RIGHT
      BmpNum=144
      FUNCTION = GetLang("Right Align")'"右对齐"
    CASE %ID_SETV_TOP
      BmpNum=145
      FUNCTION = GetLang("Top Align")'"上对齐"
    CASE %ID_SETV_CENTER
      BmpNum=134
      FUNCTION = GetLang("V-Center Align")'"垂直居中"
    CASE %ID_SETV_BOTTOM
      BmpNum=146
      FUNCTION = GetLang("Bottom Align")'"下对齐"
    CASE %MN_F_DDTCODE
      BmpNum=-1
      FUNCTION = GetLang("DDT Code")'"DDT代码"
    CASE %MN_F_SDKCODE
      BmpNum=-1
      FUNCTION = GetLang("SDK Code")'"SDK代码"
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
      BmpNum = 36 : FUNCTION = "转到最近位置"
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
' 自绘菜单项 - 由主窗口的%WM_DRAWITEM 消息调用
' *********************************************************************************************
SUB DrawMenu( BYREF lp AS LONG )
  REGISTER c          AS LONG     ' 用于 for/next 计数
  LOCAL ShortCutWidth AS LONG     ' 快捷方式文本长度(如果有)
  LOCAL ID            AS LONG     ' 临时存储菜单项ID
  LOCAL BmpNum        AS LONG     ' 位图索引值
  LOCAL hBmp          AS LONG     ' 要绘制的图标句柄
  LOCAL hBrush        AS LONG     ' 选中时的绘制刷
  LOCAL bHighlight    AS LONG     ' 高亮标识
  LOCAL drawTxtFlags  AS DWORD    ' 绘制文本方式标识
  LOCAL drawBmpFlags  AS DWORD    ' 绘制图标方式标识
  LOCAL menuBtnW      AS LONG     ' 菜单按钮象素宽度
  LOCAL sCaption      AS STRING   ' 菜单项文本字符串
  LOCAL sShortCut     AS STRING   ' 快捷方式文本
  LOCAL rc            AS RECT     ' 绘制区域
  LOCAL szl           AS SIZEL    ' 用于计算文本长度
  LOCAL lpDis AS DRAWITEMSTRUCT PTR ' we fill this from received pointer
  LOCAL bGrayed       AS LONG
  LOCAL bDisabled     AS LONG
  LOCAL nState        AS LONG     ' For ImageList Bitmap Strip ID.
  lpDis = lp    ' 从给定指针中获取结构
  menuBtnW = GETSYSTEMMETRICS( %SM_CYMENU ) - 1     ' 用于 "按钮" 尺寸
  '----------------------------------------------------------------------------
  ' 计算快捷方式最终位置。需要检查下拉的所有菜单项并计算最大宽度，
  ' 这个最大宽度即为所有快捷方式文本输出位置的起点
  '----------------------------------------------------------------------------
  FOR c = 0 TO GETMENUITEMCOUNT( @lpDis.hwndItem ) - 1 ' hwndItem 是被下拉的菜单项句柄
    ID = GETMENUITEMID( @lpDis.hwndItem, c )
    IF ID THEN
      sCaption = GetMenuTxtBmp( ID, 0 )
      IF INSTR( sCaption, $TAB ) THEN     ' 如果有快捷方式文本 (Ctrl+X, etc.)
        sShortCut = TRIM$( PARSE$( sCaption, $TAB, 2 )) '取快捷方式文本
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
  ' 获取菜单项字符串并分成最终的文本及快捷方式
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
  ' 计算菜单项的字符串尺寸
  '----------------------------------------------------------------------------
  GETTEXTEXTENTPOINT32 @lpDis.hDC, BYVAL STRPTR( sCaption ), LEN( sCaption ), szl
  '----------------------------------------------------------------------------
  ' 根据动作设置颜色
  '----------------------------------------------------------------------------
  hMenuTextBkBrush = CREATESOLIDBRUSH( GETSYSCOLOR(%COLOR_HIGHLIGHTTEXT))'SciColorsAndFonts.SubmenuTextBackColor )
  ' Creates the brush for the submenu highlighted text background 创建子菜单高亮文本背景的刷子
  hMenuHiBrush = CREATESOLIDBRUSH( GETSYSCOLOR(%COLOR_MENUHILIGHT))'SciColorsAndFonts.SubmenuHiTextBackColor )
  'hMenuHiBrush = CreateSolidBrush(RGB(64, 134, 191))
  ' Creates the brush for the menu icons background 创建菜单图标背景刷子
  hMenuIconsBrush = CREATESOLIDBRUSH( RGB( 193, 211, 239 ))
  IF ( @lpDis.itemState AND %ODS_SELECTED ) THEN    ' 菜单项被选中
    nState = %IMG_HOT
    hBrush = hMenuHiBrush
    SETBKCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_HIGHLIGHTTEXT ) '%COLOR_HIGHLIGHT )
    SETTEXTCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_HIGHLIGHTTEXT )
    bHighlight = %TRUE    ' 标识，这样在后面就能增加或移除图标周围的边框
    ' 被选中，但是是灰色的或不可用的
    IF ( @lpDis.itemState AND %ODS_GRAYED ) THEN
      nState = %IMG_DIS
      bGrayed = %TRUE
    ELSEIF ( @lpDis.itemState AND %ODS_DISABLED ) THEN
      nState = %IMG_DIS
      bDisabled = %TRUE
    END IF
  ELSEIF ( @lpDis.itemState AND %ODS_GRAYED ) THEN    ' 菜单项是灰的
    nState = %IMG_DIS
    hBrush = GETSYSCOLORBRUSH( %COLOR_MENU )
    SETBKCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENU )
    SETTEXTCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENUTEXT )
    bGrayed = %TRUE
  ELSEIF ( @lpDis.itemState AND %ODS_DISABLED ) THEN    ' 菜单项不可用
    nState = %IMG_DIS
    hBrush = GETSYSCOLORBRUSH( %COLOR_MENU )
    SETBKCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENU )
    SETTEXTCOLOR @lpDis.hDC, GETSYSCOLOR ( %COLOR_MENUTEXT )
    bDisabled = %TRUE
  ELSE    ' 未选择或不可用
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
  IF hBmp <> 0 THEN     ' 绘制位图，并放在"按钮"的中间
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
    ' 为3D边框计算需要的RECT
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
  DELETEOBJECT hBmp     ' 完成时删除位图，避免内存泄漏
  '----------------------------------------------------------------------------
  ' 绘制菜单文本到位图按钮，且纵向居中
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
'  if sCaption="新建" then
'    msgbox "test" & str$(id)
'  end if
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Sub Routine : MeasureMenu ()
' Description : 由%WM_MEASUREITEM 消息调用
' *********************************************************************************************
SUB MeasureMenu( BYVAL hWnd AS DWORD, BYVAL lParam AS DWORD )
  LOCAL hDC AS DWORD
  LOCAL hf AS DWORD
  LOCAL txt AS STRING
  LOCAL sl AS SIZEL
  LOCAL lpMis AS MEASUREITEMSTRUCT PTR
  LOCAL ncm AS NONCLIENTMETRICS
  lpMis = lParam    ' lParam 指向 MEASUREITEMSTRUCT
  txt = GetMenuTxtBmp( @lpMis.itemID, 0 )    ' 获取菜单项的文本
  hDC = GETDC( hWnd )     ' 获取对话框 DC
  ncm.cbSize = SIZEOF( ncm )    ' 获取菜单字体
  SYSTEMPARAMETERSINFO %SPI_GETNONCLIENTMETRICS, SIZEOF( ncm ), BYVAL VARPTR( ncm ), 0
  IF LEN( ncm.lfMenuFont ) THEN
    hf = CREATEFONTINDIRECT( ncm.lfMenuFont )     ' 以菜单LogFont数据创建字体
    IF hf THEN hf = SELECTOBJECT( hDC, hf )     ' 在DC中选择字体
  END IF
  GETTEXTEXTENTPOINT32 hDC, BYVAL STRPTR( txt ), LEN( txt ), sl     ' 获取文本长度
  '     @lpMis.itemWidth  = sl.cx                                 ' set width
  @lpMis.itemWidth = sl.cx + %OMENU_EXTRAWIDTH    ' 设置宽度
  @lpMis.ItemHeight = GETSYSTEMMETRICS( %SM_CYMENU ) + 1    ' 设置高度
  IF hf THEN DELETEOBJECT SELECTOBJECT( hDC, hf )     ' 删除临时菜单字体对象
  RELEASEDC hWnd, hDC     ' 释放DC
  IF @lpMis.itemId = 0 THEN 'OR @lpMis.itemId = %IDM_RECENTSEPARATOR OR @lpMis.itemId = %IDM_RECENTPROJECTSSEPARATOR THEN     ' Separator
    @lpMis.itemHeight = @lpMis.itemHeight \ 2
  END IF
END SUB
' 锁定Rebar工具条
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
  IF sign=0 THEN     '0为解锁, 1为加锁
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
' TOOLBAR.INC  结束
'------------------------------------------------------------------------------
