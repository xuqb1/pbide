'------------------------------------------------------------------------------
TYPE SED_SavePos
   Position AS LONG                ' 最后的位置
   FileName AS ASCIIZ * 255 '%MAX_PATH  ' 文件名
END TYPE
GLOBAL g_hInst        AS DWORD      '主程序的实例
GLOBAL g_hMenu        AS DWORD      '主程序菜单
GLOBAL g_hMenuFile    AS DWORD      '文件菜单
GLOBAL g_hMenuReopen  AS DWORD      '重打开文件
GLOBAL g_hMenuEdit    AS DWORD      '编辑菜单
GLOBAL g_hMenuSearch  AS DWORD      '查找菜单
GLOBAL g_hMenuRun     AS DWORD      '运行菜单
GLOBAL g_hMenuView    AS DWORD      '视图菜单
GLOBAL g_hMenuTool    AS DWORD      '工具菜单
GLOBAL g_hMenuProject AS DWORD      '项目菜单
GLOBAL g_hMenuOptions AS DWORD      '设置菜单
GLOBAL g_hMenuCoop    AS DWORD      '协作菜单
GLOBAL g_hMenuWindow  AS DWORD      '窗口菜单
GLOBAL g_hMenuHelp    AS DWORD      '帮助菜单
GLOBAL g_hStatus      AS DWORD      '状态条句柄
GLOBAL g_hWndMain     AS DWORD      '主窗口句柄
GLOBAL g_hWndClient   AS DWORD      '客户区句柄
GLOBAL g_hGuiderWin   AS DWORD      '向导窗口句柄
GLOBAL g_hGuiderTab() AS DWORD      '向导窗口中的Tab句柄数组
GLOBAL g_hOptionWin   AS DWORD      '选项窗口句柄
GLOBAL g_fClosed      AS LONG       '最后一个编辑窗口关闭标识
GLOBAL g_FreezeMenu   AS LONG       '冻结菜单
GLOBAL g_hTabMdi      AS DWORD
GLOBAL g_NewComCtl    AS CURRENCYX  '
GLOBAL g_NewDocNum    AS LONG       '
GLOBAL g_zIni         AS ASCIIZ * 500 '%MAX_PATH '
GLOBAL ghRebar    AS DWORD         ' ReBar控件句柄.
GLOBAL ghToolBar  AS DWORD         ' 工具栏控件句柄.
GLOBAL ghWindowBar AS DWORD        ' 窗口工具条句柄
GLOBAL ghWndSize AS DWORD          ' 窗口工具条后面的在编辑窗口位置大小显示部分
GLOBAL ghComboBox AS DWORD         ' ComboBox控件句柄.
GLOBAL ghDebugBar AS DWORD         ' 编译调试工具条句柄
GLOBAL ghCodeEditBar AS DWORD      '代码编辑工具栏句柄
GLOBAL ghWinArragBar AS DWORD      ' 窗口排列工具栏句柄
GLOBAL g_FileType     AS DWORD      '标识新建时的文件类型:0,bas;1,inc;2,rc

GLOBAL ghRebarMenu AS DWORD        'Rebar上的右键弹出菜单句柄
GLOBAL g_hRebar   AS DWORD
GLOBAL g_hToolbar     AS DWORD      '工具栏句柄

GLOBAL pbtempfile     AS STRING
GLOBAL language       AS STRING
GLOBAL canLog         AS INTEGER
GLOBAL ghFont         AS LONG     '设定属性列表中字的字体的句柄
GLOBAL CodeTipsFile   AS STRING
GLOBAL hCodetipsFile  AS LONG
GLOBAL hCodeKeepFile  AS LONG
GLOBAL hFont          AS LONG
GLOBAL hCurSplit      AS LONG
GLOBAL hCurSplitV     AS LONG
GLOBAL hCurNorm       AS LONG
GLOBAL strPBKeyWords  AS STRING
GLOBAL fMaximize      AS LONG
GLOBAL ConstructAutocompletion  AS LONG
GLOBAL AutoCompleteTypes        AS LONG
GLOBAL TrimTrailingBlanks       AS LONG
GLOBAL ShowProcedureName        AS LONG
GLOBAL sProjectName             AS STRING
GLOBAL sAutoCompletionTypes     AS STRING
GLOBAL bMsgWindow               AS LONG
GLOBAL g_hTabMdi                AS DWORD
GLOBAL hTabImageList            AS DWORD
'GLOBAL hCodeFinder              AS DWORD  '由ghComboBox替代
GLOBAL cUntitledFiles           AS LONG
GLOBAL gTabFilePaths()          AS STRING
GLOBAL ghEditStr                AS STRING
GLOBAL strContextMenu           AS STRING
GLOBAL SED_LastPosition         AS SED_SavePos
GLOBAL fClosed              AS LONG
GLOBAL sProjectPrimary      AS STRING
GLOBAL ghTreeView           AS DWORD
GLOBAL hPTROOT              AS DWORD
GLOBAL PrimarySourceFile    AS STRING
GLOBAL hHandles()           AS LONG         'Gripper handles
GLOBAL gZorder()            AS LONG         '保存句柄以维护我们自己的Z向列表
GLOBAL hPopupClient         AS LONG         '被编辑窗口弹出菜单句柄
GLOBAL hPopupCtl            AS LONG         '被编辑窗口上当前控件的右键弹出菜单句柄
GLOBAL hWndMbuilder         AS DWORD
GLOBAL hWndEdit             AS DWORD
GLOBAL hWndMDIEdit          AS DWORD
GLOBAL IDhitValue           AS LONG         '数字,句柄点测试指示
' 项目相关
GLOBAL hProjectTV           AS DWORD '项目窗口中的树视图
GLOBAL hPTRoot              AS DWORD         ' 树根句柄
GLOBAL hPTRCParent          AS DWORD
GLOBAL hPTBASParent         AS DWORD
GLOBAL hPTINCParent         AS DWORD
GLOBAL ghProjectTV          AS DWORD
GLOBAL RecentProjects()     AS STRING

TYPE PROJECT_PATH_NAMES
  sFName     AS STRING * 25
  sFPathName AS STRING * 255 '%MAX_PATH
END TYPE
GLOBAL udtProjectData()     AS PROJECT_PATH_NAMES
GLOBAL szProjectDir         AS STRING

GLOBAL hFind                AS DWORD '查找/查找替换窗口句柄

GLOBAL strCompilerTitle        AS STRING
GLOBAL CommandLine             AS STRING                 ' 编译运行程序时，调用的命令行参数
GLOBAL hPbIde               AS DWORD
$PBWINTITLE     = "PowerBASIC Compiler for Windows IDE"
$PBWIN80TITLE   = "PB/Win IDE"
$PBCCTITLE      = "PowerBASIC Console Compiler IDE"
$PBCC80TITLE    = "PB/CC IDE"
'global ghMsgWindow          as dword '被hCompilersTB替代
$MsgTitle = "Pre-processor"
$U01_SEDB="'# SED_Block"+":"

GLOBAL RecentFiles()           AS STRING                 ' // Recent file names
