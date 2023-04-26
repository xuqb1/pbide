$PROGRAMCLASSNAME       = "VisPowerBASIC"  ' 主窗口类
$INIFILENAME            = "config.ini"    ' 配置文件
$HELPFILE               = "PBWin.hlp"      ' 帮助文件 (可根据需要修改名字)
$DATABASENAME           = "data.db"
$CONFIGTABLENAME        = "t_config"
$DEFAULTCONFIGTABLENAME = "t_config_default"
$DATASETTABLENAME       = "t_dataset"
$STRINGLANGTABLENAME    = "t_string_lang"
$CONTROLSTYLETABLENAME  = "t_control_style"
$SCI_CODETIPSDB         = "SCI_CDTIP.TRM"
$SCI_CODEKEEPDB         = "SCI_CODKP.TRM"
$LYNXWINDOW             = "GTI_LYNX"
$OldWndProc             = "OldWndProc"
$MainEdit               = "MAINFORM"
$MainMDIEdit            = "MAINMDIFORM"
'------------------------------------------------------------------------------
'%MAX_PATH = 500
'泊坞窗口ID
%ID_CLIENTWINDOW  = 100  '%WM_USER值是1024，主窗口客户区实体窗口ID
%ID_STATUSBAR     = 101  ' 状态栏
%IDC_TABMDI       = 102  ' 标题Tab栏

%ID_PROPERTY      = 110   '属性窗口ID
%ID_CONTROLS      = 111   '控件窗口ID
%ID_PROJECT       = 112   '项目窗口ID
%ID_COMPILERS     = 113   '编译结果窗口ID
%ID_FINDRS        = 114   '查找结果窗口ID

%ID_FTOOLBAR      = 120  ' 工具箱窗口中的工具栏

'泊坞窗口对泊坞区贴边情况，组织后赋给byside
%BYLEFT       = 8
%BYTOP        = 4
%BYRIGHT      = 2
%BYBOTTOM     = 1
'上下左右四个客户区的泊坞区贴边情况
%TOLEFT       = 4
%TOTOP        = 4
%TOCENTER     = 2 '任何一个泊坞区等于0,则说明不存在；大于等于2，则存在(实际使用大于0)
%TORIGHT      = 1
%TOBOTTOM     = 1
'泊坞到的位置
%MLEFT        = 1 '主窗口的左上右下：1,2,3,4
%MTOP         = 2
%MRIGHT       = 3
%MBOTTOM      = 4
%DLEFT        = 5 '某个泊坞窗口的左上右下：5,6,7,8
%DTOP         = 6
%DRIGHT       = 7
%DBOTTOM      = 8



%IDM_NEWTEMP          = 130  ' 按模板新建文件 +20
' 文件菜单
%IDM_NEWBAS           = 160  ' 新建bas文件
%IDM_NEWINC           = 161  ' 新建inc文件
%IDM_NEWRC            = 162  ' 新建rc文件
%IDM_OPEN             = 163  ' 打开文件
%IDM_OPENPROJECT      = 164  ' 打开项目文件
%IDM_INSERTFILE       = 165  ' 插入文件
%IDM_NEW              = 166  ' 新建菜单
%IDM_REOPEN           = 167  ' 重打开菜单
%IDM_REFRESH          = 168  ' 刷新菜单
'-
%IDM_SAVE             = 170  ' 保存
%IDM_SAVEAS           = 171  ' 另存为
%IDM_SAVEALL          = 172  ' 全部保存
%IDM_SAVEPROJECT      = 173  ' 保存项目
'-
%IDM_PRINTSETTING     = 174  ' 打印设置
%IDM_PRINTPREVIEW     = 175  ' 打印预览
%IDM_PRINT            = 176  ' 打印...
'-
%IDM_CLOSE            = 177  '关闭
%IDM_CLOSEALL         = 178  '关闭全部
%IDM_CLOSEOTHERS      = 179  '关闭其他
'-
%IDM_RECENTFILES      = 180
%IDM_RECENT1          = 181  ' 最近打开文件 1
%IDM_RECENT2          = 182  ' 最近打开文件 2
%IDM_RECENT3          = 183  ' 最近打开文件 3
%IDM_RECENT4          = 184  ' 最近打开文件 4
%IDM_RECENT5          = 185  ' 最近打开文件 5
%IDM_RECENT6          = 186  ' 最近打开文件 6
%IDM_RECENT7          = 187  ' 最近打开文件 7
%IDM_RECENT8          = 188  ' 最近打开文件 8
'%IDM_RECENT9          = 189  ' 最近打开文件 9
'%IDM_RECENT10         = 190  ' 最近打开文件 10
'-
%IDM_OPENCMD          = 190  ' 打开命令行
%IDM_EXIT             = 191  ' 退出

' 编辑菜单
%IDM_UNDO             = 200   ' 撤销
%IDM_REDO             = 201   ' 重做
'-
%IDM_CUT              = 202   ' 剪切
%IDM_COPY             = 203   ' 复制
%IDM_PASTE            = 204   ' 粘贴
%IDM_PASTEIE          = 205   ' 粘贴
%IDM_SELALL           = 206   ' 全选
%IDM_DELETE           = 207   ' 删除
%IDM_LINEDELETE       = 208   ' 删除
%IDM_CLEARALL         = 209   ' 删除
'-
%IDM_COMMENT          = 210   ' 注释
%IDM_UNCOMMENT        = 211   ' 反注释
%IDM_INDENT           = 212   ' 缩进
%IDM_OUTDENT          = 213   ' 反缩进
%IDM_FORMATREGION     = 214   ' 格式化
%IDM_TABULATEREGION   = 215   ' 列表化
'-
%IDM_SELTOUPPERCASE   = 216   ' 大写化
%IDM_SELTOLOWERCASE   = 217   ' 小写化
%IDM_SELTOMIXEDCASE   = 218   ' 驼峰化
'-
%IDM_TEMPLATES        = 219   ' 模板代码
%IDM_HTMLCODE         = 220   ' 转为网页
%IDM_GUID             = 221   ' 插入GUID
'%IDM_BLOCK            = 2039&   ' 块
'%IDM_SPACEINDENT      = 2044&   ' TAB反缩进
'%IDM_SPACEOUTDENT     = 2045&   ' TAB反缩进
'-

'搜索菜单
%IDM_FIND             = 230   ' 查找
%IDM_FINDNEXT         = 231   ' 查找下一个
%IDM_FINDBACKWARDS    = 232
%IDM_REPLACE          = 233   ' 替换
%IDM_GOTOLINE         = 234   ' 转到行
'-
%IDM_TOGGLEBOOKMARK   = 235
%IDM_NEXTBOOKMARK     = 236
%IDM_PREVIOUSBOOKMARK = 237
%IDM_DELETEBOOKMARKS  = 238
'%IDM_GOTOBOOKMARK     = 2054&   ' 转到书签
%IDM_FINDINFILES      = 239
%IDM_EXPLORER         = 240
%IDM_WINDOWSFIND      = 241
%IDM_FILEFIND         = 242

'运行菜单
%IDM_COMPILE          = 250   ' 编译
%IDM_COMPILERUN       = 251   ' 编译并运行
%IDM_COMPILEDEBUG     = 252   ' 编译并调试
%IDM_EXECUTE          = 253   ' 运行
'-
%IDM_SETPRIMARY       = 254   ' 设置主源文件
%IDM_CMDPARAMETER     = 255   ' 设置命令行参数
%IDM_DEBUGTOOL        = 256   ' 调试工具

'视图菜单
%IDM_PROPERTY         = 260
%IDM_CONTROLS         = 261
%IDM_PROJECT          = 262
%IDM_COMPILERS        = 263
%IDM_FINDRS           = 264
'-
%IDM_REBAR            = 265
'-
%IDM_TOGGLE           = 266
%IDM_TOGGLEALL        = 267
%IDM_FOLDALL          = 268
%IDM_EXPANDALL        = 269
'-
%IDM_ZOOMIN           = 270
%IDM_ZOOMOUT          = 271
'-
%IDM_USETABS          = 272
%IDM_AUTOINDENT       = 273
%IDM_SHOWLINENUM      = 274
%IDM_SHOWMARGIN       = 275
%IDM_SHOWINDENT       = 276
%IDM_SHOWSPACES       = 277
%IDM_SHOWEOL          = 278
%IDM_SHOWEDGE         = 279
%IDM_SHOWPROCNAME     = 280
'-
%IDM_CVEOLTOCRLF      = 281
%IDM_CVEOLTOCR        = 282
%IDM_CVEOLTOLF        = 283
'-
%IDM_REPLSPCWITHTABS  = 284
%IDM_REPLTABSWITHSPC  = 285
'-
%IDM_FILEPROPERTIES   = 286
%IDM_SYSINFO          = 287


' 窗口菜单
%IDM_CASCADE          = 290   ' 层叠窗口
%IDM_TILEH            = 291   ' 水平排列
%IDM_TILEV            = 292   ' 垂直排列
'-
%IDM_RESTOREWSIZE     = 293   ' 重载主窗口尺寸
%IDM_SWITCHWINDOW     = 294   ' 转到窗口
%IDM_ARRANGE          = 295   ' 图标排列
'-
%IDM_CLOSEWIN         = 296   ' 全部关闭
%IDM_MAXWIN           = 297   ' 最大化窗口

'项目菜单
%IDM_NEWPROJECT       = 300   ' 新建项目
'%IDM_OPENPROJECT      = 164  ' 打开项目文件
%IDM_LOADPROJECT      = 301   ' 加载项目
'%IDM_SAVEPROJECT     = 2111&   ' 保存项目,使用文件菜单中的%IDM_SAVEPROJECT
%IDM_SAVEPROJECTas    = 302   ' 另存项目
%IDM_CLOSEPROJECT     = 303   ' 关闭项目
'-
%IDM_REGPROJECTEXT    = 304   ' 关联项目文件
'-
%IDM_ACHIVE           = 305   ' 项目归档
%IDM_NEWVISION        = 306   ' 项目新版本
%IDM_FILEVISIONSETTING= 307   ' 项目内文件版本设置
%IDM_VISIONPACKAGE    = 308   ' 版本打包
'-
%IDM_RECENTPRJ1       = 310   ' 最近项目1
%IDM_RECENTPRJ2       = 311   ' 最近项目2
%IDM_RECENTPRJ3       = 312   ' 最近项目3
%IDM_RECENTPRJ4       = 313   ' 最近项目4
%IDM_RECENTPRJ5       = 314   ' 最近项目5
%IDM_RECENTPRJ6       = 315   ' 最近项目6
%IDM_RECENTPRJ7       = 316   ' 最近项目7
%IDM_RECENTPRJ8       = 317   ' 最近项目8

'设置菜单
%IDM_INTERFACEOPT     = 320
%IDM_FILEEXTOPT       = 321
%IDM_EDITOROPT        = 322
%IDM_COLORSOPT        = 323
%IDM_FORMOROPT        = 324
%IDM_COMPILEROPT      = 325
%IDM_TOOLSOPT         = 326
%IDM_FOLDINGOPT       = 327
%IDM_PRINTOPT         = 328
%IDM_SHORTKEYOPT      = 329

'工具
%IDM_MAKEMENU         = 340   ' 制作菜单
%IDM_MAKETOOLBAR      = 341   ' 制作TOOLBAR
%IDM_GOTCHA           = 342
%IDM_FINDMULTIDIR     = 343   ' 在多个目录中查找
%IDM_REPLACEMULTIDIR  = 344   ' 在多个目录中替换
'-
%IDM_FORMATCODE       = 345   ' 格式化代码
%IDM_CODEC            = 346   ' 格式化代码
%IDM_CODECOUNT        = 347   ' 代码统计
%IDM_CODEMANAGER      = 348   ' 代码管理器
%IDM_CODEKEEPER       = 349   ' 代码保持
%IDM_KBDMACROS        = 350   ' 代码快捷键
%IDM_INCLEAN          = 351   ' 代码清洁
%IDM_WINSPY           = 352   ' 对目标窗口生成代码
'-
%IDM_CHARMAP          = 353
%IDM_MSGBOXDESIGNER   = 354
%IDM_PBFORMS          = 355
%IDM_PBCOMBR          = 356
%IDM_TYPELIBBR        = 357
%IDM_IMGEDITOR        = 358
%IDM_RCEDITOR         = 359
%IDM_CODETIPSBUILDER  = 360
%IDM_CODETYPEBUILDER  = 361
%IDM_DLGEDITOR        = 362
%IDM_POFFS            = 363
%IDM_COPYCAT          = 364
%IDM_CALCULATOR       = 365
%IDM_MORETOOLS        = 366
'-
%IDM_CUSTOMTOOL1      = 370   ' 自定义工具1
%IDM_CUSTOMTOOL2      = 371   ' 自定义工具2
%IDM_CUSTOMTOOL3      = 372   ' 自定义工具3
%IDM_CUSTOMTOOL4      = 373   ' 自定义工具4
%IDM_CUSTOMTOOL5      = 374   ' 自定义工具5
%IDM_CUSTOMTOOL6      = 375   ' 自定义工具6
%IDM_CUSTOMTOOL7      = 376   ' 自定义工具7
%IDM_CUSTOMTOOL8      = 377   ' 自定义工具8
'-
%IDM_CUSTOMTOOLSETTING= 2100&   ' 自定义工具设置

'协作
%IDM_COOPSETTING      = 2130&   ' 协作设置
%IDM_SENDINVITE       = 2131&   ' 发送邀请
%IDM_SENDVISIT        = 2132&   ' 发送访问请求
%IDM_DOWNLOADFILE     = 2133&   ' 下载文件
%IDM_UPLOADFILE       = 2134&   ' 上载文件
%IDM_CHAT             = 2135&   ' 沟通

'-
%IDM_OPTION           = 2149&   ' 选项...

' 帮助
%IDM_IDEHELP          = 2160&   ' IDE帮助
%IDM_PBWINCONTENT     = 2161&   ' PB/WIN帮助内容
%IDM_PBWININDEX       = 2162&   ' PB/WIN帮助索引
%IDM_PBWINSEARCH      = 2163&   ' PB/WIN帮助搜索
%IDM_WINDOWSSDK       = 2164&   ' WINDOWS SDK帮助
'-
%IDM_ABOUT            = 2166&   ' 关于
%IDM_UPDATE           = 2167&   ' 升级

'窗口编辑工具栏按钮命令定义
%IDM_WIN              = 2200    ' 新建窗口
%ID_TOOL_MENU         = 2201    ' 新建菜单
%ID_TOOL_TOOLBAR      = 2202    ' 新建工具栏
%ID_BARONLY_STICKY    = 2203    ' 粘贴
%ID_BARONLY_SNAP      = 2204    ' 对齐到网格
%ID_BARONLY_GRID      = 2205    ' 显示/隐藏网格
%ID_SETV_CENTER       = 2206    ' 垂直居中对齐
%ID_SETH_CENTER       = 2207    ' 水平居中对齐
%ID_SETV_EQUATE       = 2208    ' 垂直等距分布
%ID_SETH_EQUATE       = 2209    ' 水平等距分布
%ID_TOOL_COMPILE      = 2210    ' 转成DDT或SDK代码
%ID_TOOL_POS          = 2211    '位置按钮
%ID_TOOL_WH           = 2212    '宽高按钮
%ID_SETH_LEFT         = 2213    '水平左对齐按钮
%ID_SETH_RIGHT        = 2214    '水平右对齐按钮
%ID_SETV_TOP          = 2215    '垂直上对齐按钮
%ID_SETV_BOTTOM       = 2216    '垂直下对齐按钮
%MN_F_DDTCODE         = 2217    'DDT代码
%MN_F_SDKCODE         = 2218    'SDK代码

'编译，调试菜单/工具栏
%IDM_SETBRKPO         = 2230    '设置断点
%IDM_PRVBRKPO         = 2231    '上一断点
%IDM_NEXBRKPO         = 2232    '下一断点
%IDM_DELBRKPO         = 2233    '删除断点
%IDM_GOBRKPO          = 2234    '转到断点
'%IDM_COMPILE          = 2235    '编译
%IDM_CMPRUN           = 2236    '上一断点
%IDM_CMPDEBUG         = 2237    '上一断点
'%IDM_PRVBRKPO         = 2238    '上一断点
'%IDM_PRVBRKPO         = 2239    '上一断点


'代码编辑菜单/工具栏
%IDM_BLOCKCOMMIT      = 2250    '块注释
%IDM_REBLOCKCOMMIT    = 2251    '反块注释
%IDM_BLOCKINDENT      = 2252    '块缩进
%IDM_REBLOCKINDENT    = 2253    '反块缩进
%IDM_CAPWORD          = 2254    '转大写
%IDM_SMALLWORD        = 2255    '转小写
%IDM_FIRSTCAPWORD     = 2256    '首字母大写
'%IDM_FIND             = 2257    '查找
'%IDM_FINDNEXT         = 2258    '查找下一个
%IDM_FINDREPLACE      = 2259    '查找并替换
%IDM_GOLINE           = 2260    '转到行

%IDM_PBTEMPLATES    = 9000      ' 基于指示器的PB模板
%IDM_WINTEMPLATES   = 8000      ' 窗口模板

%ID_GUIDERTAB                   = 2270    '向导窗口相关ID
%IDC_GUILDNEWLIST               = 2271    '向导窗口新建TAB中的列表视图
%IDC_GUILDEXISTFINDLABEL        = 2272    '向导窗口现存TAB中的"查找范围:"标签
%IDC_GUILDEXISTFINDCOMBOBOX     = 2273    '向导窗口现存TAB中的"查找范围:"组合框
%IDC_GUILDEXISTFINDLIST         = 2274    '向导窗口现存TAB中的"查找范围:"列表视图
%IDC_GUILDEXISTFILENAMELABEL    = 2275    '向导窗口现存TAB中的"文件名:"标签
%IDC_GUILDEXISTFILENAMETEXTBOX  = 2276    '向导窗口现存TAB中的"文件名:"文本框
%IDC_GUILDEXISTFILETYPELABEL    = 2277    '向导窗口现存TAB中的"文件类型:"标签
%IDC_GUILDEXISTFILETYPECOMBOBOX = 2278    '向导窗口现存TAB中的"文件类型:"组合框
%IDC_GUILDRECENTLIST            = 2279    '向导窗口最新TAB中的列表视图
%IDC_GUILDHELPBUTTON            = 2280    '向导窗口帮助按钮
%IDC_GUILDSHOWCHECKBOX          = 2281    '向导窗口"不再显示这个对话框"检查框


%IDC_EDIT1    = 2310
%ID_ProcEnd    = 2311

%SPLITSIZE                  = 4
%IDM_INITASNEW      = 3303      ' 初始化为新的
%IDM_CLEAR          = 3304      ' 清除
'%IDM_CLEARALL       = 3305      ' 清除所有
'%IDM_SELECTALL      = 3309      ' 全选
'%IDM_PASTEIE        = 3312      ' 粘贴 HTML
'%IDM_LINEDELETE     = 3313      ' 删除行
%IDM_AUTOCOMPLETE   = 3314      ' 自动完成
'%IDM_TEMPLATES      = 3315      ' 模板
'%IDM_HTMLCODE       = 3316      ' 转换为 Html
'%IDM_SELTOUPPERCASE = 3317      ' 转换所选文本为大写
'%IDM_SELTOLOWERCASE = 3318      ' 转换所选文本为小写
'%IDM_SELTOMIXEDCASE = 3319      ' 转换所选文本为混合大小写
'%IDM_FORMATREGION   = 3320      ' 格式化所选文本
'%IDM_TABULATEREGION = 3321      ' 列表选择文本
'%IDM_BLOCKINDENT    = 3322      ' 块缩进
%IDM_BLOCKUNINDENT  = 3323      ' 块取消缩进
%IDM_GOTOBEGINPROC  = 3324      ' 跳转到程序或函数的开头
%IDM_GOTOENDPROC    = 3325      ' 跳转到程序或函数的末尾
%IDM_GOTOBEGINTHISPROC = 3326       ' 跳转到此程序或函数的开头
%IDM_GOTOENDTHISPROC = 3327         ' 跳转到此程序或函数的末尾

%IDK_COMMENT        = 3331      ' 行/块注释
%IDK_UNCOMMENT      = 3332      ' 行/块取消注释
%IDM_HELP           = 3551

'%IDM_REGPROJECTEXT  = 3609      ' 注册文件扩展名
%IDM_INSERTINPROJECT= 3610      ' 在当前项目中插入文件
%IDM_RECENTPROJECTS = 3650      ' 最近项目
'%IDM_RECENTPRJ1     = 3651      ' 最近打开项目1
'%IDM_RECENTPRJ2     = 3652      ' 最近打开项目2
'%IDM_RECENTPRJ3     = 3653      ' 最近打开项目3
'%IDM_RECENTPRJ4     = 3654      ' 最近打开项目4

%IDM_GOTOSELFILE    = 3801      ' 跳转或加载光标后具有相同文件名的文件
%IDM_GOTOSELPROC    = 3802      ' 跳转到光标后具有相同名字的函数/过程
%IDM_GOTOLASTPOSITION = 3803        ' 跳转到最近保存的点
%IDM_RECENTPROJECTSSEPARATOR = 3655         ' 分隔条

%ID_LYNX_FIND       = 110

%IDM_FORM_PRINT     = 3217      '打印窗口

'* Popup Contextmenu WM_RBUTTONDOWN on a Control被编辑窗口上的弹出菜单,关联到控件的WM_RBUTTONDOWN
%MN_UNDO            =40200      '撤销
%MN_REDO            =40201      '重做
%MN_CUT             =40202      '剪切该控件Cut this control
%MN_COPY            =40203      '复制该控件Copy this control
%MN_PASTE           =40204      '粘贴该控件Paste this control
%MN_DELETE          =40205      '删除该控件Delete this control
%MN_UPLAYER         =40206      '上移一层
%MN_DOWNLAYER       =40207      '下移一层
%MN_TOPLAYER        =40208      '移到最上面
%MN_BOTTOMLAYER     =40209      '移到最下面
%MN_CTLPROP         =40210          '控件属性
'* Popup Contextmenu WM_RBUTTONDOWN on the Client被编辑窗口上的弹出菜单,关联到窗口的WM_RBUTTONDOWN
%MN_F_UNDO          = 40220     '撤销
%MN_F_REDO          = 40221     '重做
'%MN_F_DDTCODE       = 40222     'DDT代码
'%MN_F_SDKCODE       = 40223     'SDK代码
%MN_F_PROP          = 40224     '窗口属性
%keyLen = 300
%dataLen = 1000
%use_HANDLESMALL = 1
%use_HANDLEMEDIUM = 2
%use_HANDLELARGE = 3
%use_backColor = 1
%use_backBrush = 2
%WM_POSITIONSIZE       = %WM_USER +136
%TVS_NOSCROLL = &H00002000     'OR %TVS_NOSCROLL
''' tokens

%undoDepth      = 100

%moveWindow     = 1
%CreateWindowEx = 2
%DestroyWindow  = 3
%SetWindowLong  = 4
%SetWindowText  = 5
%WS_MINIMAL = %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS

%IMAGE_PRJ  = 0
%IMAGE_FOLD = 1
%IMAGE_BAS  = 2
%IMAGE_INC  = 3
%IMAGE_RC   = 4
%IMAGE_PRI  = 5

' Keyboard Macros - Do not change unless you also modify the SED_InsertMacroText procedure
%IDM_ALT0              = 3700
%IDM_ALT1              = 3701
%IDM_ALT2              = 3702
%IDM_ALT3              = 3703
%IDM_ALT4              = 3704
%IDM_ALT5              = 3705
%IDM_ALT6              = 3706
%IDM_ALT7              = 3707
%IDM_ALT8              = 3708
%IDM_ALT9              = 3709
%IDM_CTRL0             = 3710
%IDM_CTRL1             = 3711
%IDM_CTRL2             = 3712
%IDM_CTRL3             = 3713
%IDM_CTRL4             = 3714
%IDM_CTRL5             = 3715
%IDM_CTRL6             = 3716
%IDM_CTRL7             = 3717
%IDM_CTRL8             = 3718
%IDM_CTRL9             = 3719
' 打印设置窗口
%IDCPS_CBPAPERSIZE     = 6020
%IDCPS_CBPAPERBIN      = 6021
%IDCPS_PORTRAIT        = 6022
%IDCPS_LANDSCAPE       = 6023
%IDCPS_CBFONTNAME      = 6024
%IDCPS_CBFONTSIZE      = 6025
%IDCPS_CKFONTBOLD      = 6026
%IDCPS_CKFONTITALIC    = 6027
%IDCPS_CKFONTUNDERLINE = 6028
%IDCPS_MARGINLEFT      = 6029
%IDCPS_MARGINRIGHT     = 6030
%IDCPS_MARGINTOP       = 6031
%IDCPS_MARGINBOTTOM    = 6032
%IDCPS_LINEHEIGHT      = 6033
