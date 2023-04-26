' *********************************************************************************************
' Include file for the Scintilla edit control
' *********************************************************************************************

' The following structures are defined to be exactly the same shape as the Win32
' CHARRANGE, TEXTRANGE, FINDTEXTEX, FORMATRANGE, and NMHDR structs.
' So older code that treats Scintilla as a RichSci will work.
'TYPE CharacterRange DWORD
'  cpMin AS LONG
'  cpMax AS LONG
'END TYPE
TYPE start_type
    POS           AS RECT
    fname1          AS ASCIIZ * 255
    fname2          AS ASCIIZ * 255
    fname3          AS ASCIIZ * 255
    fName4          AS ASCIIZ * 255
END TYPE
' *********************************************************************************************
TYPE PROC
    WhatIsUp      AS LONG
    WhatIsDown      AS LONG
    UpLnNo          AS LONG
    DnLnNo          AS LONG
    ProcName      AS ASCIIZ * 65
END TYPE
TYPE CHARRANGE DWORD
  cpMin AS LONG
  cpMax AS LONG
END TYPE
TYPE TEXTRANGE DWORD
  chrg AS CHARRANGE
  lpstrText AS ASCIIZ PTR                 ' allocated by caller, zero terminated by RichSci
END TYPE
TYPE TextToFind DWORD
  chrg AS CHARRANGE
  lpstrText AS ASCIIZ PTR
  chrgText AS CHARRANGE
END TYPE
' This structure is used in printing and requires some of the graphics types
' from Platform.h.  Not needed by most client code.
TYPE RangeToFormat DWORD
  hdc AS DWORD
  hdcTarget AS DWORD
  rc AS RECT
  rcPage AS RECT
  chrg AS CHARRANGE
END TYPE
TYPE NotifyHeader
  hwndFrom AS DWORD
  idfrom AS DWORD
  CODE AS LONG            ' used for messages, so needs to be LONG, not DWORD...
END TYPE
TYPE SCNotification
  hdr AS NotifyHeader
  position AS LONG
  ch AS LONG
  modifiers AS LONG
  modificationType AS LONG
  lpText AS ASCIIZ PTR
  length AS LONG
  linesAdded AS LONG
  message AS DWORD
  wParam AS DWORD
  lParam AS LONG
  nLine AS LONG
  foldLevelNow AS LONG
  foldLevelPrev AS LONG
  MARGIN AS LONG
  listType AS LONG
  x AS LONG
  y AS LONG
END TYPE
%ID_ProcNull = 0
%ID_ProcStart = 1
' Constants specific to the PowerBasic lexer
%SCLEX_VB = 8           ' // Visual Basic lexer
%SCLEX_POWERBASIC = 51          ' // Power Basic lexer Sub Folding Only!
%SCLEX_POWERBASIC2 = 99                 ' // Power Basic lexer All Folding.
%SCE_B_DEFAULT = 0              ' // in both VB and PB lexer
%SCE_B_COMMENT = 1              ' // in both VB and PB lexer
%SCE_B_NUMBER = 2               ' // in both VB and PB lexer
%SCE_B_KEYWORD = 3              ' // in both VB and PB lexer
%SCE_B_STRING = 4               ' // in both VB and PB lexer
%SCE_B_PREPROCESSOR = 5                 ' // VB lexer only, unsupported by PB lexer
%SCE_B_OPERATOR = 6             ' // in both VB and PB lexer
%SCE_B_IDENTIFIER = 7           ' // in both VB and PB lexer
%SCE_B_DATE = 8                 ' // VB lexer only, unsupported by PB lexer
%SCE_B_STRINGEOL = 9            ' // The following 4 constants were added to ver 1.58
%SCE_B_KEYWORD2 = 10            ' ""
%SCE_B_KEYWORD3 = 11            ' ""
%SCE_B_KEYWORD4 = 12            ' ""
%SCE_B_CONSTANT = 23            ' // For PowerBASIC Constants
%SCE_B_LABEL = 24               ' // For PowerBASIC Labels
TYPE EditorOptionsType      '编辑器选项类型定义
  UseTabs                 AS LONG         ' 用tab代替空格
  TabSize                 AS LONG         ' tab的相当字符数
  AutoIndent              AS LONG         ' 使用自动缩进
  IndentSize              AS LONG         ' 自动缩进相当的字符数
  LineNumbers             AS LONG         ' 显示行数
  LineNumbersWidth        AS LONG         ' 行数列的宽度（象素为单位）
  MARGIN                  AS LONG         ' 显示折叠符号的空白
  MarginWidth             AS LONG         ' 折叠空白的宽度（象素为单位）
  EdgeColumn              AS LONG         ' 显示行的宽度的边
  EdgeWidth               AS LONG         ' 边线限制(PB 限制为255个字符)
  IndentGuides            AS LONG         ' 显示缩进导引线
  WhiteSpace              AS LONG         ' 显示空格为点
  EndOfLine               AS LONG         ' 显示行末标记
  DefaultCase             AS LONG         ' 默认大小写
  KeywordCase             AS LONG         ' 关键字大小写
  Magnification           AS LONG         ' 缩放比率(+20/-20 点)
  SyntaxHighlighting      AS LONG         ' 使用语法高亮
  CodeTips                AS LONG         ' 显示代码提示
  MaximizeMainWindow      AS LONG         ' 启动时最大化主窗口
  MaximizeEditWindows     AS LONG         ' 最大化编辑窗口
  AskBeforeExit           AS LONG         ' 退出编辑器时询问
  AllowMultipleInstances  AS LONG      ' 允许同时运行多个编辑器
  ConstructAutocompletion AS LONG         ' 自动完成结构
  AutocompleteTypes       AS LONG         ' 自动完成类型
  TrimTrailingBlanks      AS LONG         ' 剪除末尾的空白
  ShowProcedureName       AS LONG         ' 显示过程名
  ShowCaretLine           AS LONG         ' 显示脱字符号行
  StartInLastFolder       AS LONG         ' 开始于最近的折叠
  LastFolder              AS ASCIIZ * %MAX_PATH
  ReloadFilesAtStartup    AS LONG        ' 启动时重载文件
  BackupEditorFiles       AS LONG         ' 保存时备份原文件
  AttachResources         AS LONG         ' 绑定资源
  CallPostProcessor       AS LONG         ' 调用预处理器
  DdocPrinting            AS LONG         ' 打印时使用ddoc
END TYPE
' *********************************************************************************************
' *********************************************************************************************
TYPE FoldingOptionsType         '折叠选项类型定义
  FoldingLevel        AS LONG         ' 折叠层次(关键字, 过程或没有)
  FoldingSymbol       AS LONG         ' 折叠符号(箭头，加/减，圆弧，方框树)
END TYPE
' *********************************************************************************************
' *********************************************************************************************
TYPE CompilerOptionsType        '编译器选项类型
  DefaultCompiler     AS LONG                         ' 使用的编译器 (PBWin 或 PBCC)
  PBWINPath           AS ASCIIZ * %MAX_PATH       ' PBWin 编译器路径
  PBWINIncPath        AS ASCIIZ * %MAX_PATH       ' PBWin 的包含文件文件夹
  PBCCPath            AS ASCIIZ * %MAX_PATH       ' PBCC 编译器路径
  PBCCIncPath         AS ASCIIZ * %MAX_PATH       ' PBCC 的包含文件文件夹
  RCPath              AS ASCIIZ * %MAX_PATH       ' 资源编译器路径
  PBRESPath           AS ASCIIZ * %MAX_PATH       ' PBRES 工具路径
  RCIncPath           AS ASCIIZ * %MAX_PATH       ' 资源编译器包含文件路径(resource.h)
  DisplayResults      AS LONG                     ' 显示编译结果
  DeleteLogFile       AS LONG                     ' 删除由编译器生成的日志文件
  BeepOnCompletion    AS LONG                     ' 结束时Beep
  DebugToolPath       AS ASCIIZ * %MAX_PATH       ' 调试工具路径
  PBWINVersion        AS LONG                     ' PBWIN编译器版本
  PBCCVersion         AS LONG                     ' PBCC编译器版本
END TYPE
' *********************************************************************************************
' *********************************************************************************************
TYPE ToolsOptionsType       '工具选项类型定义
  PBForms             AS ASCIIZ * %MAX_PATH       ' Path of PBForms路径
  PBComBrowser        AS ASCIIZ * %MAX_PATH       ' Path of the Power Basic COM browser
  TypeLibBrowser      AS ASCIIZ * %MAX_PATH       ' Path of my Typelib Browser tool
  Lynx                AS ASCIIZ * %MAX_PATH       ' Path of Lynx tool语法工具
  DialogEditor        AS ASCIIZ * %MAX_PATH       ' Path of the dialog editor对话框工具
  ImageEditor         AS ASCIIZ * %MAX_PATH       ' Path of the image editor图像编辑器
  ResourceEditor      AS ASCIIZ * %MAX_PATH       ' Path of the resource editor资源编辑器
  VisualDesigner      AS ASCIIZ * %MAX_PATH       ' Path of the visual designer you use可视化设计工具
  ToolbarDesigner     AS ASCIIZ * %MAX_PATH       ' Path of the toolbar designer工具栏设计器
  Poffs               AS ASCIIZ * %MAX_PATH       ' Path of Borje's Poffs tool
  PBWinSpy            AS ASCIIZ * %MAX_PATH       ' Path of Borje's PBWinSpy tool
  IncLean             AS ASCIIZ * %MAX_PATH       ' Path of Borje's IncLean tool
  CopyCat             AS ASCIIZ * %MAX_PATH       ' Path of Borje's CopyCat tool
  PBWinHelp           AS ASCIIZ * %MAX_PATH       ' Path of the PBWin help file  PBWin帮助文件
  PBCCHelp            AS ASCIIZ * %MAX_PATH       ' Path of the PBCC help file PBCC帮助文件
  Win32Help           AS ASCIIZ * %MAX_PATH       ' Path of the Win32 help file Win32帮助文件
  MSDN                AS ASCIIZ * %MAX_PATH       ' Path of the SDK Platform help SDK帮助文件
  RCHelp              AS ASCIIZ * %MAX_PATH       ' Path of the resource compiler help 资源编译器帮助文件
END TYPE
' *********************************************************************************************
' *********************************************************************************************
TYPE SciColorsAndFontsType      'Sci颜色和字体类型定义
  ' Default默认值
  DefaultForeColor            AS LONG                 ' Default foreground color默认前景色
  DefaultBackColor            AS LONG                 ' Default background color默认背景色
  DefaultFontName             AS ASCIIZ * 255         ' Default font name       默认字体名称
  DefaultFontCharset          AS ASCIIZ * 255         ' Default font charset    默认字符集
  DefaultFontSize             AS LONG                 ' Default font size in points默认字体尺寸（按点算）
  DefaultFontBold             AS LONG                 ' Use bold style          使用粗体风格
  DefaultFontItalic           AS LONG                 ' Use italic style        使用斜体风格
  DefaultFontUnderline        AS LONG                 ' Use underline style     使用下划线风格
  ' Comments注释
  CommentForeColor            AS LONG                 ' Comment's foreground color注释的前景色
  CommentBackColor            AS LONG                 ' Comment's background color注释的背景色
  CommentFontName             AS ASCIIZ * 255         ' Comment's font name       注释的字体名称
  CommentFontCharset          AS ASCIIZ * 255         ' Comment's charset         注释的字符集
  CommentFontSize             AS LONG                 ' Comment's font size in points注释的字体尺寸（按点算）
  CommentFontBold             AS LONG                 ' Use bold style            使用粗体风格
  CommentFontItalic           AS LONG                 ' Use italic style          使用斜体风格
  CommentFontUnderline        AS LONG                 ' Use underline style       使用下划线风格
  ' Constants常量
  ConstantForeColor           AS LONG                 ' Foreground color          前景色
  ConstantBackColor           AS LONG                 ' Background color          背景色
  ConstantFontName            AS ASCIIZ * 255         ' Font's name               字体名称
  ConstantFontCharset         AS ASCIIZ * 255         ' Charset                   字符集
  ConstantFontSize            AS LONG                 ' Font's size in points     按点算的字体尺寸
  ConstantFontBold            AS LONG                 ' Bold style                粗体风格
  ConstantFontItalic              AS LONG                 ' Italic style              斜体风格
  ConstantFontUnderline       AS LONG                 ' Underline style           下划线风格
  ' Identifier标识符
  IdentifierForeColor         AS LONG                 ' Foreground color          前景色
  IdentifierBackColor         AS LONG                 ' Background color          背景色
  IdentifierFontName          AS ASCIIZ * 255         ' Font's name               字体名称
  IdentifierFontCharset       AS ASCIIZ * 255         ' Charset                   字符集
  IdentifierFontSize          AS LONG                 ' Font's size in points     按点算的字体尺寸
  IdentifierFontBold          AS LONG                 ' Bold style                粗体风格
  IdentifierFontItalic        AS LONG                 ' Italic style              斜体风格
  IdentifierFontUnderline     AS LONG                 ' Underline style           下划线风格
  ' Keywords 关键字
  KeywordForeColor            AS LONG                 ' Foreground color          前景色
  KeywordBackColor            AS LONG                 ' Background color          背景色
  KeywordFontName             AS ASCIIZ * 255         ' Font's name               字体名称
  KeywordFontCharset          AS ASCIIZ * 255         ' Charset                   字符集
  KeywordFontSize             AS LONG                 ' Font's size               字体尺寸
  KeywordFontBold             AS LONG                 ' Bold style                粗体风格
  KeywordFontItalic           AS LONG                 ' Italic style              斜体风格
  KeywordFontUnderline        AS LONG                 ' Underline style           下划线风格
  ' Numbers数字
  NumberForeColor             AS LONG                 ' Foreground color
  NumberBackColor             AS LONG                 ' Background color
  NumberFontName              AS ASCIIZ * 255         ' Font's name
  NumberFontCharset           AS ASCIIZ * 255         ' Charset
  NumberFontSize              AS LONG                 ' Font's size
  NumberFontBold              AS LONG                 ' Bold style
  NumberFontItalic            AS LONG                 ' Italic style
  NumberFontUnderline         AS LONG                 ' Underline style
  ' Line numbers行数
  LineNumberForeColor         AS LONG                 ' Foreground color
  LineNumberBackColor         AS LONG                 ' Background color
  LineNumberFontName          AS ASCIIZ * 255         ' Font's name
  LineNumberFontCharset       AS ASCIIZ * 255         ' Charset
  LineNumberFontSize          AS LONG                 ' Font's size
  LineNumberFontBold          AS LONG                 ' Bold style
  LineNumberFontItalic        AS LONG                 ' Italic style
  LineNumberFontUnderline     AS LONG                 ' Underline style
  ' Operators 操作符
  OperatorForeColor           AS LONG                 ' Foreground color
  OperatorBackColor           AS LONG                 ' Background color
  OperatorFontName            AS ASCIIZ * 255         ' Font's name
  OperatorFontCharset         AS ASCIIZ * 255         ' Charset
  OperatorFontSize            AS LONG                 ' Font's size
  OperatorFontBold            AS LONG                 ' Bold style
  OperatorFontItalic          AS LONG                 ' Italic style
  OperatorFontUnderline       AS LONG                 ' Underline style
  ' Preprocessor预处理程序
  PreprocessorForeColor       AS LONG                 ' Foreground color
  PreprocessorBackColor       AS LONG                 ' Background color
  PreprocessorFontName        AS ASCIIZ * 255         ' Font's name
  PreprocessorFontCharset     AS ASCIIZ * 255         ' Charset
  PreprocessorFontSize        AS LONG                 ' Font's size
  PreprocessorFontBold        AS LONG                 ' Bold style
  PreprocessorFontItalic      AS LONG                 ' Italic style
  PreprocessorFontUnderline   AS LONG                 ' Underline style
  ' Strings 字符串
  StringForeColor             AS LONG                 ' Foreground color
  StringBackColor             AS LONG                 ' Background color
  StringFontName              AS ASCIIZ * 255         ' Font's name
  StringFontCharset           AS ASCIIZ * 255         ' Charset
  StringFontSize              AS LONG                 ' Font's size
  StringFontBold              AS LONG                 ' Bold style
  StringFontItalic            AS LONG                 ' Italic style
  StringFontUnderline         AS LONG                 ' Underline style
  ' Caret ^符号
  CaretForeColor              AS LONG                 ' Color of the caret^符号的颜色
  ' Edge边
  EdgeForeColor               AS LONG                 ' Edge's foreground color边的前景色
  EdgeBackColor               AS LONG                 ' Edge's background color边的背景色
  ' Fold折叠
  FoldForeColor               AS LONG                 ' Folding margin highlight color折叠空白高亮颜色
  FoldBackColor               AS LONG                 ' Folding margin background color折叠空白背景色
  ' Fold open折叠打开
  FoldOpenForeColor           AS LONG                 ' Foreground color of the folding symbols for expanded functions函数展开的折叠符号前景色
  FoldOpenBackColor           AS LONG                 ' Background color of the folding symbols for expanded functions函数展开的折叠符号背景色
  ' Fold margin折叠空白
  FoldMarginForeColor         AS LONG                 ' Foreground color of the folding symbols for folded functions函数折叠的折叠符号前景色
  FoldMarginBackColor         AS LONG                 ' Background color of the folding symbols for folded functions函数折叠的折叠符号背景色
  ' Indent guides缩进向导
  IndentGuideForeColor        AS LONG                 ' Indent guide's foreground color缩进向导前景色
  IndentGuideBackColor        AS LONG                 ' Indent guide's background color缩进向导背景色
  ' Selection选择
  SelectionForeColor          AS LONG                 ' Selection's foreground color选中部分的前景色
  SelectionBackColor          AS LONG                 ' Selection's background color选中部分的背景色
  ' Whitespace空白位置
  WhitespaceForeColor         AS LONG                 ' White space's foreground color空白位置的前景色
  WhitespaceBackColor         AS LONG                 ' White space's background color空白位置的背景色
  ' Codetip代码提示
  CodetipForeColor            AS LONG                 ' Codetip's foreground color代码提示的前景色
  CodetipBackColor            AS LONG                 ' Codetip's background color代码提示的背景色
  ' Menus 菜单
  SubMenuTextForeColor        AS LONG                 ' Submenu text foreground color子菜单文本前景色
  SubMenuTextBackColor        AS LONG                 ' Submenu text background color子菜单文本背景色
  SubMenuHiTextForeColor      AS LONG                 ' Submenu hughlighted text foreground color子菜单高亮文本前景色
  SubMenuHiTextBackColor      AS LONG                 ' Submenu hughlighted text background color子菜单高亮文本背景色
  ' Caret line ^行
  CaretLineBackColor          AS LONG                 ' Caret line background color ^行背景色
  ' Menu icons 菜单图标
  MenuIconsBackColor          AS LONG                 ' Menu icons background color 菜单图标背景色
  ' Use always the default background color 总是使用默认背景色
  UseAlwaysDefaultBackColor   AS LONG
  ' Use always the default font             总是使用默认字体
  UseAlwaysDefaultFont        AS LONG
  ' Use always the default font size        总是使用默认字体尺寸
  UseAlwaysDefaultFontSize    AS LONG
END TYPE
' *********************************************************************************************
' *********************************************************************************************
' Printer options for DDOC  DDOC的打印机选项
' *********************************************************************************************
TYPE PrinterSetupOptionsType
  PaperSize               AS LONG             ' Paper size纸张大小
  PaperBin                AS LONG             ' Paper bin 纸张屉位
  ORIENTATION             AS LONG             ' 0 = Portrait纵排, 1 = Landscape横排
  FontName                AS ASCIIZ * 255     ' Font name字体名称
  FontSize                AS LONG             ' Font size字体尺寸
  FontBold                AS LONG             ' Font bold字体粗体
  FontItalic              AS LONG             ' Font italic字体斜体
  FontUnderline           AS LONG             ' Font underline字体下划线
  MarginLeft              AS SINGLE           ' Left margin in inches左边空白
  MarginRight             AS SINGLE           ' Right margin in inches右边空白
  MarginTop               AS SINGLE           ' Top margin in inches顶上空白
  MarginBottom            AS SINGLE           ' Bottom margin in inches底边空白
  LineHeight              AS SINGLE           ' Line height行高
END TYPE
TYPE WININFOSTRUC       '窗口信息结构类型定义
  hBotR               AS DWORD        ' 右下角编辑窗口句柄
  hBotL               AS DWORD        ' 左下角编辑窗口句柄
  hTopR               AS DWORD        ' 右上角编辑窗口句柄
  hTopL               AS DWORD        ' 左上角编辑窗口句柄
  hFocus              AS DWORD        ' 具有焦点的编辑窗口句柄
  SplitY              AS INTEGER      ' 光标最后的Y坐标
  SplitXB             AS INTEGER      ' 光标最后的底X坐标
  SplitXT             AS INTEGER      ' 光标最后的顶X坐标
  OldX                AS INTEGER      ' 最后客户区宽度
  OldY                AS INTEGER      ' 最后客户区高度
  bInit               AS WORD         ' 每个控件的初始化状态标识
END TYPE
$EDITCLASSNAME = "PBEDIT"       ' MDI 子类

%IDC_EDIT             = %WM_USER      + 1026&  ' 编辑控件

'%IDC_EDIT1                  = 3110      ' Edit control    编辑控件
%IDC_EDIT2                  = 3111      ' Edit control BL 左下
%IDC_EDIT3                  = 3112      ' Edit control TR 右上
%IDC_EDIT4                  = 3113      ' Edit control TL 左上
$SCI_TYPEDB = "SCI_TYPE.TRM"    ' Types database name
%IDL_CODETYPES_LABEL = 6101
%IDB_CODETYPES_FIND = 6102
%IDB_CODETYPES_APPEND = 6103
%IDB_CODETYPES_BROWSE = 6104
%IDB_CODETYPES_HELP = 6105
%INVALID_POSITION = - 1
%SCI_START = 2000
%SCI_OPTIONAL_START = 3000
%SCI_LEXER_START = 4000
%SCI_ADDTEXT = 2001
%SCI_ADDSTYLEDTEXT = 2002
%SCI_INSERTTEXT = 2003
%SCI_CLEARALL = 2004
%SCI_CLEARDOCUMENTSTYLE = 2005
%SCI_GETLENGTH = 2006
%SCI_GETCHARAT = 2007
%SCI_GETCURRENTPOS = 2008
%SCI_GETANCHOR = 2009
%SCI_GETSTYLEAT = 2010
%SCI_REDO = 2011
%SCI_SETUNDOCOLLECTION = 2012
%SCI_SELECTALL = 2013
%SCI_SETSAVEPOINT = 2014
%SCI_GETSTYLEDTEXT = 2015
%SCI_CANREDO = 2016
%SCI_MARKERLINEFROMHANDLE = 2017
%SCI_MARKERDELETEHANDLE = 2018
%SCI_GETUNDOCOLLECTION = 2019
%SCWS_INVISIBLE = 0
%SCWS_VISIBLEALWAYS = 1
%SCWS_VISIBLEAFTERINDENT = 2
%SCI_GETVIEWWS = 2020
%SCI_SETVIEWWS = 2021
%SCI_POSITIONFROMPOINT = 2022
%SCI_POSITIONFROMPOINTCLOSE = 2023
%SCI_GOTOLINE = 2024
%SCI_GOTOPOS = 2025
%SCI_SETANCHOR = 2026
%SCI_GETCURLINE = 2027
%SCI_GETENDSTYLED = 2028
%SC_EOL_CRLF = 0
%SC_EOL_CR = 1
%SC_EOL_LF = 2
%SCI_CONVERTEOLS = 2029
%SCI_GETEOLMODE = 2030
%SCI_SETEOLMODE = 2031
%SCI_STARTSTYLING = 2032
%SCI_SETSTYLING = 2033
%SCI_GETBUFFEREDDRAW = 2034
%SCI_SETBUFFEREDDRAW = 2035
%SCI_SETTABWIDTH = 2036
%SCI_GETTABWIDTH = 2121
%SC_CP_UTF8 = 65001
%SC_CP_DBCS = 1
%SCI_SETCODEPAGE = 2037
%SCI_SETUSEPALETTE = 2039
%MARKER_MAX = 31
%SC_MARK_CIRCLE = 0
%SC_MARK_ROUNDRECT = 1
%SC_MARK_ARROW = 2
%SC_MARK_SMALLRECT = 3
%SC_MARK_SHORTARROW = 4
%SC_MARK_EMPTY = 5
%SC_MARK_ARROWDOWN = 6
%SC_MARK_MINUS = 7
%SC_MARK_PLUS = 8
%SC_MARK_VLINE = 9
%SC_MARK_LCORNER = 10
%SC_MARK_TCORNER = 11
%SC_MARK_BOXPLUS = 12
%SC_MARK_BOXPLUSCONNECTED = 13
%SC_MARK_BOXMINUS = 14
%SC_MARK_BOXMINUSCONNECTED = 15
%SC_MARK_LCORNERCURVE = 16
%SC_MARK_TCORNERCURVE = 17
%SC_MARK_CIRCLEPLUS = 18
%SC_MARK_CIRCLEPLUSCONNECTED = 19
%SC_MARK_CIRCLEMINUS = 20
%SC_MARK_CIRCLEMINUSCONNECTED = 21
%SC_MARK_BACKGROUND = 22
%SC_MARK_DOTDOTDOT = 23
%SC_MARK_ARROWS = 24
%SC_MARK_PIXMAP = 25
%SC_MARK_CHARACTER = 10000
%SC_MARKNUM_FOLDEREND = 25
%SC_MARKNUM_FOLDEROPENMID = 26
%SC_MARKNUM_FOLDERMIDTAIL = 27
%SC_MARKNUM_FOLDERTAIL = 28
%SC_MARKNUM_FOLDERSUB = 29
%SC_MARKNUM_FOLDER = 30
%SC_MARKNUM_FOLDEROPEN = 31
%SC_MASK_FOLDERS = &HFE000000
%SCI_MARKERDEFINE = 2040
%SCI_MARKERSETFORE = 2041
%SCI_MARKERSETBACK = 2042
%SCI_MARKERADD = 2043
%SCI_MARKERDELETE = 2044
%SCI_MARKERDELETEALL = 2045
%SCI_MARKERGET = 2046
%SCI_MARKERNEXT = 2047
%SCI_MARKERPREVIOUS = 2048
%SCI_MARKERDEFINEPIXMAP = 2049
%SC_MARGIN_SYMBOL = 0
%SC_MARGIN_NUMBER = 1
%SCI_SETMARGINTYPEN = 2240
%SCI_GETMARGINTYPEN = 2241
%SCI_SETMARGINWIDTHN = 2242
%SCI_GETMARGINWIDTHN = 2243
%SCI_SETMARGINMASKN = 2244
%SCI_GETMARGINMASKN = 2245
%SCI_SETMARGINSENSITIVEN = 2246
%SCI_GETMARGINSENSITIVEN = 2247
%STYLE_DEFAULT = 32
%STYLE_LINENUMBER = 33
%STYLE_BRACELIGHT = 34
%STYLE_BRACEBAD = 35
%STYLE_CONTROLCHAR = 36
%STYLE_INDENTGUIDE = 37
%STYLE_LASTPREDEFINED = 39
%STYLE_MAX = 127
%SC_CHARSET_ANSI = 0
%SC_CHARSET_DEFAULT = 1
%SC_CHARSET_BALTIC = 186
%SC_CHARSET_CHINESEBIG5 = 136
%SC_CHARSET_EASTEUROPE = 238
%SC_CHARSET_GB2312 = 134
%SC_CHARSET_GREEK = 161
%SC_CHARSET_HANGUL = 129
%SC_CHARSET_MAC = 77
%SC_CHARSET_OEM = 255
%SC_CHARSET_RUSSIAN = 204
%SC_CHARSET_SHIFTJIS = 128
%SC_CHARSET_SYMBOL = 2
%SC_CHARSET_TURKISH = 162
%SC_CHARSET_JOHAB = 130
%SC_CHARSET_HEBREW = 177
%SC_CHARSET_ARABIC = 178
%SC_CHARSET_VIETNAMESE = 163
%SC_CHARSET_THAI = 222
%SCI_STYLECLEARALL = 2050
%SCI_STYLESETFORE = 2051
%SCI_STYLESETBACK = 2052
%SCI_STYLESETBOLD = 2053
%SCI_STYLESETITALIC = 2054
%SCI_STYLESETSIZE = 2055
%SCI_STYLESETFONT = 2056
%SCI_STYLESETEOLFILLED = 2057
%SCI_STYLERESETDEFAULT = 2058
%SCI_STYLESETUNDERLINE = 2059
%SC_CASE_MIXED = 0
%SC_CASE_UPPER = 1
%SC_CASE_LOWER = 2
%SCI_STYLESETCASE = 2060
%SCI_STYLESETCHARACTERSET = 2066
%SCI_STYLESETHOTSPOT = 2409
%SCI_SETSELFORE = 2067
%SCI_SETSELBACK = 2068
%SCI_SETCARETFORE = 2069
%SCI_ASSIGNCMDKEY = 2070
%SCI_CLEARCMDKEY = 2071
%SCI_CLEARALLCMDKEYS = 2072
%SCI_SETSTYLINGEX = 2073
%SCI_STYLESETVISIBLE = 2074
%SCI_GETCARETPERIOD = 2075
%SCI_SETCARETPERIOD = 2076
%SCI_SETWORDCHARS = 2077
%SCI_BEGINUNDOACTION = 2078
%SCI_ENDUNDOACTION = 2079
%INDIC_MAX = 7
%INDIC_PLAIN = 0
%INDIC_SQUIGGLE = 1
%INDIC_TT = 2
%INDIC_DIAGONAL = 3
%INDIC_STRIKE = 4
%INDIC_HIDDEN = 5
%INDIC_BOX = 6
%INDIC0_MASK = &H20
%INDIC1_MASK = &H40
%INDIC2_MASK = &H80
%INDICS_MASK = &HE0
%SCI_INDICSETSTYLE = 2080
%SCI_INDICGETSTYLE = 2081
%SCI_INDICSETFORE = 2082
%SCI_INDICGETFORE = 2083
%SCI_SETWHITESPACEFORE = 2084
%SCI_SETWHITESPACEBACK = 2085
%SCI_SETSTYLEBITS = 2090
%SCI_GETSTYLEBITS = 2091
%SCI_SETLINESTATE = 2092
%SCI_GETLINESTATE = 2093
%SCI_GETMAXLINESTATE = 2094
%SCI_GETCARETLINEVISIBLE = 2095
%SCI_SETCARETLINEVISIBLE = 2096
%SCI_GETCARETLINEBACK = 2097
%SCI_SETCARETLINEBACK = 2098
%SCI_STYLESETCHANGEABLE = 2099
%SCI_AUTOCSHOW = 2100
%SCI_AUTOCCANCEL = 2101
%SCI_AUTOCACTIVE = 2102
%SCI_AUTOCPOSSTART = 2103
%SCI_AUTOCCOMPLETE = 2104
%SCI_AUTOCSTOPS = 2105
%SCI_AUTOCSETSEPARATOR = 2106
%SCI_AUTOCGETSEPARATOR = 2107
%SCI_AUTOCSELECT = 2108
%SCI_AUTOCSETCANCELATSTART = 2110
%SCI_AUTOCGETCANCELATSTART = 2111
%SCI_AUTOCSETFILLUPS = 2112
%SCI_AUTOCSETCHOOSESINGLE = 2113
%SCI_AUTOCGETCHOOSESINGLE = 2114
%SCI_AUTOCSETIGNORECASE = 2115
%SCI_AUTOCGETIGNORECASE = 2116
%SCI_USERLISTSHOW = 2117
%SCI_AUTOCSETAUTOHIDE = 2118
%SCI_AUTOCGETAUTOHIDE = 2119
%SCI_AUTOCSETDROPRESTOFWORD = 2270
%SCI_AUTOCGETDROPRESTOFWORD = 2271
%SCI_REGISTERIMAGE = 2405
%SCI_CLEARREGISTEREDIMAGES = 2408
%SCI_AUTOCGETTYPESEPARATOR = 2285
%SCI_AUTOCSETTYPESEPARATOR = 2286
%SCI_SETINDENT = 2122
%SCI_GETINDENT = 2123
%SCI_SETUSETABS = 2124
%SCI_GETUSETABS = 2125
%SCI_SETLINEINDENTATION = 2126
%SCI_GETLINEINDENTATION = 2127
%SCI_GETLINEINDENTPOSITION = 2128
%SCI_GETCOLUMN = 2129
%SCI_SETHSCROLLBAR = 2130
%SCI_GETHSCROLLBAR = 2131
%SCI_SETINDENTATIONGUIDES = 2132
%SCI_GETINDENTATIONGUIDES = 2133
%SCI_SETHIGHLIGHTGUIDE = 2134
%SCI_GETHIGHLIGHTGUIDE = 2135
%SCI_GETLINEENDPOSITION = 2136
%SCI_GETCODEPAGE = 2137
%SCI_GETCARETFORE = 2138
%SCI_GETUSEPALETTE = 2139
%SCI_GETREADONLY = 2140
%SCI_SETCURRENTPOS = 2141
%SCI_SETSELECTIONSTART = 2142
%SCI_GETSELECTIONSTART = 2143
%SCI_SETSELECTIONEND = 2144
%SCI_GETSELECTIONEND = 2145
%SCI_SETPRINTMAGNIFICATION = 2146
%SCI_GETPRINTMAGNIFICATION = 2147
%SC_PRINT_NORMAL = 0
%SC_PRINT_INVERTLIGHT = 1
%SC_PRINT_BLACKONWHITE = 2
%SC_PRINT_COLOURONWHITE = 3
%SC_PRINT_COLOURONWHITEDEFAULTBG = 4
%SCI_SETPRINTCOLOURMODE = 2148
%SCI_GETPRINTCOLOURMODE = 2149
%SCFIND_WHOLEWORD = 2
%SCFIND_MATCHCASE = 4
%SCFIND_WORDSTART = &H00100000
%SCFIND_REGEXP = &H00200000
%SCFIND_POSIX = &H00400000
%SCI_FINDTEXT = 2150
%SCI_FORMATRANGE = 2151
%SCI_GETFIRSTVISIBLELINE = 2152
%SCI_GETLINE = 2153
%SCI_GETLINECOUNT = 2154
%SCI_SETMARGINLEFT = 2155
%SCI_GETMARGINLEFT = 2156
%SCI_SETMARGINRIGHT = 2157
%SCI_GETMARGINRIGHT = 2158
%SCI_GETMODIFY = 2159
%SCI_SETSEL = 2160
%SCI_GETSELTEXT = 2161
%SCI_GETTEXTRANGE = 2162
%SCI_HIDESELECTION = 2163
%SCI_POINTXFROMPOSITION = 2164
%SCI_POINTYFROMPOSITION = 2165
%SCI_LINEFROMPOSITION = 2166
%SCI_POSITIONFROMLINE = 2167
%SCI_LINESCROLL = 2168
%SCI_SCROLLCARET = 2169
%SCI_REPLACESEL = 2170
%SCI_SETREADONLY = 2171
%SCI_NULL = 2172
%SCI_CANPASTE = 2173
%SCI_CANUNDO = 2174
%SCI_EMPTYUNDOBUFFER = 2175
%SCI_UNDO = 2176
%SCI_CUT = 2177
%SCI_COPY = 2178
%SCI_PASTE = 2179
%SCI_CLEAR = 2180
%SCI_SETTEXT = 2181
%SCI_GETTEXT = 2182
%SCI_GETTEXTLENGTH = 2183
%SCI_GETDIRECTFUNCTION = 2184
%SCI_GETDIRECTPOINTER = 2185
%SCI_SETOVERTYPE = 2186
%SCI_GETOVERTYPE = 2187
%SCI_SETCARETWIDTH = 2188
%SCI_GETCARETWIDTH = 2189
%SCI_SETTARGETSTART = 2190
%SCI_GETTARGETSTART = 2191
%SCI_SETTARGETEND = 2192
%SCI_GETTARGETEND = 2193
%SCI_REPLACETARGET = 2194
%SCI_REPLACETARGETRE = 2195
%SCI_SEARCHINTARGET = 2197
%SCI_SETSEARCHFLAGS = 2198
%SCI_GETSEARCHFLAGS = 2199
%SCI_CALLTIPSHOW = 2200
%SCI_CALLTIPCANCEL = 2201
%SCI_CALLTIPACTIVE = 2202
%SCI_CALLTIPPOSSTART = 2203
%SCI_CALLTIPSETHLT = 2204
%SCI_CALLTIPSETBACK = 2205
%SCI_CALLTIPSETFORE = 2206
%SCI_CALLTIPSETFOREHLT = 2207
%SCI_VISIBLEFROMDOCLINE = 2220
%SCI_DOCLINEFROMVISIBLE = 2221
%SC_FOLDLEVELBASE = &H400
%SC_FOLDLEVELWHITEFLAG = &H1000
%SC_FOLDLEVELHEADERFLAG = &H2000
%SC_FOLDLEVELBOXHEADERFLAG = &H4000
%SC_FOLDLEVELBOXFOOTERFLAG = &H8000
%SC_FOLDLEVELCONTRACTED = &H10000
%SC_FOLDLEVELUNINDENT = &H20000
%SC_FOLDLEVELNUMBERMASK = &H0FFF
%SCI_SETFOLDLEVEL = 2222
%SCI_GETFOLDLEVEL = 2223
%SCI_GETLASTCHILD = 2224
%SCI_GETFOLDPARENT = 2225
%SCI_SHOWLINES = 2226
%SCI_HIDELINES = 2227
%SCI_GETLINEVISIBLE = 2228
%SCI_SETFOLDEXPANDED = 2229
%SCI_GETFOLDEXPANDED = 2230
%SCI_TOGGLEFOLD = 2231
%SCI_ENSUREVISIBLE = 2232
%SC_FOLDFLAG_LINEBEFORE_EXPANDED = &H0002
%SC_FOLDFLAG_LINEBEFORE_CONTRACTED = &H0004
%SC_FOLDFLAG_LINEAFTER_EXPANDED = &H0008
%SC_FOLDFLAG_LINEAFTER_CONTRACTED = &H0010
%SC_FOLDFLAG_LEVELNUMBERS = &H0040
%SC_FOLDFLAG_BOX = &H0001
%SCI_SETFOLDFLAGS = 2233
%SCI_ENSUREVISIBLEENFORCEPOLICY = 2234
%SCI_SETTABINDENTS = 2260
%SCI_GETTABINDENTS = 2261
%SCI_SETBACKSPACEUNINDENTS = 2262
%SCI_GETBACKSPACEUNINDENTS = 2263
%SC_TIME_FOREVER = 10000000
%SCI_SETMOUSEDWELLTIME = 2264
%SCI_GETMOUSEDWELLTIME = 2265
%SCI_WORDSTARTPOSITION = 2266
%SCI_WORDENDPOSITION = 2267
%SC_WRAP_NONE = 0
%SC_WRAP_WORD = 1
%SCI_SETWRAPMODE = 2268
%SCI_GETWRAPMODE = 2269
%SC_CACHE_NONE = 0
%SC_CACHE_CARET = 1
%SC_CACHE_PAGE = 2
%SC_CACHE_DOCUMENT = 3
%SCI_SETLAYOUTCACHE = 2272
%SCI_GETLAYOUTCACHE = 2273
%SCI_SETSCROLLWIDTH = 2274
%SCI_GETSCROLLWIDTH = 2275
%SCI_TEXTWIDTH = 2276
%SCI_SETENDATLASTLINE = 2277
%SCI_GETENDATLASTLINE = 2278
%SCI_TEXTHEIGHT = 2279
%SCI_SETVSCROLLBAR = 2280
%SCI_GETVSCROLLBAR = 2281
%SCI_APPENDTEXT = 2282
%SCI_GETTWOPHASEDRAW = 2283
%SCI_SETTWOPHASEDRAW = 2284
%SCI_TARGETFROMSELECTION = 2287
%SCI_LINESJOIN = 2288
%SCI_LINESSPLIT = 2289
%SCI_SETFOLDMARGINCOLOUR = 2290
%SCI_SETFOLDMARGINHICOLOUR = 2291
%SCI_LINEDOWN = 2300
%SCI_LINEDOWNEXTEND = 2301
%SCI_LINEUP = 2302
%SCI_LINEUPEXTEND = 2303
%SCI_CHARLEFT = 2304
%SCI_CHARLEFTEXTEND = 2305
%SCI_CHARRIGHT = 2306
%SCI_CHARRIGHTEXTEND = 2307
%SCI_WORDLEFT = 2308
%SCI_WORDLEFTEXTEND = 2309
%SCI_WORDRIGHT = 2310
%SCI_WORDRIGHTEXTEND = 2311
%SCI_HOME = 2312
%SCI_HOMEEXTEND = 2313
%SCI_LINEEND = 2314
%SCI_LINEENDEXTEND = 2315
%SCI_DOCUMENTSTART = 2316
%SCI_DOCUMENTSTARTEXTEND = 2317
%SCI_DOCUMENTEND = 2318
%SCI_DOCUMENTENDEXTEND = 2319
%SCI_PAGEUP = 2320
%SCI_PAGEUPEXTEND = 2321
%SCI_PAGEDOWN = 2322
%SCI_PAGEDOWNEXTEND = 2323
%SCI_EDITTOGGLEOVERTYPE = 2324
%SCI_CANCEL = 2325
%SCI_DELETEBACK = 2326
%SCI_TAB = 2327
%SCI_BACKTAB = 2328
%SCI_NEWLINE = 2329
%SCI_FORMFEED = 2330
%SCI_VCHOME = 2331
%SCI_VCHOMEEXTEND = 2332
%SCI_ZOOMIN = 2333
%SCI_ZOOMOUT = 2334
%SCI_DELWORDLEFT = 2335
%SCI_DELWORDRIGHT = 2336
%SCI_LINECUT = 2337
%SCI_LINEDELETE = 2338
%SCI_LINETRANSPOSE = 2339
%SCI_LINEDUPLICATE = 2404
%SCI_LOWERCASE = 2340
%SCI_UPPERCASE = 2341
%SCI_LINESCROLLDOWN = 2342
%SCI_LINESCROLLUP = 2343
%SCI_DELETEBACKNOTLINE = 2344
%SCI_HOMEDISPLAY = 2345
%SCI_HOMEDISPLAYEXTEND = 2346
%SCI_LINEENDDISPLAY = 2347
%SCI_LINEENDDISPLAYEXTEND = 2348
%SCI_HOMEWRAP = 2349
%SCI_HOMEWRAPEXTEND = 2450
%SCI_LINEENDWRAP = 2451
%SCI_LINEENDWRAPEXTEND = 2452
%SCI_VCHOMEWRAP = 2453
%SCI_VCHOMEWRAPEXTEND = 2454
%SCI_LINECOPY = 2455
%SCI_MOVECARETINSIDEVIEW = 2401
%SCI_LINELENGTH = 2350
%SCI_BRACEHIGHLIGHT = 2351
%SCI_BRACEBADLIGHT = 2352
%SCI_BRACEMATCH = 2353
%SCI_GETVIEWEOL = 2355
%SCI_SETVIEWEOL = 2356
%SCI_GETDOCPOINTER = 2357
%SCI_SETDOCPOINTER = 2358
%SCI_SETMODEVENTMASK = 2359
%EDGE_NONE = 0
%EDGE_LINE = 1
%EDGE_BACKGROUND = 2
%SCI_GETEDGECOLUMN = 2360
%SCI_SETEDGECOLUMN = 2361
%SCI_GETEDGEMODE = 2362
%SCI_SETEDGEMODE = 2363
%SCI_GETEDGECOLOUR = 2364
%SCI_SETEDGECOLOUR = 2365
%SCI_SEARCHANCHOR = 2366
%SCI_SEARCHNEXT = 2367
%SCI_SEARCHPREV = 2368
%SCI_LINESONSCREEN = 2370
%SCI_USEPOPUP = 2371
%SCI_SELECTIONISRECTANGLE = 2372
%SCI_SETZOOM = 2373
%SCI_GETZOOM = 2374
%SCI_CREATEDOCUMENT = 2375
%SCI_ADDREFDOCUMENT = 2376
%SCI_RELEASEDOCUMENT = 2377
%SCI_GETMODEVENTMASK = 2378
%SCI_SETFOCUS = 2380
%SCI_GETFOCUS = 2381
%SCI_SETSTATUS = 2382
%SCI_GETSTATUS = 2383
%SCI_SETMOUSEDOWNCAPTURES = 2384
%SCI_GETMOUSEDOWNCAPTURES = 2385
%SC_CURSORNORMAL = - 1
%SC_CURSORWAIT = 4
%SCI_SETCURSOR = 2386
%SCI_GETCURSOR = 2387
%SCI_SETCONTROLCHARSYMBOL = 2388
%SCI_GETCONTROLCHARSYMBOL = 2389
%SCI_WORDPARTLEFT = 2390
%SCI_WORDPARTLEFTEXTEND = 2391
%SCI_WORDPARTRIGHT = 2392
%SCI_WORDPARTRIGHTEXTEND = 2393
%VISIBLE_SLOP = &H01
%VISIBLE_STRICT = &H04
%SCI_SETVISIBLEPOLICY = 2394
%SCI_DELLINELEFT = 2395
%SCI_DELLINERIGHT = 2396
%SCI_SETXOFFSET = 2397
%SCI_GETXOFFSET = 2398
%SCI_CHOOSECARETX = 2399
%SCI_GRABFOCUS = 2400
%CARET_SLOP = &H01
%CARET_STRICT = &H04
%CARET_JUMPS = &H10
%CARET_EVEN = &H08
%SCI_SETXCARETPOLICY = 2402
%SCI_SETYCARETPOLICY = 2403
%SCI_SETPRINTWRAPMODE = 2406
%SCI_GETPRINTWRAPMODE = 2407
%SCI_SETHOTSPOTACTIVEFORE = 2410
%SCI_SETHOTSPOTACTIVEBACK = 2411
%SCI_SETHOTSPOTACTIVEUNDERLINE = 2412
%SCI_SETHOTSPOTSINGLELINE = 2421
%SCI_PARADOWN = 2413
%SCI_PARADOWNEXTEND = 2414
%SCI_PARAUP = 2415
%SCI_PARAUPEXTEND = 2416
%SCI_POSITIONBEFORE = 2417
%SCI_POSITIONAFTER = 2418
%SCI_COPYRANGE = 2419
%SCI_COPYTEXT = 2420
%SC_SEL_STREAM = 0
%SC_SEL_RECTANGLE = 1
%SC_SEL_LINES = 2
%SCI_SETSELECTIONMODE = 2422
%SCI_GETSELECTIONMODE = 2423
%SCI_GETLINESELSTARTPOSITION = 2424
%SCI_GETLINESELENDPOSITION = 2425
%SCI_LINEDOWNRECTEXTEND = 2426
%SCI_LINEUPRECTEXTEND = 2427
%SCI_CHARLEFTRECTEXTEND = 2428
%SCI_CHARRIGHTRECTEXTEND = 2429
%SCI_HOMERECTEXTEND = 2430
%SCI_VCHOMERECTEXTEND = 2431
%SCI_LINEENDRECTEXTEND = 2432
%SCI_PAGEUPRECTEXTEND = 2433
%SCI_PAGEDOWNRECTEXTEND = 2434
%SCI_STUTTEREDPAGEUP = 2435
%SCI_STUTTEREDPAGEUPEXTEND = 2436
%SCI_STUTTEREDPAGEDOWN = 2437
%SCI_STUTTEREDPAGEDOWNEXTEND = 2438
%SCI_WORDLEFTEND = 2439
%SCI_WORDLEFTENDEXTEND = 2440
%SCI_WORDRIGHTEND = 2441
%SCI_WORDRIGHTENDEXTEND = 2442
%SCI_SETWHITESPACECHARS = 2443
%SCI_SETCHARSDEFAULT = 2444
%SCI_STARTRECORD = 3001
%SCI_STOPRECORD = 3002
%SCI_SETLEXER = 4001
%SCI_GETLEXER = 4002
%SCI_COLOURISE = 4003
%SCI_SETPROPERTY = 4004
%KEYWORDSET_MAX = 8
%SCI_SETKEYWORDS = 4005
%SCI_SETLEXERLANGUAGE = 4006
%SCI_LOADLEXERLIBRARY = 4007
%SC_MOD_INSERTTEXT = &H1
%SC_MOD_DELETETEXT = &H2
%SC_MOD_CHANGESTYLE = &H4
%SC_MOD_CHANGEFOLD = &H8
%SC_PERFORMED_USER = &H10
%SC_PERFORMED_UNDO = &H20
%SC_PERFORMED_REDO = &H40
%SC_LASTSTEPINUNDOREDO = &H100
%SC_MOD_CHANGEMARKER = &H200
%SC_MOD_BEFOREINSERT = &H400
%SC_MOD_BEFOREDELETE = &H800
%SC_MODEVENTMASKALL = &HF77
%SCEN_CHANGE = 768
%SCEN_SETFOCUS = 512
%SCEN_KILLFOCUS = 256
%SCK_DOWN = 300
%SCK_UP = 301
%SCK_LEFT = 302
%SCK_RIGHT = 303
%SCK_HOME = 304
%SCK_END = 305
%SCK_PRIOR = 306
%SCK_NEXT = 307
%SCK_DELETE = 308
%SCK_INSERT = 309
%SCK_ESCAPE = 7
%SCK_BACK = 8
%SCK_TAB = 9
%SCK_RETURN = 13
%SCK_ADD = 310
%SCK_SUBTRACT = 311
%SCK_DIVIDE = 312
%SCMOD_SHIFT = 1
%SCMOD_CTRL = 2
%SCMOD_ALT = 4
%SCN_STYLENEEDED = 2000
%SCN_CHARADDED = 2001
%SCN_SAVEPOINTREACHED = 2002
%SCN_SAVEPOINTLEFT = 2003
%SCN_MODIFYATTEMPTRO = 2004
%SCN_KEY = 2005
%SCN_DOUBLECLICK = 2006
%SCN_UPDATEUI = 2007
%SCN_MODIFIED = 2008
%SCN_MACRORECORD = 2009
%SCN_MARGINCLICK = 2010
%SCN_NEEDSHOWN = 2011
%SCN_PAINTED = 2013
%SCN_USERLISTSELECTION = 2014
%SCN_URIDROPPED = 2015
%SCN_DWELLSTART = 2016
%SCN_DWELLEND = 2017
%SCN_ZOOM = 2018
%SCN_HOTSPOTCLICK = 2019
%SCN_HOTSPOTDOUBLECLICK = 2020
%SCN_CALLTIPCLICK = 2021
' // SciLexer features - not in standard Scintilla
%SCLEX_CONTAINER = 0
%SCLEX_NULL = 1
%SCLEX_PYTHON = 2
%SCLEX_CPP = 3
%SCLEX_HTML = 4
%SCLEX_XML = 5
%SCLEX_PERL = 6
%SCLEX_SQL = 7
%SCLEX_VB = 8
%SCLEX_PROPERTIES = 9
%SCLEX_ERRORLIST = 10
%SCLEX_MAKEFILE = 11
%SCLEX_BATCH = 12
%SCLEX_XCODE = 13
%SCLEX_LATEX = 14
%SCLEX_LUA = 15
%SCLEX_DIFF = 16
%SCLEX_CONF = 17
%SCLEX_PASCAL = 18
%SCLEX_AVE = 19
%SCLEX_ADA = 20
%SCLEX_LISP = 21
%SCLEX_RUBY = 22
%SCLEX_EIFFEL = 23
%SCLEX_EIFFELKW = 24
%SCLEX_TCL = 25
%SCLEX_NNCRONTAB = 26
%SCLEX_BULLANT = 27
%SCLEX_VBSCRIPT = 28
%SCLEX_ASP = 29
%SCLEX_PHP = 30
%SCLEX_BAAN = 31
%SCLEX_MATLAB = 32
%SCLEX_SCRIPTOL = 33
%SCLEX_ASM = 34
%SCLEX_CPPNOCASE = 35
%SCLEX_FORTRAN = 36
%SCLEX_F77 = 37
%SCLEX_CSS = 38
%SCLEX_POV = 39
%SCLEX_LOUT = 40
%SCLEX_ESCRIPT = 41
%SCLEX_PS = 42
%SCLEX_NSIS = 43
%SCLEX_MMIXAL = 44
%SCLEX_CLW = 45
%SCLEX_CLWNOCASE = 46
%SCLEX_LOT = 47
%SCLEX_YAML = 48
%SCLEX_TEX = 49
%SCLEX_METAPOST = 50
%SCLEX_POWERBASIC = 51
%SCLEX_FORTH = 52
%SCLEX_ERLANG = 53
%SCLEX_AUTOMATIC = 1000
%SCE_P_DEFAULT = 0
%SCE_P_COMMENTLINE = 1
%SCE_P_NUMBER = 2
%SCE_P_STRING = 3
%SCE_P_CHARACTER = 4
%SCE_P_WORD = 5
%SCE_P_TRIPLE = 6
%SCE_P_TRIPLEDOUBLE = 7
%SCE_P_CLASSNAME = 8
%SCE_P_DEFNAME = 9
%SCE_P_OPERATOR = 10
%SCE_P_IDENTIFIER = 11
%SCE_P_COMMENTBLOCK = 12
%SCE_P_STRINGEOL = 13
%SCE_C_DEFAULT = 0
%SCE_C_COMMENT = 1
%SCE_C_COMMENTLINE = 2
%SCE_C_COMMENTDOC = 3
%SCE_C_NUMBER = 4
%SCE_C_WORD = 5
%SCE_C_STRING = 6
%SCE_C_CHARACTER = 7
%SCE_C_UUID = 8
%SCE_C_PREPROCESSOR = 9
%SCE_C_OPERATOR = 10
%SCE_C_IDENTIFIER = 11
%SCE_C_STRINGEOL = 12
%SCE_C_VERBATIM = 13
%SCE_C_REGEX = 14
%SCE_C_COMMENTLINEDOC = 15
%SCE_C_WORD2 = 16
%SCE_C_COMMENTDOCKEYWORD = 17
%SCE_C_COMMENTDOCKEYWORDERROR = 18
%SCE_C_GLOBALCLASS = 19
%SCE_H_DEFAULT = 0
%SCE_H_TAG = 1
%SCE_H_TAGUNKNOWN = 2
%SCE_H_ATTRIBUTE = 3
%SCE_H_ATTRIBUTEUNKNOWN = 4
%SCE_H_NUMBER = 5
%SCE_H_DOUBLESTRING = 6
%SCE_H_SINGLESTRING = 7
%SCE_H_OTHER = 8
%SCE_H_COMMENT = 9
%SCE_H_ENTITY = 10
%SCE_H_TAGEND = 11
%SCE_H_XMLSTART = 12
%SCE_H_XMLEND = 13
%SCE_H_SCRIPT = 14
%SCE_H_ASP = 15
%SCE_H_ASPAT = 16
%SCE_H_CDATA = 17
%SCE_H_QUESTION = 18
%SCE_H_VALUE = 19
%SCE_H_XCCOMMENT = 20
%SCE_H_SGML_DEFAULT = 21
%SCE_H_SGML_COMMAND = 22
%SCE_H_SGML_1ST_PARAM = 23
%SCE_H_SGML_DOUBLESTRING = 24
%SCE_H_SGML_SIMPLESTRING = 25
%SCE_H_SGML_ERROR = 26
%SCE_H_SGML_SPECIAL = 27
%SCE_H_SGML_ENTITY = 28
%SCE_H_SGML_COMMENT = 29
%SCE_H_SGML_1ST_PARAM_COMMENT = 30
%SCE_H_SGML_BLOCK_DEFAULT = 31
%SCE_HJ_START = 40
%SCE_HJ_DEFAULT = 41
%SCE_HJ_COMMENT = 42
%SCE_HJ_COMMENTLINE = 43
%SCE_HJ_COMMENTDOC = 44
%SCE_HJ_NUMBER = 45
%SCE_HJ_WORD = 46
%SCE_HJ_KEYWORD = 47
%SCE_HJ_DOUBLESTRING = 48
%SCE_HJ_SINGLESTRING = 49
%SCE_HJ_SYMBOLS = 50
%SCE_HJ_STRINGEOL = 51
%SCE_HJ_REGEX = 52
%SCE_HJA_START = 55
%SCE_HJA_DEFAULT = 56
%SCE_HJA_COMMENT = 57
%SCE_HJA_COMMENTLINE = 58
%SCE_HJA_COMMENTDOC = 59
%SCE_HJA_NUMBER = 60
%SCE_HJA_WORD = 61
%SCE_HJA_KEYWORD = 62
%SCE_HJA_DOUBLESTRING = 63
%SCE_HJA_SINGLESTRING = 64
%SCE_HJA_SYMBOLS = 65
%SCE_HJA_STRINGEOL = 66
%SCE_HJA_REGEX = 67
%SCE_HB_START = 70
%SCE_HB_DEFAULT = 71
%SCE_HB_COMMENTLINE = 72
%SCE_HB_NUMBER = 73
%SCE_HB_WORD = 74
%SCE_HB_STRING = 75
%SCE_HB_IDENTIFIER = 76
%SCE_HB_STRINGEOL = 77
%SCE_HBA_START = 80
%SCE_HBA_DEFAULT = 81
%SCE_HBA_COMMENTLINE = 82
%SCE_HBA_NUMBER = 83
%SCE_HBA_WORD = 84
%SCE_HBA_STRING = 85
%SCE_HBA_IDENTIFIER = 86
%SCE_HBA_STRINGEOL = 87
%SCE_HP_START = 90
%SCE_HP_DEFAULT = 91
%SCE_HP_COMMENTLINE = 92
%SCE_HP_NUMBER = 93
%SCE_HP_STRING = 94
%SCE_HP_CHARACTER = 95
%SCE_HP_WORD = 96
%SCE_HP_TRIPLE = 97
%SCE_HP_TRIPLEDOUBLE = 98
%SCE_HP_CLASSNAME = 99
%SCE_HP_DEFNAME = 100
%SCE_HP_OPERATOR = 101
%SCE_HP_IDENTIFIER = 102
%SCE_HPA_START = 105
%SCE_HPA_DEFAULT = 106
%SCE_HPA_COMMENTLINE = 107
%SCE_HPA_NUMBER = 108
%SCE_HPA_STRING = 109
%SCE_HPA_CHARACTER = 110
%SCE_HPA_WORD = 111
%SCE_HPA_TRIPLE = 112
%SCE_HPA_TRIPLEDOUBLE = 113
%SCE_HPA_CLASSNAME = 114
%SCE_HPA_DEFNAME = 115
%SCE_HPA_OPERATOR = 116
%SCE_HPA_IDENTIFIER = 117
%SCE_HPHP_DEFAULT = 118
%SCE_HPHP_HSTRING = 119
%SCE_HPHP_SIMPLESTRING = 120
%SCE_HPHP_WORD = 121
%SCE_HPHP_NUMBER = 122
%SCE_HPHP_VARIABLE = 123
%SCE_HPHP_COMMENT = 124
%SCE_HPHP_COMMENTLINE = 125
%SCE_HPHP_HSTRING_VARIABLE = 126
%SCE_HPHP_OPERATOR = 127
%SCE_PL_DEFAULT = 0
%SCE_PL_ERROR = 1
%SCE_PL_COMMENTLINE = 2
%SCE_PL_POD = 3
%SCE_PL_NUMBER = 4
%SCE_PL_WORD = 5
%SCE_PL_STRING = 6
%SCE_PL_CHARACTER = 7
%SCE_PL_PUNCTUATION = 8
%SCE_PL_PREPROCESSOR = 9
%SCE_PL_OPERATOR = 10
%SCE_PL_IDENTIFIER = 11
%SCE_PL_SCALAR = 12
%SCE_PL_ARRAY = 13
%SCE_PL_HASH = 14
%SCE_PL_SYMBOLTABLE = 15
%SCE_PL_REGEX = 17
%SCE_PL_REGSUBST = 18
%SCE_PL_LONGQUOTE = 19
%SCE_PL_BACKTICKS = 20
%SCE_PL_DATASECTION = 21
%SCE_PL_HERE_DELIM = 22
%SCE_PL_HERE_Q = 23
%SCE_PL_HERE_QQ = 24
%SCE_PL_HERE_QX = 25
%SCE_PL_STRING_Q = 26
%SCE_PL_STRING_QQ = 27
%SCE_PL_STRING_QX = 28
%SCE_PL_STRING_QR = 29
%SCE_PL_STRING_QW = 30
%SCE_PROPS_DEFAULT = 0
%SCE_PROPS_COMMENT = 1
%SCE_PROPS_SECTION = 2
%SCE_PROPS_ASSIGNMENT = 3
%SCE_PROPS_DEFVAL = 4
%SCE_L_DEFAULT = 0
%SCE_L_COMMAND = 1
%SCE_L_TAG = 2
%SCE_L_MATH = 3
%SCE_L_COMMENT = 4
%SCE_LUA_DEFAULT = 0
%SCE_LUA_COMMENT = 1
%SCE_LUA_COMMENTLINE = 2
%SCE_LUA_COMMENTDOC = 3
%SCE_LUA_NUMBER = 4
%SCE_LUA_WORD = 5
%SCE_LUA_STRING = 6
%SCE_LUA_CHARACTER = 7
%SCE_LUA_LITERALSTRING = 8
%SCE_LUA_PREPROCESSOR = 9
%SCE_LUA_OPERATOR = 10
%SCE_LUA_IDENTIFIER = 11
%SCE_LUA_STRINGEOL = 12
%SCE_LUA_WORD2 = 13
%SCE_LUA_WORD3 = 14
%SCE_LUA_WORD4 = 15
%SCE_LUA_WORD5 = 16
%SCE_LUA_WORD6 = 17
%SCE_LUA_WORD7 = 18
%SCE_LUA_WORD8 = 19
%SCE_ERR_DEFAULT = 0
%SCE_ERR_PYTHON = 1
%SCE_ERR_GCC = 2
%SCE_ERR_MS = 3
%SCE_ERR_CMD = 4
%SCE_ERR_BORLAND = 5
%SCE_ERR_PERL = 6
%SCE_ERR_NET = 7
%SCE_ERR_LUA = 8
%SCE_ERR_CTAG = 9
%SCE_ERR_DIFF_CHANGED = 10
%SCE_ERR_DIFF_ADDITION = 11
%SCE_ERR_DIFF_DELETION = 12
%SCE_ERR_DIFF_MESSAGE = 13
%SCE_ERR_PHP = 14
%SCE_ERR_ELF = 15
%SCE_ERR_IFC = 16
%SCE_BAT_DEFAULT = 0
%SCE_BAT_COMMENT = 1
%SCE_BAT_WORD = 2
%SCE_BAT_LABEL = 3
%SCE_BAT_HIDE = 4
%SCE_BAT_COMMAND = 5
%SCE_BAT_IDENTIFIER = 6
%SCE_BAT_OPERATOR = 7
%SCE_MAKE_DEFAULT = 0
%SCE_MAKE_COMMENT = 1
%SCE_MAKE_PREPROCESSOR = 2
%SCE_MAKE_IDENTIFIER = 3
%SCE_MAKE_OPERATOR = 4
%SCE_MAKE_TARGET = 5
%SCE_MAKE_IDEOL = 9
%SCE_DIFF_DEFAULT = 0
%SCE_DIFF_COMMENT = 1
%SCE_DIFF_COMMAND = 2
%SCE_DIFF_HEADER = 3
%SCE_DIFF_POSITION = 4
%SCE_DIFF_DELETED = 5
%SCE_DIFF_ADDED = 6
%SCE_CONF_DEFAULT = 0
%SCE_CONF_COMMENT = 1
%SCE_CONF_NUMBER = 2
%SCE_CONF_IDENTIFIER = 3
%SCE_CONF_EXTENSION = 4
%SCE_CONF_PARAMETER = 5
%SCE_CONF_STRING = 6
%SCE_CONF_OPERATOR = 7
%SCE_CONF_IP = 8
%SCE_CONF_DIRECTIVE = 9
%SCE_AVE_DEFAULT = 0
%SCE_AVE_COMMENT = 1
%SCE_AVE_NUMBER = 2
%SCE_AVE_WORD = 3
%SCE_AVE_STRING = 6
%SCE_AVE_ENUM = 7
%SCE_AVE_STRINGEOL = 8
%SCE_AVE_IDENTIFIER = 9
%SCE_AVE_OPERATOR = 10
%SCE_AVE_WORD1 = 11
%SCE_AVE_WORD2 = 12
%SCE_AVE_WORD3 = 13
%SCE_AVE_WORD4 = 14
%SCE_AVE_WORD5 = 15
%SCE_AVE_WORD6 = 16
%SCE_ADA_DEFAULT = 0
%SCE_ADA_WORD = 1
%SCE_ADA_IDENTIFIER = 2
%SCE_ADA_NUMBER = 3
%SCE_ADA_DELIMITER = 4
%SCE_ADA_CHARACTER = 5
%SCE_ADA_CHARACTEREOL = 6
%SCE_ADA_STRING = 7
%SCE_ADA_STRINGEOL = 8
%SCE_ADA_LABEL = 9
%SCE_ADA_COMMENTLINE = 10
%SCE_ADA_ILLEGAL = 11
%SCE_BAAN_DEFAULT = 0
%SCE_BAAN_COMMENT = 1
%SCE_BAAN_COMMENTDOC = 2
%SCE_BAAN_NUMBER = 3
%SCE_BAAN_WORD = 4
%SCE_BAAN_STRING = 5
%SCE_BAAN_PREPROCESSOR = 6
%SCE_BAAN_OPERATOR = 7
%SCE_BAAN_IDENTIFIER = 8
%SCE_BAAN_STRINGEOL = 9
%SCE_BAAN_WORD2 = 10
%SCE_LISP_DEFAULT = 0
%SCE_LISP_COMMENT = 1
%SCE_LISP_NUMBER = 2
%SCE_LISP_KEYWORD = 3
%SCE_LISP_STRING = 6
%SCE_LISP_STRINGEOL = 8
%SCE_LISP_IDENTIFIER = 9
%SCE_LISP_OPERATOR = 10
%SCE_EIFFEL_DEFAULT = 0
%SCE_EIFFEL_COMMENTLINE = 1
%SCE_EIFFEL_NUMBER = 2
%SCE_EIFFEL_WORD = 3
%SCE_EIFFEL_STRING = 4
%SCE_EIFFEL_CHARACTER = 5
%SCE_EIFFEL_OPERATOR = 6
%SCE_EIFFEL_IDENTIFIER = 7
%SCE_EIFFEL_STRINGEOL = 8
%SCE_NNCRONTAB_DEFAULT = 0
%SCE_NNCRONTAB_COMMENT = 1
%SCE_NNCRONTAB_TASK = 2
%SCE_NNCRONTAB_SECTION = 3
%SCE_NNCRONTAB_KEYWORD = 4
%SCE_NNCRONTAB_MODIFIER = 5
%SCE_NNCRONTAB_ASTERISK = 6
%SCE_NNCRONTAB_NUMBER = 7
%SCE_NNCRONTAB_STRING = 8
%SCE_NNCRONTAB_ENVIRONMENT = 9
%SCE_NNCRONTAB_IDENTIFIER = 10
%SCE_FORTH_DEFAULT = 0
%SCE_FORTH_COMMENT = 1
%SCE_FORTH_COMMENT_ML = 2
%SCE_FORTH_IDENTIFIER = 3
%SCE_FORTH_CONTROL = 4
%SCE_FORTH_KEYWORD = 5
%SCE_FORTH_DEFWORD = 6
%SCE_FORTH_PREWORD1 = 7
%SCE_FORTH_PREWORD2 = 8
%SCE_FORTH_NUMBER = 9
%SCE_FORTH_STRING = 10
%SCE_FORTH_LOCALE = 11
%SCE_MATLAB_DEFAULT = 0
%SCE_MATLAB_COMMENT = 1
%SCE_MATLAB_COMMAND = 2
%SCE_MATLAB_NUMBER = 3
%SCE_MATLAB_KEYWORD = 4
%SCE_MATLAB_STRING = 5
%SCE_MATLAB_OPERATOR = 6
%SCE_MATLAB_IDENTIFIER = 7
%SCE_SCRIPTOL_DEFAULT = 0
%SCE_SCRIPTOL_WHITE = 1
%SCE_SCRIPTOL_COMMENTLINE = 2
%SCE_SCRIPTOL_PERSISTENT = 3
%SCE_SCRIPTOL_CSTYLE = 4
%SCE_SCRIPTOL_COMMENTBLOCK = 5
%SCE_SCRIPTOL_NUMBER = 6
%SCE_SCRIPTOL_STRING = 7
%SCE_SCRIPTOL_CHARACTER = 8
%SCE_SCRIPTOL_STRINGEOL = 9
%SCE_SCRIPTOL_KEYWORD = 10
%SCE_SCRIPTOL_OPERATOR = 11
%SCE_SCRIPTOL_IDENTIFIER = 12
%SCE_SCRIPTOL_TRIPLE = 13
%SCE_SCRIPTOL_CLASSNAME = 14
%SCE_SCRIPTOL_PREPROCESSOR = 15
%SCE_ASM_DEFAULT = 0
%SCE_ASM_COMMENT = 1
%SCE_ASM_NUMBER = 2
%SCE_ASM_STRING = 3
%SCE_ASM_OPERATOR = 4
%SCE_ASM_IDENTIFIER = 5
%SCE_ASM_CPUINSTRUCTION = 6
%SCE_ASM_MATHINSTRUCTION = 7
%SCE_ASM_REGISTER = 8
%SCE_ASM_DIRECTIVE = 9
%SCE_ASM_DIRECTIVEOPERAND = 10
%SCE_ASM_COMMENTBLOCK = 11
%SCE_ASM_CHARACTER = 12
%SCE_ASM_STRINGEOL = 13
%SCE_ASM_EXTINSTRUCTION = 14
%SCE_F_DEFAULT = 0
%SCE_F_COMMENT = 1
%SCE_F_NUMBER = 2
%SCE_F_STRING1 = 3
%SCE_F_STRING2 = 4
%SCE_F_STRINGEOL = 5
%SCE_F_OPERATOR = 6
%SCE_F_IDENTIFIER = 7
%SCE_F_WORD = 8
%SCE_F_WORD2 = 9
%SCE_F_WORD3 = 10
%SCE_F_PREPROCESSOR = 11
%SCE_F_OPERATOR2 = 12
%SCE_F_LABEL = 13
%SCE_F_CONTINUATION = 14
%SCE_CSS_DEFAULT = 0
%SCE_CSS_TAG = 1
%SCE_CSS_CLASS = 2
%SCE_CSS_PSEUDOCLASS = 3
%SCE_CSS_UNKNOWN_PSEUDOCLASS = 4
%SCE_CSS_OPERATOR = 5
%SCE_CSS_IDENTIFIER = 6
%SCE_CSS_UNKNOWN_IDENTIFIER = 7
%SCE_CSS_VALUE = 8
%SCE_CSS_COMMENT = 9
%SCE_CSS_ID = 10
%SCE_CSS_IMPORTANT = 11
%SCE_CSS_DIRECTIVE = 12
%SCE_CSS_DOUBLESTRING = 13
%SCE_CSS_SINGLESTRING = 14
%SCE_POV_DEFAULT = 0
%SCE_POV_COMMENT = 1
%SCE_POV_COMMENTLINE = 2
%SCE_POV_NUMBER = 3
%SCE_POV_OPERATOR = 4
%SCE_POV_IDENTIFIER = 5
%SCE_POV_STRING = 6
%SCE_POV_STRINGEOL = 7
%SCE_POV_DIRECTIVE = 8
%SCE_POV_BADDIRECTIVE = 9
%SCE_POV_WORD2 = 10
%SCE_POV_WORD3 = 11
%SCE_POV_WORD4 = 12
%SCE_POV_WORD5 = 13
%SCE_POV_WORD6 = 14
%SCE_POV_WORD7 = 15
%SCE_POV_WORD8 = 16
%SCE_LOUT_DEFAULT = 0
%SCE_LOUT_COMMENT = 1
%SCE_LOUT_NUMBER = 2
%SCE_LOUT_WORD = 3
%SCE_LOUT_WORD2 = 4
%SCE_LOUT_WORD3 = 5
%SCE_LOUT_WORD4 = 6
%SCE_LOUT_STRING = 7
%SCE_LOUT_OPERATOR = 8
%SCE_LOUT_IDENTIFIER = 9
%SCE_LOUT_STRINGEOL = 10
%SCE_ESCRIPT_DEFAULT = 0
%SCE_ESCRIPT_COMMENT = 1
%SCE_ESCRIPT_COMMENTLINE = 2
%SCE_ESCRIPT_COMMENTDOC = 3
%SCE_ESCRIPT_NUMBER = 4
%SCE_ESCRIPT_WORD = 5
%SCE_ESCRIPT_STRING = 6
%SCE_ESCRIPT_OPERATOR = 7
%SCE_ESCRIPT_IDENTIFIER = 8
%SCE_ESCRIPT_BRACE = 9
%SCE_ESCRIPT_WORD2 = 10
%SCE_ESCRIPT_WORD3 = 11
%SCE_PS_DEFAULT = 0
%SCE_PS_COMMENT = 1
%SCE_PS_DSC_COMMENT = 2
%SCE_PS_DSC_VALUE = 3
%SCE_PS_NUMBER = 4
%SCE_PS_NAME = 5
%SCE_PS_KEYWORD = 6
%SCE_PS_LITERAL = 7
%SCE_PS_IMMEVAL = 8
%SCE_PS_PAREN_ARRAY = 9
%SCE_PS_PAREN_DICT = 10
%SCE_PS_PAREN_PROC = 11
%SCE_PS_TEXT = 12
%SCE_PS_HEXSTRING = 13
%SCE_PS_BASE85STRING = 14
%SCE_PS_BADSTRINGCHAR = 15
%SCE_NSIS_DEFAULT = 0
%SCE_NSIS_COMMENT = 1
%SCE_NSIS_STRINGDQ = 2
%SCE_NSIS_STRINGLQ = 3
%SCE_NSIS_STRINGRQ = 4
%SCE_NSIS_FUNCTION = 5
%SCE_NSIS_VARIABLE = 6
%SCE_NSIS_LABEL = 7
%SCE_NSIS_USERDEFINED = 8
%SCE_NSIS_SECTIONDEF = 9
%SCE_NSIS_SUBSECTIONDEF = 10
%SCE_NSIS_IFDEFINEDEF = 11
%SCE_NSIS_MACRODEF = 12
%SCE_NSIS_STRINGVAR = 13
%SCE_MMIXAL_LEADWS = 0
%SCE_MMIXAL_COMMENT = 1
%SCE_MMIXAL_LABEL = 2
%SCE_MMIXAL_OPCODE = 3
%SCE_MMIXAL_OPCODE_PRE = 4
%SCE_MMIXAL_OPCODE_VALID = 5
%SCE_MMIXAL_OPCODE_UNKNOWN = 6
%SCE_MMIXAL_OPCODE_POST = 7
%SCE_MMIXAL_OPERANDS = 8
%SCE_MMIXAL_NUMBER = 9
%SCE_MMIXAL_REF = 10
%SCE_MMIXAL_CHAR = 11
%SCE_MMIXAL_STRING = 12
%SCE_MMIXAL_REGISTER = 13
%SCE_MMIXAL_HEX = 14
%SCE_MMIXAL_OPERATOR = 15
%SCE_MMIXAL_SYMBOL = 16
%SCE_MMIXAL_INCLUDE = 17
%SCE_CLW_DEFAULT = 0
%SCE_CLW_LABEL = 1
%SCE_CLW_COMMENT = 2
%SCE_CLW_STRING = 3
%SCE_CLW_USER_IDENTIFIER = 4
%SCE_CLW_INTEGER_CONSTANT = 5
%SCE_CLW_REAL_CONSTANT = 6
%SCE_CLW_PICTURE_STRING = 7
%SCE_CLW_KEYWORD = 8
%SCE_CLW_COMPILER_DIRECTIVE = 9
%SCE_CLW_BUILTIN_PROCEDURES_FUNCTION = 10
%SCE_CLW_STRUCTURE_DATA_TYPE = 11
%SCE_CLW_ATTRIBUTE = 12
%SCE_CLW_STANDARD_EQUATE = 13
%SCE_CLW_ERROR = 14
%SCE_LOT_DEFAULT = 0
%SCE_LOT_HEADER = 1
%SCE_LOT_BREAK = 2
%SCE_LOT_SET = 3
%SCE_LOT_PASS = 4
%SCE_LOT_FAIL = 5
%SCE_LOT_ABORT = 6
%SCE_YAML_DEFAULT = 0
%SCE_YAML_COMMENT = 1
%SCE_YAML_IDENTIFIER = 2
%SCE_YAML_KEYWORD = 3
%SCE_YAML_NUMBER = 4
%SCE_YAML_REFERENCE = 5
%SCE_YAML_DOCUMENT = 6
%SCE_YAML_TEXT = 7
%SCE_YAML_ERROR = 8
%SCE_TEX_DEFAULT = 0
%SCE_TEX_SPECIAL = 1
%SCE_TEX_GROUP = 2
%SCE_TEX_SYMBOL = 3
%SCE_TEX_COMMAND = 4
%SCE_TEX_TEXT = 5
%SCE_METAPOST_DEFAULT = 0
%SCE_METAPOST_SPECIAL = 1
%SCE_METAPOST_GROUP = 2
%SCE_METAPOST_SYMBOL = 3
%SCE_METAPOST_COMMAND = 4
%SCE_METAPOST_TEXT = 5
%SCE_METAPOST_EXTRA = 6
%SCE_ERLANG_DEFAULT = 0
%SCE_ERLANG_COMMENT = 1
%SCE_ERLANG_VARIABLE = 2
%SCE_ERLANG_NUMBER = 3
%SCE_ERLANG_KEYWORD = 4
%SCE_ERLANG_STRING = 5
%SCE_ERLANG_OPERATOR = 6
%SCE_ERLANG_ATOM = 7
%SCE_ERLANG_FUNCTION_NAME = 8
%SCE_ERLANG_CHARACTER = 9
%SCE_ERLANG_MACRO = 10
%SCE_ERLANG_RECORD = 11
%SCE_ERLANG_SEPARATOR = 12
%SCE_ERLANG_NODE_NAME = 13
%SCE_ERLANG_UNKNOWN = 31
GLOBAL hMenu AS LONG
GLOBAL SciColorsAndFonts AS SciColorsAndFontsType
' *********************************************************************************************
' Declaration of the only function exported by SciLexer.dll.
' All the comunication with the control is done sending messages.
' Note: Since no function is provided to register the classes, the only way is to load the
' library with LoadLibrary (the attach process automatically register the classes) and free it
' with FreeLibrary (the detach process unregisters the classes).
' *********************************************************************************************
DECLARE FUNCTION SciMsg LIB "SCILEXER.DLL" ALIAS "Scintilla_DirectFunction" ( BYVAL pSciWndData AS DWORD, BYVAL wMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LONG ) AS LONG
' *********************************************************************************************
DECLARE SUB InitializePbKeywords( )
DECLARE SUB Scintilla_SetOptions( BYVAL pSci AS DWORD, BYVAL strFileName AS STRING )
DECLARE FUNCTION GetCurrentLine( ) AS LONG
DECLARE FUNCTION ToggleFolding( BYVAL LineNumber AS LONG ) AS LONG
DECLARE SUB ToggleAllFoldersBelow( BYVAL LineNumber AS LONG )
DECLARE SUB FoldAllProcedures( )
DECLARE SUB ExpandAllProcedures( )
DECLARE SUB BlockComment( )
DECLARE SUB BlockUncomment( )
DECLARE SUB SED_CodeFinder( )
DECLARE SUB ResetCodeFinder( )
DECLARE FUNCTION GetEdit( ) AS DWORD
DECLARE FUNCTION Sci_OnNotify( BYVAL hWnd AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS DWORD ) AS LONG
DECLARE FUNCTION AutoComplete( ) AS LONG
DECLARE FUNCTION RemoveHtml( BYVAL Buf AS STRING ) AS STRING
DECLARE SUB CvBasToHtml
DECLARE SUB SED_SetCharset( BYVAL pSci AS DWORD, BYVAL STYLE AS LONG, BYVAL CharsetName AS STRING )
DECLARE FUNCTION Sci_GetCodeTip( BYVAL strCodeTip AS STRING ) AS STRING
DECLARE SUB ShowCodetip( )
DECLARE FUNCTION SED_PrintDoc( szFilePath AS ASCIIZ, BYVAL PreviewFlag AS LONG ) AS LONG
DECLARE FUNCTION SpacesToTabs( BYVAL InString AS STRING ) AS STRING
DECLARE SUB SED_FormatRegion( )
DECLARE SUB SED_TabulateRegion( )
' *********************************************************************************************
' *********************************************************************************************
' PB Keywords
' Note: Added "|" as a dummy character so the first keyword can begin with an space.
' *********************************************************************************************
SUB InitializePbKeywords
  strPBKeyWords = "| def bloat #bloat $bloat compile #compile $compile debug #debug $debug " & _
          "dim #dim $dim if #if $if else #else $else elseif #elseif $elseif endif #endif $endif " & _
          "include #include $include option #option $option pbforms #pbforms $pbforms " & _
          "register #register $register resource #resource $resource stack #stack $stack tools #tools $tools " & _
          "$bel $bs $cr $crlf $dq $eof $esc $ff $lf $nul $spc $stac $tab $vt " & _
          "black blue green cyan red magenta yellow white gray ltgray " & _
          "vt_empty vt_null vt_i2 vt_i4 vt_r4 vt_r8 vt_cy vt_date vt_bstr vt_dispatch vt_error " & _
          "vt_bool vt_variant vt_unknown vt_i1 vt_ui1 vt_ui2 vt_ui4 vt_i8 vt_ui8 vt_int vt_uint " & _
          "vt_void vt_result vt_ptr vt_safearray vt_carray vt_userdefined vt_lpstr vt_lpwstr " & _
          "vt_filetime vt_blob vt_stream vt_storage vt_streamed_object vt_stored_object vt_cf vt_clsid vt_vector vt_array vt_byref " & _
          "abs accel accept access acode acode$ add addr alias all and any append array " & _
          "arrayattr as asc ascend asciz asciiz asm at atn attach attrib bar base baud bdecl beep bin bin$ " & _
          "binary bit bits break button bycmd bycopy byref byte byval calc call callback callstk callstk$ " & _
          "callstkcount case catch cbctl cbctlmsg cbhndl cblparam cbmsg cbwparam cbyt ccur ccux " & _
          "cd cdbl cdecl cdwd ceil cext chdir chdrive check check3state checkbox choose chr chr$ cint " & _
          "client clng close cls clsid clsid$ codeptr collate color combobox comm command command$ con connect const const$ " & _
          "control cos cqud create cset cset$ csng ctsflow cur curdir curdir$ currency currencyx cux cvbyt cvcur " & _
          "cvcux cvd cvdwd cve cvi cvl cvq cvs cvwrd cwrd data datacount date date$ declare decr default " & _
          "defbyt decur defcux defdbl defdwd defext defint deflng defqud defsng defstr defwrd delete " & _
          "descend dialog dir dir$ disable diskfree disksize dispatch dll dllmain descend dllmain " & _
          "do doevents double down draw dsrflow dsrsens strflow dtrline dword empty enable end environ environ$ " & _
          "eof eqv erase err errapi errclear error error$ exe exit exp exp10 exp2 explicit export ext " & _
          "extended extract extract$ fileattr filecopy filename filename$ filescan fill finally fix flow flush focus " & _
          "font for format format$ formfeed frac frame freefile from function funcname funcname$ get get$ getattr global " & _
          "gosub goto guid guid$ guidtxt guidtxt$ handle hex hex$ hibyt hiint hiwrd host icase icon idn iface iif iif$ " & _
          "image imagex imgbutton imgbuttonx imp in incr inp inout input input# inputbox inputbox$ insert instr int " & _
          "interface integer inv isfalse isnothing isobject istrue iterate join join$ kill label lbound " & _
          "lcase lcase$ left left$ len let lib libmain line listbox lobyt loc local lock lof log log10 log2 loint " & _
          "long loop lowrd lprint lset lset$ ltrim ltrim$ macro macrotemp main makdwd makint maklng makptr " & _
          "makwrd mat max mcase mcase$ member menu mid mid$ min mkbyt mkbyt$ mkcur mkcur$ mkcux mkcux$ " & _
          "mkd mkd$ mkdir mkdwd mkdwd$ mke mke$ mki mki$ mkl mkl$ mkq mkq$ mks mks$ mkwrd mkwrd$ mod modal " & _
          "modeless mouseptr msgbox name new next none not nothing notify objactive object objptr objresult " & _
          "oct oct$ of off on open opt option optional or out output page parity paritychar parityrepl paritytype " & _
          "parse parse$ parsecount pbd pbmain peek peek$ pixels pointer poke poke$ popup port post preserve " & _
          "print print# private profile progid progid$ ptr put put$ quad qword random randomize read read$ " & _
          "receiver records recv redim redraw regexpr regrepl rem remain remain$ remove remove$ repeat repeat$ " & _
          "replace reset resume ret16 ret32 ret87 retain$ retp16 retp32 retprn return rgb right right$ ring " & _
          "rlsd rmdir rnd rotate round rset rtrim rtrim$ rtsflow rxbuffer rxque scan scrollbar sdecl seek " & _
          "select send server set setattr seteof sgn shared shell shift show signed sin single size sizeof " & _
          "sleep sort soace space$ spc sqr state static status stdcall step stop str str$ strdelete strdelete$ " & _
          "string string$ strinsert strinsert$ strptr strreverse strreverse$ sub suspend swap switch " & _
          "tab tab$ tagarray tally tan tcp text textbox then thread threadcount threadid time time$ timeout " & _
          "timer to toggle trace trim trim$ trn try txbuffer txque type ubound ucase ucase$ ucode ucode$ udp union " & _
          "units unlock until up user using using$ val variant variant# variant$ variantvt varptr verify version3 " & _
          "version4 version5 wend while width winmain width width# with word write write# xor xinpflow xoutflow zer " & _
          "dispparams iunknown idispatch ! " & _
          "arc assign bitmap border box clear dc desktop ellipse erl field field# graphic hi key key$ lo load mak margin mix nul$ " & _
          "paint pie pixel pos ppi printercount printer$ render save sbit scale stretch style threaded xprint xprint$ window " & _
          "cancel orientation quality circle bgr device " & _
          "mix_blackness mix_copysrc mix_nop mix_not mix_masknotsrc mix_masksrc mix_masksrcnot mix_mergenotsrc " & _
          "mix_mergesrc mix_mergesrcnot mix_notcopysrc mix_notmasksrc mix_notmergesrc mix_whiteness mix_notxorsrc mix_xorsrc " & _
          "vt_record vt_decimal vt_numeric compiler #compiler pbwin pbcc "
  ' PBCC Keyword List
  strPBKeyWords = strPBKeywords & _
          "conshndl consin consname$ console consout copy cursor cursorx cursory getstderr getstdin getstdkbd " & _
          "getstdout getstdvid inshift instat locate lprint$ mouse mousebuttons mousestat mousex mousey move " & _
          "pageactive pagevisible pcopy screen screenx screeny scroll stdeof stderr stdin stdout waitkey$ waitstat " & _
          "#console $console dllmain& libmain& pblibmain view virtual "
  '= Assembler KeyWord List.
  strPBKeyWords = strPBKeyWords & _
          "asm aaa aad aam aas adc add and bound bsf bsr bswap bt btc btr bts " & _
          "call cbw ccwd cdq clc cld cli cmc cmp cmpsb cmpsd cmpsw cmpxchg cpuid cwde " & _
          "dda das dec div emms f2xm1 fabs fadd faddp fchs fclex fcom fcomp fcompp fcos " & _
          "fdecstp fdiv fdivp fddivr fdivrp ffree fiadd ficom ficomp fidiv fidivr fild " & _
          "fumul fincstp finit fist fistp fisub fisubr fld fld1 fldcw fldenv fldl2e fldl2t " & _
          "fldlg2 fldn2 fldpi fldz fmul fmulp fnclex fninit fnldcw fnop fnsave fnstcw " & _
          "fnstenv fnstsw fpatan fprem fprem1 fptan frndint frstor fsave fscale fsin fsincos " & _
          "fsqrt fst fstcw fstenv fstp fstsw fsub fsubp fsubr fsubrp ftst fucom fucomp " & _
          "fucompp fwait fxam fxch fxtract fyl2x fyl2xp1 hlt idiv imul in inc insb insd insw " & _
          "int int0 iret iretd ja jae jb jbe jc je jecxz jg jge jl jle jmp jna jnae jnb jnbe " & _
          "jnc jne jng jnge jnl jnle jn0 jnp jns jnz j0 jp jpe jp jp0 js jz lahf lar lds lea " & _
          "leave les lfs lgs lock lodsb lodsd lodsw loop loope loopne loopnz loopz lsl lss " & _
          "mov movd movq movsb movsd movsw movsx movzx mul neg nop not or out outsb outsd " & _
          "outsw packssdw packsswb packuswb paddb paddd paddsb paddsw paddusb paddusw paddw " & _
          "pand pandn pcmpeqb pcmpeqd pcmpeqw pcmpgtb pcmpgtd pcmpgtw pmaddwd pmulhw pmullw " & _
          "pop popa popad popf popfd por pslld psllq psllw psrad psraw psrld psrlq psrlw psubb " & _
          "psubd psubsb psubsw psubusb psubusw psubw punpckhbw punpckhdq punpckhwd punpcklbw " & _
          "punpckldq punpcklwd push pusha pushad pushf pushfd pzor rcl rcr rep repe repne repnz " & _
          "repz ret retf retn rol ror sahf sal sar sbb scasb scasd scasw seta setae setb setbe " & _
          "setc sete setg setge setl setle setna setnae setnb setnbe setnc setne setng setnge " & _
          "setnl setnle setn0 setnp setns setnz set0 setp setpe setp0 sets setz shl shld shr " & _
          "shrd stc std sti stosb stosd stosw sub test verr verw wait xchg xlat xlatb xor " & _
          "byte db dd dw dwd dword far near pointer ptr qwd qword short tby tbyte word wrd " & _
          "eax ax ah al ebx bx bh bl ecx cx ch cl edx dx dh dl esi si edi di esp sp ebp bp " & _
          "cs ds es ss fs gs mm mm0 mm1 mm2 mm3 mm4 mm5 mm6 mm7 st "
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Set the charsets
' *********************************************************************************************
SUB SED_SetCharset( BYVAL pSci AS DWORD, BYVAL STYLE AS LONG, BYVAL CharsetName AS STRING )
  SELECT CASE CharsetName
    CASE "Default"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_DEFAULT
    CASE "Ansi"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_ANSI
    CASE "Arabic"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_ARABIC
    CASE "Baltic"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_BALTIC
    CASE "Chinese Big 5"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_CHINESEBIG5
    CASE "East Europe"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_EASTEUROPE
    CASE "GB 2312"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_GB2312
    CASE "Greek"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_GREEK
    CASE "Hangul"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_HANGUL
    CASE "Hebrew"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_HEBREW
    CASE "Johab"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_JOHAB
    CASE "Mac"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_MAC
    CASE "OEM"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_OEM
    CASE "Russian"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_RUSSIAN
    CASE "Shiftjis"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_SHIFTJIS
    CASE "Symbol"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_SYMBOL
    CASE "Thai"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_THAI
    CASE "Turkish"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_TURKISH
    CASE "Vietnamese"
      SciMsg pSci, %SCI_STYLESETCHARACTERSET, STYLE, %SC_CHARSET_VIETNAMESE
  END SELECT
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Set Scintilla Edit Control's options
' *********************************************************************************************
SUB Scintilla_SetOptions( BYVAL pSci AS DWORD, BYVAL strFileName AS STRING )
  LOCAL EdOpt AS EditorOptionsType
  LOCAL FoldOpt AS FoldingOptionsType
  LOCAL szFont AS STRING 'ASCIIZ * 255
  LOCAL szKey AS ASCIIZ * 255
  LOCAL szValue AS ASCIIZ * 255
  LOCAL Magnification AS LONG
  LOCAL COLORMODE AS LONG
  LOCAL WrapMode AS LONG
  IF ISFALSE pSci THEN EXIT SUB
  ' Read the options
  GetEditorOptions EdOpt
  GetFoldingOptions FoldOpt
  ' Set the default style
  szFont = SciColorsAndFonts.DefaultFontName
  SciMsg pSci, %SCI_STYLESETFONT, %STYLE_DEFAULT, VARPTR( szFont )
  SciMsg pSci, %SCI_STYLESETCASE, %STYLE_DEFAULT, EdOpt.DefaultCase
  SciMsg pSci, %SCI_STYLESETSIZE, %STYLE_DEFAULT, SciColorsAndFonts.DefaultFontSize
  SED_SetCharset pSci, %STYLE_DEFAULT, SciColorsAndFonts.DefaultFontCharset
  SciMsg pSci, %SCI_STYLESETBOLD, %STYLE_DEFAULT, SciColorsAndFonts.DefaultFontBold
  SciMsg pSci, %SCI_STYLESETITALIC, %STYLE_DEFAULT, SciColorsAndFonts.DefaultFontItalic
  SciMsg pSci, %SCI_STYLESETUNDERLINE, %STYLE_DEFAULT, SciColorsAndFonts.DefaultFontUnderline
  SciMsg pSci, %SCI_STYLESETFORE, %STYLE_DEFAULT, SciColorsAndFonts.DefaultForeColor
  SciMsg pSci, %SCI_STYLESETBACK, %STYLE_DEFAULT, SciColorsAndFonts.DefaultBackColor
  ' Set all the other styles to the default
  SciMsg pSci, %SCI_STYLECLEARALL, 0, 0
  ' Print options
  GetScintillaPrintOptions( Magnification, COLORMODE, WrapMode )
  SciMsg pSci, %SCI_SETPRINTMAGNIFICATION, Magnification, 0
  SciMsg pSci, %SCI_SETPRINTCOLOURMODE, COLORMODE, 0
  SciMsg pSci, %SCI_SETPRINTWRAPMODE, WrapMode, 0
  ' Set the style for the line numbers
  IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
    szFont = SciColorsAndFonts.DefaultFontName
  ELSE
    szFont = SciColorsAndFonts.LineNumberFontName
  END IF
  SciMsg pSci, %SCI_STYLESETFONT, %STYLE_LINENUMBER, VARPTR( szFont )
  SciMsg pSci, %SCI_STYLESETCASE, %STYLE_LINENUMBER, EdOpt.DefaultCase
  IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
    SciMsg pSci, %SCI_STYLESETSIZE, %STYLE_LINENUMBER, SciColorsAndFonts.DefaultFontSize
  ELSE
    SciMsg pSci, %SCI_STYLESETSIZE, %STYLE_LINENUMBER, SciColorsAndFonts.LineNumberFontSize
  END IF
  SED_SetCharset pSci, %STYLE_LINENUMBER, SciColorsAndFonts.LineNumberFontCharset
  SciMsg pSci, %SCI_STYLESETBOLD, %STYLE_LINENUMBER, SciColorsAndFonts.LineNumberFontBold
  SciMsg pSci, %SCI_STYLESETITALIC, %STYLE_LINENUMBER, SciColorsAndFonts.LineNumberFontItalic
  SciMsg pSci, %SCI_STYLESETUNDERLINE, %STYLE_LINENUMBER, SciColorsAndFonts.LineNumberFontUnderline
  SciMsg pSci, %SCI_STYLESETFORE, %STYLE_LINENUMBER, SciColorsAndFonts.LineNumberForeColor
  SciMsg pSci, %SCI_STYLESETBACK, %STYLE_LINENUMBER, SciColorsAndFonts.LineNumberBackColor
  ' == PowerBasic specific ================================================================
  strFileName = UCASE$( strFileName )
  IF ISTRUE INSTR( strFileName, ".BAS" ) OR _
          ISTRUE INSTR( strFileName, ".INC" ) OR _
          ISTRUE INSTR( strFilename, ".BI" ) OR ISTRUE INSTR( strFileName, ".TXT" ) THEN
    IF EdOpt.SyntaxHighlighting = %BST_CHECKED THEN
      ' Set the Comments style
      IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
        szFont = SciColorsAndFonts.DefaultFontName
      ELSE
        szFont = SciColorsAndFonts.CommentFontName
      END IF
      SciMsg pSci, %SCI_STYLESETFONT, %SCE_B_COMMENT, VARPTR( szFont )
      SciMsg pSci, %SCI_STYLESETCASE, %SCE_B_COMMENT, EdOpt.DefaultCase
      IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_COMMENT, SciColorsAndFonts.DefaultFontSize
      ELSE
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_COMMENT, SciColorsAndFonts.CommentFontSize
      END IF
      SED_SetCharset pSci, %SCE_B_COMMENT, SciColorsAndFonts.CommentFontCharset
      SciMsg pSci, %SCI_STYLESETBOLD, %SCE_B_COMMENT, SciColorsAndFonts.CommentFontBold
      SciMsg pSci, %SCI_STYLESETITALIC, %SCE_B_COMMENT, SciColorsAndFonts.CommentFontItalic
      SciMsg pSci, %SCI_STYLESETUNDERLINE, %SCE_B_COMMENT, SciColorsAndFonts.CommentFontUnderline
      SciMsg pSci, %SCI_STYLESETFORE, %SCE_B_COMMENT, SciColorsAndFonts.CommentForeColor
      IF SciColorsAndFonts.UseAlwaysDefaultBackColor THEN
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_COMMENT, SciColorsAndFonts.DefaultBackColor
      ELSE
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_COMMENT, SciColorsAndFonts.CommentBackColor
      END IF
      ' Set the Keywords style
      IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
        szFont = SciColorsAndFonts.DefaultFontName
      ELSE
        szFont = SciColorsAndFonts.KeywordFontName
      END IF
      SciMsg pSci, %SCI_STYLESETFONT, %SCE_B_KEYWORD, VARPTR( szFont )
      SciMsg pSci, %SCI_STYLESETCASE, %SCE_B_KEYWORD, EdOpt.KeywordCase
      IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_KEYWORD, SciColorsAndFonts.DefaultFontSize
      ELSE
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_KEYWORD, SciColorsAndFonts.KeywordFontSize
      END IF
      SED_SetCharset pSci, %SCE_B_KEYWORD, SciColorsAndFonts.KeywordFontCharset
      SciMsg pSci, %SCI_STYLESETBOLD, %SCE_B_KEYWORD, SciColorsAndFonts.KeywordFontBold
      SciMsg pSci, %SCI_STYLESETITALIC, %SCE_B_KEYWORD, SciColorsAndFonts.KeywordFontItalic
      SciMsg pSci, %SCI_STYLESETUNDERLINE, %SCE_B_KEYWORD, SciColorsAndFonts.KeywordFontUnderline
      SciMsg pSci, %SCI_STYLESETFORE, %SCE_B_KEYWORD, SciColorsAndFonts.KeywordForeColor
      IF SciColorsAndFonts.UseAlwaysDefaultBackColor THEN
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_KEYWORD, SciColorsAndFonts.DefaultBackColor
      ELSE
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_KEYWORD, SciColorsAndFonts.KeywordBackColor
      END IF
      ' Set the Constants style
      IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
        szFont = SciColorsAndFonts.DefaultFontName
      ELSE
        szFont = SciColorsAndFonts.ConstantFontName
      END IF
      SciMsg pSci, %SCI_STYLESETFONT, %SCE_B_CONSTANT, VARPTR( szFont )
      SciMsg pSci, %SCI_STYLESETCASE, %SCE_B_CONSTANT, EdOpt.DefaultCase
      IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_CONSTANT, SciColorsAndFonts.DefaultFontSize
      ELSE
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_CONSTANT, SciColorsAndFonts.ConstantFontSize
      END IF
      SED_SetCharset pSci, %SCE_B_CONSTANT, SciColorsAndFonts.ConstantFontCharset
      SciMsg pSci, %SCI_STYLESETBOLD, %SCE_B_CONSTANT, SciColorsAndFonts.ConstantFontBold
      SciMsg pSci, %SCI_STYLESETITALIC, %SCE_B_CONSTANT, SciColorsAndFonts.ConstantFontItalic
      SciMsg pSci, %SCI_STYLESETUNDERLINE, %SCE_B_CONSTANT, SciColorsAndFonts.ConstantFontUnderline
      SciMsg pSci, %SCI_STYLESETFORE, %SCE_B_CONSTANT, SciColorsAndFonts.ConstantForeColor
      IF SciColorsAndFonts.UseAlwaysDefaultBackColor THEN
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_CONSTANT, SciColorsAndFonts.DefaultBackColor
      ELSE
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_CONSTANT, SciColorsAndFonts.ConstantBackColor
      END IF
      ' Set the Identifiers style
      IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
        szFont = SciColorsAndFonts.DefaultFontName
      ELSE
        szFont = SciColorsAndFonts.IdentifierFontName
      END IF
      SciMsg pSci, %SCI_STYLESETFONT, %SCE_B_IDENTIFIER, VARPTR( szFont )
      SciMsg pSci, %SCI_STYLESETCASE, %SCE_B_IDENTIFIER, EdOpt.DefaultCase
      IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_IDENTIFIER, SciColorsAndFonts.DefaultFontSize
      ELSE
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_IDENTIFIER, SciColorsAndFonts.IdentifierFontSize
      END IF
      SED_SetCharset pSci, %SCE_B_IDENTIFIER, SciColorsAndFonts.IdentifierFontCharset
      SciMsg pSci, %SCI_STYLESETBOLD, %SCE_B_IDENTIFIER, SciColorsAndFonts.IdentifierFontBold
      SciMsg pSci, %SCI_STYLESETITALIC, %SCE_B_IDENTIFIER, SciColorsAndFonts.IdentifierFontItalic
      SciMsg pSci, %SCI_STYLESETUNDERLINE, %SCE_B_IDENTIFIER, SciColorsAndFonts.IdentifierFontUnderline
      SciMsg pSci, %SCI_STYLESETFORE, %SCE_B_IDENTIFIER, SciColorsAndFonts.IdentifierForeColor
      IF SciColorsAndFonts.UseAlwaysDefaultBackColor THEN
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_IDENTIFIER, SciColorsAndFonts.DefaultBackColor
      ELSE
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_IDENTIFIER, SciColorsAndFonts.IdentifierBackColor
      END IF
      ' Set the Numbers style
      IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
        szFont = SciColorsAndFonts.DefaultFontName
      ELSE
        szFont = SciColorsAndFonts.NumberFontName
      END IF
      SciMsg pSci, %SCI_STYLESETFONT, %SCE_B_NUMBER, VARPTR( szFont )
      SciMsg pSci, %SCI_STYLESETCASE, %SCE_B_NUMBER, EdOpt.DefaultCase
      IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_NUMBER, SciColorsAndFonts.DefaultFontSize
      ELSE
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_NUMBER, SciColorsAndFonts.NumberFontSize
      END IF
      SED_SetCharset pSci, %SCE_B_NUMBER, SciColorsAndFonts.NumberFontCharset
      SciMsg pSci, %SCI_STYLESETBOLD, %SCE_B_NUMBER, SciColorsAndFonts.NumberFontBold
      SciMsg pSci, %SCI_STYLESETITALIC, %SCE_B_NUMBER, SciColorsAndFonts.NumberFontItalic
      SciMsg pSci, %SCI_STYLESETUNDERLINE, %SCE_B_NUMBER, SciColorsAndFonts.NumberFontUnderline
      SciMsg pSci, %SCI_STYLESETFORE, %SCE_B_NUMBER, SciColorsAndFonts.NumberForeColor
      IF SciColorsAndFonts.UseAlwaysDefaultBackColor THEN
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_NUMBER, SciColorsAndFonts.DefaultBackColor
      ELSE
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_NUMBER, SciColorsAndFonts.NumberBackColor
      END IF
      ' Set the Operators style
      IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
        szFont = SciColorsAndFonts.DefaultFontName
      ELSE
        szFont = SciColorsAndFonts.OperatorFontName
      END IF
      SciMsg pSci, %SCI_STYLESETFONT, %SCE_B_OPERATOR, VARPTR( szFont )
      SciMsg pSci, %SCI_STYLESETCASE, %SCE_B_OPERATOR, EdOpt.DefaultCase
      IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_OPERATOR, SciColorsAndFonts.DefaultFontSize
      ELSE
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_OPERATOR, SciColorsAndFonts.OperatorFontSize
      END IF
      SED_SetCharset pSci, %SCE_B_OPERATOR, SciColorsAndFonts.OperatorFontCharset
      SciMsg pSci, %SCI_STYLESETBOLD, %SCE_B_OPERATOR, SciColorsAndFonts.OperatorFontBold
      SciMsg pSci, %SCI_STYLESETITALIC, %SCE_B_OPERATOR, SciColorsAndFonts.OperatorFontItalic
      SciMsg pSci, %SCI_STYLESETUNDERLINE, %SCE_B_OPERATOR, SciColorsAndFonts.OperatorFontUnderline
      SciMsg pSci, %SCI_STYLESETFORE, %SCE_B_OPERATOR, SciColorsAndFonts.OperatorForeColor
      IF SciColorsAndFonts.UseAlwaysDefaultBackColor THEN
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_OPERATOR, SciColorsAndFonts.DefaultBackColor
      ELSE
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_OPERATOR, SciColorsAndFonts.OperatorBackColor
      END IF
      ' Set the Preprocessor style
      IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
        szFont = SciColorsAndFonts.DefaultFontName
      ELSE
        szFont = SciColorsAndFonts.PreprocessorFontName
      END IF
      SciMsg pSci, %SCI_STYLESETFONT, %SCE_B_PREPROCESSOR, VARPTR( szFont )
      SciMsg pSci, %SCI_STYLESETCASE, %SCE_B_PREPROCESSOR, EdOpt.DefaultCase
      IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_PREPROCESSOR, SciColorsAndFonts.DefaultFontSize
      ELSE
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_PREPROCESSOR, SciColorsAndFonts.NumberFontSize
      END IF
      SED_SetCharset pSci, %SCE_B_PREPROCESSOR, SciColorsAndFonts.NumberFontCharset
      SciMsg pSci, %SCI_STYLESETBOLD, %SCE_B_PREPROCESSOR, SciColorsAndFonts.NumberFontBold
      SciMsg pSci, %SCI_STYLESETITALIC, %SCE_B_PREPROCESSOR, SciColorsAndFonts.NumberFontItalic
      SciMsg pSci, %SCI_STYLESETUNDERLINE, %SCE_B_PREPROCESSOR, SciColorsAndFonts.NumberFontUnderline
      SciMsg pSci, %SCI_STYLESETFORE, %SCE_B_PREPROCESSOR, SciColorsAndFonts.NumberForeColor
      IF SciColorsAndFonts.UseAlwaysDefaultBackColor THEN
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_PREPROCESSOR, SciColorsAndFonts.DefaultBackColor
      ELSE
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_PREPROCESSOR, SciColorsAndFonts.NumberBackColor
      END IF
      ' Set the Strings style
      IF SciColorsAndFonts.UseAlwaysDefaultFont THEN
        szFont = SciColorsAndFonts.DefaultFontName
      ELSE
        szFont = SciColorsAndFonts.StringFontName
      END IF
      SciMsg pSci, %SCI_STYLESETFONT, %SCE_B_STRING, VARPTR( szFont )
      SciMsg pSci, %SCI_STYLESETCASE, %SCE_B_STRING, EdOpt.DefaultCase
      IF SciColorsAndFonts.UseAlwaysDefaultFontSize THEN
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_STRING, SciColorsAndFonts.DefaultFontSize
      ELSE
        SciMsg pSci, %SCI_STYLESETSIZE, %SCE_B_STRING, SciColorsAndFonts.StringFontSize
      END IF
      SED_SetCharset pSci, %SCE_B_STRING, SciColorsAndFonts.StringFontCharset
      SciMsg pSci, %SCI_STYLESETBOLD, %SCE_B_STRING, SciColorsAndFonts.StringFontBold
      SciMsg pSci, %SCI_STYLESETITALIC, %SCE_B_STRING, SciColorsAndFonts.StringFontItalic
      SciMsg pSci, %SCI_STYLESETUNDERLINE, %SCE_B_STRING, SciColorsAndFonts.StringFontUnderline
      SciMsg pSci, %SCI_STYLESETFORE, %SCE_B_STRING, SciColorsAndFonts.StringForeColor
      IF SciColorsAndFonts.UseAlwaysDefaultBackColor THEN
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_STRING, SciColorsAndFonts.DefaultBackColor
      ELSE
        SciMsg pSci, %SCI_STYLESETBACK, %SCE_B_STRING, SciColorsAndFonts.StringBackColor
      END IF
      ' Set the lexer
      IF FoldOpt.FoldingLevel = 1 THEN
        SciMsg pSci, %SCI_SETLEXER, %SCLEX_POWERBASIC2, 0
      ELSE
        SciMsg pSci, %SCI_SETLEXER, %SCLEX_POWERBASIC, 0
      END IF
      IF FoldOpt.FoldingLevel THEN
        ' Enable folding of the procedures and functions
        szKey = "fold" : szValue = "1"
        SciMsg pSci, %SCI_SETPROPERTY, VARPTR( szKey ), VARPTR( szValue )
      END IF
      ' Set PB Keywords
      SciMsg pSci, %SCI_SETKEYWORDS, 0, BYVAL STRPTR( strPbKeyWords )
    END IF
  END IF
  ' =======================================================================================
  ' Codetips
  SciMsg pSci, %SCI_CALLTIPSETFORE, SciColorsAndFonts.CodetipForeColor, 0
  SciMsg pSci, %SCI_CALLTIPSETBACK, SciColorsAndFonts.CodetipBackColor, 0
  ' Caret
  SciMsg pSci, %SCI_SETCARETFORE, SciColorsAndFonts.CaretForeColor, 0
  ' Show caret line
  IF EdOpt.ShowCaretLine = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETCARETLINEVISIBLE, %TRUE, 0
  ELSE
    SciMsg pSci, %SCI_SETCARETLINEVISIBLE, %FALSE, 0
  END IF
  ' Caret line color
  SciMsg pSci, %SCI_SETCARETLINEBACK, SciColorsAndFonts.CaretLineBackColor, 0
  ' Selection
  SciMsg pSci, %SCI_SETSELFORE, %TRUE, SciColorsAndFonts.SelectionForeColor
  SciMsg pSci, %SCI_SETSELBACK, %TRUE, SciColorsAndFonts.SelectionBackColor
  ' Tabs
  IF EdOpt.UseTabs = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETUSETABS, %TRUE, 0
    CHECKMENUITEM hMenu, %IDM_USETABS, %MF_CHECKED
  ELSE
    SciMsg pSci, %SCI_SETUSETABS, %FALSE, 0
    CHECKMENUITEM hMenu, %IDM_USETABS, %MF_UNCHECKED
  END IF
  SciMsg pSci, %SCI_SETTABWIDTH, EdOpt.TabSize, 0
  ' Auto indentation
  IF EdOpt.AutoIndent = %BST_CHECKED THEN
    CHECKMENUITEM hMenu, %IDM_AUTOINDENT, %MF_CHECKED
  ELSE
    CHECKMENUITEM hMenu, %IDM_AUTOINDENT, %MF_UNCHECKED
  END IF
  SciMsg pSci, %SCI_SETINDENT, EdOpt.IndentSize, 0
  ' Indentation guides
  IF EdOpt.IndentGuides = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETINDENT, EdOpt.TabSize, 0
    CHECKMENUITEM hMenu, %IDM_SHOWINDENT, %MF_CHECKED
  ELSE
    SciMsg pSci, %SCI_SETINDENT, 0, 0
    CHECKMENUITEM hMenu, %IDM_SHOWINDENT, %MF_UNCHECKED
  END IF
  IF EdOpt.IndentGuides = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETINDENTATIONGUIDES, %TRUE, 0
  ELSE
    SciMsg pSci, %SCI_SETINDENTATIONGUIDES, %FALSE, 0
  END IF
  SciMsg pSci, %SCI_STYLESETFORE, %STYLE_INDENTGUIDE, SciColorsAndFonts.IndentGuideForeColor
  SciMsg pSci, %SCI_STYLESETBACK, %STYLE_INDENTGUIDE, SciColorsAndFonts.IndentGuideBackColor
  ' Set edge column and mode
  IF EdOpt.EdgeColumn = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETEDGEMODE, %EDGE_LINE, 0
    CHECKMENUITEM hMenu, %IDM_SHOWEDGE, %MF_CHECKED
  ELSE
    SciMsg pSci, %SCI_SETEDGEMODE, %EDGE_NONE, 0
    CHECKMENUITEM hMenu, %IDM_SHOWEDGE, %MF_UNCHECKED
  END IF
  SciMsg pSci, %SCI_SETEDGECOLUMN, EdOpt.EdgeWidth, 0
  SciMsg pSci, %SCI_SETEDGECOLOUR, SciColorsAndFonts.EdgeForeColor, 0
  ' Show white spaces as dots
  IF GETMENUSTATE( hMenu, %IDM_SHOWSPACES, %MF_BYCOMMAND = %MF_CHECKED ) THEN
    SciMsg pSci, %SCI_SETVIEWWS, %SCWS_VISIBLEALWAYS, 0
    CHECKMENUITEM hMenu, %IDM_SHOWSPACES, %MF_CHECKED
  ELSE
    SciMsg pSci, %SCI_SETVIEWWS, %SCWS_INVISIBLE, 0
    CHECKMENUITEM hMenu, %IDM_SHOWSPACES, %MF_UNCHECKED
  END IF
  ' White space
  IF EdOpt.WhiteSpace = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETVIEWWS, %SCWS_VISIBLEALWAYS, 0
    CHECKMENUITEM hMenu, %IDM_SHOWSPACES, %MF_CHECKED
  ELSE
    SciMsg pSci, %SCI_SETVIEWWS, %SCWS_INVISIBLE, 0
    CHECKMENUITEM hMenu, %IDM_SHOWSPACES, %MF_UNCHECKED
  END IF
  SciMsg pSci, %SCI_SETWHITESPACEFORE, %TRUE, SciColorsAndFonts.WhitespaceForeColor
  SciMsg pSci, %SCI_SETWHITESPACEBACK, %TRUE, SciColorsAndFonts.WhitespaceBackColor
  ' End of line
  IF EdOpt.EndOfLine = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETVIEWEOL, %TRUE, 0
    CHECKMENUITEM hMenu, %IDM_SHOWEOL, %MF_CHECKED
  ELSE
    SciMsg pSci, %SCI_SETVIEWEOL, %FALSE, 0
    CHECKMENUITEM hMenu, %IDM_SHOWEOL, %MF_UNCHECKED
  END IF
  ' Magnification
  SciMsg pSci, %SCI_SETZOOM, EdOpt.Magnification, 0
  ' Margin 0 for numbers
  SciMsg pSci, %SCI_SETMARGINTYPEN, 0, %SC_MARGIN_NUMBER
  SciMsg pSci, %SCI_SETMARGINSENSITIVEN, 0, 1
  IF EdOpt.LineNumbers = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETMARGINWIDTHN, 0, EdOpt.LineNumbersWidth
    CHECKMENUITEM hMenu, %IDM_SHOWLINENUM, %MF_CHECKED
  ELSE
    SciMsg pSci, %SCI_SETMARGINWIDTHN, 0, 0
    CHECKMENUITEM hMenu, %IDM_SHOWLINENUM, %MF_UNCHECKED
  END IF
  ' Margin 1 for non-folding symbols
  SciMsg pSci, %SCI_SETMARGINTYPEN, 1, %SC_MARGIN_SYMBOL
  SciMsg pSci, %SCI_SETMARGINMASKN, 1, 1
  SciMsg pSci, %SCI_SETMARGINSENSITIVEN, 1, 1
  SciMsg pSci, %SCI_SETMARGINWIDTHN, 1, 16
  SciMsg pSci, %SCI_SETFOLDMARGINCOLOUR, 2, RGB( 200, 0, 200 )
  SciMsg pSci, %SCI_SETFOLDMARGINHICOLOUR, 2, RGB( 100, 0, 100 )
  ' Margin 2 for folding symbols
  SciMsg pSci, %SCI_SETMARGINTYPEN, 2, %SC_MARGIN_SYMBOL
  SciMsg pSci, %SCI_SETMARGINMASKN, 2, %SC_MASK_FOLDERS
  SciMsg pSci, %SCI_SETMARGINSENSITIVEN, 2, 1
  IF EdOpt.Margin = %BST_CHECKED THEN
    SciMsg pSci, %SCI_SETMARGINWIDTHN, 2, EdOpt.MarginWidth
    CHECKMENUITEM hMenu, %IDM_SHOWMARGIN, %MF_CHECKED
  ELSE
    SciMsg pSci, %SCI_SETMARGINWIDTHN, 2, 0
    CHECKMENUITEM hMenu, %IDM_SHOWMARGIN, %MF_UNCHECKED
  END IF
  SELECT CASE FoldOpt.FoldingSymbol
    CASE 1
      ' Initialize fold symbols for folding - Arrow (mimics MacIntosh)
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEROPEN, %SC_MARK_ARROWDOWN
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDER, %SC_MARK_ARROW
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERSUB, %SC_MARK_EMPTY
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERTAIL, %SC_MARK_EMPTY
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEREND, %SC_MARK_EMPTY
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEROPENMID, %SC_MARK_EMPTY
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERMIDTAIL, %SC_MARK_EMPTY
    CASE 2
      ' Initialize fold symbols for folding - Plus/minus
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEROPEN, %SC_MARK_MINUS
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDER, %SC_MARK_PLUS
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERSUB, %SC_MARK_EMPTY
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERTAIL, %SC_MARK_EMPTY
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEREND, %SC_MARK_EMPTY
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEROPENMID, %SC_MARK_EMPTY
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERMIDTAIL, %SC_MARK_EMPTY
    CASE 3
      ' Initialize fold symbols for folding - Circle
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEROPEN, %SC_MARK_CIRCLEMINUS
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDER, %SC_MARK_CIRCLEPLUS
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERSUB, %SC_MARK_VLINE
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERTAIL, %SC_MARK_LCORNERCURVE
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEREND, %SC_MARK_CIRCLEPLUSCONNECTED
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEROPENMID, %SC_MARK_CIRCLEMINUSCONNECTED
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERMIDTAIL, %SC_MARK_TCORNERCURVE
    CASE 4
      ' Initialize fold symbols for folding - Box tree
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEROPEN, %SC_MARK_BOXMINUS
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDER, %SC_MARK_BOXPLUS
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERSUB, %SC_MARK_VLINE
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERTAIL, %SC_MARK_LCORNER
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEREND, %SC_MARK_BOXPLUSCONNECTED
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDEROPENMID, %SC_MARK_BOXMINUSCONNECTED
      SciMsg pSci, %SCI_MARKERDEFINE, %SC_MARKNUM_FOLDERMIDTAIL, %SC_MARK_TCORNER
  END SELECT
  ' Draw line below if not expanded
  SciMsg pSci, %SCI_SETFOLDFLAGS, 16, 0
  ' Fold margin colors.
  SciMsg pSci, %SCI_SETFOLDMARGINCOLOUR, 2, SciColorsAndFonts.FoldMarginForeColor
  SciMsg pSci, %SCI_SETFOLDMARGINHICOLOUR, 2, SciColorsAndFonts.FoldMarginBackColor
  ' Colors for folders closed and folders opened
  SciMsg pSci, %SCI_MARKERSETFORE, %SC_MARKNUM_FOLDER, SciColorsAndFonts.FoldForeColor
  SciMsg pSci, %SCI_MARKERSETBACK, %SC_MARKNUM_FOLDER, SciColorsAndFonts.FoldBackColor
  SciMsg pSci, %SCI_MARKERSETFORE, %SC_MARKNUM_FOLDEROPEN, SciColorsAndFonts.FoldOpenForeColor
  SciMsg pSci, %SCI_MARKERSETBACK, %SC_MARKNUM_FOLDEROPEN, SciColorsAndFonts.FoldOpenBackColor
  ' Add "." and "_" to the list of allowable characters for word identification
  szValue = "._:\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
  SciMsg pSci, %SCI_SETWORDCHARS, 0, VARPTR( szValue )
  ' Disable right click popup menu
  SciMsg pSci, %SCI_USEPOPUP, %FALSE, %FALSE
  ' For some reason it works with %SCMOD_SHIFT and %SCMOD_ALT but not with %SCMOD_CTRL
  '   Local KeyMod as long
  '   Local KeyId as long
  '   KeyMod = %SCMOD_CTRL
  '   Shift LEFT KeyMod, 16
  '   KeyId = %SCK_INSERT
  '   KeyId = KeyId or KeyMod
  '   SciMsg pSci, %SCI_ASSIGNCMDKEY, KeyId, %SCI_COPY
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Retrieves the line number at the caret position
' *********************************************************************************************
FUNCTION GetCurrentLine( ) AS LONG
  LOCAL hSci AS DWORD
  LOCAL curPos AS DWORD
  hSci = GetEdit
  curPos = SENDMESSAGE( hSci, %SCI_GETCURRENTPOS, 0, 0 )
  FUNCTION = SENDMESSAGE( hSci, %SCI_LINEFROMPOSITION, curPos, 0 )
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' Toggle procedure/function folding
' *********************************************************************************************
FUNCTION ToggleFolding( BYVAL LineNumber AS LONG ) AS LONG
  LOCAL hSci AS DWORD
  LOCAL pSci AS DWORD
  hSci = GetEdit
  ' Get direct pointer for faster access
  pSci = SENDMESSAGE( hSci, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT FUNCTION
  IF ( SciMsg( pSci, %SCI_GETFOLDLEVEL, LineNumber, 0 ) AND %SC_FOLDLEVELHEADERFLAG ) = 0 THEN            ' If is not the head line...
    LineNumber = SciMsg( pSci, %SCI_GETFOLDPARENT, LineNumber, 0 )          ' Get the number of the head line of the procedure or function
  END IF
  IF LineNumber > - 1 THEN
    SciMsg pSci, %SCI_TOGGLEFOLD, LineNumber, 0             ' Toggle the sub or function
    SciMsg pSci, %SCI_GOTOLINE, LineNumber, 0               ' Set the caret position
  END IF
  FUNCTION = LineNumber           ' Return the current line
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' Toggle all functions and procedures below
' *********************************************************************************************
SUB ToggleAllFoldersBelow( BYVAL LineNumber AS LONG )
  LOCAL hSci AS DWORD
  LOCAL pSci AS DWORD
  LOCAL FoldState AS LONG
  LOCAL i AS LONG
  LOCAL LineCount AS LONG
  hSci = GetEdit
  ' Get direct pointer for faster access
  pSci = SENDMESSAGE( hSci, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT SUB
  ' Force the lexer to style the whole document
  SciMsg( pSci, %SCI_COLOURISE, 0, - 1 )
  ' Toggle the first sub or function
  LineNumber = ToggleFolding( LinenUmber )
  ' Determine wether the fold is expanded or not
  FoldState = SciMsg( pSci, %SCI_GETFOLDEXPANDED, LineNumber, 0 )
  ' Toggle the rest of functions/procedures
  LineCount = SciMsg( pSci, %SCI_GETLINECOUNT, 0, 0 )             ' Number of lines
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_APPSTARTING )            ' Show hourglass mouse
  FOR i = LineNumber TO LineCount
    IF ( SciMsg( pSci, %SCI_GETFOLDLEVEL, i, 0 ) AND %SC_FOLDLEVELHEADERFLAG ) THEN                 ' If we are in the head line ...
      IF SciMsg( pSci, %SCI_GETFOLDEXPANDED, i, 0 ) <> FoldState THEN                 ' If the state is different ...
        SciMsg pSci, %SCI_TOGGLEFOLD, i, 0              ' Toggle it
      END IF
    END IF
  NEXT
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_ARROW )          ' Show standard mouse
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Folds all functions and procedures
' *********************************************************************************************
SUB FoldAllProcedures
  LOCAL hSci AS DWORD
  LOCAL pSci AS DWORD
  LOCAL i AS LONG
  LOCAL LineCount AS LONG
  hSci = GetEdit
  ' Get direct pointer for faster access
  pSci = SENDMESSAGE( hSci, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT SUB
  ' Force the lexer to style the whole document
  SciMsg( pSci, %SCI_COLOURISE, 0, - 1 )
  LineCount = SciMsg( pSci, %SCI_GETLINECOUNT, 0, 0 )             ' Number of lines
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_APPSTARTING )            ' Show hourglass mouse
  FOR i = 0 TO LineCount
    IF ( SciMsg( pSci, %SCI_GETFOLDLEVEL, i, 0 ) AND %SC_FOLDLEVELHEADERFLAG ) THEN                 ' If we are in the head line ...
      IF ISTRUE SciMsg( pSci, %SCI_GETFOLDEXPANDED, i, 0 ) THEN               ' If it is expanded...
        SciMsg pSci, %SCI_TOGGLEFOLD, i, 0              ' fold it
      END IF
    END IF
  NEXT
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_ARROW )          ' Show standard mouse
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Expands all functions and procedures
' *********************************************************************************************
SUB ExpandAllProcedures
  LOCAL hSci AS DWORD
  LOCAL pSci AS DWORD
  LOCAL i AS LONG
  LOCAL LineCount AS LONG
  hSci = GetEdit
  ' Get direct pointer for faster access
  pSci = SENDMESSAGE( hSci, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT SUB
  ' Force the lexer to style the whole document
  SciMsg( pSci, %SCI_COLOURISE, 0, - 1 )
  LineCount = SciMsg( pSci, %SCI_GETLINECOUNT, 0, 0 )             ' Number of lines
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_APPSTARTING )            ' Show hourglass mouse
  FOR i = 1 TO LineCount
    IF ( SciMsg( pSci, %SCI_GETFOLDLEVEL, i, 0 ) AND %SC_FOLDLEVELHEADERFLAG ) THEN                 ' If we are in the head line ...
      IF ISFALSE SciMsg( pSci, %SCI_GETFOLDEXPANDED, i, 0 ) THEN              ' If it is not expanded...
        SciMsg pSci, %SCI_TOGGLEFOLD, i, 0              ' expand it
      END IF
    END IF
  NEXT
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_ARROW )          ' Show standard mouse
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Block comment
' *********************************************************************************************
SUB BlockComment
  LOCAL hEdit AS DWORD            ' // Scintilla control's handle
  LOCAL pSci AS DWORD             ' // Scintilla control's direct pointer
  LOCAL curPos AS LONG            ' // Current position
  LOCAL startPos AS LONG          ' // Starting position
  LOCAL endPos AS LONG            ' // Ending position
  LOCAL startLine AS LONG                 ' // Starting line
  LOCAL endLine AS LONG           ' // Ending line
  LOCAL nCol AS LONG              ' // Column
  LOCAL POS AS LONG               ' // Position
  LOCAL nLen AS LONG              ' // Length of the line
  LOCAL i AS LONG                 ' // Loop counter
  LOCAL x AS LONG                 ' // Loop counter
  LOCAL buffer AS STRING          ' // Buffer
  LOCAL szText AS ASCIIZ * 255            ' // Portion of the line to replace
  LOCAL IndentSize AS LONG                ' // Size of indentation
  hEdit = GetEdit
  IF ISFALSE hEdit THEN EXIT SUB
  ' Get direct pointer for faster access
  pSci = SENDMESSAGE( hEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT SUB
  curPos = SciMsg( pSci, %SCI_GETCURRENTPOS, 0, 0 )
  startPos = SciMsg( pSci, %SCI_GETSELECTIONSTART, 0, 0 )
  startLine = SciMsg( pSci, %SCI_LINEFROMPOSITION, startPos, 0 )
  endPos = SciMsg( pSci, %SCI_GETSELECTIONEND, 0, 0 )
  endLine = SciMsg( pSci, %SCI_LINEFROMPOSITION, endPos, 0 )
  nCol = SciMsg( pSci, %SCI_GETCOLUMN, endPos, 0 )
  IF nCol = 0 AND endLine > startLine THEN endLine = endLine - 1
  SciMsg( pSci, %SCI_GOTOLINE, startLine, 0 )
  IF ( SciMsg( pSci, %SCI_GETFOLDLEVEL, startLine, 0 ) AND %SC_FOLDLEVELHEADERFLAG ) <> 0 THEN            ' If is the head line...
    IF ISFALSE SciMsg( pSci, %SCI_GETFOLDEXPANDED, startLine, 0 ) THEN              ' If it is folded...
      ToggleFolding startLine                 ' ...toggle it
    END IF
  END IF
  FOR i = startLine TO endLine
    szText = ""
    nLen = SciMsg( pSci, %SCI_LINELENGTH, i, 0 )
    buffer = SPACE$( nLen )
    SciMsg pSci, %SCI_GETLINE, i, STRPTR( buffer )
    IndentSize = SciMsg( pSci, %SCI_GETINDENT, 0, 0 )
    FOR x = 1 TO LEN( buffer )
      IF MID$( buffer, x, 1 ) <> " " AND MID$( buffer, x, 1 ) <> CHR$( 9 ) THEN
        szText = LEFT$( buffer, x - 1 )
        EXIT FOR
      END IF
    NEXT
    buffer = LTRIM$( buffer, ANY CHR$( 32, 9, 13, 10 ))
    POS = SciMsg( pSci, %SCI_POSITIONFROMLINE, i, 0 )
    IF LEN( buffer ) THEN
      szText = "'" & szText
      SciMsg pSci, %SCI_SETTARGETSTART, POS, 0
      SciMsg pSci, %SCI_SETTARGETEND, POS + LEN( szText ) - 1, 0
      SciMsg pSci, %SCI_REPLACETARGET, - 1, VARPTR( szText )
      IF startLine <> endLine THEN curPos = curPos + 1
    END IF
  NEXT
  SciMsg pSci, %SCI_SETCURRENTPOS, curPos, 0
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Block uncomment
' *********************************************************************************************
SUB BlockUncomment
  LOCAL hEdit AS DWORD            ' // Scintilla control's handle
  LOCAL pSci AS DWORD             ' // Scintilla control's direct pointer
  LOCAL curPos AS LONG            ' // Current position
  LOCAL startPos AS LONG          ' // Starting position
  LOCAL endPos AS LONG            ' // Ending position
  LOCAL startLine AS LONG                 ' // Starting line
  LOCAL endLine AS LONG           ' // Ending line
  LOCAL nCol AS LONG              ' // Column
  LOCAL POS AS LONG               ' // Position
  LOCAL nLen AS LONG              ' // Length of the line
  LOCAL i AS LONG                 ' // Loop counter
  LOCAL x AS LONG                 ' // Loop counter
  LOCAL buffer AS STRING          ' // Buffer
  LOCAL szText AS ASCIIZ * 255            ' // Portion of the line to replace
  LOCAL IndentSize AS LONG                ' // Size of indentation
  hEdit = GetEdit
  IF ISFALSE hEdit THEN EXIT SUB
  ' Get direct pointer for faster access
  pSci = SENDMESSAGE( hEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT SUB
  curPos = SciMsg( pSci, %SCI_GETCURRENTPOS, 0, 0 )
  startPos = SciMsg( pSci, %SCI_GETSELECTIONSTART, 0, 0 )
  startLine = SciMsg( pSci, %SCI_LINEFROMPOSITION, startPos, 0 )
  endPos = SciMsg( pSci, %SCI_GETSELECTIONEND, 0, 0 )
  endLine = SciMsg( pSci, %SCI_LINEFROMPOSITION, endPos, 0 )
  nCol = SciMsg( pSci, %SCI_GETCOLUMN, endPos, 0 )
  IF nCol = 0 AND endLine > startLine THEN endLine = endLine - 1
  SciMsg( pSci, %SCI_GOTOLINE, startLine, 0 )
  IF ( SciMsg( pSci, %SCI_GETFOLDLEVEL, startLine, 0 ) AND %SC_FOLDLEVELHEADERFLAG ) <> 0 THEN            ' If is the head line...
    IF ISFALSE SciMsg( pSci, %SCI_GETFOLDEXPANDED, startLine, 0 ) THEN              ' If it is folded...
      ToggleFolding startLine                 ' ...toggle it
    END IF
  END IF
  FOR i = startLine TO endLine
    szText = ""
    nLen = SciMsg( pSci, %SCI_LINELENGTH, i, 0 )
    buffer = SPACE$( nLen )
    SciMsg pSci, %SCI_GETLINE, i, STRPTR( buffer )
    IndentSize = SciMsg( pSci, %SCI_GETINDENT, 0, 0 )
    FOR x = 1 TO LEN( buffer )
      IF MID$( buffer, x, 1 ) <> " " AND MID$( buffer, x, 1 ) <> CHR$( 9 ) THEN
        szText = LEFT$( buffer, x - 1 )
        EXIT FOR
      END IF
    NEXT
    buffer = LTRIM$( buffer, ANY CHR$( 32, 9, 13, 10 ))
    POS = SciMsg( pSci, %SCI_POSITIONFROMLINE, i, 0 )
    IF LEN( buffer ) THEN
      IF LEFT$( buffer, 1 ) = "'" THEN
        SciMsg pSci, %SCI_SETTARGETSTART, POS, 0
        SciMsg pSci, %SCI_SETTARGETEND, POS + LEN( szText ) + 1, 0
        SciMsg pSci, %SCI_REPLACETARGET, - 1, VARPTR( szText )
        IF startLine <> endLine THEN curPos = curPos - 1
      END IF
    END IF
  NEXT
  SciMsg pSci, %SCI_SETCURRENTPOS, curPos, 0
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Reset the CodeFinder combobox content
' *********************************************************************************************
SUB ResetCodeFinder
  LOCAL buffer AS STRING
  SENDMESSAGE ghComboBox, %CB_RESETCONTENT, 0, 0
  buffer = "(代码快速定位)"
  SENDMESSAGE ghComboBox, %CB_ADDSTRING, 0, BYVAL STRPTR( buffer )
  SENDMESSAGE ghComboBox, %CB_SETITEMDATA, 0, 1
  SENDMESSAGE ghComboBox, %CB_SETCURSEL, 0, 0
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Find subs and functions
' *********************************************************************************************
SUB SED_CodeFinder
  LOCAL hSci AS DWORD             ' // Handle of the Scintilla control
  LOCAL pSci AS DWORD             ' // Scintilla control's direct pointer
  LOCAL nLines AS LONG            ' // Number of lines
  LOCAL nLen AS LONG              ' // Length of the line
  LOCAL i AS LONG                 ' // Loop counter
  LOCAL buffer AS STRING          ' // Buffer
  LOCAL strProcName AS STRING             ' // Sub/Function name
  LOCAL p AS LONG                 ' // Position
  LOCAL strPrevName AS STRING             ' // Previous name
  ' Reset the combobox content
  ResetCodeFinder
  ' Get the handle of the edit control
  hSci = GetEdit
  IF ISFALSE hSci THEN EXIT SUB
  ' Get direct pointer for faster access
  pSci = SENDMESSAGE( hSci, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT SUB
  ' Number of lines
  nLines = SciMsg( pSci, %SCI_GETLINECOUNT, 0, 0 )
  ' Parse the text
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_APPSTARTING )            ' Show hourglass mouse
  FOR i = 0 TO nLines - 1
    nLen = SciMsg( pSci, %SCI_LINELENGTH, i, 0 )
    buffer = SPACE$( nLen )
    SciMsg pSci, %SCI_GETLINE, i, STRPTR( buffer )
    buffer = LTRIM$( buffer, ANY CHR$( 32, 9 ))
    strProcName = ""
    ' ---------------------------------------------------------------------------------
    IF LEN( strPrevName ) THEN buffer = strPrevName & buffer
    strPrevName = ""
    IF UCASE$( LEFT$( buffer, 9 )) = "FUNCTION " AND _
            MID$( LTRIM$( MID$( buffer, 10 )), 1, 1 ) = "_" THEN
      strPrevName = "FUNCTION "
      ITERATE FOR
    ELSEIF UCASE$( LEFT$( buffer, 4 )) = "SUB " AND _
            MID$( LTRIM$( MID$( buffer, 5 )), 1, 1 ) = "_" THEN
      strPrevName = "SUB "
      ITERATE FOR
    ELSEIF UCASE$( LEFT$( buffer, 9 )) = "CALLBACK " THEN
      buffer = LTRIM$( MID$( buffer, 10 ), ANY CHR$( 32, 9 ))
      IF UCASE$( LEFT$( buffer, 9 )) = "FUNCTION " AND _
              MID$( LTRIM$( MID$( buffer, 10 )), 1, 1 ) = "_" THEN
        strPrevName = "FUNCTION "
        ITERATE FOR
      END IF
    ELSEIF UCASE$( LEFT$( buffer, 7 )) = "STATIC " THEN
      buffer = LTRIM$( MID$( buffer, 8 ), ANY CHR$( 32, 9 ))
      IF UCASE$( LEFT$( buffer, 9 )) = "FUNCTION " AND _
              MID$( LTRIM$( MID$( buffer, 10 )), 1, 1 ) = "_" THEN
        strPrevName = "FUNCTION "
        ITERATE FOR
      ELSEIF UCASE$( LEFT$( buffer, 4 )) = "SUB " AND _
              MID$( LTRIM$( MID$( buffer, 5 )), 1, 1 ) = "_" THEN
        strPrevName = "SUB "
        ITERATE FOR
      END IF
    END IF
    ' ---------------------------------------------------------------------------------
    IF UCASE$( LEFT$( buffer, 9 )) = "FUNCTION " AND _
            INSTR( 10, buffer, "=" ) = 0 THEN
      buffer = LTRIM$( MID$( buffer, 10 ))
      p = INSTR( buffer, "(" )
      IF p = 0 THEN p = INSTR( buffer, " " )
      IF p = 0 THEN p = INSTR( buffer, $CR )
      IF p THEN strProcName = LEFT$( buffer, p - 1 )
    ELSEIF UCASE$( LEFT$( buffer, 4 )) = "SUB " THEN
      buffer = LTRIM$( MID$( buffer, 5 ))
      p = INSTR( buffer, "(" )
      IF p = 0 THEN p = INSTR( buffer, " " )
      IF p = 0 THEN p = INSTR( buffer, $CR )
      IF p THEN strProcName = LEFT$( buffer, p - 1 )
    ELSEIF UCASE$( LEFT$( buffer, 9 )) = "CALLBACK " THEN
      buffer = LTRIM$( MID$( buffer, 10 ), ANY CHR$( 32, 9 ))
      IF UCASE$( LEFT$( buffer, 9 )) = "FUNCTION " THEN
        buffer = LTRIM$( MID$( buffer, 10 ))
        p = INSTR( buffer, "(" )
        IF p = 0 THEN p = INSTR( buffer, " " )
        IF p = 0 THEN p = INSTR( buffer, $CR )
        IF p THEN strProcName = LEFT$( buffer, p - 1 )
      END IF
    ELSEIF UCASE$( LEFT$( buffer, 7 )) = "STATIC " THEN
      buffer = LTRIM$( MID$( buffer, 8 ), ANY CHR$( 32, 9 ))
      IF UCASE$( LEFT$( buffer, 9 )) = "FUNCTION " THEN
        buffer = LTRIM$( MID$( buffer, 10 ))
        p = INSTR( buffer, "(" )
        IF p = 0 THEN p = INSTR( buffer, " " )
        IF p = 0 THEN p = INSTR( buffer, $CR )
        IF p THEN strProcName = LEFT$( buffer, p - 1 )
      ELSEIF UCASE$( LEFT$( buffer, 4 )) = "SUB " THEN
        buffer = LTRIM$( MID$( buffer, 5 ))
        p = INSTR( buffer, "(" )
        IF p = 0 THEN p = INSTR( buffer, " " )
        IF p = 0 THEN p = INSTR( buffer, $CR )
        IF p THEN strProcName = LEFT$( buffer, p - 1 )
      END IF
    END IF
    strProcName = TRIM$( strProcName )
    IF LEN( strProcName ) THEN
      p = SENDMESSAGE( ghComboBox, %CB_ADDSTRING, 0, BYVAL STRPTR( strProcName ))
      SENDMESSAGE ghComboBox, %CB_SETITEMDATA, p, i
    END IF
  NEXT
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_ARROW )
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Returns the handle of the active edit control
' *********************************************************************************************
FUNCTION GetEdit( ) AS DWORD
  LOCAL wi AS WININFOSTRUC PTR
  ' Get the handles of the windows
  wi = GETWINDOWLONG( MdiGetActive( g_hWndClient ), %GWL_USERDATA )
  IF wi THEN
    IF @wi.hFocus THEN
      FUNCTION = @wi.hFocus
      EXIT FUNCTION
    END IF
  END IF
  FUNCTION = GETDLGITEM( MdiGetActive( g_hWndClient ), %IDC_EDIT1 )
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' Function    : Sci_GetCodeTip ()
' Description : Retrieves the CodeTip and returns it in as the result of the function.
' *********************************************************************************************
FUNCTION Sci_GetCodeTip( BYVAL strCodeTip AS STRING ) AS STRING
  LOCAL Record AS STRING * 576            ' // Record string
  LOCAL p AS LONG                 ' // Position
  LOCAL IniFile AS STRING
  IniFile=EXE.PATH$ & "config.ini"
  IF VAL( IniRead( IniFile, "Editor options", "CodeTips", "" )) <> %BST_CHECKED THEN EXIT FUNCTION
  IF ISFALSE hCodetipsFile THEN EXIT FUNCTION
  ' Get the Tsunami CodeTip Record Equal to strCodeTip.
  Record = trm_GetEqualOrGreater( hCodetipsFile, 1, strCodeTip )
  ' Return the Found CodeTip.
  IF UCASE$( TRIM$( LEFT$( Record, 64 ))) = UCASE$( strCodeTip ) THEN
    strCodeTip = MID$( Record, 65, LEN( Record ) - 64 )
    REPLACE "  " WITH " " IN strCodeTip
    p = INSTR( strCodeTip, "(" )
    IF p THEN
      IF MID$( strCodeTip, p + 1, 1 ) <> ")" AND MID$( strCodeTip, p + 1, 2 ) <> " )" THEN
        REPLACE "(" WITH "(" & $LF IN strCodeTip
        REPLACE ")" WITH $LF & ")" IN strCodeTip
      END IF
    END IF
    REPLACE ", " WITH "," IN strCodeTip
    REPLACE "," WITH "," & $LF IN strCodeTip
    FUNCTION = strCodeTip
  END IF
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' Show codetips
' *********************************************************************************************
SUB ShowCodetip
  LOCAL txtrg AS TEXTRANGE                ' // Text range
  LOCAL x AS LONG                 ' // Start of the word
  LOCAL y AS LONG                 ' // End of the word
  LOCAL curPos AS LONG            ' // Current position
  LOCAL buffer AS STRING          ' // Buffer
  LOCAL tipPos AS LONG            ' // Tip position
  LOCAL p AS LONG                 ' // Position
  LOCAL CodetipStr AS STRING              ' // Codetip string
  ' Retrieve the word under the caret
  x = 0 : y = 0
  ' Retrieve the current position
  curPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )
  tipPos = curPos
  ' Skip the "("
  DECR curPos
  DO
    ' Retrieve the starting position of the word
    x = SENDMESSAGE( GetEdit, %SCI_WORDSTARTPOSITION, curPos, %TRUE )
    ' Retrieve the ending position of the word
    y = SENDMESSAGE( GetEdit, %SCI_WORDENDPOSITION, curPos, %TRUE )
    IF x = 0 AND y = 0 THEN EXIT DO
    ' Prepare the buffer
    buffer = SPACE$( y - x + 1 )
    ' Text range
    txtrg.chrg.cpMin = x
    txtrg.chrg.cpMax = y
    txtrg.lpstrText = STRPTR( buffer )
    SENDMESSAGE GetEdit, %SCI_GETTEXTRANGE, 0, BYVAL VARPTR( txtrg )
    ' Remove the $NUL
    p = INSTR( buffer, CHR$( 0 ))
    IF p THEN buffer = LEFT$( buffer, p - 1 )
    ' Allow for spaces and tabs
    IF LEN( buffer ) = 0 THEN
      DECR curPos
    ELSE
      EXIT DO
    END IF
  LOOP
  IF LEN( buffer ) THEN
    CodetipStr = TRIM$( Sci_GetCodetip( buffer ))
    IF LEN( CodetipStr ) THEN
      SENDMESSAGE GetEdit, %SCI_CALLTIPSHOW, tipPos - LEN( buffer ), STRPTR( CodetipStr )
    END IF
  END IF
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Replaces spaces with tabs
' Code by Steve Miltiadous
' *********************************************************************************************
FUNCTION SpacesToTabs( BYVAL InString AS STRING ) AS STRING
  LOCAL TabSize AS LONG           ' // Size of the tab
  LOCAL buffer AS STRING          ' // Buffer
  LOCAL lLooper AS LONG           ' // Loop counter
  LOCAL sSpc AS STRING            ' // Spaces
  LOCAL sTab AS STRING            ' // Tabs
  LOCAL sOutString AS STRING              ' // Output string
  IF ( GETMENUSTATE( hMenu, %IDM_USETABS, %MF_BYCOMMAND ) AND %MF_CHECKED ) <> %MF_CHECKED THEN
    FUNCTION = InString
    EXIT FUNCTION
  END IF
  IF LEN( InString ) = 0 THEN EXIT FUNCTION
  TabSize = SENDMESSAGE( GetEdit, %SCI_GETTABWIDTH, 0, 0 )
  IF TabSize < 1 THEN
    FUNCTION = InString
    EXIT FUNCTION
  END IF
  buffer = InString
  FOR lLooper = 1 TO LEN( buffer )
    IF INSTR( MID$( buffer, lLooper, 1 ), CHR$( 32 )) THEN
      sSpc = sSpc + MID$( buffer, lLooper, 1 )
    ELSEIF INSTR( MID$( buffer, lLooper, 1 ), CHR$( 9 )) THEN
      sTab = sTab + MID$( buffer, lLooper, 1 )
    ELSE
      EXIT FOR
    END IF
  NEXT
  IF LEN( sSpc ) > = TabSize THEN
    REPLACE SPACE$( TabSize ) WITH $TAB IN sSpc
    sOutString = sTab + sSpc + MID$( buffer, lLooper )
  ELSE
    sOutString = InString
  END IF
  FUNCTION = sOutString
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' Process the Scintilla Edit Control notification messges
' *********************************************************************************************
FUNCTION Sci_OnNotify( BYVAL hWnd AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS DWORD ) AS LONG
  LOCAL pNSC AS SCNotification PTR                ' // Scintilla notification messages structure
  LOCAL LineNumber AS LONG                ' // Line number
  LOCAL curPos AS LONG            ' // Current position
  LOCAL LineLen AS LONG           ' // Line length
  LOCAL buffer AS STRING          ' // Buffer
  LOCAL IndentSize AS LONG                ' // Size of the indent
  LOCAL TabSize AS LONG           ' // Tab size
  LOCAL nSpaces AS LONG           ' // Number of spaces
  LOCAL strFill AS STRING                 ' // Filler string
  LOCAL i AS LONG                 ' // Loop counter
  LOCAL rightWord AS STRING               ' // Word on the right side of the line
  LOCAL strEnd AS STRING          ' // End of the construct
  LOCAL pSci AS DWORD             ' // Scintilla control's direct pointer
  LOCAL tmpPROC AS PROC           ' // PROC structure
  LOCAL szText AS ASCIIZ * 256            ' // General purpose variable
  pNSC = lParam
  SELECT CASE @pNSC.hdr.code
    CASE %SCN_UPDATEUI
      ' Show line and column
      ShowLinCol
      ' Show the sub/function name in the codefinder
      '         IF ShowProcedureName THEN
      '            SED_WithinProc(tmpPROC)
      '            buffer = tmpPROC.ProcName
      '            SendMessage hCodeFinder, %CB_SELECTSTRING , -1, BYVAL STRPTR(buffer)
      '         END IF
      ' Show the sub/function name in the statusbar
      IF ShowProcedureName THEN
        SED_WithinProc( tmpPROC )
        szText = tmpPROC.ProcName
        SENDMESSAGE g_hStatus, %SB_SETTEXT, 4, VARPTR( szText )
      END IF
    CASE %SCN_MARGINCLICK           ' Margin mouse click
      IF @pNSC.margin = 2 THEN                ' Folder margin
        LineNumber = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, @pNSC.position, 0 )
        IF ( SENDMESSAGE( GetEdit, %SCI_GETFOLDLEVEL, LineNumber, 0 ) AND %SC_FOLDLEVELHEADERFLAG ) <> 0 THEN           ' If is the head line...
          IF ToggleFolding( LineNumber ) = - 1 THEN
            SENDMESSAGE( GetEdit, %SCI_GOTOLINE, LineNumber, 0 )
            curPos = SENDMESSAGE( GetEdit, %SCI_POSITIONFROMLINE, LineNumber, 0 )
            SENDMESSAGE GetEdit, %SCI_SETSELECTIONSTART, curPos, 0
            SENDMESSAGE GetEdit, %SCI_SETSELECTIONEND, SENDMESSAGE( GetEdit, %SCI_GETLINEENDPOSITION, LineNumber, 0 ), 0
          END IF
        ELSE
          SENDMESSAGE( GetEdit, %SCI_GOTOLINE, LineNumber, 0 )
          curPos = SENDMESSAGE( GetEdit, %SCI_POSITIONFROMLINE, LineNumber, 0 )
          SENDMESSAGE GetEdit, %SCI_SETSELECTIONSTART, curPos, 0
          SENDMESSAGE GetEdit, %SCI_SETSELECTIONEND, SENDMESSAGE( GetEdit, %SCI_GETLINEENDPOSITION, LineNumber, 0 ), 0
        END IF
      ELSE
        LineNumber = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, @pNSC.position, 0 )
        SENDMESSAGE( GetEdit, %SCI_GOTOLINE, LineNumber, 0 )
        curPos = SENDMESSAGE( GetEdit, %SCI_POSITIONFROMLINE, LineNumber, 0 )
        SENDMESSAGE GetEdit, %SCI_SETSELECTIONSTART, curPos, 0
        SENDMESSAGE GetEdit, %SCI_SETSELECTIONEND, SENDMESSAGE( GetEdit, %SCI_GETLINEENDPOSITION, LineNumber, 0 ), 0
      END IF
    CASE %SCN_CHARADDED
      IF @pNSC.ch = 40 THEN ShowCodetip               ' Show codetip
      IF @pNSC.ch = 41 THEN SENDMESSAGE GetEdit, %SCI_CALLTIPCANCEL, 0, 0
      ' Auto indentation - since SCN_KEY isn't send in the Windows version,
      ' we detect the new line if the charadded is a carriage return
      IF @pNSC.ch = 13 THEN           ' carriage return
        IF GETMENUSTATE( hMenu, %IDM_AUTOINDENT, %MF_BYCOMMAND = %MF_CHECKED ) THEN
          curPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )               ' current position
          LineNumber = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, curPos, 0 )           ' line number
          LineLen = SENDMESSAGE( GetEdit, %SCI_LINELENGTH, LineNumber - 1, 0 )            ' length of the line
          buffer = SPACE$( LineLen )              ' size the buffer
          SENDMESSAGE( GetEdit, %SCI_GETLINE, LineNumber - 1, STRPTR( buffer ))           ' get the text of the line
          TabSize = SENDMESSAGE( GetEdit, %SCI_GETTABWIDTH, 0, 0 )                ' size of the tab
          nSpaces = 0             ' number of spaces on the left
          FOR i = 1 TO LEN( buffer )
            IF MID$( buffer, i, 1 ) <> " " THEN
              IF MID$( buffer, i, 1 ) = $TAB THEN
                nSpaces = nSpaces + TabSize
              ELSE
                EXIT FOR
              END IF
            ELSE
              nSpaces = nSpaces + 1
            END IF
          NEXT
          buffer = REMOVE$( buffer, ANY CHR$( 13, 10 ))           ' removes $CRLF
          buffer = TRIM$( UCASE$( buffer ), ANY CHR$( 32, 9 ))            ' removes spaces and tabs and converts to uppercase
          IF ( LEFT$( buffer, 3 ) = "IF " AND RIGHT$( buffer, 5 ) = " THEN" ) OR _
                  LEFT$( buffer, 4 ) = "ELSE" OR _
                  LEFT$( buffer, 7 ) = "SELECT " OR _
                  LEFT$( buffer, 5 ) = "CASE " OR _
                  LEFT$( buffer, 4 ) = "FOR " OR _
                  LEFT$( buffer, 3 ) = "DO " OR _
                  buffer = "DO" OR _
                  LEFT$( buffer, 6 ) = "WHILE " OR _
                  buffer = "WHILE" THEN
            IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )               ' size of the indent
            strFill = SPACE$( nSpaces + IndentSize )                ' add spaces to indent the line
          ELSE
            strFill = SPACE$( nSpaces )             ' add the same spaces on the left that the line above
          END IF
          strFill = SpacesToTabs( strFill )
          SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))          ' indents the line
        END IF
      END IF
      '*************************SM AUTOCOMPLETE START***************************************
      IF @pNSC.ch = 32 THEN
        IF ISTRUE ConstructAutocompletion OR ISTRUE AutocompleteTypes THEN
          curPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )               ' current position
          LineNumber = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, curPos, 0 )           ' line number
          LineLen = SENDMESSAGE( GetEdit, %SCI_LINELENGTH, LineNumber, 0 )                ' length of the line
          buffer = SPACE$( LineLen )              ' size the buffer
          SENDMESSAGE( GetEdit, %SCI_GETLINE, LineNumber, STRPTR( buffer ))               ' get the text of the line
          TabSize = SENDMESSAGE( GetEdit, %SCI_GETTABWIDTH, 0, 0 )                ' size of the tab
          nSpaces = 0             ' number of spaces on the left
          FOR i = 1 TO LEN( buffer )
            IF MID$( buffer, i, 1 ) <> " " THEN
              IF MID$( buffer, i, 1 ) = $TAB THEN
                nSpaces = nSpaces + TabSize
              ELSE
                EXIT FOR
              END IF
            ELSE
              nSpaces = nSpaces + 1
            END IF
          NEXT
          buffer = REMOVE$( buffer, ANY CHR$( 13, 10 ))           ' removes $CRLF
          buffer = TRIM$( UCASE$( buffer ), ANY CHR$( 32, 9 ))            ' removes spaces and tabs and converts to uppercase
          FOR i = LEN( buffer ) TO 1 STEP - 1
            IF MID$( buffer, i, 1 ) <> CHR$( 32 ) THEN
              rightWord = MID$( buffer, i, 1 ) + rightWord
            ELSE
              EXIT FOR
            END IF
          NEXT
          ' MSGBOX buffer & $crlf & rightWord
          IF ISTRUE ConstructAutocompletion THEN
            SELECT CASE buffer
              CASE "FUNCTION"
                SED_WithinProc( tmpPROC )
                IF tmpPROC.WhatIsDown <> %ID_ProcEnd THEN
                  strFill = "FunName()" & $CRLF
                  SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                  IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )
                  strFill = SpacesToTabs( SPACE$( nSpaces + IndentSize ))
                  SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                  strEnd = "END FUNCTION"
                  strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces )) & strEnd
                  SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                  SENDMESSAGE( GetEdit, %SCI_GOTOPOS, curPos, 0 )
                  SENDMESSAGE( GetEdit, %SCI_SETANCHOR, curPos + 7, 0 )
                END IF
              CASE "SUB"
                strFill = "SubName()" & $CRLF
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )
                strFill = SpacesToTabs( SPACE$( nSpaces + IndentSize ))
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                strEnd = "END SUB"
                strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces )) & strEnd
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                SENDMESSAGE( GetEdit, %SCI_GOTOPOS, curPos, 0 )
                SENDMESSAGE( GetEdit, %SCI_SETANCHOR, curPos + 7, 0 )
              CASE "IF"
                strFill = "x THEN" & $CRLF
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )
                strFill = SpacesToTabs( SPACE$( nSpaces + IndentSize ))
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                strEnd = "END IF"
                strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces )) & strEnd
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                SENDMESSAGE( GetEdit, %SCI_GOTOPOS, curPos, 0 )
                SENDMESSAGE( GetEdit, %SCI_SETANCHOR, curPos + 1, 0 )
              CASE "WHILE"
                strFill = "x" & $CRLF
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )
                strFill = SpacesToTabs( SPACE$( nSpaces + IndentSize ))
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                strEnd = "WEND"
                strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces )) & strEnd
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                SENDMESSAGE( GetEdit, %SCI_GOTOPOS, curPos, 0 )
                SENDMESSAGE( GetEdit, %SCI_SETANCHOR, curPos + 1, 0 )
              CASE "FOR"
                strFill = "x" & $CRLF
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )
                strFill = SpacesToTabs( SPACE$( nSpaces + IndentSize ))
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                strEnd = "NEXT"
                strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces )) & strEnd
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                SENDMESSAGE( GetEdit, %SCI_GOTOPOS, curPos, 0 )
                SENDMESSAGE( GetEdit, %SCI_SETANCHOR, curPos + 1, 0 )
              CASE "SELECT"
                strFill = "CASE x" & $CRLF
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )
                '                        strEnd = "CASE y"
                '                        strFill = SpacesToTabs(SPACE$(nSpaces + 3)) & strEnd
                '                        SendMessage(GetEdit, %SCI_ADDTEXT, LEN(strFill), STRPTR(strFill))
                strEnd = "CASE ELSE"
                strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces + 3 )) & strEnd
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                strEnd = "END SELECT"
                strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces )) & strEnd
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                SENDMESSAGE( GetEdit, %SCI_GOTOPOS, curPos + 5, 0 )
                SENDMESSAGE( GetEdit, %SCI_SETANCHOR, curPos + 6, 0 )
              CASE "TYPE"
                strFill = "TypeName" & $CRLF
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )
                strFill = SpacesToTabs( SPACE$( nSpaces + IndentSize ))
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                strEnd = "END TYPE"
                strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces )) & strEnd
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                SENDMESSAGE( GetEdit, %SCI_GOTOPOS, curPos, 0 )
                SENDMESSAGE( GetEdit, %SCI_SETANCHOR, curPos + 8, 0 )
              CASE "UNION"
                strFill = "UnionName" & $CRLF
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                IndentSize = SENDMESSAGE( GetEdit, %SCI_GETINDENT, 0, 0 )
                strFill = SpacesToTabs( SPACE$( nSpaces + IndentSize ))
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                strEnd = "END UNION"
                strFill = $CRLF & SpacesToTabs( SPACE$( nSpaces )) & strEnd
                SENDMESSAGE( GetEdit, %SCI_ADDTEXT, LEN( strFill ), STRPTR( strFill ))
                SENDMESSAGE( GetEdit, %SCI_GOTOPOS, curPos, 0 )
                SENDMESSAGE( GetEdit, %SCI_SETANCHOR, curPos + 9, 0 )
            END SELECT
          END IF
          IF ISTRUE AutocompleteTypes THEN
            IF UCASE$( rightWord ) = "AS" THEN
              IF LEN( sAutoCompletionTypes ) THEN
                pSci = SENDMESSAGE( GetEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
                SciMsg pSci, %SCI_AUTOCSETIGNORECASE, %TRUE, 0          ' Ignore case
                SciMsg pSci, %SCI_AUTOCSHOW, 0, STRPTR( sAutoCompletionTypes )
              END IF
            END IF
          END IF
        END IF
      END IF
      DIM tmpStr AS STRING
      LOCAL tmpLng AS LONG
      'tmpStr=GetTabName(TabCtrl_GetCurSel(ghTabMdi))
      tmpLng=SendMessage(g_hTabMdi, %TCM_GETCURSEL, 0, 0)
      tmpStr=GetTabName(tmpLng)
      IF tmpStr<>"" AND RIGHT$(tmpStr,2)<>" *" THEN
        tmpLng=SendMessage(g_hTabMdi, %TCM_GETCURSEL, 0, 0)
        SetTabName(tmpLng,tmpStr & " *")
      END IF
      '*******************************SM AUTOCOMPLETE END**********************************
  END SELECT
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' Autocompletion listbox
' *********************************************************************************************
FUNCTION AutoComplete( ) AS LONG
  LOCAL pSci AS DWORD             ' // Scintilla control's direct pointer
  LOCAL oldpSci AS DWORD          ' // Saved Scintilla direct pointer
  LOCAL hr AS LONG                ' // Result code
  LOCAL x AS LONG                 ' // Starting position
  LOCAL y AS LONG                 ' // Ending position
  LOCAL curPos AS LONG            ' // Current position
  LOCAL buffer AS STRING          ' // Buffer
  LOCAL txtrg AS TEXTRANGE                ' // Text range
  LOCAL findFlags AS LONG                 ' // Find flags
  LOCAL startPos AS LONG          ' // Starting position
  LOCAL endPos AS LONG            ' // Ending position
  LOCAL szFindText AS ASCIIZ * 256                ' // Text to find
  LOCAL numItems AS LONG          ' // Number of items
  LOCAL iPos AS LONG              ' // Position
  LOCAL strWordList AS STRING             ' // List of words
  LOCAL CodetipRecord AS STRING * 576             ' // Record string
  LOCAL hWndActive AS DWORD               ' // Handle of the active window
  LOCAL p AS LONG                 ' // String position
  LOCAL p1 AS LONG                ' // String position
  LOCAL p2 AS LONG                ' // String position
  '   local p2 as long                         ' // String position
  DIM WordList( 0 ) AS STRING             ' // Array of words found
  IF ISFALSE GetEdit THEN EXIT FUNCTION
  ' Get direct pointer for faster access
  pSci = SENDMESSAGE( GetEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT FUNCTION
  oldpSci = pSci
  curPos = SciMsg( pSci, %SCI_GETCURRENTPOS, 0, 0 )
  x = SciMsg( pSci, %SCI_WORDSTARTPOSITION, curPos, %TRUE )
  y = SciMsg( pSci, %SCI_WORDENDPOSITION, curPos, %TRUE )
  IF y < = x THEN EXIT FUNCTION
  buffer = SPACE$( y - x + 1 )
  txtrg.chrg.cpMin = x
  txtrg.chrg.cpMax = y
  txtrg.lpstrText = STRPTR( buffer )
  SciMsg pSci, %SCI_GETTEXTRANGE, 0, BYVAL VARPTR( txtrg )
  buffer = LEFT$( buffer, LEN( buffer ) - 1 )
  IF LEN( buffer ) < 2 THEN
    MESSAGEBOX BYVAL g_hWndMain, "Word too short to search   ", _
            " Exit", %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL
    EXIT FUNCTION
  END IF
  szFindText = buffer
  findFlags = %SCFIND_WORDSTART
  SciMsg pSci, %SCI_AUTOCSETIGNORECASE, %TRUE, 0          ' Ignore case
  ' Count the number of child windows
  p = 0
  p1 = GETWINDOW( g_hWndClient, %GW_CHILD )
  DO WHILE p1 <> 0
    p1 = GETWINDOW( p1, %GW_HWNDNEXT )
    INCR p
  LOOP
  DO WHILE p <> 0
    hWndActive = SendMessage(g_hWndClient, %WM_MDIGETACTIVE, 0, 0) 'MdiGetActive( hWndClient )
    pSci = SENDMESSAGE( GetEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
    IF ISFALSE pSci THEN EXIT DO
    ' Search whole document
    startPos = 0
    endPos = SciMsg( pSci, %SCI_GETTEXTLENGTH, 0, 0 )
    DO
      ' Set the find flags and the starting and ending positions
      SciMsg( pSci, %SCI_SETSEARCHFLAGS, findFlags, 0 )
      SciMsg( pSci, %SCI_SETTARGETSTART, startPos, 0 )
      SciMsg( pSci, %SCI_SETTARGETEND, endPos, 0 )
      ' Search for the text to replace
      hr = SciMsg( pSci, %SCI_SEARCHINTARGET, LEN( szFindText ), VARPTR( szFindText ))
      ' Increase the starting position
      startPos = hr + 1
      ' If hr = -1 there are no more text to scan
      IF CINT( hr ) = - 1 THEN
        EXIT DO
      ELSE
        x = SciMsg( pSci, %SCI_WORDSTARTPOSITION, hr, %TRUE )
        y = SciMsg( pSci, %SCI_WORDENDPOSITION, hr, %TRUE )
        buffer = SPACE$( y - x + 1 )
        txtrg.chrg.cpMin = x
        txtrg.chrg.cpMax = y
        txtrg.lpstrText = STRPTR( buffer )
        SciMsg pSci, %SCI_GETTEXTRANGE, 0, BYVAL VARPTR( txtrg )
        buffer = LEFT$( buffer, LEN( buffer ) - 1 )
        IF UCASE$( buffer ) <> UCASE$( szFindText ) THEN                ' Skip our own word
          ARRAY SCAN WordList( ), COLLATE UCASE, = buffer, TO iPos
          IF ISFALSE iPos THEN
            INCR numItems
            REDIM PRESERVE WordList( numItems - 1 )
            WordList( UBOUND( WordList )) = buffer
          END IF
        END IF
      END IF
    LOOP
    'MdiNext g_hWndClient, hWndActive, 0               ' Set focus to next child window
    SendMessage(g_hWndClient, %WM_MDINEXT, hWndActive, 0)
    DECR p          ' Decrement window counter
  LOOP
  ' Include PB keywords in the autocomplete list
   p2 = 1
   DO
     p = INSTR(p2, strPBKeyWords, " " & LCASE$(szFindText))
     IF p = 0 THEN EXIT DO
     p2 = INSTR(p + 1, strPBKeyWords, " ")
     IF p2 = 0 THEN EXIT DO
     buffer = MID$(strPBKeyWords, p + 1, p2 - p - 1)
     ARRAY SCAN WordList(), COLLATE UCASE, = buffer, TO iPos
     IF ISFALSE iPos THEN
       INCR numItems
       REDIM PRESERVE WordList(numItems - 1)
       WordList(UBOUND(WordList)) = UCASE$(buffer)
     END IF
   LOOP
  ' Include the function names of the Codetips in the autocomplete list
  IF ISTRUE hCodetipsFile THEN
    CodetipRecord = trm_GetEqualOrGreater( hCodetipsFile, 1, BYCOPY szFindText )
    DO
      IF CodetipRecord = "" THEN EXIT DO
      buffer = TRIM$( LEFT$( CodetipRecord, 64 ))
      IF UCASE$( LEFT$( buffer, LEN( szFindText ))) <> UCASE$( szFindText ) THEN EXIT DO
      ARRAY SCAN WordList( ), COLLATE UCASE, = buffer, TO iPos
      IF ISFALSE iPos THEN
        INCR numItems
        REDIM PRESERVE WordList( numItems - 1 )
        WordList( UBOUND( WordList )) = buffer
      END IF
      CodetipRecord = trm_GetNext( hCodetipsFile )
    LOOP
  END IF
  IF ISTRUE numItems THEN
    ARRAY SORT WordList( ), COLLATE UCASE
    strWordList = JOIN$( WordList( ), $SPC )
    pSci = oldpSci
    SciMsg pSci, %SCI_AUTOCSHOW, LEN( szFindText ), STRPTR( strWordList )
    FUNCTION = %TRUE
  END IF
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' Remove HTML from text - even handles any eventual JAVA scripts..
' Based in code posted by Borje Hagsten: http://www.powerbasic.com/support/forums/Forum6/HTML/001550.html
' *********************************************************************************************
FUNCTION RemoveHtml( BYVAL Buf AS STRING ) AS STRING
  LOCAL XWord AS STRING
  LOCAL BlockSize AS LONG
  LOCAL Ac AS LONG
  LOCAL Bc AS LONG
  LOCAL IsJava AS INTEGER
  LOCAL Letter AS BYTE PTR
  LOCAL AcLetter AS BYTE PTR
  LOCAL tmpLetter AS BYTE PTR
  LOCAL ckLet1 AS BYTE PTR
  LOCAL p AS LONG
  ' Skip all the header and title information
  p = INSTR( Buf, "<!--StartFragment" )
  IF p THEN Buf = MID$( Buf, p )
  'some necessary adjustments
  REPLACE "<p>" WITH $CRLF IN Buf                 'line feeds
  REPLACE "<P>" WITH $CRLF IN Buf
  REPLACE "<br>" WITH $CRLF IN Buf
  REPLACE "<BR>" WITH $CRLF IN Buf
  REPLACE "& n b s p ;" WITH " " IN Buf           '<- NOTE! Trim away spaces in first word, Forum auto-converts it, so I had to split it up
  REPLACE "&NBSP;" WITH " " IN Buf                'non-breaking spaces
  REPLACE "&nbsp;" WITH " " IN Buf
  Letter = STRPTR( Buf )
  BlockSize = LEN( Buf )
  XWord = SPACE$( BlockSize )
  AcLetter = STRPTR( XWord )
  DO WHILE @Letter > 0
    IF @Letter = 60 THEN            ' "<"
      tmpLetter = Letter
      Bc = Ac
      ckLet1 = Letter
      IsJava = 0              'javascript = <!- and ->
      IF @ckLet1 > 0 AND Ac < BlockSize THEN          ' is it Java?
        INCR ckLet1             'check for <!
        IF @ckLet1 > 0 AND @ckLet1 = 33 AND Ac < BlockSize - 1 THEN
          INCR ckLet1             'check for <!-  = prob. java
          IF @ckLet1 > 0 AND @ckLet1 = 45 THEN IsJava = 1
        END IF
      END IF
      DO WHILE @Letter > 0
        INCR Letter
        INCR Ac
        IF IsJava = 0 AND @Letter = 60 OR @Letter = 0 THEN              'if < or end
          Letter = tmpLetter
          Ac = Bc                 'start over
          @AcLetter = @Letter
          INCR AcLetter
          EXIT DO
        END IF
        IF IsJava = 1 THEN              'end javascript = ->
          IF @Letter = 62 AND Ac > 1 THEN
            ckLet1 = Letter
            DECR ckLet1
            IF @ckLet1 = 45 THEN EXIT DO            '-> end java
          END IF
        ELSE
          IF @Letter = 62 THEN EXIT DO            '> end html
        END IF
      LOOP
      INCR Letter : INCR Ac
    ELSE
      INCR Letter : INCR Ac
      INCR AcLetter
    END IF
    @AcLetter = @Letter
  LOOP
  @AcLetter = 0           'cut the text here..
  XWord = LEFT$( XWord, INSTR( XWORD, CHR$( 0 )) - 1 )
  REPLACE "& g t ;" WITH ">" IN XWord             '<- NOTE! Trim away spaces in first word, Forum auto-converts it, so I had to split it up
  REPLACE "& l t ;" WITH "<" IN XWord             '<- found in a few places, so..
  REPLACE "&gt;" WITH ">" IN XWord
  REPLACE "&lt;" WITH "<" IN XWord
  REPLACE "&amp;" WITH "&" IN XWord
  REPLACE "&AMP;" WITH "&" IN XWord
  XWord = LTRIM$( XWord, ANY CHR$( 13, 10 ))
  XWord = RTRIM$( XWord, ANY CHR$( 32, 13, 10 ))
  XWord = XWord & $LF
  FUNCTION = XWord
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
' BAS to colored HTM converter - Based ob Borje's WebCoder utility.
' Borje's standalone utility uses an indexed array of uppercase keywords to find them faster.
' I use INSTR to avoid to duplicate the list of PB keywords.
' *********************************************************************************************
SUB CvBasToHtml
  LOCAL ii AS LONG
  LOCAL stopPos AS LONG
  LOCAL Result AS LONG
  LOCAL wFlag AS LONG
  LOCAL remFlag AS LONG
  LOCAL dqFlag AS LONG
  LOCAL isRem AS LONG
  LOCAL pLet AS BYTE PTR
  LOCAL pLet2 AS BYTE PTR
  LOCAL tmpWord AS STRING
  LOCAL outBuf AS STRING
  LOCAL lCaseBuf AS STRING
  LOCAL htmlPrefix AS STRING
  LOCAL htmlPostfix AS STRING
  LOCAL greenStr AS STRING
  LOCAL blueStr AS STRING
  LOCAL redStr AS STRING
  LOCAL PBFstr AS STRING
  LOCAL endBlock AS STRING
  LOCAL endBlue AS STRING
  LOCAL endGreen AS STRING
  LOCAL txt AS STRING
  LOCAL hMdi AS DWORD
  LOCAL pSci AS DWORD
  LOCAL szKey AS ASCIIZ * 255
  LOCAL szValue AS ASCIIZ * 255
  LOCAL nTab AS LONG
  IF ISFALSE GetEdit THEN EXIT SUB
  txt = SPACE$( GETWINDOWTEXTLENGTH( GetEdit ) + 1 )
  GETWINDOWTEXT GetEdit, BYVAL STRPTR( txt ), LEN( txt )
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_APPSTARTING )            ' Show hourglass mouse
  htmlPrefix = "<HTML>" + $CRLF + _
          "<HEAD>" + $CRLF + _
          "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=iso-8859-1"">" + $CRLF + _
          "<META NAME=""GENERATOR"" CONTENT=""WebCoder 1.0"">" + $CRLF + _
          "<TITLE>WebCoder result</TITLE>" + $CRLF + _
          "</HEAD>" + $CRLF + $CRLF + _
          "<BODY BGCOLOR=""#FFFFFF"">" + $CRLF + _
          "<PRE>" + $CRLF
  htmlPostfix = "</PRE>" + $CRLF + "</BODY>" + $CRLF + "</HTML>" + $CRLF
  greenStr = "<FONT COLOR=""#008000"">"
  blueStr = "<FONT COLOR=""#0000FF"">"
  PBFStr = "<FONT COLOR=""#C06400"">"
  redStr = "<FONT COLOR=""#FF0000"">"
  endBlock = "</FONT>"
  endBlue = "</FONT2>"
  endGreen = "</FONT3>"
  REPLACE "<" WITH "&lt;" IN txt
  txt = txt + " "                 'add a space to ensure last word will be checked if nothing follows it
  OutBuf = STRING$( MAX& ( 1000, 5 * LEN( txt )), 0 )             '5 times bigger mem for result should be enough
  lCaseBuf = LCASE$( txt )                'use lowercase string for compare
  pLet = STRPTR( txt )            'pointer to global string (input)
  pLet2 = STRPTR( OutBuf )                'pointer to output buffer
  FOR ii = 1 TO LEN( txt )
    SELECT CASE @pLet               'the characters we need to inlude in a word
      CASE 65 TO 90, 97 TO 122, 35 TO 38, 48 TO 57, 63, 95
        IF wFlag = 0 AND remFlag = 0 AND dqFlag = 0 THEN
          wFlag = 1 : stopPos = ii
        END IF
      CASE 34                 ' double quote -> "
        IF dqFlag = 0 AND remFlag = 0 THEN              'if start of string literal
          POKE$ pLet2, redStr             'poke RTF code into output string
          pLet2 = pLet2 + 22              'and move pointer forward
          dqFlag = 1 : wFlag = 0          'set flags - since now inside DQ, wordflag is off
        ELSEIF dqFlag = 1 THEN          'should be end of DQ block
          @pLet2 = @pLet          'set value in output string
          INCR pLet2              'move one character ahead
          POKE$ pLet2, endBlock           'poke RTF end block string into output
          pLet2 = pLet2 + 7               'and move pointer forward
          dqFlag = 3              'end of DQ - set DQ flag
        END IF
      CASE 59                 ' asm uncomment character -> ;
        IF remFlag = 0 AND dqFlag = 2 THEN
          POKE$ pLet2, endBlock           'poke RTF end block string into output
          pLet2 = pLet2 + 7               'and move pointer forward
          POKE$ pLet2, greenStr
          pLet2 = pLet2 + 22
          remFlag = 1 : wFlag = 0
        END IF
      CASE 39                 ' uncomment character -> '
        IF remFlag = 0 AND dqFlag <> 1 THEN
          IF dqFlag = 2 THEN
            POKE$ pLet2, endBlock           'poke RTF end block string into output
            pLet2 = pLet2 + 7               'and move pointer forward
          END IF
          POKE$ pLet2, greenStr
          pLet2 = pLet2 + 22
          remFlag = 1 : wFlag = 0 : isRem = 1
        END IF
      CASE 33                 ' asm character -> !
        IF remFlag = 0 AND dqFlag = 0 THEN
          POKE$ pLet2, redStr
          pLet2 = pLet2 + 22
          dqFlag = 2 : wFlag = 0
        END IF
      CASE ELSE               'word is ready
        IF @pLet = 13 THEN              'if CRLF - end of line
          IF remFlag OR dqFlag THEN               'in rem or asm
            IF isRem = 0 THEN
              POKE$ pLet2, endBlock
              pLet2 = pLet2 + 7
            ELSE
              POKE$ pLet2, endGreen
              pLet2 = pLet2 + 8
            END IF
            remFlag = 0 : wFlag = 0 : dqFlag = 0 : isRem = 0                'reset flags
          END IF
        END IF
        IF wFlag = 1 THEN               'if we have a word
          tmpWord = MID$( lCaseBuf, stopPos, ii - stopPos )               'Get word
          tmpWord = LCASE$( tmpWord )
          Result = INSTR( strPBKeyWords, " " & tmpWord & " " )
          tmpWord = UCASE$( tmpWord )             ' Convert to uppercase
          IF tmpWord = "REM" THEN                 'extra for REM keyword
            pLet2 = pLet2 - 3               'set position to start of word
            POKE$ pLet2, greenStr
            pLet2 = pLet2 + 22
            POKE$ pLet2, tmpWord
            pLet2 = pLet2 + 3
            remFlag = 1 : isRem = 1
          ELSEIF tmpWord = "#PBFORMS" THEN                'extra for #PBFORMS statement
            pLet2 = pLet2 - 8               'set position to start of word
            POKE$ pLet2, PBFstr
            pLet2 = pLet2 + 22
            POKE$ pLet2, tmpWord
            pLet2 = pLet2 + 8
            remFlag = 1
          ELSEIF tmpWord = "ASM" THEN             'extra for ASM keyword
            pLet2 = pLet2 - 3               'set position to start of word
            POKE$ pLet2, RedStr
            pLet2 = pLet2 + 22
            POKE$ pLet2, tmpWord
            pLet2 = pLet2 + 3
            dqFlag = 2
          ELSE
            IF Result THEN          'if match was found, it's a PB keyword
              pLet2 = pLet2 - LEN( tmpWord )          'set position to start of word
              POKE$ pLet2, blueStr            'and poke RTF string for blue color into output string
              pLet2 = pLet2 + 22              'move pointer ahead
              POKE$ pLet2, tmpWord            'poke the word into output string
              pLet2 = pLet2 + LEN( tmpWord )          'move pointer ahead
              POKE$ pLet2, endBlue            'and finally poke RTF end block string into output-
              pLet2 = pLet2 + 8               'move pointer ahead
              Result = 0              'and reset result
            END IF
          END IF
          wFlag = 0
        END IF
    END SELECT
    IF dqFlag <> 3 THEN             'if not handled matching double-quote
      @pLet2 = @pLet          'copy original character to output
      INCR pLet2              'and increase pos in output
    ELSE
      dqFlag = 0              'else reset DQ flag
    END IF
    INCR pLet               'move ahead to next character
  NEXT ii
  OutBuf = EXTRACT$( OutBuf, CHR$( 32, 0 ))               'extract result (and remove the added space)
  REPLACE endBlue + " " + blueStr WITH " " IN OutBuf              'Trim size: If keywords are next to each other,
  REPLACE endBlue WITH endBlock IN OutBuf                 'replace remaining blue endblocks with proper HTML
  REPLACE endGreen + $CRLF + GreenStr WITH $CRLF IN OutBuf                'remmed out lines..
  REPLACE endGreen WITH endBlock IN OutBuf                'replace remaining green endblocks with proper HTML
  OutBuf = htmlPrefix + OutBuf + htmlPostfix
  SETCURSOR LOADCURSOR ( BYVAL %NULL, BYVAL %IDC_ARROW )          ' Show standard mouse
  hMdi = CreateMdiChild($EDITCLASSNAME, g_hWndClient, "", IIF( fMaximize, %WS_MAXIMIZE, 0 ))
  IF ISFALSE hMdi THEN EXIT SUB
  SHOWWINDOW hMdi, %SW_SHOW
  ' Get a direct pointer for faster access
  pSci = SENDMESSAGE( GetEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT SUB
  ' Insert the text at the current position
  SciMsg pSci, %SCI_INSERTTEXT, 0, BYVAL STRPTR( OutBuf )
  ' Change the title of the window
  SETWINDOWTEXT hMdi, "Untitled" & FORMAT$( cUntitledFiles ) & ".html"
  ' Change the tab name
  nTab = TabCtrl_GetCurSel( g_hTabMdi )
  CALL SetTabName( nTab, "Untitled" & FORMAT$( cUntitledFiles ) & ".html" )
  ' Change the name in the array
  IF nTab < = UBOUND( gTabFilePaths ) THEN
    gTabFilePaths( nTab ) = "Untitled" & FORMAT$( cUntitledFiles ) & ".html"
  END IF
  ' Set styles to the default
  SciMsg pSci, %SCI_STYLECLEARALL, 0, 0
  ' Disable folding of the procedures and functions
  szKey = "fold" : szValue = "0"
  SciMsg pSci, %SCI_SETPROPERTY, VARPTR( szKey ), VARPTR( szValue )
  ' Change statusbar information
  ShowLinCol
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Function    : SED_FormatRegion
' Description : Formats the selected region removing/adding spaces where nessesary
' Code by Steve Miltiadous
' *********************************************************************************************
SUB SED_FormatRegion
  LOCAL startPos AS LONG          ' // Start of selected text
  LOCAL endPos AS LONG            ' // End of selected text
  LOCAL startLine AS LONG                 ' // Starting line
  LOCAL endLine AS LONG           ' // Ending line
  LOCAL LineLen AS LONG           ' // Length of the line
  LOCAL buffer AS STRING          ' // Buffer
  LOCAL i AS LONG                 ' // Loop counter
  LOCAL strFill AS STRING                 ' // Filler string
  LOCAL RemLine AS STRING                 ' // REMed line
  LOCAL lWithinString AS LONG             ' // Flag
  startPos = SENDMESSAGE( GetEdit, %SCI_GETSELECTIONSTART, 0, 0 )
  endPos = SENDMESSAGE( GetEdit, %SCI_GETSELECTIONEND, 0, 0 )
  startLine = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, startPos, 0 )
  endLine = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, endPos, 0 )
  lWithinString = %FALSE
  WHILE startLine < = endLine
    RemLine = ""
    LineLen = SENDMESSAGE( GetEdit, %SCI_LINELENGTH, startLine, 0 )                 ' length of the line
    buffer = SPACE$( LineLen )              ' size the buffer
    SENDMESSAGE( GetEdit, %SCI_GETLINE, startLine, STRPTR( buffer ))                ' get the text of the line
    IF LEN( REMOVE$( buffer, ANY CHR$( 13, 10, 32 ))) THEN
      FOR i = 1 TO LEN( buffer )
        IF INSTR( MID$( buffer, i, 1 ), ANY CHR$( 32, 9 )) = 0 THEN EXIT FOR
      NEXT
      strFill = MID$( buffer, 1, i - 1 )
      WHILE i < LEN( buffer )
        IF ASC( MID$( buffer, i, 1 )) = 34 THEN
          lWithinString = IIF( lWithinString = %TRUE, %FALSE, %TRUE )
        END IF
        IF lWithinString = %FALSE THEN
          SELECT CASE MID$( buffer, i, 1 )
            CASE ">"                'asc 62
              IF ASC( MID$( strFill, - 1 )) = 32 AND ( ASC( MID$( strFill, - 2 )) = 61 OR ASC( MID$( strFill, - 2 )) = 60 ) THEN
                strFill = MID$( strFill, 1, LEN( strFill ) - 1 ) + MID$( buffer, i, 1 ) + " "
              ELSE
                IF ASC( MID$( strFill, - 1 )) = 32 THEN
                  strFill = strFill + MID$( buffer, i, 1 ) + " "
                ELSE
                  strFill = strFill + " " + MID$( buffer, i, 1 ) + " "
                END IF
              END IF
            CASE "<"                'asc 60
              IF ASC( MID$( strFill, - 1 )) = 32 AND ( ASC( MID$( strFill, - 2 )) = 61 OR ASC( MID$( strFill, - 2 )) = 62 ) THEN
                strFill = MID$( strFill, 1, LEN( strFill ) - 1 ) + MID$( buffer, i, 1 ) + " "
              ELSE
                IF ASC( MID$( strFill, - 1 )) = 32 THEN
                  strFill = strFill + MID$( buffer, i, 1 ) + " "
                ELSE
                  strFill = strFill + " " + MID$( buffer, i, 1 ) + " "
                END IF
              END IF
            CASE "'"
              RemLine = REMOVE$( MID$( buffer, i ), $CRLF )
              EXIT LOOP
            CASE " "
              IF MID$( strFill, - 1 ) <> " " AND _
                      MID$( strFill, - 1 ) <> "(" THEN strFill = strFill + " "
            CASE "-"
              IF MID$( strFill, - 1 ) <> " " THEN
                IF MID$( strFill, - 2, 1 ) <> "=" THEN
                  strFill = strFill + " " + MID$( buffer, i, 1 ) + " "
                ELSE
                  strFill = strFill + " " + MID$( buffer, i, 1 )
                END IF
              ELSE
                IF MID$( strFill, - 2, 1 ) <> "=" THEN
                  strFill = strFill + MID$( buffer, i, 1 ) + " "
                ELSE
                  strFill = strFill + MID$( buffer, i, 1 )
                END IF
              END IF
            CASE "="
              SELECT CASE MID$( strFill, - 1 )
                CASE ">", "<" : strFill = strFill + MID$( buffer, i, 1 ) + " "
                CASE " "
                  IF ASC( MID$( strFill, - 2 )) = 60 OR ASC( MID$( strFill, - 2 )) = 62 THEN
                    strFill = MID$( strFill, 1, LEN( strFill ) - 1 ) + MID$( buffer, i, 1 ) + " "
                  ELSE
                    strFill = strFill + MID$( buffer, i, 1 ) + " "
                  END IF
                CASE ELSE
                  strFill = strFill + " " + MID$( buffer, i, 1 ) + " "
              END SELECT
              'IF MID$(strFill, -1) <> " " THEN
              '   strFill = strFill + " " + MID$(buffer, i, 1) + " "
              'ELSE
              '   strFill = strFill + MID$(buffer, i, 1) + " "
              'END IF
            CASE "+", "*", "/"
              IF MID$( strFill, - 1 ) <> " " THEN
                strFill = strFill + " " + MID$( buffer, i, 1 ) + " "
              ELSE
                strFill = strFill + MID$( buffer, i, 1 ) + " "
              END IF
            CASE ","
              IF MID$( strFill, - 1 ) = " " THEN
                strFill = MID$( strFill, 1, LEN( strFill ) - 1 ) + MID$( buffer, i, 1 )
              ELSE
                strFill = strFill + MID$( buffer, i, 1 ) + " "
              END IF
            CASE ")"
              IF MID$( strFill, - 1 ) = " " THEN
                strFill = MID$( strFill, 1, LEN( strFill ) - 1 ) + MID$( buffer, i, 1 )
              ELSE
                IF MID$( buffer, i + 1, 1 ) = "." THEN
                  strFill = strFill + MID$( buffer, i, 1 )
                ELSE
                  strFill = strFill + MID$( buffer, i, 1 ) + " "
                END IF
              END IF
            CASE ELSE
              strFill = strFill + MID$( buffer, i, 1 )
          END SELECT
        ELSE
          IF MID$( buffer, i, 1 ) = "_" THEN              ' line continuation
            IF MID$( strFill, - 1 ) <> " " THEN strFill = strFill + " "
          END IF
          strFill = strFill + MID$( buffer, i, 1 )
        END IF
        INCR i
      WEND
      strFill = REMOVE$( strFill, CHR$( 13 ))
      strFill = RTRIM$( strFill )             ' Remove trailing spaces
      IF LEN( RemLine ) > 0 THEN strFill = strFill + SPACE$(( i - 1 ) - LEN( strFill )) + RemLine
      strFill = SpacesToTabs( strFill )
      SENDMESSAGE( GetEdit, %SCI_GOTOPOS, SENDMESSAGE( GetEdit, %SCI_POSITIONFROMLINE, startLine, 0 ), 0 )
      SENDMESSAGE( GetEdit, %SCI_SETANCHOR, SENDMESSAGE( GetEdit, %SCI_GETLINEENDPOSITION, startLine, 0 ), 0 )
      SENDMESSAGE( GetEdit, %SCI_REPLACESEL, LEN( strFill ), STRPTR( strFill ))
    END IF
    INCR startLine
  WEND
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Function    : SED_TabulateRegion
' Description : Tabulate the selected region removing/adding spaces where nessesary
' Code by Steve Miltiadous
' *********************************************************************************************
SUB SED_TabulateRegion
  LOCAL startPos AS LONG          ' // Start of selected text
  LOCAL endPos AS LONG            ' // End of selected text
  LOCAL startLine AS LONG
  LOCAL startLineSave AS LONG
  LOCAL endLine AS LONG
  LOCAL LineLen AS LONG
  LOCAL buffer AS STRING
  LOCAL i AS LONG
  LOCAL j AS LONG
  LOCAL strFill AS STRING
  LOCAL TabSize AS LONG
  LOCAL nSpaces AS LONG
  LOCAL indent AS LONG
  LOCAL WordNumber AS LONG
  LOCAL arLine( ) AS STRING
  LOCAL arRemLine( ) AS STRING
  LOCAL arWords( ) AS STRING
  LOCAL arPLine( ) AS LONG
  LOCAL nCol AS LONG
  TabSize = SENDMESSAGE( GetEdit, %SCI_GETTABWIDTH, 0, 0 )
  startPos = SENDMESSAGE( GetEdit, %SCI_GETSELECTIONSTART, 0, 0 )
  endPos = SENDMESSAGE( GetEdit, %SCI_GETSELECTIONEND, 0, 0 )
  startLine = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, startPos, 0 )
  endLine = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, endPos, 0 )
  nCol = SENDMESSAGE( GetEdit, %SCI_GETCOLUMN, endPos, 0 )
  IF nCol = 0 AND endLine > startLine THEN endLine = endLine - 1
  startLineSave = startLine
  indent = 0
  WordNumber = 0
  indent = 10000
  WHILE startLine < = endLine
    strFill = ""
    LineLen = SENDMESSAGE( GetEdit, %SCI_LINELENGTH, startLine, 0 )                 ' length of the line
    buffer = SPACE$( LineLen )              ' size the buffer
    SENDMESSAGE( GetEdit, %SCI_GETLINE, startLine, STRPTR( buffer ))                ' get the text of the line
    nSpaces = 0
    IF LEN( REMOVE$( buffer, ANY CHR$( 13, 10, 32, 9 ))) THEN
      FOR i = 1 TO LEN( buffer )
        IF MID$( buffer, i, 1 ) <> $SPC THEN
          IF MID$( buffer, i, 1 ) = $TAB THEN
            nSpaces = nSpaces + TabSize
          ELSE
            EXIT FOR
          END IF
        ELSE
          INCR nSpaces
        END IF
      NEXT
      IF LEN( REMOVE$( buffer, $CRLF )) THEN
        IF indent > nSpaces THEN indent = nSpaces
      END IF
      REDIM PRESERVE arRemLine( UBOUND( arRemLine ) + 1 )
      arRemLine( UBOUND( arRemLine )) = ""
      WHILE i < LEN( buffer )
        SELECT CASE MID$( buffer, i, 1 )
          CASE "'"
            arRemLine( UBOUND( arRemLine )) = REMOVE$( MID$( buffer, i ), $CRLF )
            EXIT LOOP
          CASE " "
            IF MID$( strFill, - 1 ) <> " " AND _
                    IIF( i + 1 < LEN( buffer ), IIF( INSTR( MID$( buffer, i + 1, 1 ), ANY CHR$( 32, 13, 10 )) <> 0, %FALSE, %TRUE ), %FALSE ) AND _
                    MID$( strFill, - 1 ) <> "(" THEN strFill = strFill + " "
          CASE "-"
            IF MID$( strFill, - 1 ) <> " " THEN
              IF MID$( strFill, - 2, 1 ) <> "=" THEN
                strFill = strFill + " " + MID$( buffer, i, 1 ) + " "
              ELSE
                strFill = strFill + " " + MID$( buffer, i, 1 )
              END IF
            ELSE
              IF MID$( strFill, - 2, 1 ) <> "=" THEN
                strFill = strFill + MID$( buffer, i, 1 ) + " "
              ELSE
                strFill = strFill + MID$( buffer, i, 1 )
              END IF
            END IF
          CASE "=", "+", "*", "/"
            IF MID$( strFill, - 1 ) <> " " THEN
              strFill = strFill + " " + MID$( buffer, i, 1 ) + " "
            ELSE
              strFill = strFill + MID$( buffer, i, 1 ) + " "
            END IF
          CASE ",", ")"
            IF MID$( strFill, - 1 ) = " " THEN
              strFill = MID$( strFill, 1, LEN( strFill ) - 1 ) + MID$( buffer, i, 1 )
            ELSE
              strFill = strFill + MID$( buffer, i, 1 ) + " "
            END IF
          CASE ELSE
            strFill = strFill + MID$( buffer, i, 1 )
        END SELECT
        INCR i
      WEND
      strFill = REMOVE$( strFill, CHR$( 13 ))
      IF WordNumber < PARSECOUNT( strFill, " " ) THEN
        WordNumber = PARSECOUNT( strFill, " " )
        REDIM PRESERVE arPLine( WordNumber )
      END IF
      REDIM PRESERVE arLine( UBOUND( arLine ) + 1 )
      arLine( UBOUND( arLine )) = strFill
      FOR i = 1 TO PARSECOUNT( strFill, " " )
        IF arPLine( i - 1 ) < LEN( PARSE$( strFill, " ", i )) THEN arPLine( i - 1 ) = LEN( PARSE$( strFill, " ", i ))
      NEXT
    ELSE
      REDIM PRESERVE arLine( UBOUND( arLine ) + 1 )
      arLine( UBOUND( arLine )) = REMOVE$( buffer, CHR$( 13, 10 ))
      REDIM PRESERVE arRemLine( UBOUND( arRemLine ) + 1 )
      arRemLine( UBOUND( arRemLine )) = ""
    END IF
    INCR startLine
  WEND
  FOR i = 0 TO UBOUND( arLine )
    IF LEN( REMOVE$( arLine( i ), ANY CHR$( 13, 10, 32, 9 ))) THEN
      strFill = SPACE$( indent )
      REDIM arWords( PARSECOUNT( arLine( i ), ANY " " ))
      PARSE arLine( i ), arWords( ), ANY " "
      FOR j = 0 TO UBOUND( arWords )
        IF j < UBOUND( arWords ) THEN
          strFill = strFill + arWords( j ) + SPACE$( arPLine( j ) - LEN( arWords( j )) + 1 )
        ELSE
          IF arWords( j ) = $CRLF THEN
            IF MID$( strFill, - 1, 1 ) = " " THEN
              strFill = MID$( strFill, 1, LEN( strFill ) - 1 ) + arWords( j )
            ELSE
              strFill = strFill + arWords( j )
            END IF
          ELSE
            strFill = strFill + arWords( j )
          END IF
        END IF
      NEXT
      nSpaces = 0
      IF i < = UBOUND( arRemLine ) THEN
        FOR j = 0 TO UBOUND( arPLine )
          nSpaces = nSpaces + arPLine( j ) + 1
        NEXT
        strFill = strFill + SPACE$(( nSpaces + indent ) - LEN( strFill )) + arRemLine( i )
        strFill = RTRIM$( strFill )             ' Trim trailing spaces
      END IF
    ELSE
      ITERATE FOR
    END IF
    strFill = SpacesToTabs( strFill )
    SENDMESSAGE( GetEdit, %SCI_GOTOPOS, SENDMESSAGE( GetEdit, %SCI_POSITIONFROMLINE, startLineSave + i, 0 ), 0 )
    SENDMESSAGE( GetEdit, %SCI_SETANCHOR, SENDMESSAGE( GetEdit, %SCI_GETLINEENDPOSITION, startLineSave + i, 0 ), 0 )
    SENDMESSAGE( GetEdit, %SCI_REPLACESEL, LEN( strFill ), STRPTR( strFill ))
  NEXT
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Function    : SED_PrintDoc
' Description : Prints the SciEditor Document using the DDoc Print Engine.
'               PreviewFlag = %TRUE for Print Preview.
'               PreviewFlag = %FALSE for direct printing showing the Print dialog
'               PreviewFlag = -1 for direct printing without showing the Print dialog
' *********************************************************************************************
FUNCTION SED_PrintDoc( szFilePath AS ASCIIZ, BYVAL PreviewFlag AS LONG ) AS LONG
  LOCAL hFileDoc    AS INTEGER                         ' // ddoc file handle
  LOCAL lh          AS SINGLE                          ' // Line height
  LOCAL PrintBuffer AS STRING                          ' // Print buffer
  LOCAL szText      AS ASCIIZ                  * 255   ' // General purpose variable
  LOCAL PrnOpt      AS PrinterSetupOptionsType         ' // Printer setup options
  LOCAL PaperSize   AS LONG                            ' // Paper size
  LOCAL PaperHeight AS SINGLE                          ' // Paper height
  LOCAL PaperWidth  AS SINGLE                          ' // Paper with
  LOCAL FontSize    AS LONG                            ' // Font size
  LOCAL chPrinted   AS LONG                            ' // Characters printed
  LOCAL bufferLen   AS LONG                            ' // Length of the buffer
  LOCAL nPage       AS LONG                            ' // Page number
  LOCAL FontStyle   AS LONG                            ' // Font style
  LOCAL startPos    AS LONG                            ' // Start of selected text
  LOCAL endPos      AS LONG                            ' // End of selected text
  LOCAL EdOpt       AS EditorOptionsType               ' // Editor options
  IF ISFALSE GetEdit THEN EXIT FUNCTION
  szText = "SED_TEMP.DOC"                 ' Document name for ddoc
  GetPrinterSetupOptions PrnOpt           ' Printer options
  GetEditorOptions EdOpt          ' Editor options
  ' Amount of space I allow between lines.
  lh = PrnOpt.LineHeight
  IF lh < = 0 THEN lh = 0.17!
  ' Establish the paper sizes
  IF PrnOpt.PaperSize = 1 THEN
    PaperSize = %DDOC_PAPER_LETTER
    IF PrnOpt.Orientation = 1 THEN
      PaperWidth = 11
      PaperHeight = 8.5
    ELSE
      PaperWidth = 8.5
      PaperHeight = 11
    END IF
  ELSEIF PrnOpt.PaperSize = 2 THEN
    PaperSize = %DDOC_PAPER_A4
    IF PrnOpt.Orientation = 1 THEN
      PaperWidth = 11.69
      PaperHeight = 8.27
    ELSE
      PaperWidth = 8.27
      PaperHeight = 11.69
    END IF
  ELSEIF PrnOpt.PaperSize = 3 THEN
    PaperSize = %DDOC_PAPER_LEGAL
    IF PrnOpt.Orientation = 1 THEN
      PaperWidth = 14
      PaperHeight = 8.5
    ELSE
      PaperWidth = 8.5
      PaperHeight = 14
    END IF
  ELSEIF PrnOpt.PaperSize = 4 THEN
    PaperSize = %DDOC_PAPER_FOLIO
    IF PrnOpt.Orientation = 1 THEN
      PaperWidth = 13
      PaperHeight = 8.5
    ELSE
      PaperWidth = 8.5
      PaperHeight = 13
    END IF
  END IF
  ' Establish the font size
  SELECT CASE PrnOpt.FontSize
    CASE 1 : FontSize = 8
    CASE 2 : FontSize = 9
    CASE 3 : FontSize = 10
    CASE 4 : FontSize = 11
    CASE 5 : FontSize = 12
    CASE 6 : FontSize = 14
  END SELECT
  ' Get the text to print. If there is selected text, print it
  ' otherwise, print all the document
  startPos = SENDMESSAGE( GetEdit, %SCI_GETSELECTIONSTART, 0, 0 )
  endPos = SENDMESSAGE( GetEdit, %SCI_GETSELECTIONEND, 0, 0 )
  IF endPos > startPos THEN
    PrintBuffer = SPACE$( endPos - startPos + 1 )
    SENDMESSAGE( GetEdit, %SCI_GETSELTEXT, 0, STRPTR( PrintBuffer ))
    PrintBuffer = REMOVE$( PrintBuffer, CHR$( 0 ))
  ELSE
    PrintBuffer = SPACE$( GETWINDOWTEXTLENGTH( GetEdit ) + 1 )
    SENDMESSAGE GetEdit, %SCI_GETTEXT, BYVAL LEN( PrintBuffer ), BYVAL STRPTR( PrintBuffer )
    PrintBuffer = REMOVE$( PrintBuffer, CHR$( 0 ))
  END IF
  bufferLen = LEN( PrintBuffer )
  IF bufferLen = 0 THEN EXIT FUNCTION
  REPLACE $TAB WITH SPACE$( EdOpt.tabSize ) IN PrintBuffer
  ' Start the ddoc document
  IF PreviewFlag = 0 OR PreviewFlag = - 1 THEN
    hFileDoc = dpStartDoc( 0, szText, "", %DDOC_INCH, _
            PaperSize, PrnOpt.Orientation, PrnOpt.PaperBin, %DDOC_ZOOMFIT )
  ELSE
    hFileDoc = dpStartDoc( 0, szText, "", %DDOC_INCH, _
            PaperSize, PrnOpt.Orientation, PrnOpt.PaperBin, _
            %DDOC_ZOOM100 OR %DDOC_VIEWBUILD OR %DDOC_ALLOWFAX OR %DDOC_ALLOWMAPI _
            OR %DDOC_ALLOWSAVE OR %DDOC_ALLOWSMTP )
  END IF
  IF hFileDoc < 1 THEN
    MSGBOX "Error:" + STR$( hFileDoc ) + " printing the file"
    EXIT FUNCTION
  END IF
  ' Establish the font.
  dpFont hFileDoc, %DDOC_FONTNORMAL, FontSize, %BLACK, PrnOpt.FontName
  DO
    ' Increment the page number
    INCR nPage
    ' Create a bookmark
    dpBookmark hFileDoc, "Page " & FORMAT$( nPage )
    ' Select font bold
    dpFont hFileDoc, %DDOC_FONTBOLD, FontSize, %BLACK, PrnOpt.FontName
    ' Print the header
    dpText hFileDoc, PrnOpt.MarginLeft, PrnOpt.MarginTop, %DDOC_LEFT, "File: " & szFilePath
    dpText hFileDoc, PrnOpt.MarginLeft, PrnOpt.MarginTop + lh, %DDOC_LEFT, "Date: " & DATE$ & " - Time: " & TIME$ & " - Page: " & FORMAT$( nPage )
    ' Select the font
    FontStyle = %DDOC_FONTNORMAL
    IF PrnOpt.FontBold THEN FontStyle = FontStyle OR %DDOC_FONTBOLD
    IF PrnOpt.FontItalic THEN FontStyle = FontStyle OR %DDOC_FONTITALIC
    IF PrnOpt.FontUnderline THEN FontStyle = FontStyle OR %DDOC_FONTUNDERLINE
    dpFont hFileDoc, FontStyle, FontSize, %BLACK, PrnOpt.FontName
    ' Print wrapped text taking into account the orientation (portrait or landscape)
    chPrinted = dpWrapText( hFileDoc, PrnOpt.MarginLeft, PrnOpt.MarginTop + lh * 3, _
            PaperWidth - PrnOpt.MarginRight, PaperHeight - PrnOpt.MarginBottom, _
            lh, BYVAL STRPTR( PrintBuffer ))
    ' Exit if not characters printed
    IF chPrinted = 0 THEN EXIT DO
    ' Show the progress bar
    dpSetProgress hFileDoc, nPage
    ' Select font bold
    dpFont hFileDoc, %DDOC_FONTBOLD, FontSize, %BLACK, PrnOpt.FontName
    ' Strip the characters already printed
    bufferLen = bufferLen - chPrinted
    IF bufferLen < 1 THEN EXIT DO
    PrintBuffer = MID$( PrintBuffer, chPrinted )
    IF LEN( PrintBuffer ) = 0 THEN EXIT DO
    IF PrintBuffer = $CRLF OR PrintBuffer = $CR OR PrintBuffer = $LF THEN EXIT DO
    ' Start a new page
    dpNewPage hFileDoc, PaperSize, PrnOpt.Orientation, PrnOpt.PaperBin
    '- See if the user aborted this operation
    IF dpCheckAbort( hFileDoc ) THEN
      ' User aborted document creation
      dpEndDoc hFileDoc, %DDOC_END_CLOSE + %DDOC_END_DELETE
      IF PreviewFlag = 0 THEN
        dpEndDoc hFileDoc, %DDOC_END_PRINT + %DDOC_END_DELETE
      ELSEIF PreviewFlag = 1 THEN
        dpEndDoc hFileDoc, %DDOC_END_VIEW + %DDOC_END_DELETE
      ELSEIF PreviewFlag = - 1 THEN
        dpEndDoc hFileDoc, %DDOC_END_PRINT_NODIALOG + %DDOC_END_DELETE
      END IF
      EXIT FUNCTION
    END IF
  LOOP
  ' Show or print the document and tell ddoc to kill the temporary file on exit
  IF PreviewFlag = 0 THEN
    dpEndDoc hFileDoc, %DDOC_END_PRINT + %DDOC_END_DELETE
  ELSEIF PreviewFlag = 1 THEN
    dpEndDoc hFileDoc, %DDOC_END_VIEW + %DDOC_END_DELETE
  ELSEIF PreviewFlag = - 1 THEN
    dpEndDoc hFileDoc, %DDOC_END_PRINT_NODIALOG + %DDOC_END_DELETE
  END IF
  FUNCTION = %TRUE
END FUNCTION
SUB GetEditorOptions( EdOpt AS EditorOptionsType )
  LOCAL rs AS STRING
  LOCAL IniFile AS STRING
  IniFile=EXE.PATH$ & "config.ini"
  rs = IniRead( IniFile, "Editor options", "UseTabs", "" )
  IF LEN( rs ) THEN EdOpt.UseTabs = VAL( rs ) ELSE EdOpt.UseTabs = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "TabSize", "" )
  IF LEN( rs ) THEN EdOpt.TabSize = VAL( rs ) ELSE EdOpt.TabSize = 3
  rs = IniRead( IniFile, "Editor options", "AutoIndent", "" )
  IF LEN( rs ) THEN EdOpt.AutoIndent = VAL( rs ) ELSE EdOpt.AutoIndent = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "IndentSize", "" )
  IF LEN( rs ) THEN EdOpt.IndentSize = VAL( rs ) ELSE EdOpt.IndentSize = 3
  rs = IniRead( IniFile, "Editor options", "LineNumbers", "" )
  IF LEN( rs ) THEN EdOpt.LineNumbers = VAL( rs ) ELSE EdOpt.LineNumbers = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "LineNumbersWidth", "" )
  IF LEN( rs ) THEN EdOpt.LineNumbersWidth = VAL( rs ) ELSE EdOpt.LineNumbersWidth = 50
  rs = IniRead( IniFile, "Editor options", "Margin", "" )
  IF LEN( rs ) THEN EdOpt.Margin = VAL( rs ) ELSE EdOpt.Margin = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "MarginWidth", "" )
  IF LEN( rs ) THEN EdOpt.MarginWidth = VAL( rs ) ELSE EdOpt.MarginWidth = 20
  rs = IniRead( IniFile, "Editor options", "EdgeColumn", "" )
  IF LEN( rs ) THEN EdOpt.EdgeColumn = VAL( rs ) ELSE EdOpt.EdgeColumn = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "EdgeWidth", "" )
  IF LEN( rs ) THEN EdOpt.EdgeWidth = VAL( rs ) ELSE EdOpt.EdgeWidth = 255
  rs = IniRead( IniFile, "Editor options", "IndentGuides", "" )
  IF LEN( rs ) THEN EdOpt.IndentGuides = VAL( rs ) ELSE EdOpt.IndentGuides = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "Magnification", "" )
  EdOpt.Magnification = VAL( rs )
  rs = IniRead( IniFile, "Editor options", "WhiteSpace", "" )
  IF LEN( rs ) THEN EdOpt.WhiteSpace = VAL( rs ) ELSE EdOpt.WhiteSpace = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "EndOfLine", "" )
  IF LEN( rs ) THEN EdOpt.EndOfLine = VAL( rs ) ELSE EdOpt.EndOfLine = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "SyntaxHighlighting", "" )
  IF LEN( rs ) THEN EdOpt.SyntaxHighlighting = VAL( rs ) ELSE EdOpt.SyntaxHighlighting = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "CodeTips", "" )
  EdOpt.CodeTips = VAL( rs )
  rs = IniRead( IniFile, "Editor options", "MaximizeMainWindow", "" )
  EdOpt.MaximizeMainWindow = VAL( rs )
  rs = IniRead( IniFile, "Editor options", "MaximizeEditWindows", "" )
  IF LEN( rs ) THEN EdOpt.MaximizeEditWindows = VAL( rs ) ELSE EdOpt.MaximizeEditWindows = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "AskBeforeExit", "" )
  IF LEN( rs ) THEN EdOpt.AskBeforeExit = VAL( rs ) ELSE EdOpt.AskBeforeExit = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "AllowMultipleInstances", "" )
  IF LEN( rs ) THEN EdOpt.AllowMultipleInstances = VAL( rs ) ELSE EdOpt.AllowMultipleInstances = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "DefaultCase", "" )
  IF LEN( rs ) THEN EdOpt.DefaultCase = VAL( rs ) ELSE EdOpt.DefaultCase = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "KeywordCase", "" )
  IF LEN( rs ) THEN EdOpt.KeywordCase = VAL( rs ) ELSE EdOpt.KeywordCase = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "ConstructAutocompletion", "" )
  IF LEN( rs ) THEN EdOpt.ConstructAutocompletion = VAL( rs ) ELSE EdOpt.ConstructAutocompletion = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "AutocompleteTypes", "" )
  IF LEN( rs ) THEN EdOpt.AutocompleteTypes = VAL( rs ) ELSE EdOpt.AutocompleteTypes = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "TrimTrailingBlanks", "" )
  IF LEN( rs ) THEN EdOpt.TrimTrailingBlanks = VAL( rs ) ELSE EdOpt.TrimTrailingBlanks = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "ShowProcedureName", "" )
  IF LEN( rs ) THEN EdOpt.ShowProcedureName = VAL( rs ) ELSE EdOpt.ShowProcedureName = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "ShowCaretLine", "" )
  IF LEN( rs ) THEN EdOpt.ShowCaretLine = VAL( rs ) ELSE EdOpt.ShowCaretLine = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "StartInLastFolder", "" )
  IF LEN( rs ) THEN EdOpt.StartInLastFolder = VAL( rs ) ELSE EdOpt.StartInLastFolder = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "LastFolder", "" )
  IF LEN( rs ) THEN EdOpt.LastFolder = rs ELSE EdOpt.LastFolder = CURDIR$
  rs = IniRead( IniFile, "Editor options", "ReloadFilesAtStartup", "" )
  IF LEN( rs ) THEN EdOpt.ReloadFilesAtStartup = VAL( rs ) ELSE EdOpt.ReloadFilesAtStartup = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "BackupEditorFiles", "" )
  IF LEN( rs ) THEN EdOpt.BackupEditorFiles = VAL( rs ) ELSE EdOpt.BackupEditorFiles = %BST_CHECKED
  rs = IniRead( IniFile, "Editor options", "AttachResources", "" )
  IF LEN( rs ) THEN EdOpt.AttachResources = VAL( rs ) ELSE EdOpt.AttachResources = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "CallPostProcessor", "" )
  IF LEN( rs ) THEN EdOpt.CallPostProcessor = VAL( rs ) ELSE EdOpt.CallPostProcessor = %BST_UNCHECKED
  rs = IniRead( IniFile, "Editor options", "DdocPrinting", "" )
  IF LEN( rs ) THEN EdOpt.DdocPrinting = VAL( rs ) ELSE EdOpt.DdocPrinting = %BST_UNCHECKED
  IF EdOpt.MaximizeEditWindows = %BST_CHECKED THEN fMaximize = %TRUE ELSE fMaximize = %FALSE
  IF EdOpt.ConstructAutocompletion = %BST_CHECKED THEN ConstructAutocompletion = %TRUE ELSE ConstructAutocompletion = %FALSE
  IF EdOpt.AutocompleteTypes = %BST_CHECKED THEN AutoCompleteTypes = %TRUE ELSE AutoCompleteTypes = %FALSE
  IF EdOpt.TrimTrailingBlanks = %BST_CHECKED THEN TrimTrailingBlanks = %TRUE ELSE TrimTrailingBlanks = %FALSE
  IF EdOpt.ShowProcedureName = %BST_CHECKED THEN ShowProcedureName = %TRUE ELSE ShowProcedureName = %FALSE
  'CheckMenuOptions EdOpt
END SUB
SUB GetCompilerOptions( CpOpt AS CompilerOptionsType )
  LOCAL IniFile AS STRING
  IniFile=EXE.PATH$ & "config.ini"
  CpOpt.DefaultCompiler = VAL( IniRead( IniFile, "Compiler options", "DefaultCompiler", "" ))
  CpOpt.PBWINPath = IniRead( IniFile, "Compiler options", "PBWINPath", "" )
  CpOpt.PBWINIncPath = IniRead( IniFile, "Compiler options", "PBWINIncPath", "" )
  CpOpt.PBCCPath = IniRead( IniFile, "Compiler options", "PBCCPath", "" )
  CpOpt.PBCCIncPath = IniRead( IniFile, "Compiler options", "PBCCIncPath", "" )
  CpOpt.RCPath = IniRead( IniFile, "Compiler options", "RCPath", "" )
  CpOpt.PBRESPath = IniRead( IniFile, "Compiler options", "PBRESPath", "" )
  CpOpt.RCIncPath = IniRead( IniFile, "Compiler options", "RCIncPath", "" )
  CpOpt.DisplayResults = VAL( IniRead( IniFile, "Compiler options", "DisplayResults", "" ))
  CpOpt.DeleteLogFile = VAL( IniRead( IniFile, "Compiler options", "DeleteLogFile", "" ))
  CpOpt.BeepOnCompletion = VAL( IniRead( IniFile, "Compiler options", "BeepOnCompletion", "" ))
  CpOpt.DebugToolPath = IniRead( IniFile, "Compiler options", "DebugToolPath", "" )
  CpOpt.PBWINVersion = VAL( IniRead( IniFile, "Compiler options", "PBWINVersion", "" ))
  CpOpt.PBCCVersion = VAL( IniRead( IniFile, "Compiler options", "PBCCVersion", "" ))
END SUB
SUB GetColorOptions( ColOpt AS SciColorsAndFontsType )
  LOCAL rs AS STRING
  LOCAL IniFile AS STRING
  IniFile=EXE.PATH$ & "config.ini"
  ' Default
  rs = IniRead( IniFile, "Color options", "DefaultForeColor", "" )
  IF LEN( rs ) THEN ColOpt.DefaultForeColor = VAL( rs ) ELSE ColOpt.DefaultForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "DefaultBackColor", "" )
  IF LEN( rs ) THEN ColOpt.DefaultBackColor = VAL( rs ) ELSE ColOpt.DefaultBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "DefaultFontName", "" )
  IF LEN( rs ) THEN ColOpt.DefaultFontName = rs ELSE ColOpt.DefaultFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "DefaultFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.DefaultFontCharset = rs ELSE ColOpt.DefaultFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "DefaultFontSize", "" )
  IF LEN( rs ) THEN ColOpt.DefaultFontSize = VAL( rs ) ELSE ColOpt.DefaultFontSize = 8
  ColOpt.DefaultFontBold = VAL( IniRead( IniFile, "Color options", "DefaultFontBold", "" ))
  ColOpt.DefaultFontItalic = VAL( IniRead( IniFile, "Color options", "DefaultFontItalic", "" ))
  ColOpt.DefaultFontUnderline = VAL( IniRead( IniFile, "Color options", "DefaultFontUnderline", "" ))
  ' Comment
  rs = IniRead( IniFile, "Color options", "CommentForeColor", "" )
  IF LEN( rs ) THEN ColOpt.CommentForeColor = VAL( rs ) ELSE ColOpt.CommentForeColor = RGB( 0, 128, 0 )
  rs = IniRead( IniFile, "Color options", "CommentBackColor", "" )
  IF LEN( rs ) THEN ColOpt.CommentBackColor = VAL( rs ) ELSE ColOpt.CommentBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "CommentFontName", "" )
  IF LEN( rs ) THEN ColOpt.CommentFontName = rs ELSE ColOpt.CommentFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "CommentFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.CommentFontCharset = rs ELSE ColOpt.CommentFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "CommentFontSize", "" )
  IF LEN( rs ) THEN ColOpt.CommentFontSize = VAL( rs ) ELSE ColOpt.CommentFontSize = 8
  ColOpt.CommentFontBold = VAL( IniRead( IniFile, "Color options", "CommentFontBold", "" ))
  ColOpt.CommentFontItalic = VAL( IniRead( IniFile, "Color options", "CommentFontItalic", "" ))
  ColOpt.CommentFontUnderline = VAL( IniRead( IniFile, "Color options", "CommentFontUnderline", "" ))
  ' Constants
  rs = IniRead( IniFile, "Color options", "ConstantForeColor", "" )
  IF LEN( rs ) THEN ColOpt.ConstantForeColor = VAL( rs ) ELSE ColOpt.ConstantForeColor = RGB( 230, 50, 0 )
  rs = IniRead( IniFile, "Color options", "ConstantBackColor", "" )
  IF LEN( rs ) THEN ColOpt.ConstantBackColor = VAL( rs ) ELSE ColOpt.ConstantBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "ConstantFontName", "" )
  IF LEN( rs ) THEN ColOpt.ConstantFontName = rs ELSE ColOpt.ConstantFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "ConstantFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.ConstantFontCharset = rs ELSE ColOpt.ConstantFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "ConstantFontSize", "" )
  IF LEN( rs ) THEN ColOpt.ConstantFontSize = VAL( rs ) ELSE ColOpt.ConstantFontSize = 8
  ColOpt.ConstantFontBold = VAL( IniRead( IniFile, "Color options", "ConstantFontBold", "" ))
  ColOpt.ConstantFontItalic = VAL( IniRead( IniFile, "Color options", "ConstantFontItalic", "" ))
  ColOpt.ConstantFontUnderline = VAL( IniRead( IniFile, "Color options", "ConstantFontUnderline", "" ))
  ' Identifier
  rs = IniRead( IniFile, "Color options", "IdentifierForeColor", "" )
  IF LEN( rs ) THEN ColOpt.IdentifierForeColor = VAL( rs ) ELSE ColOpt.IdentifierForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "IdentifierBackColor", "" )
  IF LEN( rs ) THEN ColOpt.IdentifierBackColor = VAL( rs ) ELSE ColOpt.IdentifierBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "IdentifierFontName", "" )
  IF LEN( rs ) THEN ColOpt.IdentifierFontName = rs ELSE ColOpt.IdentifierFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "IdentifierFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.IdentifierFontCharset = rs ELSE ColOpt.IdentifierFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "IdentifierFontSize", "" )
  IF LEN( rs ) THEN ColOpt.IdentifierFontSize = VAL( rs ) ELSE ColOpt.IdentifierFontSize = 8
  ColOpt.IdentifierFontBold = VAL( IniRead( IniFile, "Color options", "IdentifierFontBold", "" ))
  ColOpt.IdentifierFontItalic = VAL( IniRead( IniFile, "Color options", "IdentifierFontItalic", "" ))
  ColOpt.IdentifierFontUnderline = VAL( IniRead( IniFile, "Color options", "IdentifierFontUnderline", "" ))
  ' Keywords
  rs = IniRead( IniFile, "Color options", "KeywordForeColor", "" )
  IF LEN( rs ) THEN ColOpt.KeywordForeColor = VAL( rs ) ELSE ColOpt.KeywordForeColor = RGB( 0, 0, 255 )
  rs = IniRead( IniFile, "Color options", "KeywordBackColor", "" )
  IF LEN( rs ) THEN ColOpt.KeywordBackColor = VAL( rs ) ELSE ColOpt.KeywordBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "KeywordFontName", "" )
  IF LEN( rs ) THEN ColOpt.KeywordFontName = rs ELSE ColOpt.KeywordFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "KeywordFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.KeywordFontCharset = rs ELSE ColOpt.KeywordFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "KeywordFontSize", "" )
  IF LEN( rs ) THEN ColOpt.KeywordFontSize = VAL( rs ) ELSE ColOpt.KeywordFontSize = 8
  ColOpt.KeywordFontBold = VAL( IniRead( IniFile, "Color options", "KeywordFontBold", "" ))
  ColOpt.KeywordFontItalic = VAL( IniRead( IniFile, "Color options", "KeywordFontItalic", "" ))
  ColOpt.KeywordFontUnderline = VAL( IniRead( IniFile, "Color options", "KeywordFontUnderline", "" ))
  ' Number
  rs = IniRead( IniFile, "Color options", "NumberForeColor", "" )
  IF LEN( rs ) THEN ColOpt.NumberForeColor = VAL( rs ) ELSE ColOpt.NumberForeColor = RGB( 192, 100, 0 )
  rs = IniRead( IniFile, "Color options", "NumberBackColor", "" )
  IF LEN( rs ) THEN ColOpt.NumberBackColor = VAL( rs ) ELSE ColOpt.NumberBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "NumberFontName", "" )
  IF LEN( rs ) THEN ColOpt.NumberFontName = rs ELSE ColOpt.NumberFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "NumberFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.NumberFontCharset = rs ELSE ColOpt.NumberFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "NumberFontSize", "" )
  IF LEN( rs ) THEN ColOpt.NumberFontSize = VAL( rs ) ELSE ColOpt.NumberFontSize = 8
  ColOpt.NumberFontBold = VAL( IniRead( IniFile, "Color options", "NumberFontBold", "" ))
  ColOpt.NumberFontItalic = VAL( IniRead( IniFile, "Color options", "NumberFontItalic", "" ))
  ColOpt.NumberFontUnderline = VAL( IniRead( IniFile, "Color options", "NumberFontUnderline", "" ))
  ' Line numbers
  rs = IniRead( IniFile, "Color options", "LineNumberForeColor", "" )
  IF LEN( rs ) THEN ColOpt.LineNumberForeColor = VAL( rs ) ELSE ColOpt.LineNumberForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "LineNumberBackColor", "" )
  IF LEN( rs ) THEN ColOpt.LineNumberBackColor = VAL( rs ) ELSE ColOpt.LineNumberBackColor = %LTGRAY
  rs = IniRead( IniFile, "Color options", "LineNumberFontName", "" )
  IF LEN( rs ) THEN ColOpt.LineNumberFontName = rs ELSE ColOpt.LineNumberFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "LineNumberFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.LineNumberFontCharset = rs ELSE ColOpt.LineNumberFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "LineNumberFontSize", "" )
  IF LEN( rs ) THEN ColOpt.LineNumberFontSize = VAL( rs ) ELSE ColOpt.LineNumberFontSize = 8
  ColOpt.LineNumberFontBold = VAL( IniRead( IniFile, "Color options", "LineNumberFontBold", "" ))
  ColOpt.LineNumberFontItalic = VAL( IniRead( IniFile, "Color options", "LineNumberFontItalic", "" ))
  ColOpt.LineNumberFontUnderline = VAL( IniRead( IniFile, "Color options", "LineNumberFontUnderline", "" ))
  ' Operators
  rs = IniRead( IniFile, "Color options", "OperatorForeColor", "" )
  IF LEN( rs ) THEN ColOpt.OperatorForeColor = VAL( rs ) ELSE ColOpt.OperatorForeColor = RGB( 0, 128, 128 )
  rs = IniRead( IniFile, "Color options", "OperatorBackColor", "" )
  IF LEN( rs ) THEN ColOpt.OperatorBackColor = VAL( rs ) ELSE ColOpt.OperatorBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "OperatorFontName", "" )
  IF LEN( rs ) THEN ColOpt.OperatorFontName = rs ELSE ColOpt.OperatorFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "OperatorFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.OperatorFontCharset = rs ELSE ColOpt.OperatorFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "OperatorFontSize", "" )
  IF LEN( rs ) THEN ColOpt.OperatorFontSize = VAL( rs ) ELSE ColOpt.OperatorFontSize = 8
  ColOpt.OperatorFontBold = VAL( IniRead( IniFile, "Color options", "OperatorFontBold", "" ))
  ColOpt.OperatorFontItalic = VAL( IniRead( IniFile, "Color options", "OperatorFontItalic", "" ))
  ColOpt.OperatorFontUnderline = VAL( IniRead( IniFile, "Color options", "OperatorFontUnderline", "" ))
  ' Preprocessor
  rs = IniRead( IniFile, "Color options", "PreprocessorForeColor", "" )
  IF LEN( rs ) THEN ColOpt.PreprocessorForeColor = VAL( rs ) ELSE ColOpt.PreprocessorForeColor = RGB( 255, 0, 0 )
  rs = IniRead( IniFile, "Color options", "PreprocessorBackColor", "" )
  IF LEN( rs ) THEN ColOpt.PreprocessorBackColor = VAL( rs ) ELSE ColOpt.PreprocessorBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "PreprocessorFontName", "" )
  IF LEN( rs ) THEN ColOpt.PreprocessorFontName = rs ELSE ColOpt.PreprocessorFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "PreprocessorFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.PreprocessorFontCharset = rs ELSE ColOpt.PreprocessorFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "PreprocessorFontSize", "" )
  IF LEN( rs ) THEN ColOpt.PreprocessorFontSize = VAL( rs ) ELSE ColOpt.PreprocessorFontSize = 8
  ColOpt.PreprocessorFontBold = VAL( IniRead( IniFile, "Color options", "PreprocessorFontBold", "" ))
  ColOpt.PreprocessorFontItalic = VAL( IniRead( IniFile, "Color options", "PreprocessorFontItalic", "" ))
  ColOpt.PreprocessorFontUnderline = VAL( IniRead( IniFile, "Color options", "PreprocessorFontUnderline", "" ))
  ' Strings
  rs = IniRead( IniFile, "Color options", "StringForeColor", "" )
  IF LEN( rs ) THEN ColOpt.StringForeColor = VAL( rs ) ELSE ColOpt.StringForeColor = RGB( 255, 0, 255 )
  rs = IniRead( IniFile, "Color options", "StringBackColor", "" )
  IF LEN( rs ) THEN ColOpt.StringBackColor = VAL( rs ) ELSE ColOpt.StringBackColor = %WHITE
  rs = IniRead( IniFile, "Color options", "StringFontName", "" )
  IF LEN( rs ) THEN ColOpt.StringFontName = rs ELSE ColOpt.StringFontName = "Courier New"
  rs = IniRead( IniFile, "Color options", "StringFontCharset", "" )
  IF LEN( rs ) THEN ColOpt.StringFontCharset = rs ELSE ColOpt.StringFontCharset = "Default"
  rs = IniRead( IniFile, "Color options", "StringFontSize", "" )
  IF LEN( rs ) THEN ColOpt.StringFontSize = VAL( rs ) ELSE ColOpt.StringFontSize = 8
  ColOpt.StringFontBold = VAL( IniRead( IniFile, "Color options", "StringFontBold", "" ))
  ColOpt.StringFontItalic = VAL( IniRead( IniFile, "Color options", "StringFontItalic", "" ))
  ColOpt.StringFontUnderline = VAL( IniRead( IniFile, "Color options", "StringFontUnderline", "" ))
  ' Caret
  rs = IniRead( IniFile, "Color options", "CaretForeColor", "" )
  IF LEN( rs ) THEN ColOpt.CaretForeColor = VAL( rs ) ELSE ColOpt.CaretForeColor = %BLACK
  ' Edge
  rs = IniRead( IniFile, "Color options", "EdgeForeColor", "" )
  IF LEN( rs ) THEN ColOpt.EdgeForeColor = VAL( rs ) ELSE ColOpt.EdgeForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "EdgeBackColor", "" )
  IF LEN( rs ) THEN ColOpt.EdgeBackColor = VAL( rs ) ELSE ColOpt.EdgeBackColor = %WHITE
  ' Fold
  rs = IniRead( IniFile, "Color options", "FoldForeColor", "" )
  IF LEN( rs ) THEN ColOpt.FoldForeColor = VAL( rs ) ELSE ColOpt.FoldForeColor = %RED
  rs = IniRead( IniFile, "Color options", "FoldBackColor", "" )
  IF LEN( rs ) THEN ColOpt.FoldBackColor = VAL( rs ) ELSE ColOpt.FoldBackColor = %WHITE
  ' Fold open
  rs = IniRead( IniFile, "Color options", "FoldOpenForeColor", "" )
  IF LEN( rs ) THEN ColOpt.FoldOpenForeColor = VAL( rs ) ELSE ColOpt.FoldOpenForeColor = %RED
  rs = IniRead( IniFile, "Color options", "FoldOpenBackColor", "" )
  IF LEN( rs ) THEN ColOpt.FoldOpenBackColor = VAL( rs ) ELSE ColOpt.FoldOpenBackColor = %WHITE
  ' Fold margin
  rs = IniRead( IniFile, "Color options", "FoldMarginForeColor", "" )
  IF LEN( rs ) THEN ColOpt.FoldMarginForeColor = VAL( rs ) ELSE ColOpt.FoldMarginForeColor = RGB( 200, 0, 200 )
  rs = IniRead( IniFile, "Color options", "FoldMarginBackColor", "" )
  IF LEN( rs ) THEN ColOpt.FoldMarginBackColor = VAL( rs ) ELSE ColOpt.FoldMarginBackColor = RGB( 100, 0, 100 )
  ' Indent guides
  rs = IniRead( IniFile, "Color options", "IndentGuideForeColor", "" )
  IF LEN( rs ) THEN ColOpt.IndentGuideForeColor = VAL( rs ) ELSE ColOpt.IndentGuideForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "IndentGuideBackColor", "" )
  IF LEN( rs ) THEN ColOpt.IndentGuideBackColor = VAL( rs ) ELSE ColOpt.IndentGuideBackColor = %WHITE
  ' Selection
  rs = IniRead( IniFile, "Color options", "SelectionForeColor", "" )
  IF LEN( rs ) THEN ColOpt.SelectionForeColor = VAL( rs ) ELSE ColOpt.SelectionForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "SelectionBackColor", "" )
  IF LEN( rs ) THEN ColOpt.SelectionBackColor = VAL( rs ) ELSE ColOpt.SelectionBackColor = %LTGRAY
  ' Whitespace
  rs = IniRead( IniFile, "Color options", "WhitespaceForeColor", "" )
  IF LEN( rs ) THEN ColOpt.WhitespaceForeColor = VAL( rs ) ELSE ColOpt.WhitespaceForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "WhitespaceBackColor", "" )
  IF LEN( rs ) THEN ColOpt.WhitespaceBackColor = VAL( rs ) ELSE ColOpt.WhitespaceBackColor = %WHITE
  ' Codetips
  rs = IniRead( IniFile, "Color options", "CodetipForeColor", "" )
  IF LEN( rs ) THEN ColOpt.CodetipForeColor = VAL( rs ) ELSE ColOpt.CodetipForeColor = %LTGRAY
  rs = IniRead( IniFile, "Color options", "CodetipBackColor", "" )
  IF LEN( rs ) THEN ColOpt.CodetipBackColor = VAL( rs ) ELSE ColOpt.CodetipBackColor = %WHITE
  ' Submenus
  rs = IniRead( IniFile, "Color options", "SubmenuTextForeColor", "" )
  IF LEN( rs ) THEN ColOpt.SubmenuTextForeColor = VAL( rs ) ELSE ColOpt.SubmenuTextForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "SubmenuTextBackColor", "" )
  IF LEN( rs ) THEN ColOpt.SubmenuTextBackColor = VAL( rs ) ELSE ColOpt.SubmenuTextBackColor = RGB( 255, 249, 242 )
  rs = IniRead( IniFile, "Color options", "SubmenuHiTextForeColor", "" )
  IF LEN( rs ) THEN ColOpt.SubmenuHiTextForeColor = VAL( rs ) ELSE ColOpt.SubmenuHiTextForeColor = %BLACK
  rs = IniRead( IniFile, "Color options", "SubmenuHiTextBackColor", "" )
  IF LEN( rs ) THEN ColOpt.SubmenuHiTextBackColor = VAL( rs ) ELSE ColOpt.SubmenuHiTextBackColor = &H00EFD3C1
  ' Caret line background color
  rs = IniRead( IniFile, "Color options", "CaretLineBackColor", "" )
  IF LEN( rs ) THEN ColOpt.CaretLineBackColor = VAL( rs ) ELSE ColOpt.CaretLineBackColor = %YELLOW
  ' Menu icons background color
  rs = IniRead( IniFile, "Color options", "MenuIconsBackColor", "" )
  IF LEN( rs ) THEN ColOpt.MenuIconsBackColor = VAL( rs ) ELSE ColOpt.MenuIconsBackColor = RGB( 193, 211, 239 )
  ' Use always the default backgroundcolor
  rs = IniRead( IniFile, "Color options", "UseAlwaysDefaultBackColor", "" )
  IF LEN( rs ) THEN ColOpt.UseAlwaysDefaultBackColor = VAL( rs ) ELSE ColOpt.UseAlwaysDefaultBackColor = %BST_UNCHECKED
  rs = IniRead( IniFile, "Color options", "UseAlwaysDefaultFont", "" )
  IF LEN( rs ) THEN ColOpt.UseAlwaysDefaultFont = VAL( rs ) ELSE ColOpt.UseAlwaysDefaultFont = %BST_UNCHECKED
  rs = IniRead( IniFile, "Color options", "UseAlwaysDefaultFontSize", "" )
  IF LEN( rs ) THEN ColOpt.UseAlwaysDefaultFontSize = VAL( rs ) ELSE ColOpt.UseAlwaysDefaultFontSize = %BST_UNCHECKED
END SUB
' *********************************************************************************************
SUB GetFoldingOptions( FoldOpt AS FoldingOptionsType )
  LOCAL rs AS STRING
  LOCAL IniFile AS STRING
  IniFile=EXE.PATH$ & "config.ini"
  rs = IniRead( IniFile, "Folding options", "FoldingLevel", "" )
  IF LEN( rs ) THEN FoldOpt.FoldingLevel = VAL( rs ) ELSE FoldOpt.FoldingLevel = 1
  rs = IniRead( IniFile, "Folding options", "FoldingSymbol", "" )
  IF LEN( rs ) THEN FoldOpt.FoldingSymbol = VAL( rs ) ELSE FoldOpt.FoldingSymbol = 4
END SUB
' *********************************************************************************************
' Reads the Scintilla print options
' *********************************************************************************************
SUB GetScintillaPrintOptions( Magnification AS LONG, COLORMODE AS LONG, WrapMode AS LONG )
  LOCAL rs AS STRING
  LOCAL IniFile AS STRING
  IniFile=EXE.PATH$ & "config.ini"
  rs = IniRead( IniFile, "Scintilla Print Options", "Magnification", "" )
  Magnification = VAL( rs )
  rs = IniRead( IniFile, "Scintilla Print Options", "ColorMode", "" )
  COLORMODE = VAL( rs )
  rs = IniRead( IniFile, "Scintilla Print Options", "WrapMode", "" )
  WrapMode = VAL( rs )
END SUB
SUB ShowLinCol
  LOCAL curPos AS LONG        ' // Current position 当前位置
  LOCAL nLine AS LONG         ' // Number of line 行数
  LOCAL nCol AS LONG      ' // Number of column 列数
  LOCAL nLines AS LONG        ' // Number of lines in the document 文档的总行数
  LOCAL nTextLen AS LONG      ' // Length in characters of the document文档的字符长度
  LOCAL szText AS ASCIIZ * 255        ' // General purpose variable 一般目标变量
  LOCAL hSci AS DWORD         ' // Handle of the Scintilla control Scintilla控件句柄
  LOCAL pSci AS DWORD         ' // Direct pointer to the Scintilla control指向Scintilla控件的指针
  ' Retrieve the handle of the edit window 得到编辑窗口的句柄
  hSci = GetEdit
  ' If there is not any file being edited, clear the status bar如果没有编辑任何文件，清除状态栏
  IF ISFALSE hSci THEN
    szText = ""
    SENDMESSAGE g_hStatus, %SB_SETTEXT, 1, VARPTR( szText )
    SENDMESSAGE g_hStatus, %SB_SETTEXT, 2, VARPTR( szText )
    'ChangeButtonsState
    EXIT SUB
  END IF
  ' Get direct pointer to the Scintilla control for faster access 为最近访问获得Scintilla控件的指针
  pSci = SENDMESSAGE( hSci, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT SUB
  ' Retrieve the information and show it in the status bar 得到信息并显示在状态栏里
  curPos = SciMsg( pSci, %SCI_GETCURRENTPOS, 0, 0 )
  nLine = SciMsg( pSci, %SCI_LINEFROMPOSITION, curPos, 0 ) + 1
  nCol = SciMsg( pSci, %SCI_GETCOLUMN, curPos, 0 ) + 1
  szText = " " & FORMAT$( nLine ) & ":" & FORMAT$( nCol )
  SENDMESSAGE g_hStatus, %SB_SETTEXT, 1, VARPTR( szText )
  nLines = SciMsg( pSci, %SCI_GETLINECOUNT, 0, 0 )
  nTextLen = SciMsg( pSci, %SCI_GETTEXTLENGTH, 0, 0 )
  szText = " " & FORMAT$( nLines ) & " 行, " & FORMAT$( nTextLen ) & " 个字符"
  SENDMESSAGE g_hStatus, %SB_SETTEXT, 2, VARPTR( szText )
  ' Change the toolbar buttons state if needed 如果需要，改变状态栏按钮状态
  'ChangeButtonsState
END SUB
' *********************************************************************************************
' Clear the status bar and disable the toolbar buttons if there is not any file being edited
' 如果没有文件被编辑则清除状态栏并且不可用工具的按钮
' *********************************************************************************************
SUB ClearStatusbar( )
  LOCAL szText AS ASCIIZ * 255
  szText = ""
  SENDMESSAGE g_hStatus, %SB_SETTEXT, 1, VARPTR( szText )
  SENDMESSAGE g_hStatus, %SB_SETTEXT, 2, VARPTR( szText )
  SENDMESSAGE g_hStatus, %SB_SETTEXT, 4, VARPTR( szText )
  'DisableToolbarButtons
END SUB
' *********************************************************************************************
' Returns true if cursor within SUB/FUNCTION如果光标在SUB/FUNCTION内则返回真
' *********************************************************************************************
SUB SED_WithinProc( BYREF tmpPROC AS PROC )
  LOCAL curPos            AS LONG         ' // Current position当前位置
  LOCAL LineNumber        AS LONG         ' // Number of line行数
  LOCAL LineLen           AS LONG         ' // Length of the line行长度
  LOCAL i                 AS LONG         ' // Loop counter循环计数
  LOCAL buffer            AS STRING   ' // buffer缓冲
  LOCAL Namebuffer        AS STRING   ' // buffer to keep the function name缓冲保持函数名
  LOCAL LineCount         AS LONG         ' // Number of lines行总数
  curPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )       ' Current position当前位置
  LineNumber = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, curPos, 0 )       ' Line number行数
  LineCount = SENDMESSAGE( GetEdit, %SCI_GETLINECOUNT, 0, 0 )         ' Number of lines行总数
  tmpPROC.WhatIsUp = %ID_ProcNull         ' Start with false开始时是false
  tmpPROC.UpLnNo = 0
  tmpPROC.WhatIsDown = %ID_ProcNull       ' Start with false开始时是False
  tmpPROC.DnLnNo = LineCount
  '   tmpPROC.ProcName = "(Code finder)"
  tmpPROC.ProcName = "(代码快速定位)"
  ' Check current line检查当前行
  LineLen = SENDMESSAGE( GetEdit, %SCI_LINELENGTH, LineNumber, 0 )        ' Length of the line行的长度
  buffer = SPACE$( LineLen )      ' Size the buffer缓冲大小
  SENDMESSAGE( GetEdit, %SCI_GETLINE, LineNumber, STRPTR( buffer ))       ' Get the text of the line得到行的文本
  DO
    IF ISFALSE INSTR( buffer, "  " ) THEN EXIT DO
    REPLACE "  " WITH " " IN buffer
  LOOP
  Namebuffer = LTRIM$( buffer )
  buffer = LTRIM$( UCASE$( buffer ))
  IF LEFT$( buffer, 4 ) = "SUB " OR LEFT$( buffer, 9 ) = "FUNCTION " OR _
        ( LEFT$( buffer, 7 ) = "STATIC " AND MID$( buffer, 8, 3 ) = "SUB" ) OR _
        ( LEFT$( buffer, 7 ) = "STATIC " AND MID$( buffer, 8, 8 ) = "FUNCTION" ) OR _
        ( LEFT$( buffer, 9 ) = "CALLBACK " AND MID$( buffer, 10, 8 ) = "FUNCTION" ) THEN
    IF LEFT$( buffer, 9 ) = "FUNCTION " THEN
      Namebuffer = LTRIM$( MID$( Namebuffer, 10 ))
      buffer = LTRIM$( MID$( buffer, 10 ))
      IF LEFT$( buffer, 1 ) <> "=" THEN
        tmpPROC.WhatIsUp = %ID_ProcStart        ' Proc. Begins
        tmpPROC.UpLnNo = LineNumber
      END IF
    ELSE
      tmpPROC.WhatIsUp = %ID_ProcStart        ' Proc. Begins
      tmpPROC.UpLnNo = LineNumber
    END IF
    Namebuffer = TRIM$( MID$( Namebuffer, 1, INSTR( Namebuffer, ANY CHR$( 40, 13, 10 )) - 1 ))
    tmpPROC.ProcName = PARSE$( Namebuffer, CHR$( 32 ), PARSECOUNT( Namebuffer, CHR$( 32 )))
  ELSEIF ( LEFT$( buffer, 4 ) = "END " AND MID$( buffer, 5, 3 ) = "SUB" ) OR _
        ( LEFT$( buffer, 4 ) = "END " AND MID$( buffer, 5, 8 ) = "FUNCTION" ) THEN
    tmpPROC.WhatIsDown = %ID_ProcEnd        ' Proc. Ends过程末尾
    tmpPROC.DnLnNo = LineNumber
  END IF
  ' Check for begining if not found如果没找到检查开始
  IF tmpPROC.WhatIsUp = %ID_ProcNull THEN
    FOR i = LineNumber - 1 TO 0 STEP - 1
      LineLen = SENDMESSAGE( GetEdit, %SCI_LINELENGTH, i, 0 )         ' Length of the line行长度
      buffer = SPACE$( LineLen )      ' Size the buffer缓冲尺寸
      SENDMESSAGE( GetEdit, %SCI_GETLINE, i, STRPTR( buffer ))        ' Get the text of the line得到行的文本
      DO
        IF ISFALSE INSTR( buffer, "  " ) THEN EXIT DO
        REPLACE "  " WITH " " IN buffer
      LOOP
      Namebuffer = LTRIM$( buffer )
      buffer = LTRIM$( UCASE$( buffer ))
      IF LEFT$( buffer, 4 ) = "SUB " OR LEFT$( buffer, 9 ) = "FUNCTION " OR _
          ( LEFT$( buffer, 7 ) = "STATIC " AND MID$( buffer, 8, 3 ) = "SUB" ) OR _
          ( LEFT$( buffer, 7 ) = "STATIC " AND MID$( buffer, 8, 8 ) = "FUNCTION" ) OR _
          ( LEFT$( buffer, 9 ) = "CALLBACK " AND MID$( buffer, 10, 8 ) = "FUNCTION" ) THEN
        IF LEFT$( buffer, 9 ) = "FUNCTION " THEN
          Namebuffer = LTRIM$( MID$( Namebuffer, 10 ))
          buffer = LTRIM$( MID$( buffer, 10 ))
          IF LEFT$( buffer, 1 ) <> "=" THEN
            tmpPROC.WhatIsUp = %ID_ProcStart        ' Proc. Begins过程开始
            tmpPROC.UpLnNo = i
            Namebuffer = TRIM$( MID$( Namebuffer, 1, INSTR( Namebuffer, ANY CHR$( 40, 13, 10 )) - 1 ))
            tmpPROC.ProcName = PARSE$( Namebuffer, CHR$( 32 ), PARSECOUNT( Namebuffer, CHR$( 32 )))
            EXIT FOR
          END IF
        ELSE
          tmpPROC.WhatIsUp = %ID_ProcStart        ' Proc. Begins过程开始
          tmpPROC.UpLnNo = i
          Namebuffer = TRIM$( MID$( Namebuffer, 1, INSTR( Namebuffer, ANY CHR$( 40, 13, 10 )) - 1 ))
          tmpPROC.ProcName = PARSE$( Namebuffer, CHR$( 32 ), PARSECOUNT( Namebuffer, CHR$( 32 )))
          EXIT FOR
        END IF
      ELSEIF ( LEFT$( buffer, 4 ) = "END " AND MID$( buffer, 5, 3 ) = "SUB" ) OR _
          ( LEFT$( buffer, 4 ) = "END " AND MID$( buffer, 5, 8 ) = "FUNCTION" ) THEN
        tmpPROC.WhatIsUp = %ID_ProcEnd      ' Proc. Ends过程末尾
        tmpPROC.UpLnNo = i
        EXIT FOR
      END IF
    NEXT
  END IF
  ' Check for End检查末尾
  IF tmpPROC.WhatIsDown = %ID_ProcNull THEN
    FOR i = LineNumber + 1 TO LineCount
      LineLen = SENDMESSAGE( GetEdit, %SCI_LINELENGTH, i, 0 )         ' Length of the line行长度
      buffer = SPACE$( LineLen )      ' Size the buffer缓冲尺寸
      SENDMESSAGE( GetEdit, %SCI_GETLINE, i, STRPTR( buffer ))        ' Get the text of the line得到行的文本
      DO
        IF ISFALSE INSTR( buffer, "  " ) THEN EXIT DO
        REPLACE "  " WITH " " IN buffer
      LOOP
      buffer = LTRIM$( UCASE$( buffer ))
      IF ( LEFT$( buffer, 4 ) = "END " AND MID$( buffer, 5, 3 ) = "SUB" ) OR _
          ( LEFT$( buffer, 4 ) = "END " AND MID$( buffer, 5, 8 ) = "FUNCTION" ) THEN
        tmpPROC.WhatIsDown = %ID_ProcEnd
        tmpPROC.DnLnNo = i
        EXIT FOR
      ELSE
        IF LEFT$( buffer, 4 ) = "SUB " OR LEFT$( buffer, 9 ) = "FUNCTION " OR _
            ( LEFT$( buffer, 7 ) = "STATIC " AND MID$( buffer, 8, 3 ) = "SUB" ) OR _
            ( LEFT$( buffer, 7 ) = "STATIC " AND MID$( buffer, 8, 8 ) = "FUNCTION" ) OR _
            ( LEFT$( buffer, 9 ) = "CALLBACK " AND MID$( buffer, 10, 8 ) = "FUNCTION" ) THEN
          IF LEFT$( buffer, 9 ) = "FUNCTION " THEN
            buffer = LTRIM$( MID$( buffer, 10 ))
            IF LEFT$( buffer, 1 ) <> "=" THEN
              tmpPROC.WhatIsDown = %ID_ProcStart
              tmpPROC.DnLnNo = i
              EXIT FOR
            END IF
          ELSE
            tmpPROC.WhatIsDown = %ID_ProcStart
            tmpPROC.DnLnNo = i
            EXIT FOR
          END IF
        END IF
      END IF
    NEXT
  END IF
END SUB
