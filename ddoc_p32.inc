'  ddoc_p32.inc
'
'  Power Basic 32-bit call interface for the ddoc Print and Preview Engine
'  Copyright 1998, Don Dickinson, All Rights Reserved.
'  ddoc打印及预览机制,基于32位调用接口
'  ddoc打印预览基于用VB3写成的VB3报表视图WinRep软件。

'  由dpStartDoc返回的错误
'==================================================================
%DDOC_NOHANDLES = - 1
%DDOC_BADFILE = - 2
%DDOC_FILEEXISTS = - 3

'  由dpMailInit返回的状态代码
'==================================================================
%DDOC_MAIL_SUCCESS = 0
%DDOC_MAIL_FNF = 1
%DDOC_MAIL_LNF = 2
%DDOC_MAIL_NOH = 3
%DDOC_MAIL_HEIGHT_INVALID = 4
%DDOC_MAIL_WIDTH_INVALID = 5
%DDOC_MAIL_LABHEIGHT_INVALID = 6
%DDOC_MAIL_LABWIDTH_INVALID = 7
%DDOC_MAIL_COL_INVALID = 8
%DDOC_MAIL_ROW_INVALID = 9
%DDOC_MAIL_LINES_INVALID = 10
'
'  测量单位常量，传递到dpStartDoc
'==================================================================
%DDOC_INCH = 0
%DDOC_CM = 1
'- 以下会传递到dpStartDoc或dpEndDoc中的iOption参数
'     Paper / Orientation / Bin
'
'  如果被传递，则ddoc就要遵守用户对方向的选择并执行
'  (如果用户未指定，则使用系统默认值).
'
%DDOC_SYSTEM_DEFAULT = - 1
%DDOC_SAME_AS_FIRST = - 2
%DDOC_PAPER_CUSTOM = - 3		'use custom paper size使用自定义纸张尺寸
'
'  Orientation Constants for use with dpStartDoc 用于dpStartDoc的方向常量
'==================================================================
%DDOC_PORTRAIT = 0
%DDOC_LANDSCAPE = 1
'
'  Paper Bin Constants used in dpStartDoc and dpNewPage
'==================================================================
%DDOC_BIN_UPPER = 1
%DDOC_BIN_FIRST = %DDOC_BIN_UPPER
%DDOC_BIN_ONLYONE = 1
%DDOC_BIN_LOWER = 2
%DDOC_BIN_MIDDLE = 3
%DDOC_BIN_MANUAL = 4
%DDOC_BIN_ENVELOPE = 5
%DDOC_BIN_ENVMANUAL = 6
%DDOC_BIN_AUTO = 7
%DDOC_BIN_TRACTOR = 8
%DDOC_BIN_SMALLFMT = 9
%DDOC_BIN_LARGEFMT = 10
%DDOC_BIN_LARGECAPACITY = 11
%DDOC_BIN_CASSETTE = 14
%DDOC_BIN_LAST = %DDOC_BIN_CASSETTE
'
'  Paper Size Constants used in dpStartDoc and dpNewPage
'==================================================================
%DDOC_PAPER_LETTER = 1		'  Letter 8 1/2 x 11 in
%DDOC_PAPER_LETTERSMALL = 2		'  Letter Small 8 1/2 x 11 in
%DDOC_PAPER_TABLOID = 3		'  Tabloid 11 x 17 in
%DDOC_PAPER_LEDGER = 4		'  Ledger 17 x 11 in
%DDOC_PAPER_LEGAL = 5		'  Legal 8 1/2 x 14 in
%DDOC_PAPER_STATEMENT = 6		'  Statement 5 1/2 x 8 1/2 in
%DDOC_PAPER_EXECUTIVE = 7		'  Executive"7 1/2 x 10 in
%DDOC_PAPER_A3 = 8		'  A3 297 x 420 mm
%DDOC_PAPER_A4 = 9		'  A4 210 x 297 mm
%DDOC_PAPER_A4SMALL = 10		'  A4 Small 210 x 297 mm
%DDOC_PAPER_A5 = 11		'  A5 148 x 210 mm
%DDOC_PAPER_B4 = 12		'  B4 250 x 354
%DDOC_PAPER_B5 = 13		'  B5 182 x 257 mm
%DDOC_PAPER_FOLIO = 14		'  Folio 8 1/2 x 13 in
%DDOC_PAPER_QUARTO = 15		'  Quarto 215 x 275 mm
%DDOC_PAPER_10x14 = 16		'  10x14 in
%DDOC_PAPER_11X17 = 17		'  11x17 in
%DDOC_PAPER_NOTE = 18		'  Note 8 1/2 x 11 in
%DDOC_ENV_9 = 19		'  Envelope #9 3 7/8 x 8 7/8
%DDOC_ENV_10 = 20		'  Envelope #10 4 1/8 x 9 1/2
%DDOC_ENV_11 = 21		'  Envelope #11 4 1/2 x 10 3/8
%DDOC_ENV_12 = 22		'  Envelope #12 4 \276 x 11
%DDOC_ENV_14 = 23		'  Envelope #14 5 x 11 1/2
%DDOC_ENV_DL = 27		'  Envelope DL 110 x 220mm
%DDOC_ENV_C5 = 28		'  Envelope C5 162 x 229 mm
%DDOC_ENV_C3 = 29		'  Envelope C3  324 x 458 mm
%DDOC_ENV_C4 = 30		'  Envelope C4  229 x 324 mm
%DDOC_ENV_C6 = 31		'  Envelope C6  114 x 162 mm
%DDOC_ENV_C65 = 32		'  Envelope C65 114 x 229 mm
%DDOC_ENV_B4 = 33		'  Envelope B4  250 x 353 mm
%DDOC_ENV_B5 = 34		'  Envelope B5  176 x 250 mm
%DDOC_ENV_B6 = 35		'  Envelope B6  176 x 125 mm
%DDOC_ENV_ITALY = 36		'  Envelope 110 x 230 mm
%DDOC_ENV_MONARCH = 37		'  Envelope Monarch 3.875 x 7.5 in
%DDOC_ENV_PERSONAL = 38		'  6 3/4 Envelope 3 5/8 x 6 1/2 in
%DDOC_FANFOLD_US = 39		'  US Std Fanfold 14 7/8 x 11 in
%DDOC_FANFOLD_STD_GERMAN = 40		'  German Std Fanfold 8 1/2 x 12 in
%DDOC_FANFOLD_LGL_GERMAN = 41		'  German Legal Fanfold 8 1/2 x 13 in
'
'  dpStartDoc Option Constants
'==================================================================
%DDOC_VIEWBUILD = 1
%DDOC_ALLOWFAX = 2
%DDOC_ALLOWMAPI = 4
%DDOC_ALLOWSAVE = 8
%DDOC_ZOOMWIDTH = 64
%DDOC_ZOOMFIT = 128
%DDOC_ZOOM75 = 256
%DDOC_ZOOM100 = 512
%DDOC_ZOOM125 = 1024
%DDOC_SCALETEXT = 2048
%DDOC_FASTTEXT = 0
%DDOC_SAVE_EXE_ONLY = 4096
%DDOC_ALLOWSMTP = 8192
'
' Constants for dpDuplex
'==================================================================
%DDOC_DUPLEX_HORIZONTAL = 16384
%DDOC_DUPLEX_VERTICAL = 32768
'
'  Font Style Bit-wise constats for the dpFont call
'==================================================================
%DDOC_FONTNORMAL = 0
%DDOC_FONTBOLD = 1
%DDOC_FONTITALIC = 2
%DDOC_FONTUNDERLINE = 4
%DDOC_FONTSTRIKEOUT = 8
%DDOC_FONTBASELINE = 16
'
'  dpEndDoc Option constants
'==================================================================
%DDOC_END_VIEW = 0
%DDOC_END_PRINT = 1
%DDOC_END_CLOSE = 2
%DDOC_END_DELETE = 4
%DDOC_END_PRINT_NODIALOG = 8
%DDOC_END_FAX = 16
%DDOC_END_EMAIL = 32
%DDOC_END_SPECIFIC_PRINTER = 64
%DDOC_END_CREATE_EXE = 128
'
'  Text alignment constants for dpText, dpPageNo, dpPageCount
'==================================================================
%DDOC_LEFT = 0
%DDOC_CENTER = 1
%DDOC_RIGHT = 2
%DDOC_DECIMAL = 3
'
'  Error codes for dpSelfExtract
'==================================================================
%DDOC_INPUT_NOT_FOUND = - 5
%DDOC_OUTPUT_EXISTS = - 4
%DDOC_ERR_OPEN_INPUT = - 3
%DDOC_ERROR_OPEN_OUTPUT = - 2
%DDOC_ERROR_OPEN_STUB = - 1
%DDOC_SUCCESS = 0
'
'  Button codes to Assign hints (captions) to buttons
'  these will normally be used only by International users
'  to replace the English words on the screen. Used by
'  the dpCaption function
'============================================================================
%DDOC_INIT_WORDS = 0		' Initialization flag
%DDOC_MENU_FILE = 1		' File Menu
%DDOC_MENU_DOC = 2		' Document Menu
%DDOC_BTN_EXIT = 3		' Exit button hint and exit menu
%DDOC_BTN_FIRST = 4		' First Button hint and menu
%DDOC_BTN_PREV = 5		' Previous Button hint and menu
%DDOC_BTN_NEXT = 6		' Next button hint and menu
%DDOC_BTN_LAST = 7		' Last Button hint and menu
%DDOC_BTN_JUMP = 8		' Jump Button hint and menu
%DDOC_BTN_PRINT = 9		' Print hint and menu
%DDOC_WORD_ZOOM = 10		' Zoom label caption
%DDOC_WORD_ZOOMFIT = 11		' Zoom Fit word in drop down
%DDOC_WORD_ZOOMWIDTH = 12		' Zoom Width in drop down
%DDOC_WORD_DOCINDEX = 13		' Word Document index
%DDOC_WORD_PAGE = 14		' Page
%DDOC_WORD_OF = 15		' of
%DDOC_DLG_JUMP_TITLE = 16		' jump dialog title
%DDOC_DLG_JUMP_TEXT = 17		' dialog jump text
%DDOC_BTN_CANCEL = 18		' cancel button
%DDOC_BTN_OK = 19		' ok button
%DDOC_DLG_PW_TITLE = 20		' password box title
%DDOC_DLG_PW_TEXT = 21		' password text in dialog
%DDOC_DLG_PW_INVALID = 22		' password is invalid (try again)
%DDOC_DLG_PW_CANCEL = 23		' password invalid (3 strikes and you're out)
%DDOC_BTN_EMAIL = 24		' email button
%DDOC_BTN_FAX = 25		' fax button
%DDOC_BTN_SAVE = 26		' save button
'
'  Document generation
'==================================================================
DECLARE FUNCTION dpStartDoc LIB "ddoc32.dll" ALIAS "dpStartDoc" _
				( BYVAL hParent AS LONG, zTitle AS ASCIIZ, _
				zFile AS ASCIIZ, BYVAL iUOM%, BYVAL iPaper%, _
				BYVAL iOrient%, BYVAL iBin%, BYVAL iOptions% ) AS INTEGER
DECLARE SUB dpNewPage LIB "ddoc32.dll" ALIAS "dpNewPage" _
				( BYVAL iHandle%, BYVAL iPaper%, BYVAL iOrient%, BYVAL iBin% )
DECLARE SUB dpBookmark LIB "ddoc32.dll" ALIAS "dpBookmark" _
				( BYVAL iHandle%, zBookmark AS ASCIIZ )
DECLARE SUB dpEndDoc LIB "ddoc32.dll" ALIAS "dpEndDoc" _
				( BYVAL iHandle%, BYVAL iOptions% )
DECLARE FUNCTION dpEndDocInst LIB "ddoc32.dll" ALIAS "dpEndDocInst" AS LONG
DECLARE SUB dpGetFileName LIB "ddoc32.dll" ALIAS "dpGetFileName" _
				( BYVAL iHandle%, zDoc AS ASCIIZ, BYVAL iLen% )
DECLARE SUB dpChangeWords LIB "ddoc32.dll" ALIAS "dpChangeWords" _
				( BYVAL iHandle%, BYVAL iWhich%, zText AS ASCIIZ )
DECLARE FUNCTION dpSelfExtract LIB "ddoc32.dll" ALIAS "dpSelfExtract" _
				( zInputFile AS ASCIIZ, zDDOC AS ASCIIZ, _
				zOutEXE AS ASCIIZ ) AS INTEGER
DECLARE FUNCTION dpCheckAbort LIB "ddoc32.dll" ALIAS "dpCheckAbort" _
				( BYVAL iHandle% ) AS LONG
DECLARE SUB dpSetProgress LIB "ddoc32.dll" ALIAS "dpSetProgress" _
				( BYVAL iHandle%, BYVAL iProgress% )
DECLARE SUB dpSetTabs LIB "ddoc32.dll" ALIAS "dpSetTabs" _
				( BYVAL iHandle%, zTabStops AS ASCIIZ )
DECLARE SUB dpTabText LIB "ddoc32.dll" ALIAS "dpTabText" _
				( BYVAL iHandle%, BYVAL y!, zText AS ASCIIZ )
'
'  Document Options
'==================================================================
DECLARE SUB dpSetPassword LIB "ddoc32.dll" ALIAS "dpSetPassword" _
				( BYVAL iHandle%, zPassword AS ASCIIZ )
DECLARE SUB dpSetScreenOptions LIB "ddoc32.dll" ALIAS "dpSetScreenOptions" _
				( BYVAL iHandle%, BYVAL x&, BYVAL y&, BYVAL x2&, _
				BYVAL y2&, BYVAL iMaximize& )
DECLARE SUB dpSetAttachExt LIB "ddoc32.dll" ALIAS "dpSetAttachExt" _
				( BYVAL iHandle%, zExt AS ASCIIZ )
DECLARE SUB dpSetEMailOptions LIB "ddoc32.dll" ALIAS "dpSetEMailOptions" _
				( BYVAL iHandle%, zLogon AS ASCIIZ, zPassword AS ASCIIZ, _
				zNames AS ASCIIZ, zAddress AS ASCIIZ, zSubject AS ASCIIZ, _
				zNote AS ASCIIZ, BYVAL iShowDialog%, BYVAL iMailSelfExtract% )
DECLARE SUB dpSetFaxOptions LIB "ddoc32.dll" ALIAS "dpSetFaxOptions" _
				( BYVAL iHandle%, zLogon AS ASCIIZ, zPassword AS ASCIIZ, _
				zRecip AS ASCIIZ, zSubject AS ASCIIZ, zNote AS ASCIIZ, _
				BYVAL iShowDialog% )
DECLARE FUNCTION dpSMTPOptions LIB "ddoc32.dll" ALIAS "dpSMTPOptions" _
				( BYVAL hParent& ) AS LONG
DECLARE SUB dpDuplex LIB "ddoc32.dll" ALIAS "dpDuplex" _
				( BYVAL iHandle%, BYVAL iHow& )
DECLARE SUB dpPaperSize LIB "ddoc32.dll" ALIAS "dpPaperSize" _
				( BYVAL iHandle AS INTEGER, BYVAL sgWidth AS SINGLE, BYVAL sgHeight AS SINGLE )
'
'  Font and Text Functions
'==================================================================
DECLARE SUB dpFont LIB "ddoc32.dll" ALIAS "dpFont" _
				( BYVAL iHandle%, BYVAL iStyle%, BYVAL Size!, _
				BYVAL iColor&, zFontName AS ASCIIZ )
DECLARE SUB dpText LIB "ddoc32.dll" ALIAS "dpText" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL iAlign%, _
				zText AS ASCIIZ )
DECLARE SUB dpSpecialText LIB "ddoc32.dll" ALIAS "dpSpecialText" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL iAlign%, _
				zText AS ASCIIZ )
DECLARE SUB dpFontSpacing LIB "ddoc32.dll" ALIAS "dpFontSpacing" _
				( BYVAL iHandle AS INTEGER, BYVAL dblPoints AS DOUBLE )
DECLARE FUNCTION dpWrapText LIB "ddoc32.dll" ALIAS "dpWrapText" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, _
				BYVAL x2!, BYVAL y2!, BYVAL leading!, _
				zText AS ASCIIZ ) AS LONG
DECLARE FUNCTION dpWrapContinue LIB "ddoc32.dll" ALIAS "dpWrapContinue" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL x2!, BYVAL y2!, _
				BYVAL leading! ) AS LONG
DECLARE FUNCTION dpWrapCount LIB "ddoc32.dll" ALIAS "dpWrapCount" _
				( BYVAL iHandle% ) AS LONG
DECLARE SUB dpAngleText LIB "ddoc32.dll" ALIAS "dpAngleText" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, _
				BYVAL iAngle%, zText AS ASCIIZ )
DECLARE SUB dpClipText LIB "ddoc32.dll" ALIAS "dpClipText" _
				( BYVAL iHandle AS INTEGER, BYVAL x AS SINGLE, _
				BYVAL y AS SINGLE, BYVAL w AS SINGLE, _
				BYVAL iAlign AS INTEGER, zText AS ASCIIZ )
DECLARE SUB dpPageOffset LIB "ddoc32.dll" ALIAS "dpPageOffset" _
				( BYVAL iHandle AS INTEGER, BYVAL iOffset AS LONG )
DECLARE SUB dpPageNo LIB "ddoc32.dll" ALIAS "dpPageNo" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL iAlign%, _
				BYVAL iAngle% )
DECLARE SUB dpPageCount LIB "ddoc32.dll" ALIAS "dpPageCount" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL iAlign%, _
				BYVAL iAngle% )
DECLARE FUNCTION dpTextWidth LIB "ddoc32.dll" ALIAS "dpTextWidth" _
				( BYVAL iHandle%, zText AS ASCIIZ ) AS SINGLE
'
'  Lines and Graphics
'==================================================================
DECLARE SUB dpLine LIB "ddoc32.dll" ALIAS "dpLine" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, _
				BYVAL x2!, BYVAL y2!, BYVAL w!, BYVAL c AS LONG )
DECLARE SUB dpArc LIB "ddoc32.dll" ALIAS "dpArc" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL x2!, BYVAL y2!, _
				BYVAL aStart!, BYVAL aEnd!, BYVAL w!, BYVAL iColor& )
DECLARE SUB dpPie LIB "ddoc32.dll" ALIAS "dpPie" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL x2!, BYVAL y2!, _
				BYVAL aStart!, BYVAL aEnd!, BYVAL w!, BYVAL lineColor&, _
				BYVAL fillColor& )
DECLARE SUB dpRect LIB "ddoc32.dll" ALIAS "dpRect" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL x2!, BYVAL y2!, _
				BYVAL w!, BYVAL iFillColor&, BYVAL iBorderColor& )
DECLARE SUB dpEllipse LIB "ddoc32.dll" ALIAS "dpEllipse" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL x2!, BYVAL y2!, _
				BYVAL w!, BYVAL iFillColor&, BYVAL iBorderColor& )
DECLARE SUB dpGraphic LIB "ddoc32.dll" ALIAS "dpGraphic" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL x2!, BYVAL y2!, _
				zFileName AS ASCIIZ )
DECLARE SUB dpEmbedGraphic LIB "ddoc32.dll" ALIAS "dpEmbedGraphic" _
				( BYVAL iHandle%, BYVAL x!, BYVAL y!, BYVAL x2!, BYVAL y2!, _
				zFileName AS ASCIIZ )
DECLARE FUNCTION dpAddGraphic LIB "ddoc32.dll" ALIAS "dpAddGraphic" _
				( BYVAL iHandle%, zFile AS ASCIIZ ) AS LONG
DECLARE FUNCTION dpDrawGraphic LIB "ddoc32.dll" ALIAS "dpDrawGraphic" _
				( BYVAL iHandle%, BYVAL hBitmap&, BYVAL x!, BYVAL y!, _
				BYVAL x2!, BYVAL y2! ) AS LONG
DECLARE SUB dpStretchMode LIB "ddoc32.dll" ALIAS "dpStretchMode" _
				( BYVAL iHandle%, BYVAL iStretch& )
'
'  Printer Handling
'==================================================================
DECLARE FUNCTION dpPrinterCount LIB "ddoc32.dll" ALIAS "dpPrinterCount" AS LONG
DECLARE FUNCTION dpGetPrinter LIB "ddoc32.dll" ALIAS "dpGetPrinter" _
				( BYVAL iWhich AS LONG ) AS STRING
DECLARE FUNCTION dpIsPrinter LIB "ddoc32.dll" ALIAS "dpIsPrinter" _
				( zPrinter AS ASCIIZ ) AS LONG
DECLARE FUNCTION dpSpecifyPrinter LIB "ddoc32.dll" ALIAS "dpSpecifyPrinter" _
				( BYVAL iWhich AS INTEGER, zPrinter AS ASCIIZ ) AS LONG
DECLARE SUB dpPrinterCopies LIB "ddoc32.dll" ALIAS "dpPrinterCopies" _
				( BYVAL iWhich AS INTEGER, BYVAL iCopies AS LONG )
'
'  Mailing Label API
'==================================================================
DECLARE FUNCTION dpMailInit LIB "ddoc32.dll" ALIAS "dpMailInit" _
				( BYVAL iHandle AS INTEGER, szIniFile AS ASCIIZ, _
				szIniSection AS ASCIIZ, BYVAL nAlign AS INTEGER ) AS INTEGER
DECLARE FUNCTION dpLabelLines LIB "ddoc32.dll" ALIAS "dpLabelLines" _
				( BYVAL iHandle% ) AS INTEGER
DECLARE FUNCTION dpLabelsPerPage LIB "ddoc32.dll" ALIAS "dpLabelsPerPage" _
				( BYVAL iHandle% ) AS INTEGER
DECLARE SUB dpLabelText LIB "ddoc32.dll" ALIAS "dpLabelText" _
				( BYVAL iHandle%, szLabelLine AS ASCIIZ )
DECLARE SUB dpNextLabel LIB "ddoc32.dll" ALIAS "dpNextLabel" _
				( BYVAL iHandle% )
DECLARE SUB dpSetLabelLine LIB "ddoc32.dll" ALIAS "dpSetLabelLine" _
				( BYVAL iHandle%, BYVAL nLabelLine% )
DECLARE SUB dpSetLabel LIB "ddoc32.dll" ALIAS "dpSetLabel" _
				( BYVAL iHandle%, BYVAL nLabelNumber% )
DECLARE SUB dpPrintTemplate LIB "ddoc32.dll" ALIAS "dpPrintTemplate" _
				( BYVAL iHandle%, BYVAL LabAlign% )
DECLARE SUB dpLabelX LIB "ddoc32.dll" ALIAS "dpLabelX" _
				( BYVAL iHandle%, sngX AS SINGLE )
DECLARE SUB dpLabelY LIB "ddoc32.dll" ALIAS "dpLabelY" _
				( BYVAL iHandle%, sngY AS SINGLE )
