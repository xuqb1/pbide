''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' file : acompile.bas
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
$plus   = " OR "
$RegKey = "SOFTWARE\PowerBasic\PB/DLL\"
$CComment = "//"
$DlgBlock = "DIALOG "
$Begin    = "BEGIN"
$END      = "END"
$BLANK    = " "

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' file : lexer.bas
'
''   API         ''  ''    DDT        ''   ''     VB        ''  ''       RC          ''
%Api_BUTTON         = 1
%DDT_Button         = 1
%VB_COMMAND         = 1
%RC_PUSHBUTTON      = 1
%RC_DEFPUSHBUTTON   = 2

%DDT_Checkbox       = 3
%VB_CHECKBOX        = 3
%RC_CHECKBOX        = 3
%RC_AUTOCHECKBOX    = 4

%DDT_Radio          = 5
%VB_OPTION          = 5
%RC_RADIOBUTTON     = 5
%RC_AUTORADIOBUTTON = 6

%DDT_IMGBUTTON      = 7
%DDT_IMGBUTTONX     = 8

%DDT_CHECK3STATE    = 9
%RC_STATE3          = 9
%RC_AUTO3STATE      = 10

%DDT_FRAME          = 11
%VB_FRAME           = 11
%RC_GROUPBOX        = 11

%Api_STATIC         = 12
%DDT_Static         = 12

%DDT_LINE           = 13
%VB_LINE            = 13

%DDT_Label          = 14
%VB_LABEL           = 14
%RC_LTEXT           = 14

%RC_CTEXT           = 15

%RC_RTEXT           = 16

%DDT_Image          = 17
%VB_IMAGE           = 17
%RC_ICON            = 17
%DDT_ImageX         = 18
%VB_PICTUREBOX      = 18

%Api_SCROLLBAR      = 19
%DDT_ScrollBar      = 19
%VB_HSCROLL         = 19
%RC_SCROLLBAR       = 19
%VB_VSCROLL         = 20

%api_EDIT           = 21
%DDT_Edit           = 21
%VB_TEXTBOX         = 21
%RC_EDITTEXT        = 21

%Api_LISTBOX        = 22
%DDT_ListBox        = 22
%VB_LISTBOX         = 22
%RC_LISTBOX         = 22

%Api_COMBOBOX       = 23
%DDT_ComboBox       = 23
%VB_COMBOBOX        = 23
%RC_COMBOBOX        = 23

%VB_MENU            = 24
%VB_GRID            = 25
%RC_CONTROL         = 26
%CC_IMAGELIST       = 27 '<- commctl
%CC_LISTVIEW        = 28
%CC_PROGRESS        = 29
%CC_SLIDER          = 30
%CC_STATUSBAR       = 31
%CC_TABSTRIP        = 32
%CC_TOOLBAR         = 33
%CC_TREEVIEW        = 34
%CC_UPDOWN          = 35
'''''''''''''''''''''''''''''''''''''
' createwindow base classes
'$WC_BUTTON    = "BUTTON"
'$WC_COMBOBOX  = "COMBOBOX"
'$WC_EDIT      = "EDIT"
'$WC_LISTBOX   = "LISTBOX"
'$WC_MDICLIENT = "MDICLIENT"
'$WC_SCROLLBAR = "SCROLLBAR"
'$WC_STATIC    = "STATIC"


'''''''''''' pbdll DDT calls ''''''''''
$CW_EDIT        = "TEXTBOX"    'to create a edit
$CW_LINE        = "LINE"       'to create a static
$CW_LABEL       = "LABEL"
$CW_IMAGE       = "IMAGE"
$CW_IMAGEX      = "IMAGEX"
$CW_FRAME       = "FRAME"      'to create a button
$CW_OPTION      = "OPTION"
$CW_CHECKBOX    = "CHECKBOX"
$CW_IMGBUTTON   = "IMGBUTTON"
$CW_IMGBUTTONX  = "IMGBUTTONX"
$CW_CHECK3STATE = "CHECK3STATE"
$CW_RICHEDIT    = "RICHEDIT"

$Blank         = " "                ' single space
$vbTab         = "   "              ' triple space
$myTab         = "    "             ' quad space
$digi5         = "#####"

$Call          = "CALL "
$DialogNew     = "DIALOG NEW"
$DialogShow    = "DIALOG SHOW"
$ControlAdd    = "CONTROL ADD "
$BASECont      = $ControlAdd & "/classname, hDlg, /id, /txt, /x, /y, /h, /w, /style, /exstyle /" & $Call
$basedlg       = $DialogNew & " 0, /txt, /x, /y, /h, /w, /style TO hDlg"


'enum rc control para
%ctr_txt   = 1
%ctr_id    = 2
%ctr_class = 3
%ctr_style = 4
%ctr_x     = 5
%ctr_y     = 6
%ctr_xx    = 7
%ctr_yy    = 8
TYPE Const_type
    szConst AS ASCIIZ * 35
    dwConst AS DWORD
END TYPE

TYPE dialog_dims
    x        AS LONG
    y        AS LONG
    nwidth   AS LONG
    height   AS LONG
END TYPE

TYPE dialog_font
    lSize  AS LONG
    szName AS ASCIIZ * 80
END TYPE

TYPE dialog_block
    rcID      AS ASCIIZ * 80
    dims      AS dialog_dims
    STYLE     AS LONG
    caption   AS ASCIIZ * 80
    szMENU    AS ASCIIZ * 80
    szClass   AS ASCIIZ * 80
    FONT      AS dialog_font
    startLine AS LONG
    beginLine AS LONG
    endLine   AS LONG
END TYPE
GLOBAL localdir AS STRING
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' get api className
'
FUNCTION myGetClassName( BYVAL hWnd AS LONG ) AS STRING

  LOCAL lpClassName  AS ASCIIZ * 256

  GetClassName  hwnd, lpClassName,  SIZEOF(lpClassName)

  FUNCTION = lpClassName
END FUNCTION

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' translate window classes to strings
'-----------------------------------------------------------------------------
FUNCTION EnumClassName( BYVAL hWnd AS LONG ) AS LONG

  LOCAL lpClassName  AS ASCIIZ * 256
  LOCAL lRet AS LONG

  IF ISFALSE GetClassName ( hwnd, lpClassName,  SIZEOF(lpClassName) ) THEN EXIT FUNCTION

  SELECT CASE UCASE$(lpClassName)
    CASE $WC_EDIT      : lRet = 1
    CASE $WC_BUTTON    : lRet = 2
    CASE "RADIO"       : lRet = 3           ''' dialog style
    CASE $CW_CHECKBOX  : lRet = 4           ''' dialog style
    CASE $WC_STATIC    : lRet = 5
    CASE $WC_SCROLLBAR : lRet = 6
    CASE $WC_COMBOBOX  : lRet = 7
    CASE $WC_LISTBOX   : lRet = 8
    CASE "HANDLES"     : lRet = 1000  ''' skip this
  END SELECT

  FUNCTION = lRet
END FUNCTION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' exist checks fileName or Directory Name
' return error on filespec = $nul
'FUNCTION Exist( BYVAL FileSpec AS STRING ) AS LONG
'
'    LOCAL hFind AS DWORD
'    LOCAL fd    AS WIN32_FIND_DATA
'
'    hFind = FindFirstFile(BYCOPY FileSpec, fd)
'
'    IF hFind <> BITS???(%INVALID_HANDLE_VALUE) THEN
'  FindClose hFind
'  FUNCTION = %TRUE
'    END IF
'
'END FUNCTION
FUNCTION GetFile(fname AS STRING) AS STRING
    LOCAL a$
    LOCAL InF AS LONG

    ON ERROR GOTO GetFile_error
    InF = FREEFILE
    OPEN fname$ FOR BINARY AS #InF
  GET$  #InF, LOF(InF), a$
    CLOSE #InF

    FUNCTION = a$
    EXIT FUNCTION

GetFile_error:
    IF InF THEN CLOSE #InF
    FUNCTION = ""
    EXIT FUNCTION

END FUNCTION
SUB statAdd( sLine AS STRING, cArr() AS Const_type )

    LOCAL sConst AS STRING
    LOCAL sVal   AS STRING
    LOCAL sWrd   AS STRING

    LOCAL i      AS LONG
    LOCAL a      AS LONG
    LOCAL e      AS LONG

    sConst = TRIM$( PARSE$( sLine, "=", 1 ) )
    IF ISFALSE LEN(sConst) THEN EXIT SUB
    i = UBOUND( cArr() ) + 1
    REDIM PRESERVE cArr(i) AS Const_type

    cArr(i).szConst = sConst
    sVal  = PARSE$( sLine, "=", 2 )

    IF INSTR( sVal, "%" ) THEN
   sVal = TRIM$(sVal)
   FOR a = 1 TO PARSECOUNT (sVal, " " )
       sWrd = UCASE$( TRIM$( PARSE$( sVal, " ", a ) ) )
       IF sWrd = "OR" THEN ITERATE
       IF LEFT$( sWrd, 1 ) = "%" THEN
    ARRAY SCAN cArr(), FROM 1 TO LEN(sWrd), = sWrd, TO e
    IF e THEN DECR e : cArr(i).dwConst = cArr(i).dwConst OR cArr(e).dwConst
       ELSE
    cArr(i).dwConst = cArr(i).dwConst OR VAL (sWrd)
       END IF
   NEXT a
    ELSE
   cArr(i).dwConst = VAL( sVal )
    END IF

END SUB
FUNCTION statReadConst( BYVAL fName AS STRING, cArr() AS Const_type ) AS LONG

    LOCAL sFile    AS STRING
    LOCAL sLine    AS STRING

    LOCAL i        AS LONG
    LOCAL hfile    AS LONG
    LOCAL current  AS LONG
    LOCAL PREVIOUS AS LONG

    LOCAL letter   AS BYTE PTR

    IF exist( fName ) = 0 THEN EXIT FUNCTION           ''' exit if no source file
    REDIM cArr(-1) AS Const_type                       ''' empty existing cArr()
    sFile  = getFile( fName )                          ''' get text
    letter = STRPTR( sFile )                           ''' set pointer
    FOR i= 1 TO LEN( sFile ) - 1                       ''' scan file
       IF @Letter = 13 AND PEEK(letter+1) = 10 THEN    '''  for CRLF
    PREVIOUS = current                           ''' shift old pos
    current  = i                                 ''' line end & start
    sLine    = MID$( sFile, PREVIOUS+1, current-PREVIOUS-1 )
    IF LEFT$(sLine,1) = "%" THEN                 ''' check for static
    sLine = TRIM$( EXTRACT$( sLine, "'" ) )      ''' remove comment
    IF LEN( sLine ) THEN statAdd sLine, cArr()   ''' add to rslt array
    END IF                                       ''' %
    INCR current                                 ''' match zero based strptr to mid$ fct
       END IF
       INCR Letter
    NEXT i

    FUNCTION = UBOUND ( cArr )
END FUNCTION
FUNCTION Reg_Present ( BYVAL Hive AS LONG, KeyName AS ASCIIZ, ValName AS ASCIIZ) AS LONG

    LOCAL hKey AS LONG
    LOCAL lBuf AS LONG

    FUNCTION  = %True

    RegOpenKeyEx Hive, KeyName, 0, %KEY_QUERY_VALUE, hKey
       IF ISFALSE hKey THEN FUNCTION = %false : EXIT FUNCTION
       IF ISFALSE LEN(ValName) THEN GOTO Reg_PresentDone                 ' No ValName Specified, KeyName Exists
       RegQueryValueEx hKey, ValName, BYVAL 0, 0, BYVAL 0, lBuf
       IF lBuf = 0 THEN FUNCTION = %FALSE                                ' ValName NOT Present
       Reg_PresentDone:
    RegCloseKey hKey

END FUNCTION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' REGISTRY read wrappers
'
FUNCTION Reg_Get( Key AS ASCIIZ ) AS STRING

    LOCAL hKey    AS LONG
    LOCAL Buffer  AS ASCIIZ * 2048
    LOCAL lSIZE   AS LONG

    ' ** Open the section  #section
    RegOpenKeyEx %HKEY_CURRENT_USER, Key , 0, %KEY_QUERY_VALUE, hKey
  IF ISFALSE hKey THEN EXIT FUNCTION
  lSIZE = SIZEOF(Buffer)  ' ** Get the key value
  RegQueryValueEx hKey, "RcInclude", 0, 0, Buffer, lSIZE
    RegCloseKey hKey

    FUNCTION = Buffer
END FUNCTION
FUNCTION Reg_ApiIncDir( fName AS STRING ) AS STRING

    LOCAL sRet AS STRING

    IF Reg_Present(%HKEY_CURRENT_USER,$RegKey+"6.10","") THEN
       sRet = Reg_Get($RegKey+"6.10\Compiler") & fName
       IF exist( sRet ) THEN GOTO Reg_ApiIncDirDone
       EXIT FUNCTION
    END IF

    IF Reg_Present(%HKEY_CURRENT_USER,$RegKey+"6.00","") THEN
       sRet = Reg_Get($RegKey+"6.00\Compiler") & fName
       IF exist( sRet ) THEN GOTO Reg_ApiIncDirDone
       EXIT FUNCTION
    END IF

    IF Reg_Present(%HKEY_CURRENT_USER,$RegKey+"5.00","") THEN
       sRet = Reg_Get($RegKey+"5.00\Compiler") & fName
    END IF

    Reg_ApiIncDirDone:
    FUNCTION = sRet
END FUNCTION
FUNCTION ini_selectInclude( BYVAL hWnd AS LONG ) AS STRING

    LOCAL f AS STRING
    LOCAL fType AS STRING
    LOCAL sBuff AS STRING
    LOCAL Filter AS STRING


    fType  = "BAS"
    f      = "*." & fType
    Filter = "PB 文件 (*.INC, *.BAS)|*.INC;*.BAS|" & _
       "所有文件 (*.*)|*.*"

    IF ISFALSE OpenFileDialog( hWnd, "选择win32api.inc文件", f, _
             localdir, Filter, fType, _
             %OFN_EXPLORER OR %OFN_HIDEREADONLY) THEN EXIT FUNCTION


    FUNCTION = f
END FUNCTION
FUNCTION ini_getInclude() AS STRING

    LOCAL sp     AS settings

    Ini_Get "","",sp
    IF ISFALSE ISFILE ( sp.Include ) THEN sp.Include = Reg_ApiIncDir("win32api.inc")
    IF ISFALSE ISFILE ( sp.Include ) THEN sp.Include = ini_selectInclude( 0 )
    IF ISFALSE ISFILE ( sp.Include ) THEN EXIT FUNCTION
    Ini_Set "","",sp     ''

    FUNCTION = sp.Include
END FUNCTION
FUNCTION statExtract( cArr() AS Const_type, eArr() AS Const_type, BYVAL sWrd AS STRING, uBorder AS LONG ) AS LONG

    LOCAL lFound AS LONG
    LOCAL i      AS LONG
    LOCAL lCount AS LONG

    ARRAY SCAN cArr(), FROM 1 TO LEN(sWrd), = sWrd, TO lFound
    IF ISFALSE lFound THEN EXIT FUNCTION
    i = lfound - 1

    FOR i = lfound - 1 TO UBOUND( cArr )
       IF LEFT$( cArr(i).szConst, LEN(sWrd) ) <> sWrd THEN ITERATE
       REDIM PRESERVE eArr(lCount) AS Const_type
       eArr(lCount) = cArr(i)
       INCR lCount
       IF Uborder THEN IF lCount > Uborder THEN EXIT FOR
    NEXT i

    FUNCTION = lCount
END FUNCTION
FUNCTION statNum2Style( BYVAL dwStyle AS DWORD, eArr() AS Const_type ) AS STRING
  LOCAL i    AS LONG
  LOCAL sRet AS STRING

  IF dwStyle = 0 THEN EXIT FUNCTION

  FOR i = 0 TO UBOUND( eArr )
    IF eArr(i).DWConst <> dwStyle THEN ITERATE
    sRet = TRIM$( eArr(i).szConst )
  NEXT i

  FUNCTION = sRet
END FUNCTION
FUNCTION statEnumText( dwStyle AS DWORD, eArr() AS Const_type ) AS STRING
    LOCAL i AS LONG
    LOCAL sText AS STRING

    FOR i = 0 TO UBOUND( eArr )
  IF eArr(i).DWConst = 0 THEN ITERATE
  IF unmask ( dwStyle, eArr(i).DWConst  ) THEN _
     sText = sText & $plus & statNum2Style( eArr(i).DWConst, eArr() )
    NEXT i

    '''' trim leading/trailing OR OR ..
    WHILE LEFT$( sText, 4 ) = $plus : sText = MID$(sText, 5 )             : WEND
    WHILE RIGHT$( sText, 4 )= $plus : sText = LEFT$(sText, LEN(sText)-4 ) : WEND

    FUNCTION = sText
END FUNCTION

FUNCTION statStyle2Num( BYVAL sStyle AS STRING, eArr() AS Const_type  ) AS DWORD

    LOCAL I     AS LONG
    LOCAL dwRet AS DWORD

    FOR i = 0 TO UBOUND( eArr )
  IF LEFT$( eArr(i).szConst, LEN(sStyle) ) <> sStyle THEN ITERATE   '' compare left
  IF LEN( TRIM$( eArr(i).szConst ) ) > LEN( sStyle ) THEN ITERATE   '' compare right %SS_Right # %SS_RightJust
  dwRet = eArr(i).DWConst
    NEXT i

    FUNCTION = dwRet
END FUNCTION

FUNCTION statEnumStyle( sStyle AS STRING, eArr() AS Const_type ) AS DWORD


    LOCAL i       AS LONG
    LOCAL iMax    AS LONG
    LOCAL dwStyle AS DWORD

    iMax = PARSECOUNT( sStyle, " OR " )

    FOR i = 1 TO iMax
       dwStyle = dwStyle OR statStyle2Num( TRIM$(PARSE$(sStyle," OR ",i )), eArr() )
    NEXT i

    FUNCTION = dwStyle
END FUNCTION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CodegenProc helper fct
' wraps calls to unmask.bas
FUNCTION getWinStyles( BYVAL hWnd AS LONG, BYVAL WS_EX_flag AS LONG ) AS STRING


  DIM cArr(-1) AS STATIC Const_type             ''' compiler dummy - all constants
  DIM eArr(-1) AS LOCAL Const_type              ''' compiler dummy - work array
  LOCAL STYLE  AS DWORD
  LOCAL sStyle AS STRING


  IF hWnd = 0 THEN
    IF ISFALSE statReadConst(ini_getInclude,cArr()) THEN MSGBOX "Can't read win32api.inc"
    EXIT FUNCTION
  END IF

  IF hWnd = -1 THEN
     REDIM cArr(-1) AS STATIC Const_type
     EXIT FUNCTION
  END IF

  IF WS_EX_flag THEN
     STYLE = GetWindowLong(hwnd, %GWL_EXSTYLE)
     statExtract cArr(), eArr(), "%WS_EX_", 0
     sStyle = statEnumText( STYLE,  eArr() )
     ''' unknown
     IF STYLE THEN sStyle = sStyle + $plus + FORMAT$(STYLE)
     IF sStyle = "" THEN sStyle = "0"
     FUNCTION  = sStyle
     EXIT FUNCTION
  END IF

''    %WS_MAXIMIZEBOX     = &H10000&

  STYLE = GetWindowLong(hwnd, %GWL_STYLE)
  STYLE = STYLE AND NOT  %WS_CLIPSIBLINGS    '' remove %WS_CLIPSIBLINGS set by vfe
  statExtract cArr(), eArr(), "%WS_", 27     '' copy all %WS_ stats to eArr() - preparing 1st enumWinStyles call
  sStyle = statEnumText( STYLE,  eArr() )    '' removes all %WS_ ... equ's

  IF STYLE THEN
    SELECT CASE EnumClassName( hwnd )
'
    CASE 1     : statExtract cArr(), eArr(), "%ES_",  %null     '' prepare eArr()
    CASE 2,3,4 : statExtract cArr(), eArr(), "%BS_",  %null
    CASE 5     : statExtract cArr(), eArr(), "%SS_",  %null
    CASE 6     : statExtract cArr(), eArr(), "%SBS_", %null
    CASE 7     : statExtract cArr(), eArr(), "%CBS_", %null
    CASE 8     : statExtract cArr(), eArr(), "%LBS_", %null

    END SELECT

    sStyle = sStyle + " OR _" + $CRLF + statEnumText( STYLE,  eArr() )

  END IF

  FUNCTION = sStyle
END FUNCTION




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' api compile functions
'
FUNCTION AC_buildUniqueCtlId( BYVAL hWnd AS LONG, lpKeyName AS ASCIIZ ) AS LONG

  STATIC iButton     AS INTEGER
  STATIC iCheckbox   AS INTEGER
  STATIC iRadio      AS INTEGER

  ''' neu
  STATIC iIMGBUTTON  AS INTEGER
  STATIC iIMGBUTTONX AS INTEGER
  STATIC iCHECK3STATE AS INTEGER
  STATIC iFRAME      AS INTEGER

  STATIC iStatic     AS INTEGER

  ''' neu
  STATIC iLINE      AS INTEGER
  STATIC iLabel     AS INTEGER
  STATIC iImage     AS INTEGER
  STATIC iImageX    AS INTEGER

  STATIC iScrollBar  AS INTEGER
  STATIC iEdit       AS INTEGER
  STATIC iListBox    AS INTEGER
  STATIC iComboBox   AS INTEGER

  STATIC iControl    AS INTEGER

  IF ISFALSE hWnd THEN
     iButton    = 0
     iCheckbox  = 0
     iRadio     = 0
     iStatic    = 0
     iScrollBar = 0
     iEdit      = 0
     iListBox   = 0
     iComboBox  = 0
     iControl   = 0
     EXIT FUNCTION
  END IF

  SELECT CASE EnumClassName( hwnd )
    CASE 1 :    INCR iEdit      : lpKeyName = FORMAT$(iEdit)
    CASE 2 :    INCR iButton    : lpKeyName = FORMAT$(iButton)
    CASE 3 :    INCR iRadio     : lpKeyName = FORMAT$(iRadio)
    CASE 4 :    INCR iCheckbox  : lpKeyName = FORMAT$(iCheckbox)
    CASE 5 :    INCR iStatic    : lpKeyName = FORMAT$(iStatic)
    CASE 6 :    INCR iScrollBar : lpKeyName = FORMAT$(iScrollBar)
    CASE 7 :    INCR iComboBox  : lpKeyName = FORMAT$(iComboBox)
    CASE 8 :    INCR iListBox   : lpKeyName = FORMAT$(iListBox)
    CASE 1000 : EXIT FUNCTION                                               ' skip Grippers
    CASE ELSE : INCR iControl   : lpKeyName = FORMAT$(iControl)             ' unknown class
  END SELECT

  FUNCTION = %true
END FUNCTION



#IF 0

"BUTTON"

"CHECK3STATE"
    "CHECKBOX"
    "COMBOBOX"
    "FRAME"
    "IMAGE"
    "IMAGEX"
    "IMGBUTTON"
    "IMGBUTTONX"
    "LABEL"
    "LINE"

"LISTBOX"
    "OPTION"
    "SCROLLBAR"
    "TEXTBOX"
#ENDIF







FUNCTION AC_EnumCodeProc ( BYVAL hwnd AS LONG, lparam AS ASCIIZ  ) AS LONG

  STATIC iTotal      AS INTEGER

  LOCAL rc           AS RECT
  LOCAL rcParent     AS RECT

  LOCAL lpKeyName    AS ASCIIZ * 40
  LOCAL lpTemp       AS ASCIIZ * 1024
  LOCAL lpClassName  AS ASCIIZ * 256
  LOCAL lpWindowName AS ASCIIZ * 256

  FUNCTION = %TRUE       ' always continue enum
  IF ISFALSE AC_buildUniqueCtlId( hWnd, lpKeyName ) THEN EXIT FUNCTION

  INCR iTotal                                                   ' incr on every run to get unique ctl nums
'    GetClassName  hwnd, lpClassName,  SIZEOF(lpClassName)         ' registered class name
  GetWindowRect hwnd, rc                                        ' get control
  GetWindowRect GetParent(hwnd), rcParent                       ' get parent
  lpClassName = myGetClassName(hwnd)                            ' registered class name
  lpKeyName   = lpClassName & lpKeyName                         ' name controls handle
  lpTemp      = "GLOBAL h\Name AS LONG/" & _
          "%ID_\Name = \ID/" & _
          "h\Name = CreateWindowEx(\ExStyles, \Class, \Caption, \Styles, _/\x, \y, \w, \h, Form1, %ID_\Name, hInst, %Null)/" & _
          "SendMessage h\Name, %WM_SETFONT, GetStockObject(%DEFAULT_GUI_FONT), 0//"

  REPLACE "\Name"     WITH lpKeyName                                          IN lpTemp     ' handle Name
  REPLACE "\ID"       WITH FORMAT$(iTotal)                                    IN lpTemp     ' window id
  REPLACE "\Class"    WITH $DQ & lpClassName & $DQ                            IN lpTemp     ' class
  REPLACE "\Caption"  WITH $DQ & lpWindowName & $DQ                           IN lpTemp     ' caption
  REPLACE "\ExStyles" WITH getWinStyles( hWnd, %true )                        IN lpTemp     ' extended window style
  REPLACE "\Styles"   WITH getWinStyles( hWnd, %false  )                      IN lpTemp     ' window style
  REPLACE "\x"        WITH FORMAT$((rc.nleft - rcParent.nleft) - 2 )          IN lpTemp     ' left
  REPLACE "\y"        WITH FORMAT$((rc.ntop - rcParent.ntop) - 20 )           IN lpTemp     ' top
  REPLACE "\w"        WITH FORMAT$( rc.nRight-rc.nLeft )                      IN lpTemp     ' window width
  REPLACE "\h"        WITH FORMAT$( rc.nBottom-rc.nTop )                      IN lpTemp     ' window height
  REPLACE "/"         WITH $CRLF                                              IN lpTemp     ' insert linefeed
  REPLACE $NUL        WITH ""                                                 IN lpTemp     ' just to be shure

  lparam = lpTemp & lparam


END FUNCTION




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' enumerates all child windows of hWnd
'
FUNCTION AC_Compile(  BYVAL hVisualEdit AS LONG, szTemp AS ASCIIZ ) AS LONG


  LOCAL sHandle AS STRING
  LOCAL szForm  AS ASCIIZ * 2048

  IF ISFALSE hVisualEdit THEN MSGBOX "找不到可视化编辑窗口" : EXIT FUNCTION

  getWinStyles        0, %false                                                ' build win32api array
  AC_buildUniqueCtlId 0, BYVAL 0                                               ' reset ctl ID enum
  EnumChildWindows    hVisualEdit, CODEPTR(AC_EnumCodeProc), VARPTR( szTemp )  ' get all child controls next
  AC_EnumCodeProc     hVisualEdit, szForm                                      ' get parent's info
  getWinStyles        -1, %false                                               ' delete win32api array


  sHandle = PARSE$(szForm, $CRLF, 1)
  REPLACE "GLOBAL h" WITH "" IN sHandle                  ' post process remove var defs
  REPLACE " AS LONG" WITH "" IN sHandle
  REPLACE sHandle WITH "Wnd" IN szForm
  REPLACE "Form1" WITH "%HWND_DESKTOP" IN szForm
  REPLACE "Form1" WITH "hWnd"          IN szTemp         ' write correct parent win handle to childs
  szTemp = szForm & szTemp                               ' build source string parent win & all childs
  szForm = ""                                            ' clear for reuse

  LOCAL i AS LONG                                        ' CTL IDs & GLOBAL defs mixed with Code

  FOR i = 1 TO PARSECOUNT( szTemp, $CRLF )               ' extract GLOBAL vars to szForm
    sHandle = PARSE$(szTemp, $CRLF, i)
    IF LEFT$(sHandle,6) = "GLOBAL" THEN
       REPLACE sHandle & $CRLF WITH "" IN szTemp
       szForm = szForm + sHandle + $CRLF
    END IF
  NEXT i
  szTemp = szForm & szTemp
  szForm = ""

  FOR i = 1 TO PARSECOUNT( szTemp, $CRLF )               ' extract CTL IDs to szForm
    sHandle = PARSE$(szTemp, $CRLF, i)
    IF LEFT$(sHandle,4) = "%ID_" THEN
       REPLACE sHandle & $CRLF WITH "" IN szTemp
       szForm = szForm + sHandle + $CRLF
    END IF
  NEXT i
  szTemp = szForm & szTemp
  FUNCTION = LEN ( szTemp )

END FUNCTION



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' generate DDT stsyle code from Visual Editor Window
'
FUNCTION DC_ApiClass2DDT( szClass AS ASCIIZ, lStyle AS DWORD ) AS LONG

  LOCAL extraStyle AS LONG

  szClass = TRIM$(UCASE$(szClass))

  SELECT CASE szClass
    CASE $WC_BUTTON

       IF isMask( lStyle, %BS_CHECKBOX )       THEN szClass = $CW_CHECKBOX :    unmask lStyle, %BS_CHECKBOX
       IF isMask( lStyle, %BS_AUTORADIOBUTTON) THEN szClass = $CW_OPTION :      unmask lStyle, %BS_AUTORADIOBUTTON
       IF isMask( lStyle, %BS_AUTOCHECKBOX )   THEN szClass = $CW_CHECKBOX :    unmask lStyle, %BS_AUTOCHECKBOX
       IF isMask( lStyle, %BS_ICON )           THEN szClass = $CW_IMGBUTTON :   unmask lStyle, %BS_ICON
       IF isMask( lStyle, %BS_BITMAP )         THEN szClass = $CW_IMGBUTTONX :  unmask lStyle, %BS_BITMAP
       IF isMask( lStyle, %BS_AUTO3STATE )     THEN szClass = $CW_CHECK3STATE : unmask lStyle, %BS_AUTO3STATE
       IF isMask( lStyle, %BS_GROUPBOX )       THEN szClass = $CW_FRAME :       unmask lStyle, %BS_GROUPBOX


    CASE $WC_STATIC
       IF isMask( lStyle, %SS_ETCHEDFRAME )    THEN szClass = $CW_LINE :   unmask lStyle, %SS_ETCHEDFRAME
       IF isMask( lStyle, %SS_LEFT )           THEN szClass = $CW_LABEL :  unmask lStyle, %SS_LEFT
       IF isMask( lStyle, %SS_ICON )           THEN szClass = $CW_IMAGE :  unmask lStyle, %SS_ICON
       IF isMask( lStyle, %SS_BITMAP )         THEN szClass = $CW_IMAGEX : unmask lStyle, %SS_BITMAP


    CASE $WC_COMBOBOX, _  ''' unchanged
       $WC_LISTBOX, _
       $WC_SCROLLBAR

    CASE $WC_EDIT         : szClass = $CW_EDIT
    CASE ELSE             : szClass = $DQ & szClass & $DQ

  END SELECT


END FUNCTION


FUNCTION enumClass2Equ( szClass AS ASCIIZ ) AS STRING
  LOCAL sTemp AS STRING

  SELECT CASE szClass
      CASE $WC_BUTTON, _
         $CW_OPTION, _
         $CW_CHECKBOX, _
         $CW_IMGBUTTON, _
         $CW_IMGBUTTONX, _
         $CW_CHECK3STATE,_
         $CW_FRAME        : sTemp = "%BS_"
      CASE $WC_STATIC, _
         $CW_LINE, _
         $CW_LABEL, _
         $CW_IMAGE, _
         $CW_IMAGEX       : sTemp = "%SS_"
      CASE $WC_COMBOBOX     : sTemp = "%CBS_"
      CASE $WC_LISTBOX      : sTemp = "%LBS_"
      CASE $WC_SCROLLBAR    : sTemp = "%SBS_"
      CASE $WC_EDIT         : sTemp = "%ES_"
  END SELECT

  FUNCTION = sTemp
END FUNCTION



SUB newDummyDlg( hDummy AS LONG )
  ' build dummy dialog
  DIALOG NEW 0, "",,,0,0 TO hDummy
  DIALOG SHOW MODELESS hDummy
END SUB


SUB pixel2unit( hDummy AS LONG, r AS rect, dims AS dialog_dims )

  IF ISFALSE hDummy THEN EXIT SUB
  ' translate pixels to units
  DIALOG PIXELS hDummy, r.nLeft, r.nTop         TO UNITS dims.x, dims.y
  DIALOG PIXELS hDummy, r.nRight-r.nLeft, r.nBottom-r.nTop TO UNITS dims.nWidth, dims.Height

END SUB




SUB DDT_ClientRect( r AS rect )
  r.nBottom = r.nBottom - GetSystemMetrics( %SM_CYCAPTION )
  r.nBottom = r.nBottom - GetSystemMetrics( %SM_CYFRAME ) * 2
  r.nRight  = r.nRight  - GetSystemMetrics( %SM_CXFRAME ) * 2
END SUB


SUB DDT_topLeft( r AS rect )
  r.ntop  = r.ntop + GetSystemMetrics( %SM_CYCAPTION ) + GetSystemMetrics( %SM_CXFRAME )
  r.nleft = r.nleft + GetSystemMetrics( %SM_CYFRAME )
END SUB



SUB DDT_MoveRect( r AS rect, rPrnt AS rect )
  r.nleft   = r.nleft - rPrnt.nleft
  r.nRight  = r.nRight - rPrnt.nleft
  r.ntop    = r.ntop  - rPrnt.ntop
  r.nBottom = r.nBottom - rPrnt.ntop
END SUB


SUB KillVfeStyles( dwStyle AS DWORD )
  dwStyle = dwStyle AND NOT %WS_CLIPCHILDREN        ''' remove style paras set by visual editor
  dwStyle = dwStyle AND NOT %WS_CLIPSIBLINGS
END SUB


SUB insertDims( sTemp AS STRING, dims AS dialog_dims )
  REPLACE "/x"     WITH FORMAT$( dims.x )      IN sTemp
  REPLACE "/y"     WITH FORMAT$( dims.y  )     IN sTemp
  REPLACE "/h"     WITH FORMAT$( dims.nWidth ) IN sTemp
  REPLACE "/w"     WITH FORMAT$( dims.height ) IN sTemp
END SUB


FUNCTION DC_Compile(  BYVAL hVisualEdit AS LONG, szTemp AS ASCIIZ ) AS LONG

  LOCAL rPrnt     AS rect
  LOCAL r         AS rect
  LOCAL dims      AS dialog_dims

  LOCAL szForm    AS ASCIIZ * 2048
  LOCAL szClass   AS ASCIIZ * 256

  LOCAL sTemp     AS STRING
  LOCAL sStyle    AS STRING
  LOCAL sStyle2   AS STRING
  LOCAL sExSytle  AS STRING
  LOCAL Ids       AS STRING

  LOCAL i         AS LONG
  LOCAL hDlg      AS LONG
  LOCAL hCtl      AS LONG
  LOCAL dwStyle   AS DWORD
  LOCAL dwExStyle AS DWORD

  DIM cArr(-1) AS LOCAL Const_type             ''' compiler dummy - all constants
  DIM eArr(-1) AS LOCAL Const_type             ''' compiler dummy - work array

  '''' check paras
  IF ISFALSE hVisualEdit THEN MSGBOX "找不到可视化编辑窗口"                    : EXIT FUNCTION
  IF ISFALSE statReadConst(ini_getInclude,cArr()) THEN MSGBOX "无法读取win32api.inc" : EXIT FUNCTION

  '''' collect data
  newDummyDlg    hDlg
  getWindowRect  hVisualEdit, rPrnt
  DDT_ClientRect rPrnt
  pixel2unit     hDlg, rPrnt, dims
  DDT_topLeft    rPrnt

  GetWindowText  hVisualEdit, szForm, 2048
  GetWindowLong  hVisualEdit, %GWL_STYLE TO dwStyle
  KillVfeStyles  dwStyle

  '''' get dlg window equates
  statExtract cArr(), eArr(), "%WS_", 27
  sStyle = statEnumText( dwStyle, eArr() )          '' WS_ styles
  statExtract cArr(), eArr(), "%DS_", %null
  sStyle2 = statEnumText( dwStyle, eArr() )         '' DS_ styles
  IF LEN( sStyle2 ) AND LEN( sStyle ) THEN sStyle2 = $plus & sStyle2
  sStyle = sStyle & sStyle2

  ''''' build dailog new ...
  sTemp = $basedlg
  insertDims sTemp, dims
  REPLACE "/txt"   WITH $DQ & szForm & $DQ IN sTemp
  REPLACE "/style" WITH sStyle             IN sTemp
  szTemp = $myTab & sTemp & $CRLF


  FOR I = LBOUND( gZorder() ) TO UBOUND( gZorder() )
    hCtl =  gZorder(i) : IF ISFALSE hCtl THEN ITERATE

    szClass = myGetClassName(hCtl)                     ''' get class
    GetWindowLong   hCtl, %GWL_STYLE TO dwStyle        ''' get style
    DC_ApiClass2DDT szClass, dwStyle
    KillVfeStyles   dwStyle
    GetWindowText   hCtl, szForm, 2048                 ''' get Text
    getWindowRect   hCtl, r                            ''' get size
    DDT_MoveRect    r, rPrnt
    pixel2unit      hDlg, r, dims
    statExtract cArr(), eArr(), "%WS_", 27             ''' WS_ styles
    sStyle = statEnumText( dwStyle, eArr() )
    sTemp  = enumClass2Equ( szClass )
    IF LEN(sTemp) THEN statExtract cArr(), eArr(), sTemp, %null
    sStyle2 = statEnumText( dwStyle, eArr() )         '' control styles
    IF LEN( sStyle2 ) AND LEN( sStyle ) THEN sStyle2 = $plus & sStyle2
    sStyle = sStyle & sStyle2

    dwExStyle = GetWindowLong(  hCtl, %GWL_EXSTYLE )    ''' get extended win styles
    IF dwExStyle THEN
       statExtract cArr(), eArr(), "%WS_EX_", %null
       sExSytle = statEnumText( dwExStyle, eArr() )
    END IF

    Ids   = "%" & szClass & FORMAT$(i)
    sTemp = LEFT$( $BASECont, 73 )
    insertDims sTemp, dims
    REPLACE "/classname"  WITH szClass            IN sTemp
    REPLACE "/id"         WITH Ids                IN sTemp
    REPLACE "/txt"        WITH $DQ & szForm & $DQ IN sTemp
    REPLACE "/style"      WITH sStyle             IN sTemp
    REPLACE "/exstyle"    WITH sExSytle           IN sTemp

    szTemp = szTemp & $myTab & sTemp & $CRLF

  NEXT i

  DIALOG END hDlg                                            ''' remove dummy

  FUNCTION = LEN ( szTemp )
END FUNCTION
FUNCTION cbMain_APICOMPILE  (BYVAL hWnd AS LONG) AS LONG

    LOCAL szTemp AS ASCIIZ * 32000
  LOCAL hMdi AS DWORD
  LOCAL pSci AS LONG
  LOCAL szFileName AS STRING
    LOCAL nTab AS DWORD
    LOCAL findFlags AS DWORD
    LOCAL startPos AS LONG
    LOCAL endPos AS LONG
  LOCAL strTemp AS STRING

    MOUSEPTR 11
    apisleep 0
    IF ISFALSE AC_Compile( hWndEdit, szTemp ) THEN                   EXIT FUNCTION
'    sChld.hCodegen = NewCodeWindow( hWnd )
'    IF ISFALSE sChld.hCodegen THEN EXIT FUNCTION
'    ' color off, new text, default font in edit window
'    FUNCTION = sendmessage( sChld.hCodegen, %WM_COMMAND, %IDBUILD, VARPTR( szTemp ) )
  IF GetEdit THEN UPDATEWINDOW GetEdit
  hMdi = CreateMdiChild( $EDITCLASSNAME, g_hWndClient, "", IIF( fMaximize, %WS_MAXIMIZE, 0 ))
  IF ISFALSE hMdi THEN EXIT FUNCTION
  SHOWWINDOW hMdi, %SW_SHOW
  ' Get a direct pointer for faster access
  pSci = SENDMESSAGE( GetEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT FUNCTION
  ' Insert the text at the current position
  strTemp=szTemp
  SciMsg pSci, %SCI_INSERTTEXT, 0, BYVAL STRPTR( strTemp )
  ' Change the title of the window
  szFileName = "Untitled" & FORMAT$( cUntitledFiles ) & ".BAS"
  SETWINDOWTEXT hMdi,"Untitled" & FORMAT$( cUntitledFiles ) & ".BAS" 'szFileName
  ' Change the tab name
  nTab = TabCtrl_GetCurSel( g_hTabMdi )
  CALL SetTabName( nTab, szFileName )
  ' Change the name in the array
  IF nTab < = UBOUND( gTabFilePaths ) THEN
    gTabFilePaths( nTab ) = szFileName
  END IF
  ' Change the editor options
  Scintilla_SetOptions pSci, szFileName
  ' Find flags
  findFlags = %SCFIND_MATCHCASE OR %SCFIND_WHOLEWORD
  ' Begin to search from the current position
  startPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )
  ' End position = length of the document
  endPos = SENDMESSAGE( GetEdit, %SCI_GETTEXTLENGTH, 0, 0 )

  ' Change statusbar information
  ShowLinCol
  MOUSEPTR 0
END FUNCTION



FUNCTION cbMain_DDTCOMPILE  (BYVAL hWnd AS LONG) AS LONG

    'LOCAL edit   AS edit_type
    LOCAL szTemp AS ASCIIZ * 32000
    'EDIT_Get "", "", edit
  LOCAL hMdi AS DWORD
  LOCAL pSci AS LONG
  LOCAL szFileName AS STRING
    LOCAL nTab AS DWORD
    LOCAL findFlags AS DWORD
    LOCAL startPos AS LONG
    LOCAL endPos AS LONG
  LOCAL strTemp AS STRING
    MOUSEPTR 11
    apisleep 0
    IF ISFALSE DC_Compile( hWndEdit, szTemp ) THEN EXIT FUNCTION
'    sChld.hCodegen = NewCodeWindow( hWnd )
'    IF ISFALSE sChld.hCodegen THEN EXIT FUNCTION
'    FUNCTION = sendmessage( sChld.hCodegen, %WM_COMMAND, %IDBUILD, VARPTR( szTemp ) )
  IF GetEdit THEN UPDATEWINDOW GetEdit
  hMdi = CreateMdiChild( $EDITCLASSNAME, g_hWndClient, "", IIF( fMaximize, %WS_MAXIMIZE, 0 ))
  IF ISFALSE hMdi THEN EXIT FUNCTION
  SHOWWINDOW hMdi, %SW_SHOW
  ' Get a direct pointer for faster access
  pSci = SENDMESSAGE( GetEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
  IF ISFALSE pSci THEN EXIT FUNCTION
  ' Insert the text at the current position
  strTemp=szTemp
  SciMsg pSci, %SCI_INSERTTEXT, 0, BYVAL STRPTR( strTemp )
  ' Change the title of the window
  szFileName = "Untitled" & FORMAT$( cUntitledFiles ) & ".BAS"
  SETWINDOWTEXT hMdi,"Untitled" & FORMAT$( cUntitledFiles ) & ".BAS" 'szFileName
  ' Change the tab name
  nTab = TabCtrl_GetCurSel( g_hTabMdi )
  CALL SetTabName( nTab, szFileName )
  ' Change the name in the array
  IF nTab < = UBOUND( gTabFilePaths ) THEN
    gTabFilePaths( nTab ) = szFileName
  END IF
  ' Change the editor options
  Scintilla_SetOptions pSci, szFileName
  ' Find flags
  findFlags = %SCFIND_MATCHCASE OR %SCFIND_WHOLEWORD
  ' Begin to search from the current position
  startPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )
  ' End position = length of the document
  endPos = SENDMESSAGE( GetEdit, %SCI_GETTEXTLENGTH, 0, 0 )

  ' Change statusbar information
  ShowLinCol
  MOUSEPTR 0
END FUNCTION
FUNCTION isMask( dwValue AS DWORD, dwMask AS DWORD ) AS LONG
    FUNCTION = (dwValue AND dwMask) = dwMask
END FUNCTION


FUNCTION unmask( dwValue AS DWORD, dwMask AS DWORD ) AS LONG
    IF ISFALSE isMask (dwValue, dwMask) THEN EXIT FUNCTION

    dwValue  = dwValue XOR dwMask
    FUNCTION = %true
'    END IF
END FUNCTION
