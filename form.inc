' 窗口编辑实现
SUB NormalizeRect(BYREF ptFirst AS POINTAPI,BYREF ptSecond AS POINTAPI)
  LOCAL nTemp AS INTEGER
  IF ptFirst.x > ptSecond.x THEN
    nTemp      = ptFirst.x
    ptFirst.x  = ptSecond.x
    ptSecond.x = nTemp
  END IF
  IF ptFirst.y > ptSecond.y THEN
    nTemp      = ptFirst.y
    ptFirst.y  = ptSecond.y
    ptSecond.y = nTemp
  END IF
END SUB
FUNCTION HandlesProc( BYVAL hWnd AS LONG, _
          BYVAL wMsg AS LONG, BYVAL wParam AS LONG, _
          BYVAL lParam AS LONG ) AS LONG
  LOCAL pt AS POINTAPI
  SELECT CASE wMsg
    CASE %WM_LBUTTONDOWN, %WM_MOUSEMOVE
      pt.x = LOWRD( lParam )
      pt.y = HIWRD( lParam )
      CLIENTTOSCREEN hWnd, pt
      SCREENTOCLIENT GETPARENT( hWnd ), pt
      POSTMESSAGE GETPARENT( hWnd ), wMsg, wParam, MAKLNG( pt.x, pt.y )
    CASE %WM_MOUSEACTIVATE : FUNCTION = %MA_NOACTIVATE              'Defeat this message so that mouse clicks do not activate the control
    CASE %WM_SETCURSOR : FUNCTION = %TRUE           'Defeat the system changing cursors on us. We do it based on our own hit testing
  END SELECT
  FUNCTION = DEFWINDOWPROC( hWnd, wMsg, wParam, lParam )
END FUNCTION
SUB InitializeHandlesClass( BYVAL hWnd AS LONG, hHandle( ) AS LONG )
  LOCAL i AS LONG
  LOCAL wc AS WNDCLASSEX
  LOCAL szAppName AS ASCIIZ * 20
  '**Register Handles Window Class
  szAppName = "HANDLES"
  wc.cbSize = SIZEOF( wc )
  wc.style = %CS_SAVEBITS
  wc.lpfnWndProc = CODEPTR( HandlesProc )
  wc.hInstance = GETMODULEHANDLE( BYVAL 0 )
  wc.hCursor = LOADCURSOR ( %NULL, BYVAL %IDC_ARROW )
  wc.hbrBackground = CREATESOLIDBRUSH( RGB( 255, 0, 0 ))
  wc.lpszClassName = VARPTR( szAppName )
  REGISTERCLASSEX wc
  FOR i = 0 TO 7          '' UBOUND( sh.hHandles )
    '*Create the Handle Windows -------------------------------------------------
    hHandles( i ) = CREATEWINDOWEX(0, szAppName, "", _
                    %WS_MINIMAL OR %WS_OVERLAPPED, _
                    0, 0, 6, 6, hWnd, BYVAL %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL)
  NEXT
END SUB
FUNCTION WndClientProc_Create( BYVAL hWnd AS LONG, sf AS flag_type, sp AS setting_type ) AS LONG
  LOCAL IniFile AS STRING
  IniFile=EXE.PATH$ & "config.ini"
  sf.Dragging = %OFF              ' Initialize some flags
  sf.Sizing = %OFF
  sf.cType = 0            ' Turn on selection tool
  sf.DrawNewControl = %OFF                'Turn off drawing now
  sf.LostCapture = %OFF
  sf.changed = %FALSE
  hWndControl = 0                 'No controls selected yet
'  hPopupCtl      = LoadMenu(GetModuleHandle(BYVAL 0), "CONTROLMN")   ' load the Control Popup Context Edit Menu
'  hPopupCtl      = GetSubMenu(hPopupCtl, 0)
'  hPopupClient   = LoadMenu(GetModuleHandle(BYVAL 0), "CLIENTMN")    ' load the Clients Popup Reserved Menu
'  hPopupClient   = GetSubMenu(hPopupClient, 0)
  MENU NEW POPUP TO hPopupCtl
  MENU ADD STRING,hPopupCtl,"撤销",%MN_UNDO,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"重做",%MN_REDO,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"剪切",%MN_CUT,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"复制",%MN_COPY,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"粘贴",%MN_PASTE,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"删除",%MN_DELETE,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"上移一层",%MN_UPLAYER,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"下移一层",%MN_DOWNLAYER,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"移到顶层",%MN_TOPLAYER,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"移到底层",%MN_BOTTOMLAYER,%MF_ENABLED
  MENU ADD STRING,hPopupCtl,"属性",%MN_CTLPROP,%MF_ENABLED

  MENU NEW POPUP TO hPopupClient
  MENU ADD STRING,hPopupClient,"撤销",%MN_F_UNDO,%MF_ENABLED
  MENU ADD STRING,hPopupClient,"重做",%MN_F_REDO,%MF_ENABLED
  MENU ADD STRING,hPopupClient,"DDT代码",%MN_F_DDTCODE,%MF_ENABLED
  MENU ADD STRING,hPopupClient,"SDK代码",%MN_F_SDKCODE,%MF_ENABLED
  MENU ADD STRING,hPopupClient,"属性",%MN_F_PROP,%MF_ENABLED

  REDIM hHandles( 8 )
  InitializeHandlesClass hWnd, hHandles( )                ' Initialize the 8 Sizing Handles
  Ini_Get IniFile, "设计器", sp           ' load ini file
END FUNCTION
FUNCTION WndClientProc_CAPTURECHANGED( BYVAL hWnd AS LONG, BYVAL wParam AS LONG, BYVAL lParam AS LONG, sf AS flag_type ) AS LONG
  'Used for Windows version 4.0 and up only
  'Don't do it if WE are Releasing the mouse capture!
  IF sf.Drawing OR sf.Dragging OR sf.Sizing THEN EXIT FUNCTION
  IF sf.LostCapture = %ON THEN
    sf.LostCapture = %OFF
    'Clean up any GDI garbage in some certain circumstances
    INVALIDATERECT lparam, BYVAL %NULL, %TRUE
    FUNCTION = DEFWINDOWPROC( hWnd, %WM_CAPTURECHANGED, wParam, lParam )
  END IF
END FUNCTION
FUNCTION pickPrinter( BYVAL hWnd AS LONG, BYVAL hWndNotify AS LONG, BYVAL pdFlags AS LONG ) AS LONG
  LOCAL szText AS ASCIIZ * 64
  LOCAL PD AS PRINTDLGAPI
  LOCAL di AS DOCINFO
  pd.lStructSize = SIZEOF( pd )
  pd.Flags = pdFlags
  pd.hWndOwner = hWnd
  IF ISFALSE PRINTDLG( PD ) THEN EXIT FUNCTION
  GETWINDOWTEXT hWnd, szText, 64
  di.cbsize = SIZEOF( di )
  di.lpszDocName = VARPTR( szText )
  IF ISFALSE STARTDOC( PD.hDC, di ) THEN EXIT FUNCTION
  IF ISFALSE STARTPAGE( PD.hDC ) THEN EXIT FUNCTION
  CALL SENDMESSAGE( hWndNotify, %WM_USER OR %WM_PRINT, PD.hDC, %PRF_CLIENT OR %PRF_ERASEBKGND OR %PRF_NONCLIENT )
  ENDPAGE PD.hDC
  ENDDOC PD.hDC
  DELETEDC PD.hDC
END FUNCTION
FUNCTION WndClientProc_COMMAND ( BYVAL hWnd AS LONG, BYVAL wParam AS LONG ) AS LONG
  SELECT CASE wParam
    'CASE %IDM_FORM_PRINT : pickPrinter hWnd, hWnd, %PD_NOSELECTION OR %PD_NOPAGENUMS OR %PD_USEDEVMODECOPIES OR %PD_RETURNDC
    'Control popup menu
    CASE %MN_UNDO
      MSGBOX "Control Undo"
      EXIT FUNCTION
    CASE %MN_REDO
    CASE %MN_CUT
      MSGBOX "Control Cut"
      EXIT FUNCTION
    CASE %MN_COPY
    CASE %MN_PASTE
    CASE %MN_DELETE
    CASE %MN_UPLAYER
    CASE %MN_DOWNLAYER
    CASE %MN_TOPLAYER
    CASE %MN_BOTTOMLAYER
    CASE %MN_CTLPROP
    'Forms client popup menu
    CASE %MN_F_UNDO     'reserved
    CASE %MN_F_REDO     'reserved
    CASE %MN_F_DDTCODE
      cbMain_DDTCOMPILE hWnd
    CASE %MN_F_SDKCODE
      cbMain_APICOMPILE hWnd
    CASE %MN_F_PROP     'reserved
  END SELECT
END FUNCTION
SUB SetZorder(BYVAL hWnd AS LONG) 'Pass the currently selected control handle

  LOCAL POS AS LONG   'Position number
  LOCAL i   AS LONG   'Index number

  'The window is already on top, don't do anything
  IF hWnd = gZorder(0) THEN EXIT SUB

  'Get the relative index position of this window
  ARRAY SCAN gZorder(0), = hWnd ,TO POS
  'change to zero based
  DECR POS

  'Shift each element in the array down by 1 to the position of the
  'window we will change to the top of the z-order
  FOR i = 0 TO POS-1
    gZorder(POS-i) = gZorder(POS-1-i)
  NEXT

  'Last step, put the window into the top of the z-order list
  gZorder(0) = hWnd

END SUB
'-------------------------------------------------------------
' Load up the new Handles coordinates/size into an array for
' further processing by DrawHandles() and FigureHandles()
' called from FigureHandleSizes only
'-------------------------------------------------------------
SUB MakeHandle(DialogHandle AS GRIPPERS, _
         BYVAL x  AS LONG,BYVAL y  AS LONG, _
         BYVAL dx AS LONG,BYVAL dy AS LONG, _
         BYVAL handleSize AS LONG, _
         BYVAL dwHitValue AS WORD)


  DialogHandle.rcArea.nleft   = x
  DialogHandle.rcArea.ntop    = y
  DialogHandle.rcArea.nright  = x+dx + handleSize
  DialogHandle.rcArea.nbottom = y+dy + handleSize
  DialogHandle.dwNCHITVALUE   = dwHitValue

END SUB



'--------------------------------------------------------------
' Determine the location of the 8 Grip Handles
' on the currently selected Control.
' called from DrawHandles only
'--------------------------------------------------------------
SUB FigureHandleSizes(BYVAL handleSize AS LONG )

  LOCAL rc         AS RECT
  LOCAL dx         AS LONG
  LOCAL dy         AS LONG

  '**Get the width and height of the window frame
  '  we want to put handles on.
  dx = GetSystemMetrics(%SM_CXFRAME)
  dy = GetSystemMetrics(%SM_CYFRAME)

  GetWindowRect hWndControl, rc

  MakeHandle DialogHandles(0), rc.nleft, (rc.ntop+rc.nbottom) \ 2 - dy \ 2-1, dx, dy, handleSize, %HTLEFT
  MakeHandle DialogHandles(1), rc.nright-dx-handleSize, (rc.ntop+rc.nbottom) \ 2-dy \ 2-1, dx, dy, handleSize, %HTRIGHT
  MakeHandle DialogHandles(2), (rc.nleft+rc.nright) \ 2-dx \ 2 -1, rc.ntop, dx, dy, handleSize, %HTTOP
  MakeHandle DialogHandles(3), rc.nleft, rc.ntop, dx, dy, handleSize, %HTTOPLEFT
  MakeHandle DialogHandles(4), rc.nright-dx-handleSize, rc.ntop, dx, dy, handleSize, %HTTOPRIGHT
  MakeHandle DialogHandles(5), (rc.nleft+rc.nright) \ 2-dx \ 2 -1, rc.nbottom-dy-handleSize, dx, dy,handleSize, %HTBOTTOM
  MakeHandle DialogHandles(6), rc.nleft, rc.nbottom-dy-handleSize, dx, dy, handleSize, %HTBOTTOMLEFT
  MakeHandle DialogHandles(7), rc.nright-dy-handleSize, rc.nbottom-dy-handleSize, dx, dy,handleSize, %HTBOTTOMRIGHT


END SUB



SUB HideHandles( )
   LOCAL i AS INTEGER
   FOR i = 0 TO 7
     ShowWindow hHandles(i), %SW_HIDE
   NEXT
END SUB
'-------------------------------------------------------------------------
' Draw the 8 sizing handles on top of the currently selected control
' Handle 3 cases, 1. Controls/Windows
'                 2. Objects/Shapes     <-- ToDo:
'                 3. Form/Dialog        <-- ToDo:
' Entry: the hWnd to the Control to put the Sizing handles on.
'-------------------------------------------------------------------------
SUB DrawHandles ( handleSize AS LONG  )

  LOCAL i   AS LONG
  LOCAL rc  AS RECT

  REDIM DialogHandles(8)
  'redim hHandles(8)

  IF ISFALSE hWndControl THEN EXIT SUB
  FigureHandleSizes handleSize  '', sh.DialogHandles()


  FOR  i = 0 TO 7  '' UBOUND( sh.hHandles() )
    rc = DialogHandles(i).rcArea
    MapWindowPoints %HWND_DESKTOP, getparent(hWndControl), rc,2
    SetWindowPos    hHandles(i),%HWND_TOP, rc.nLeft, rc.nTop, rc.nRight-rc.nLeft, rc.nBottom-rc.nTop, %SWP_SHOWWINDOW
    InvalidateRect  hHandles(i), BYVAL %NULL,%TRUE
  NEXT

END SUB

'---------------------------------------------------------------------------
' After a control has been selected, we first clean the handles from the
' previous control then repaint them on the new control after the Move or
' Sizing handler is finished (or if select Form or draw new control).
'--------------------------------------------------------------------------
SUB ClearHandles( )

  HideHandles                    '<--Use new HideHandles routine
  IF ISFALSE  hWndControl THEN EXIT SUB
  'This Call helps to clean off the attached objects, ie EditControl+ScrollBar,
  'specifically this switch->%SWP_DRAWFRAME
  SetWindowPos   hWndControl,%HWND_TOP,0,0,0,0,%SWP_DRAWFRAME OR %SWP_NOSIZE OR %SWP_NOMOVE OR %SWP_NOZORDER
  InvalidateRect hWndControl, BYVAL %NULL, %TRUE                                'Give the Control an Update
  UpdateWindow   hWndControl

END SUB

FUNCTION IsGroupBox(BYVAL hWnd AS LONG) AS LONG

    LOCAL ClassName AS ASCIIZ*40

    GetClassName hWnd,BYVAL VARPTR(ClassName),SIZEOF(ClassName)
    ClassName = UCASE$( ClassName )

    IF ClassName = $WC_BUTTON  THEN
       IF isMask( GetWindowLong(hWnd, %GWL_STYLE), %BS_GROUPBOX )    THEN FUNCTION = 1 :  EXIT FUNCTION
    END IF
    IF ClassName = $WC_STATIC THEN
       IF isMask( GetWindowLong(hWnd, %GWL_STYLE), %SS_ETCHEDFRAME ) THEN FUNCTION = 2 : EXIT FUNCTION
       IF isMask( GetWindowLong(hWnd, %GWL_STYLE), %SS_BLACKFRAME )  THEN FUNCTION = 3
    END IF

END FUNCTION
'-------------------------------------------------------------------------
' Here we get all the controls inside of the GroupBox and reposition
' them relative to the GroupBox we just moved, making it look like they
' were stuck to it.
' This is a User option gSticky = On/Off
' Entry: -Handle to the Control, Should always be a GroupBox control
'        -Rect structure of new position, rcDrag <relative to Client>
'-------------------------------------------------------------------------
SUB GetStickyControls( BYVAL hWnd AS LONG, BYVAL hWndEdit AS LONG, rc AS RECT)

    LOCAL rcGroupBox  AS RECT  'GroupBox position
    LOCAL rcControl   AS RECT  'Control position
    LOCAL rcDestRect  AS RECT  'needed as parameter
    LOCAL Sticky()    AS LONG  'hWnd's to controls found
    LOCAL result AS LONG
    LOCAL a AS LONG
    LOCAL i AS LONG
    LOCAL m AS LONG
    LOCAL x AS LONG
    LOCAL y AS LONG
    LOCAL dx AS LONG
    LOCAL dy AS LONG

    'Called from WM_LBUTTONUP, Drag Block
    'We determined already if sp.stickFrames = On AND if this is a GroupBox
    'Locate old position of this GroupBox, (we have not moved it yet!)
    GetWindowRect  hWnd,rcGroupBox
    MapWindowPoints %HWND_DESKTOP,hWndEdit,rcGroupBox,2

    'We Exit this Sub if we did not move, IOW just activating the handles
    IF rcGroupBox.nLeft = rc.nLeft AND rcGroupBox.nTop = rc.nTop THEN EXIT SUB

    'Enumerate all the controls to find all the intersecting controls
    'using our gZorder list.
    GetWindowRect hWnd, rcGroupBox 'we will use DeskTop coordinates

    FOR i = 0 TO UBOUND(gZorder)
  IF gZorder(i) = hWnd THEN 'skip this Control
     EXIT IF
  ELSE
      GetWIndowRect gZorder(i),rcControl
      IF IntersectRect(rcDestRect,rcGroupBox,rcControl) <> 0 THEN
    'record this controls handle for further processing
    a = UBOUND(Sticky) + 1
    REDIM PRESERVE Sticky(a)
    Sticky(a) = gZorder(i)
      END IF
  END IF
    NEXT

    'Did we find any controls intesecting this GroupBox?
    IF UBOUND(Sticky) = -1 THEN EXIT SUB

    'Now move the sticky controls to their new position.
    'rcGroupBox = old position relative to Desktop, map points to client
    'rc = new position (rcDrag) relative to client already

    MapWindowPoints %HWND_DESKTOP,hWndEdit,rcGroupBox,2

    'Get the delta of the GroupBoxes x/y points
    dx = rc.nLeft - rcGroupBox.nLeft '<new x - old x>
    dy = rc.nTop  - rcGroupBox.nTop  '<new y - old y>

    'NOTE:  I will need to omit moving other GroupBoxes, I get that old
    '       Magic Ring Trick, they are stuck together until I toggle
    '       sp.stickFrames = Off.

    FOR m = 0 TO UBOUND(Sticky)

      'Do not drag any other controls that are GroupBoxes
      'or EtchedFrames or BlackFrames
      '1= groupbox :2= EtchedFrame :3= BlackFrame
      result = IsGroupBox(Sticky(m))
      IF  result = 1 OR result = 2  OR result = 3 THEN
    EXIT IF

      ELSE
    'Get the position of Controls to move and Map to client points
    GetWindowRect Sticky(m&),rcControl
    MapWindowPoints %HWND_DESKTOP,hWndEdit,rcControl,2

    'add the Delta points
    x= rcControl.nLeft + dx
    y= rcControl.nTop  + dy

    'position the control
    SetWindowPos Sticky(m),0,x,y,0,0,%SWP_NOSIZE OR %SWP_NOZORDER

      END IF
    NEXT

END SUB
'SUB UpdateSizeStatus( BYVAL hWnd AS LONG )

'    STATIC hSizeStatus AS LONG

'    LOCAL  rc          AS RECT
'    LOCAL  sz1         AS ASCIIZ * 40
'    LOCAL  myTBB       AS TBBUTTONINFO
'   LOCAL rcWidth           AS LONG
'   LOCAL rcHeight          AS LONG
'    'LOCAL  lp_sChld    AS childWin_types PTR


'    IF ISFALSE hSizeStatus THEN
'       'lp_sChld    = SendMessage(GetMainWinHndl(hWnd),%WM_USER+140,0,0)
'       hSizeStatus = hToolbar'@lp_sChld.
'    END IF

'    GetWindowRect hWnd, rc

'    myTBB.cbSize    = LEN( TBBUTTONINFO )
'    myTBB.dwMask    = %TBIF_TEXT
'    myTBB.pszText   = VARPTR(sz1)

'    sz1 = FORMAT$(rc.nLeft) & " ," & FORMAT$(rc.nTop)
'    SendMessage hSizeStatus, %TB_SETBUTTONINFO, 20000, VARPTR(myTBB)
'   rcWidth=rc.nRight-rc.nLeft
'   rcHeight=rc.nbottom-rc.ntop
'    sz1 = FORMAT$(rcWidth) & " x" & FORMAT$(rcHeight)
'    SendMessage hSizeStatus, %TB_SETBUTTONINFO, 20001, VARPTR(myTBB)

'END SUB
'SUB UpdateSizeStatus2( BYVAL hWnd AS LONG, rc AS RECT)

'    LOCAL  sz1  AS ASCIIZ * 40
'    LOCAL myTBB AS TBBUTTONINFO
'    LOCAL  lp_sChld    AS LONG PTR
'   LOCAL rcWidth AS LONG
'   LOCAL rcHeight AS LONG

'    STATIC hSizeStatus AS LONG

'    IF ISFALSE hSizeStatus THEN
'       lp_sChld    = SendMessage(GetMainWinHndl(hWnd),%WM_USER+140,0,0)
'       hSizeStatus = hToolbar
'    END IF

'    myTBB.cbSize    = LEN( TBBUTTONINFO )
'    myTBB.dwMask    = %TBIF_TEXT
'    myTBB.pszText   = VARPTR(sz1)


'    sz1 = FORMAT$(rc.nLeft) & " ," & FORMAT$(rc.nTop)
'    SendMessage hSizeStatus, %TB_SETBUTTONINFO, 20000, VARPTR(myTBB)

'   rcWidth=rc.nRight-rc.nLeft
'   rcHeight=rc.nBottom-rc.nTop
'    sz1 = FORMAT$(rcWidth) & " x" & FORMAT$(rcHeight)
'    SendMessage hSizeStatus, %TB_SETBUTTONINFO, 20001, VARPTR(myTBB)
'END SUB
'------------------------------------------------------------------------------
' Update the Form or Control's Position & Size on the Status window
'------------------------------------------------------------------------------
SUB UpdateSizeStatus( BYVAL hWndControl AS LONG,BYVAL lpRect AS RECT PTR)

    LOCAL rc2 AS RECT
    LOCAL rc3 AS RECT

    rc3 = @lpRect

    LOCAL hForm AS LONG
    hForm = hWndEdit

    IF hWndControl = 0 THEN
       GetClientRect hForm,rc2
    ELSE
       rc2 = rc3
    END IF

    SendMessage ghWndSize, %WM_UPDATESIZEWINDOW, BYREF rc2, hWndControl

END SUB
SUB RefreshDmsClient(BYVAL hWnd AS LONG)
  LOCAL  sp AS setting_type
  LOCAL rc AS rect
  LOCAL tmp AS LONG
  LOCAL ny AS LONG,nx AS LONG
  LOCAL x  AS LONG,y  AS LONG,cx AS LONG,cy AS LONG
  LOCAL fSizingControl AS LONG

    IF ISFALSE  hWnd THEN EXIT SUB
    InvalidateRect hWnd  , BYVAL %NULL,%TRUE
    UpdateWindow   hWnd
    'DrawHandles    hWnd, 4 '' should be handleSize
    GetWindowRect hWnd,rc
    GetStickyControls hWndControl, hWnd, rc

    DrawHandles sp.handles
    UpdateSizeStatus hWnd,BYVAL VARPTR(rc)

END SUB
'------------------------------------------------------------------------------
' MAIN SUBCLASS PROCEDURE FOR ALL CONTROLS:
' All controls are sublcassed and all the messages are passed here to process
' the necessary messages required by the DMS engine.
' We also subclass any children of the control and pass those messages here.
'------------------------------------------------------------------------------
FUNCTION ControlSubClassProc(BYVAL hWnd AS LONG, BYVAL wMsg AS LONG, _
               BYVAL wParam AS LONG, BYVAL lParam AS LONG) AS LONG

  LOCAL pt      AS  POINTAPI
  LOCAL hParent AS LONG

  SELECT CASE wMsg

    CASE %WM_NCPAINT, _                                 'Allow the control to paint itself first.
       %WM_PAINT         : CallWindowProc GetProp(hWnd, $OldWndProc), hWnd, wMsg, wParam, lParam

    CASE %WM_NCHITTEST     : FUNCTION = %HTCLIENT      'let any Object respond with a hit test message.
    CASE %WM_MOUSEACTIVATE : FUNCTION = %MA_NOACTIVATE 'Defeat this message so that mouse clicks do not activate the control
    CASE %WM_SETCURSOR     : FUNCTION = %TRUE          'block system changing cursors. do it later on our own hit testing
    CASE %WM_SETFOCUS      : FUNCTION = 0              'block Caret activation for the RE Control, %MA_NOACTIVATE fails

    CASE %WM_LBUTTONDOWN,%WM_LBUTTONUP,%WM_MOUSEMOVE,_ 'Pass all mouse messages on through to parent in parent coordinates
       %WM_LBUTTONDBLCLK,%WM_RBUTTONDOWN,%WM_RBUTTONUP
       hParent = GetParent(hWnd) : IF ISFALSE hParent THEN EXIT FUNCTION
       IF wMsg = %WM_RBUTTONUP THEN
'                               gfWM_RBUTTONUP = 1
'                               gIDhWndCtl = hWnd
        sendmessage hParent, %WM_USER + 152, 1, hWnd  'special flag for context menu
      END IF
       pt.x = LOWRD ( lParam )
       pt.y = HIWRD ( lParam )
       ClientToScreen hWnd,    pt
       ScreenToClient hParent, pt
       PostMessage    hParent, wMsg, wParam, MAKLNG (pt.x, pt.y )


    CASE %WM_DESTROY
        SetWindowLong hWnd, %GWL_WNDPROC, GetProp(hWnd,$OldWndProc)   'UnSubClass the control and restore its default WndProc
        RemoveProp hWnd, $OldWndProc                                  'Remove our associated property
        EXIT FUNCTION

  END SELECT

  FUNCTION = CallWindowProc( GetProp(hWnd, $OldWndProc), hWnd, wMsg, wParam, lParam)

END FUNCTION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' subclasses Visual Editors's childs ( our controls ) which are created external
' called by form interpreter
'
FUNCTION WndClientProc_AddControl( BYVAL hWnd AS LONG, BYVAL hChild AS LONG , _
                   sp AS setting_type) AS LONG

  LOCAL szOldWndProc AS ASCIIZ * 20        ''' SetProp ASCIIZ
  LOCAL i            AS LONG               ''' zOrder array bounds
  szOldWndProc= $OldWndProc                                                'SubClass the Control.
  SetProp       hChild, szOldWndProc,SetWindowLong(hChild, %GWL_WNDPROC, BYVAL CODEPTR(ControlSubClassProc) )
  SetClassLong  hChild, %GCL_STYLE,  GetClassLong (hChild,%GCL_STYLE) OR %CS_DBLCLKS  'accept D.Clicks
  SetWindowLong hChild, %GWL_STYLE,  GetWindowLong(hChild,%GWL_STYLE) OR %WS_TABSTOP  'add WS_TABSTOP style
  SendMessage   hChild, %WM_SETFONT, GetStockObject(%DEFAULT_GUI_FONT), 0             'use dlg's font for control


  i = UBOUND(gZorder) + 1        'Bump up the Array count
  REDIM PRESERVE gZorder(i)
  gZorder(i) = hChild
  SetZorder  gZorder(i)          'Put it on top of the List

  hWndControl = gZorder(0)
  DrawHandles sp.Handles

END FUNCTION
FUNCTION WndClientProc_PAINT( BYVAL hWnd AS LONG, BYVAL wParam AS LONG, BYVAL lParam AS LONG, sp AS setting_type ) AS LONG

  LOCAL hdc AS LONG
  LOCAL i   AS LONG
  LOCAL j   AS LONG
  LOCAL ps  AS PAINTSTRUCT

  hDC = BeginPaint(hWnd,ps)

    IF sp.showGrid = %ON THEN
    FOR  i = 0 TO ps.rcPaint.nRight STEP sp.GridSize
      FOR j= 0 TO ps.rcPaint.nBottom STEP sp.GridSize
        SetPixel hDC, i, j, RGB(0,0,0)
      NEXT
    NEXT
    END IF

  FUNCTION = EndPaint(hWnd,ps)
END FUNCTION


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' wm_paint msg handler, needs a device context and a window handle
'

FUNCTION WndClientProc_PRINT( BYVAL hWnd AS LONG, BYVAL PrinterDC AS LONG, BYVAL flag AS LONG ) AS LONG
  LOCAL hDC    AS LONG
  LOCAL hBmp   AS LONG
  LOCAL hBmpDC AS LONG
  LOCAL r      AS RECT
  LOCAL x      AS LONG
  LOCAL y      AS LONG
  LOCAL rcWidth AS LONG
  LOCAL rcHeight AS LONG
  IF ISFALSE hWnd         THEN EXIT FUNCTION
  IF ISFALSE PrinterDC    THEN EXIT FUNCTION
  IF IsWindow( hWnd ) = 0 THEN EXIT FUNCTION
  GetWindowRect hWnd, r
  hDC    = GetWindowDC(hWnd)
  hBmpDC = CreateCompatibleDC( PrinterDC )
  rcWidth=r.nRight-r.nLeft
  rcHeight=r.nBottom-r.nTop
  hBmp   = CreateCompatibleBitmap(PrinterDC, rcWidth, rcHeight )
  RedrawWindow hWnd, BYVAL 0, BYVAL 0, %RDW_INTERNALPAINT OR %RDW_UPDATENOW OR %RDW_ALLCHILDREN
  SelectObject hBmpDC, hBmp
  BitBlt       hBmpDC, 0,0,rcWidth,rcHeight,  hDC,0,0, %SRCCOPY   '' flag ?
  x = GetDeviceCaps(PrinterDC, %HORZRES) / 2                            '' scale the bitmap to half pagewidth
  y = CLNG((x / rcWidth ) * rcHeight )
  SelectObject PrinterDC, hBmp
  StretchBlt   PrinterDC, 0,0,x,y, hBmpDC,0,0,rcWidth,rcHeight,%SRCCOPY
  DeleteObject hBmp                                                     '' clean mem
  DeleteDC     hDC
  DeleteDC     hBmpDC
END FUNCTION
'---------------------------------------------------------------------
' Normalize the Rectangle:
'---------------------------------------------------------------------
SUB rcEmpty ( rcSetZero AS rect )
    LOCAL rcBlank AS rect
    rcSetZero = rcBlank
END SUB
SUB ptEmpty ( ptSetZero AS POINTAPI )
    LOCAL ptBlank AS POINTAPI
    ptSetZero = ptBlank
END SUB
FUNCTION WndClientProc_CANCELMODE( BYVAL hWnd AS LONG, st AS st_types, sf AS flag_type ) AS LONG


    IF sf.Dragging THEN                                 '--- Cancel the dragging operation ---
       sf.Dragging = %OFF
       SetCursor LoadCursor(%NULL,BYVAL %IDC_ARROW)   'Restore the cursor, the Client window will do this for us anyway.
       rcEmpty st.rcDrag                             'If was Dragging, Clear these before we use them again.
       ptEmpty st.ptDrawEnd
       ptEmpty st.ptDrawStart
    END IF

    IF sf.Sizing THEN                                   '--- Cancel the Sizing operation ---
       sf.Sizing = %OFF
       ptEmpty st.ptDragStart                          'If was Sizing, Clear these before we use them again.
       ptEmpty st.ptDragStart
       ptEmpty st.ptDragLast
       rcEmpty st.nRect
    END IF

    IF sf.Drawing  THEN                                 '--- Cancel the Drawing operation ---
       sf.Drawing = %OFF
       SetCursor LoadCursor(%NULL,BYVAL %IDC_ARROW)   'Restore the cursor, the Client window will do this for us anyway.
       rcEmpty st.rcDrag                             'If was Drawing, Clear these before we use them again
       ptEmpty st.ptDrawEnd
       ptEmpty st.ptDrawStart
    END IF

    ClipCursor BYVAL %NULL
    sf.LostCapture = %ON                                'Clean up any  GDI garbage left over

END FUNCTION
FUNCTION WndClientProc_LBUTTONDBLCLK( BYVAL hWnd AS LONG) AS LONG
  'Determine if the Hit was on a control
  IF hWndControl <> 0 THEN
    'Yes, its a control
        'Now call our necessary Dialog Box  <Control properties>
      MSGBOX "Double Clicked on Control",%MB_SYSTEMMODAL  ,GetLang("prompt")
  ELSE
    'No, its for the client
        'Now call our necssary Dialog Box  <Main Form properties>
      MSGBOX "Double Clicked on a Client",%MB_SYSTEMMODAL   ,GetLang("prompt")
  END IF

END FUNCTION
FUNCTION WndClientProc_RBUTTONDOWN( BYVAL hWnd AS LONG, sf AS flag_type ) AS LONG
  'Note: If any operation Draw/Drag/Size is in process, ignore
  '      any right clicks.
  IF sf.Dragging OR sf.Sizing OR sf.Drawing THEN EXIT FUNCTION
  'We must always activate this in case the frame/client window was
  'not the active window previously and we click on a control first.
'    FUNCTION = SetFocus( hWnd )
END FUNCTION





FUNCTION WndClientProc_CONTEXTMENU( BYVAL hWnd AS LONG, BYVAL lParam AS LONG, sf AS flag_type) AS LONG

  'Note: If any operation Draw/Drag/Size is in process, ignore any right clicks, cancel the Context PopUp Menu
  IF sf.Dragging OR sf.Sizing OR sf.Drawing THEN EXIT FUNCTION

  IF sf.fWM_RBUTTONUP = 0 THEN
    TrackPopupMenu hPopupClient, _                      'Display the context menu for the Client/Main Form
             %TPM_LEFTALIGN OR %TPM_LEFTBUTTON, _
             LOWRD(lParam), HIWRD(lParam), _
             0, hWnd, BYVAL %NULL
    ELSE
     IF sf.IDhWndCtl = hWndControl THEN               'Is this control the Selected Active control?
      TrackPopupMenu hPopupCtl, _                     'Display the context menu for the Control
               %TPM_LEFTALIGN OR %TPM_LEFTBUTTON, _
               LOWRD(lParam), _
               HIWRD(lParam), _
               0, hWnd, BYVAL %NULL
     END IF
     sf.fWM_RBUTTONUP = 0            'Clear Flag
  END IF

END FUNCTION
'-------------------------------------------------------------------------
' Just bring the new selected window to the front
' Entry: hWnd = sh.hWndControl
'-------------------------------------------------------------------------
SUB BringToFront(BYVAL hWnd AS LONG)

    SetWindowPos   hWnd,%HWND_TOP,0,0,0,0,%SWP_DRAWFRAME OR %SWP_NOSIZE OR %SWP_NOMOVE
    InvalidateRect hWnd, BYVAL %NULL, %TRUE
    UpdateWindow   hWnd

END SUB
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' enums up from any app window created by pgm or user to apps main window
'
FUNCTION GetMainWinHndl( BYVAL anyProcessWin AS LONG ) AS LONG

    LOCAL hWndOwnwer AS LONG
    LOCAL sz1        AS ASCIIZ * 256
    hWndOwnwer  = anyProcessWin
    DO WHILE hWndOwnwer
       hWndOwnwer = getwindow( hWndOwnwer, %GW_OWNER )
       GetClassName  hWndOwnwer, sz1, 256
       IF sz1 = $PROGRAMCLASSNAME THEN EXIT LOOP
    LOOP
    FUNCTION = hWndOwnwer
END FUNCTION



'------------------------------------------------------------
' Draw our sizing Rectangle.
' And update the coordinates status window
'
' ToDo: add FitRectToBounds()
'------------------------------------------------------------
SUB DrawRectangle2 ( BYVAL hWnd AS LONG, ptStart AS POINTAPI, ptEnd AS POINTAPI)
    LOCAL hPen      AS LONG
    LOCAL hOldPen   AS LONG
    LOCAL hBrush    AS LONG
    LOCAL hOldBrush AS LONG
    LOCAL hDC       AS LONG
    LOCAL nDrawMode AS LONG
    hDC = GetDC(hWnd)                                          'now use only client dc
  hPen      = CreatePen(%PS_DASH, 0, 0)                  '*Select GDI object
  hOldPen   = SelectObject(hDC, hPen)
  hBrush    = GetStockObject(%NULL_BRUSH)
  hOldBrush = SelectObject(hDC, hBrush)
  nDrawMode = SetROP2(hDC, %R2_NOT)
  Rectangle hDC, ptStart.x, ptStart.y, ptEnd.x, ptEnd.y  '*Draw rectangle
  SetROP2 hDC, nDrawMode                                 '*Restore DC
  SelectObject hDC, hOldBrush
  SelectObject hDC, hOldPen
  DeleteObject hPen                                      '*Delete GDI objects
    ReleaseDC hWnd, hDC


    '-------Show Coordinates-----
    LOCAL pt1 AS POINTAPI        '
    LOCAL Pt2 AS POINTAPI        '
    LOCAL rc AS RECT             '

    pt1 = ptStart                '
    pt2 = ptEnd                  '
    NormalizeRect pt1,pt2        '
    rc.nLeft =    pt1.x          '
    rc.nTop  =    pt1.y          '
    rc.nRight =   pt2.x          '
    rc.nBottom =  pt2.y          '
         '
    UpdateSizeStatus hWnd, BYVAL VARPTR(rc)         '
    '------------end------------ '

END SUB

FUNCTION WndClientProc_LBUTTONDOWN( BYVAL hWnd AS LONG, BYVAL wParam AS LONG, BYVAL lParam AS LONG, _
                   st AS st_types, sf AS flag_type, sp AS setting_type ) AS LONG
  LOCAL  rcControl   AS RECT
  LOCAL  rcClient    AS RECT
  LOCAL  rc          AS RECT
  LOCAL  pt          AS POINTAPI
  LOCAL  dnc         AS POINTAPI
  LOCAL  dwc         AS POINTAPI
  LOCAL rcWidth   AS LONG
  LOCAL rcHeight  AS LONG
  SetFocus  hWnd              'always activate edit window.
  'ptFromLng lParam, pt
  pt.x = LOWRD ( lParam )
    pt.y = HIWRD ( lParam )
    '-------------------------------------------------------------------
    ' Hit-Test Every thing, LeftMouseDown handler
    '-------------------------------------------------------------------
    ' 1. test 8 Handles first
    ' 2. then objects/shapes [maybe should be after #4?]   <--ToDo:
    ' 3. then groupbox/frames
    ' 4. then controls/windows
    ' 5. else it's for the Client
    ' 6. start to draw a new control
    ' 7. start to draw an outline for a "Group of Controls" move
    '    <we should be able to use a similar routine for this like we do for GroupBoxes>
    '-------------------------------------------------------------------
   IF sf.cType = 0 THEN '<this will allow us to draw on top of any controls>
    '-------------------------------------------------------------------
    '**Test #1, Start with testing the Sizing Handles
    '-------------------------------------------------------------------
    IF hWndControl THEN
    LOCAL h AS LONG, i AS LONG, x AS LONG, y AS LONG, cx AS LONG, cy AS LONG
'        FigureHandleSizes sh.hWndControl, sp.handles
    FOR h = 0 TO 7
      rc = DialogHandles(h&).rcArea
      MapWindowPoints %HWND_DESKTOP,hWnd,rc,2
      IF  PtInRect( rc, pt.x,pt.y ) = %TRUE THEN
        'We ID'd the Handle
        IDhitValue = DialogHandles(h&).dwNCHITVALUE
        SetCapture hWnd
        'Let THE %WM_MOUSEMOVE block know we are sizing
        sf.Sizing = %ON
        'Get the coordinates RECT for the current control
        GetWindowRect   hWndControl, rc
        MapWindowPoints %HWND_DESKTOP, hWnd, rc, 2
        x  = rc.nLeft
        y  = rc.nTop
        rcWidth=rc.nRight-rc.nLeft
        rcHeight=rc.nBottom-rc.nTop
        cx = rcWidth '' rc.nRight   - rc.nLeft
        cy = rcHeight '' rc.nBottom  - rc.nTop
        st.nRect.nTop    = x   'save these reference points
        st.nRect.nleft   = y   'relative to Client
        st.nRect.nRight  = cx
        st.nRect.nBottom = cy
        'The cursor is maintain until %WM_LBUTTONUP for Controls,
        'except for Custom windows. So we better maintain the current cursor.
        SELECT CASE IDhitValue                                      '**Set the intial points for each side and corners
          CASE %HTTOPLEFT
            st.ptDragStart.x = x + cx
            st.ptDragStart.y = y + cy
            pt.x = x                                             'Move the cursor to top left corner of control
            pt.y = y
            i    = %IDC_SIZENWSE

          CASE %HTTOP
            st.ptDragStart.x = x
            st.ptDragStart.y = y + cy
            pt.x = x + (pt.x - x)                                'Move the cursor to the top edge of the control
            pt.y = y
            i    = %IDC_SIZENS

          CASE %HTTOPRIGHT
            st.ptDragStart.x = x
            st.ptDragStart.y = y  + cy
            pt.x = x + cx                                        'Move the cursor to the top right corner of the control
            pt.y = y
            i    = %IDC_SIZENESW

          CASE %HTLEFT
            st.ptDragStart.x = x + cx
            st.ptDragStart.y = y + cy
            pt.x = x                                             'Move the cursor to the left edge of the control
            pt.y = y + (pt.y - y )
            i    = %IDC_SIZEWE

          CASE %HTRIGHT
            st.ptDragStart.x = x
            st.ptDragStart.y = y + cy
            pt.x = x + cx                                        'Move the cursor to the right edge of the control
            pt.y = cy + (pt.y - cy )
            i    = %IDC_SIZEWE

          CASE %HTBOTTOMLEFT
            st.ptDragStart.x = x + cx
            st.ptDragStart.y = y
            pt.x = x                                             'Move the cursor to the bottom left corner of the control
            pt.y = y + cy
            i    = %IDC_SIZENESW

          CASE %HTBOTTOM
            st.ptDragStart.x = x
            st.ptDragStart.y = y
            pt.x = x + (pt.x -x)                                 'Move the cursor to the bottom edge of the control
            pt.y = y + cy
            i    = %IDC_SIZENS

          CASE %HTBOTTOMRIGHT
            st.ptDragStart.x = x
            st.ptDragStart.y = y
            pt.x = x + cx                                        'Move the cursor to the bottom right corner of the control
            pt.y = y + cy
            i    = %IDC_SIZENWSE
        END SELECT

        ClientToScreen hWnd,pt
        SetCursorPos   pt.x,pt.y
        SetCursor      LoadCursor(%NULL,BYVAL i)         'Set the appropiate Sizing cursor
        ptEmpty st.ptDragEnd
        ClientToScreen hWnd, st.ptDragEnd        'Send these coordinates/points to the Client
        ScreenToClient hWnd, st.ptDragEnd
        st.ptDragLast = st.ptDragEnd
        DrawRectangle2 hWnd,st.ptDragStart, st.ptDragLast      'Start the drawing of the Rectangle  '
        GetWindowRect hWnd,rc                     'Keep the mouse within the client area
        ClipCursor    rc
        FUNCTION = %true
        EXIT FUNCTION                                    'Now let the %WM_MOUSEMOVE Block take over from here

      END IF

    NEXT h

    END IF
    '--------------------------------------------------------------------------


    '--------------------------------------------------------------------------
    'Test #3/4 Hit Test Controls and Windows LeftMouseDown handler.
    'Control Selection and Control Drag operations
    '--------------------------------------------------------------------------

    'NOTE: We must update the Z-ORDER list so we always start from the
    '      top of the Z-ORDER then work our way down for Controls with
    '      Handles. Objects/shapes will get handled elsewhere.

    LOCAL g     AS LONG
    LOCAL tmp   AS LONG
    LOCAL delta AS LONG

    FOR i = 0 TO UBOUND(gZorder)

    'Convert RECT structure from Screen to Client coordinates
    GetWindowRect   gZorder(i), rc
    MapWindowPoints %HWND_DESKTOP, hWnd, rc, 2

    'Test #3/4  Control/Window & GroupBox/Frame
    IF PtInRect( rc, pt.x, pt.y ) THEN

        '---
        tmp = IsGroupBox(gZorder(i))
        '1= GroupBox : 2= EtchedFrame : 3= BlackRect
        IF tmp = 1 OR tmp = 2  OR tmp = 3 THEN

            'We dedicate 20 pixels at the Top Edge for a drag
            delta      = rc.nBottom -(rc.ntop + 20)
            rc.nBottom = rc.nBottom - delta

            'Do we want to drag it again?
            IF PtInRect( rc, pt.x, pt.y ) = %TRUE THEN
               EXIT IF
            ELSE 'If not, look for controls inside the groupbox
               'skip the groupbox in Zorder
               FOR g = i + 1 TO UBOUND(gZorder)
                 GetWindowRect gZorder(g&),rc
                 MapWindowPoints %HWND_DESKTOP,hWnd,rc,2
                 IF PtInRect( rc, pt.x, pt.y ) = %TRUE THEN
                   i = g
                   EXIT FOR
                 END IF
               NEXT
            END IF
        END IF
         '---

      '---
      'I added this "IF" to reduce the flicker from redrawing the same
      'control over if it has already been selected. [I still have some
      'more optimization other places, to minimize the number of times
      'I am Updating all the controls to satisfy the GroupBox needs.]
      IF hWndControl = gZorder(i) THEN
        EXIT IF
      ELSE

        ClearHandles               ' Clear the handles from the previous selection
        hWndControl = gZorder(i)   ' Set the New Selection, so everyone knows we have one.
        SetZorder    hWndControl
        BringToFront hWndControl
        DrawHandles  sp.handles   ' Now draw the handles on the new selection

      END IF
      '---

      'Capture the Mouse
      SetCapture hWnd

      'Keep the Cursor inside of the Client area
      GetWindowRect hWnd, rc
      ClipCursor    rc

      'DRAG RECTANGLE USING ROP------------------------------
      'Get the coordinates RECT for the current control
      GetWindowRect   hWndControl, rc
      MapWindowPoints %HWND_DESKTOP, hWnd, rc, 2

      'This is our Virtual Window we will maintain throughout
      'this operation. We will use the Final st.rcDraw RECT to
      'position our control
      st.rcDrag = rc


      'Set starting points
      IF sp.snapGrid = %ON THEN
        st.Start_dwc.x = LOWRD ( lParam ) - LOWRD ( lParam ) MOD sp.GridSize
        st.Start_dwc.y  = HIWRD ( lParam ) - HIWRD ( lParam ) MOD sp.GridSize
      ELSE
         'ptFromLng lParam, st.Start_dwc
         st.Start_dwc.x=LOWRD(lParam)
         st.Start_dwc.y=HIWRD(lParam)
      END IF


      DrawRectangle2 hWnd, st.ptDrawStart,st.ptDrawEnd  'Start with the first Rectangle
      UpdateSizeStatus hWnd,BYVAL VARPTR(rc)                    'Show the initial coordinates
      'UpdateSizeStatus wParam, BYVAL lParam
      sf.fOneShot = %ON       'change the cursor once until we start moving and Keep the Handles visible.
      sf.Dragging = %ON         'Let everyone know we are dragging a window
      FUNCTION = %true
      EXIT FUNCTION
    END IF

    NEXT

   END IF '<this allows us to draw on top of any controls> ' if sf.cType = 0
   '***

    '---------------------------------------------------------------------
    'Test #5, Fall thru...
    'No Hits on Handles, Objects/Shapes, GroupBox/Frames or
    'Controls/Windows so it's for Client!!!!
    '---------------------------------------------------------------------

    'Did we just hit the main client area only?
    'Then we should remove the Sizing Handles.

    IF sf.Dragging = %OFF THEN

     ClearHandles    'Then NO active selection, hide the Handles
     hWndControl    = 0         'No active selection
     sf.DrawNewControl = %ON     'allow to draw rectangle on main client <todo: add group move>

    END IF


    '----------------------------------------------------------------------
    'Action #6: Are we drawing a new control or maybe a Group-Move ???
    ' 1. Starting to draw a new control LeftMouseDown handler
    ' 2. Draw a RECT around controls for a Group move.
    '----------------------------------------------------------------------
    IF sf.DrawNewControl THEN

      'Set starting points to nearest grid points
      IF sp.snapGrid  = %ON THEN
        st.Start_dnc.x = LOWRD ( lParam ) - LOWRD ( lParam ) MOD sp.GridSize
        st.Start_dnc.y = HIWRD ( lParam ) - HIWRD ( lParam ) MOD sp.GridSize
      ELSE
        'ptFromLng lParam, st.Start_dnc
        st.Start_dnc.x=LOWRD(lParam)
        st.Start_dnc.y=HIWRD(lParam)
      END IF

      SetCapture hwnd
      SetCursor  LoadCursor(%NULL,BYVAL %IDC_CROSS)
      getWindowRect hWnd, rc
      ClipCursor    rc

      DrawRectangle2 hWnd, st.ptDrawStart, st.ptDrawEnd         'Start with the first Rectangle
      rc.nLeft   =  st.Start_dnc.x                    '---Show coordinates---
      rc.nTop    =  st.Start_dnc.y
      rc.nRight  =  st.Start_dnc.x
      rc.nBottom =  st.Start_dnc.y
      UpdateSizeStatus hwnd, BYVAL VARPTR(rc)                         '--------end-----------


      sf.Drawing = %ON
      sf.DrawNewControl = %OFF

      'Reset the Tool Pallete to default Select Tool button
'            sendMessage ghFtoolbar,%TB_CHECKBUTTON,%IDP_ACT0, MAKLNG(%TRUE,0)
'            sf.ctype = gSelectTool
      'WM_ MOUSEMOVE takes over from here.

    END IF
    '----------------------------------------------------------------------
END FUNCTION



FUNCTION WndClientProc_MOUSEMOVE( BYVAL hWnd AS LONG, BYVAL wParam AS LONG , BYVAL lParam AS LONG, _
                  st AS st_types, sf AS flag_type, sp AS setting_type ) AS LONG

  LOCAL  h   AS LONG
  LOCAL  rc  AS RECT
  LOCAL  pt  AS POINTAPI
  LOCAL  dnc AS POINTAPI
  LOCAL  dwc AS POINTAPI
    '-------------------------------------------------------------------
    'Dragging a control using a Focus Rectangle, MouseMove handler.
    '
    '-------------------------------------------------------------------
    'Are we dragging a control? Then Update its new position
    IF sf.Dragging = %ON THEN
      'Gridize if necessary?
      IF sp.snapGrid = %ON THEN
        dwc.x = LOWRD ( lParam ) - LOWRD ( lParam ) MOD sp.GridSize
        dwc.y = HIWRD ( lParam ) - HIWRD ( lParam ) MOD sp.GridSize
      ELSE
        'ptFromLng lParam, dwc
        dwc.x=LOWRD(lParam)
        dwc.y=HIWRD(lParam)
      END IF
      'Move the Rectangle if we moved the mouse
      IF (( dwc.x <> st.Start_dwc.x ) OR ( dwc.y <> st.Start_dwc.y  )) THEN
        'Show the Four Corner Moving Arrows while we Drag it
        IF sf.fOneShot  = %ON THEN
          SetCursor LoadCursor(%NULL,BYVAL %IDC_SIZEALL)
          sf.fOneShot = %OFF
        END IF
        DrawRectangle2 hWnd, st.ptDrawStart, st.ptDrawEnd                      'Erasing the previous Rectangle
        st.rcDrag.nLeft   = st.rcDrag.nLeft   + dwc.x - st.Start_dwc.x  'Get the new points for new location
        st.rcDrag.nRight  = st.rcDrag.nRight  + dwc.x - st.Start_dwc.x
        st.rcDrag.nTop    = st.rcDrag.nTop    + dwc.y - st.Start_dwc.y
        st.rcDrag.nBottom = st.rcDrag.nBottom + dwc.y - st.Start_dwc.y
        st.ptDrawEnd.x   = st.rcDrag.nLeft                              'Convert to st.rcDraw to points
        st.ptDrawEnd.y   = st.rcDrag.nTop
        st.ptDrawStart.x = st.rcDrag.nRight
        st.ptDrawStart.y = st.rcDrag.nBottom
        DrawRectangle2 hWnd, st.ptDrawStart, st.ptDrawEnd                      'Move Rect to new location
        st.Start_dwc = dwc                                              'Remember the previous location
      END IF
    END IF
    '---------------------------------------------------------------------
    IF sf.Sizing THEN                                   'Sizing Control MouseMove handler
        DrawRectangle2 hWnd, st.ptDragStart, st.ptDragLast    'Clear the previous Rectangle
        SELECT CASE IDhitValue
          CASE  %HTTOPLEFT,%HTTOPRIGHT,%HTBOTTOMLEFT,%HTBOTTOMRIGHT '*Corners
            IF sp.snapGrid = %ON THEN
              st.ptDragEnd.x = LOWRD(lParam )-LOWRD(lParam )MOD sp.GridSize
              st.ptDragEnd.y = HIWRD(lParam )-HIWRD(lParam )MOD sp.GridSize
            ELSE
              'ptFromLng lParam, st.ptDragEnd
          st.ptDragEnd.x = LOWRD(lParam)
          st.ptDragEnd.y = HIWRD(lParam)
            END IF
          CASE %HTLEFT, %HTRIGHT '*Left, Right
             IF sp.snapGrid = %ON THEN
              st.ptDragEnd.x = LOWRD(lParam )-LOWRD(lParam )MOD sp.GridSize
              st.ptDragEnd.y = st.nRect.nleft
             ELSE
              st.ptDragEnd.x = INT(LOWRD(lParam))
              st.ptDragEnd.y = st.nRect.nleft
             END IF
          CASE %HTTOP, %HTBOTTOM '*Bottom, Top
            IF sp.snapGrid = %ON THEN
              st.ptDragEnd.x = st.ptDragStart.x + st.nRect.nRight
              st.ptDragEnd.y = HIWRD(lParam )-HIWRD(lParam )MOD sp.GridSize
            ELSE
              st.ptDragEnd.x = st.ptDragStart.x + st.nRect.nRight
              st.ptDragEnd.y = INT(HIWRD(lParam))
            END IF
        END SELECT
        DrawRectangle2 hWnd, st.ptDragStart, st.ptDragEnd          'Draw the rectangle at the new location
        st.ptDragLast = st.ptDragEnd
    END IF
    IF sf.Dragging = %OFF AND sf.Sizing = %OFF THEN       'Size Grip Handles MouseMove handler
      IF ISFALSE hWndControl THEN EXIT IF          'Do we have a valid selection?
      pt.x = LOWRD(lParam)
      pt.y = HIWRD(lParam)
'            FigureHandleSizes hWndControl, sp.handles
      FOR h = 0 TO  7
        rc = DialogHandles(h).rcArea
        MapWindowPoints %HWND_DESKTOP, hWnd, rc,2
        IF PtInRect( rc, pt.x,pt.y ) = %TRUE THEN 'Are we over any of the 8 "Handles", change NSWE cursors
           SELECT CASE DialogHandles(h&).dwNCHITVALUE
              CASE %HTLEFT,%HTRIGHT          : SetCursor LoadCursor(%NULL, BYVAL %IDC_SIZEWE  )
              CASE %HTTOP,%HTBOTTOM          : SetCursor LoadCursor(%NULL, BYVAL %IDC_SIZENS  )
              CASE %HTTOPLEFT,%HTBOTTOMRIGHT : SetCursor LoadCursor(%NULL, BYVAL %IDC_SIZENWSE)
              CASE %HTTOPRIGHT,%HTBOTTOMLEFT : SetCursor LoadCursor(%NULL, BYVAL %IDC_SIZENESW)
           END SELECT
           EXIT FUNCTION
        ELSE
           'NOTE: The main client area will take care of changing back
           'the cursor, but we have to do it for controls -- Use the Normal cursor
           SetCursor LoadCursor(%NULL,BYVAL %IDC_ARROW)
        END IF
      NEXT h
    END IF
    IF sf.Drawing THEN                          'Draw New Control MouseMove handler
      st.ptDrawEnd.x   = st.rcDraw.nLeft          'Convert to st.rcDraw to points
      st.ptDrawEnd.y   = st.rcDraw.nTop
      st.ptDrawStart.x = st.rcDraw.nRight
      st.ptDrawStart.y = st.rcDraw.nBottom
      DrawRectangle2 hWnd, st.ptDrawStart,st.ptDrawEnd   'Erase the previous rectangle
      IF sp.snapGrid = %ON THEN
        dnc.x = LOWRD ( lParam ) - LOWRD ( lParam ) MOD sp.GridSize
        dnc.y = HIWRD ( lParam ) - HIWRD ( lParam ) MOD sp.GridSize
      ELSE
        'ptFromLng lParam, dnc
        dnc.x=LOWRD(lParam)
        dnc.y=HIWRD(lParam)
      END IF
      IF st.Start_dnc.x  > dnc.x THEN             'Going left?
        st.rcDraw.nLeft  = dnc.x
        st.rcDraw.nRight = st.Start_dnc.x
      ELSE                                        'going right
        st.rcDraw.nLeft  = st.Start_dnc.x
        st.rcDraw.nRight = dnc.x
      END IF
      IF st.Start_dnc.y < dnc.y THEN              'Going down?
        st.rcDraw.nBottom = dnc.y
        st.rcDraw.nTop    = st.Start_dnc.y
      ELSE                                        'Going up
        st.rcDraw.nBottom = st.Start_dnc.y
        st.rcDraw.nTop    = dnc.y
      END IF
      st.ptDrawEnd.x   = st.rcDraw.nLeft          'Convert st.rcDraw to points
      st.ptDrawEnd.y   = st.rcDraw.nTop
      st.ptDrawStart.x = st.rcDraw.nRight
      st.ptDrawStart.y = st.rcDraw.nBottom
      DrawRectangle2 hWnd, st.ptDrawStart,st.ptDrawEnd   'Update the new position
    END IF
END FUNCTION





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' api wrapper
'
'
FUNCTION myCreateWindowEx( BYVAL dwExStyle AS LONG, lpClassName AS ASCIIZ, lpWindowName AS ASCIIZ, BYVAL dwStyle AS LONG, _
         BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, _
         BYVAL hWndParent AS LONG, BYVAL hMenu AS LONG, BYVAL hInstance AS LONG,BYVAL lpParam AS LONG, _
         BYVAL lpMem AS LONG ) AS LONG

    LOCAL done AS undoDMS_type

    done.undo.CreateW.hWnd = CreateWindowEx ( dwExStyle, lpClassName, lpWindowName, dwStyle, _
                x, y, nWidth, nHeight, _
                hWndParent, hMenu, hInstance, lpParam)

    IF ISFALSE done.undo.CreateW.hWnd THEN EXIT FUNCTION  ''' create error - don't add to undo buffer

    done.action                    = %CreateWindowEx
    done.undo.CreateW.dwExStyle    = dwExStyle
    done.undo.CreateW.szClassName  = lpClassName
    done.undo.CreateW.szWindowName = lpWindowName
    done.undo.CreateW.dwStyle      = dwStyle
    done.undo.CreateW.x            = x
    done.undo.CreateW.y            = y
    done.undo.CreateW.nWidth       = nWidth
    done.undo.CreateW.nHeight      = nHeight
    done.undo.CreateW.hWndParent   = hWndParent
    done.undo.CreateW.hMenu        = hMenu
    done.undo.CreateW.hInstance    = hInstance
    done.undo.CreateW.lpParam      = lpParam

    SendMessage sendmessage( GetMainWinHndl( hWndParent), %WM_USER + 145, 0, 0 ), %TB_ENABLEBUTTON, %IDM_Undo, %TRUE

    FUNCTION = done.undo.CreateW.hWnd
END FUNCTION
'---------------------------------------------------------------------------
' Create New Tab controls for the Visual Designer
' Entry:  Handle of Parent, rcDraw
' Return: Handle of New Control
' Also subclasses the UpDown child control so it passes all mouse hits to
' its parent Tab window.
'---------------------------------------------------------------------------
FUNCTION CreateNewTabCtl(BYVAL hWnd AS LONG,rc AS RECT) AS LONG

  LOCAL hWndNew AS LONG
  LOCAL Txt     AS ASCIIZ * 10
  LOCAL Tei     AS TC_ITEM
  LOCAL rcWidth AS LONG
  LOCAL rcHeight AS LONG

  rcWidth=rc.nRight-rc.nLeft
  rcHeight=rc.nBottom-rc.nTop

  hWndNew = myCreateWindowEx(0,$WC_TABCONTROL, "", %WS_MINIMAL OR %TCS_TABS , _
               rc.nLeft, rc.nTop, rcWidth, rcHeight, hWnd, BYVAL %NULL, _
               GetModuleHandle(BYVAL 0), BYVAL %NULL, 0 )
  IF ISFALSE hWndNew THEN EXIT FUNCTION

  Tei.pszText = VARPTR(Txt)            'same for all
  Tei.iImage = -1                      'no image
  Tei.Mask = %TCIF_TEXT                'Initialize to default of 3 tabs

  Txt = "Tab 1" : TabCtrl_InsertItem hWndNew, 0, Tei
  Txt = "Tab 2" : TabCtrl_InsertItem hWndNew, 1, Tei
  Txt = "Tab 3" : TabCtrl_InsertItem hWndNew, 2, Tei

  FUNCTION   = hWndNew
END FUNCTION
SUB SetCol(hlist AS LONG, dList AS ListData)
    LOCAL lvc AS LV_COLUMN
    LOCAL iStatus AS INTEGER
    LOCAL szStr AS ASCIIZ * 32
    LOCAL i AS INTEGER

    lvc.mask = %LVCF_FMT OR %LVCF_Width OR %LVCF_Text OR %LVCF_SUBITEM
    lvc.fmt = %LVCFMT_LEFT
    lvc.PszText = VARPTR(szStr)

    FOR i = 1 TO dList.Columns
  szStr = dList.Fields(i).fName
  lvc.cx = dList.Fields(i).SIZE

  IF lvc.cx > 200 THEN lvc.cx = 200

  lvc.iSubItem = i
  lvc.fmt = dList.Fields(i).Alignment
  iStatus = ListView_InsertColumn (hList, i, lvc)

    NEXT i

END SUB
SUB SetListView (hlist AS LONG, dList AS ListData)
    LOCAL lvc AS LV_COLUMN
    LOCAL iStatus AS INTEGER
    LOCAL szStr AS ASCIIZ * 32

    lvc.mask     = %LVCF_FMT OR %LVCF_Width OR %LVCF_Text OR %LVCF_SUBITEM
    lvc.fmt      = %LVCFMT_LEFT
    lvc.PszText  = VARPTR(szStr)
    szStr        = dList.Fields(0).fName
    lvc.cx       = dList.Fields(0).SIZE
    lvc.iSubItem = 0

    ''' Test
    iStatus = listView_InsertColumn (hList, %Null, lvc)
    SetCol hList, dList

END SUB
SUB AppendListView (hList AS LONG, Rec() AS STRING )
    LOCAL z AS INTEGER
    LOCAL iStatus AS INTEGER
    LOCAL szStr AS ASCIIZ * 300
    LOCAL lvi AS LV_ITEM
    LOCAL x AS LONG

    'this will be the next record
    lvi.iItem     = ListView_GetItemCount(hList) '+ 1
    lvi.mask      = %LVIF_TEXT
    lvi.stateMask = %LVIS_FOCUSED '%LVIS_SELECTED
    lvi.pszText   = VARPTR(szStr)

    FOR z = 0 TO UBOUND(Rec)
  szStr = Rec(z)
  lvi.iSubItem = z

  lvi.lParam = lvi.iItem
  IF z = 0 THEN
      lvi.mask = %LVIF_TEXT OR %LVIF_PARAM OR %LVIF_STATE
      iStatus  = ListView_InsertItem (hList, lvi)
  ELSE
      lvi.mask = %LVIF_TEXT
      iStatus  = ListView_SetItem (hList, lvi)
  END IF

    NEXT

END SUB

'---------------------------------------------------------------------------
' Create a ListView Control for the Visual Designer
' Entry:  Handle of Parent, rcDraw data
' Return: Handle of new control
' We also subclass the Header child control so we can pass all mouse hits to
' the ListView Parent.
' See ListView.BI for more information regarding equates, help subs, UDT's.
'--------------------------------------------------------------------------
FUNCTION CreateNewListView(BYVAL hWnd AS LONG, rc AS RECT) AS LONG

  LOCAL hWndNew            AS LONG
  LOCAL lData              AS LISTDATA
  LOCAL Rec()              AS STRING
  LOCAL hWndListViewHeader AS LONG

  'Create the ListView Control using our default settings
  hWndNew = myCreateWindowEx(0,$WC_LISTVIEW,"", %WS_BORDER OR %WS_MINIMAL _
                 OR %LVS_REPORT OR %LVS_SINGLESEL OR %LVS_SHOWSELALWAYS, _
                 rc.nLeft, rc.nTop, rc.nRight-rc.nLeft, rc.nBottom-rc.nTop, hWnd, BYVAL %NULL, _
                 GetModuleHandle(BYVAL 0),BYVAL %Null, 0)

  IF ISFALSE hWndNew THEN EXIT FUNCTION

  SendMessage hWndNew, %LVM_SETEXTENDEDLISTVIEWSTYLE, 0, SendMessage(hWndNew, %LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) OR %LVS_EX_FULLROWSELECT OR %LVS_EX_GRIDLINES

  'populate the listview with sample data
  'create 5 columns in report view
  lData.Columns = 4 'zero based!
  lData.Fields(0).fName = "ID"
  lData.Fields(0).SIZE = 40
  lData.Fields(0).Alignment = 1
  lData.Fields(1).fName = "标题"
  lData.Fields(1).SIZE = 160
  lData.Fields(1).Alignment = 0
  lData.Fields(2).fName = "描述"
  lData.Fields(2).SIZE = 200
  lData.Fields(2).Alignment = 0
  lData.Fields(3).fName = "分类"
  lData.Fields(3).SIZE = 100
  lData.Fields(3).Alignment = 0
  lData.Fields(4).fName = "价格"
  lData.Fields(4).SIZE = 60
  lData.Fields(4).Alignment = 1
  SetListView hWndNew, lData

  'add some records
  REDIM Rec(0 TO 4)
  Rec(0) = "0022" : Rec(1) = "直接打印可用"
  Rec(2) = "绕过打印机驱动。打印.doc, .rtf 和 .wri 文件。"
  Rec(3) = "打印可用" : Rec(4) = "$10.00"
  AppendListView hWndNew , Rec()
  Rec(0) = "0023" : Rec(1) = "直接打印的 DLL"
  Rec(2) = "用32位实模式跳过打印机驱动，发送格式代码，设置默认打印机。"
  Rec(3) = "编程库" : Rec(4) = "$79.50"
  AppendListView hWndNew, Rec()
  Rec(0) = "0024" : Rec(1) = "ACTION Process Automator"
  Rec(2) = "Unattended batch processor with built in netware and compression support."
  Rec(3) = "Automation utilities" : Rec(4) = "$79.50"
  AppendListView hWndNew, Rec()
  Rec(0) = "0025" : Rec(1) = "Zipserver"
  Rec(2) = "Zip compatible drag & drop utility & ActiveX."
  Rec(3) = "Compression utilities" : Rec(4) = "$39.50"
  AppendListView hWndNew, Rec()

  FUNCTION = hWndNew
END FUNCTION
                                                                     _
                                                             _

'---------------------------------------------------------------------------
' Create a simple TreeView Control for the Visual Designer
' Entry:  Handle of Parent, rcDraw data
' Return: Handle of new control
' ToDo: impliment ImageList control to show some bitmaps
'--------------------------------------------------------------------------
FUNCTION CreateNewTreeView(BYVAL hWnd AS LONG, rc AS RECT) AS LONG

  LOCAL hWndNew AS LONG

  '---TreeView
  LOCAL tvs          AS TV_INSERTSTRUCT
  LOCAL tvi          AS TV_ITEM
  LOCAL sText        AS ASCIIZ * 40
  LOCAL I AS LONG



    'Create the TreeView control
    hWndNew = myCreateWindowEx(%WS_EX_CLIENTEDGE, $WC_TREEVIEW,"", _
                   %WS_MINIMAL OR %TVS_HASBUTTONS OR %TVS_HASLINES  _
                   OR %TVS_LINESATROOT OR %TVS_SHOWSELALWAYS, _
                   rc.nLeft, rc.nTop, rc.nRight-rc.nLeft, rc.nBottom-rc.nTop,  _
                   hWnd,BYVAL %NULL,GetModuleHandle(BYVAL 0),BYVAL %Null, 0)


    'Initialize the tree List
    tvs.hInsertAfter = %TVI_LAST
    tvi.mask         = %TVIF_TEXT
    tvi.pszText      = VARPTR(sText)

    DIM hTreeWnd(0 TO 4) AS LOCAL LONG
    sText       = "One"
    tvs.item.hitem = VARPTR(tvi)
    'tvs.item.item = tvi
    hTreeWnd(0) = TreeView_InsertItem(hWndNew,tvs)

    sText       = "Two"
    tvs.hParent = hTreeWnd(0)
    hTreeWnd(1) = TreeView_InsertItem(hWndNew,tvs)

    sText       = "Buckle"
    tvs.hParent = hTreeWnd(1)
    hTreeWnd(2) = TreeView_InsertItem(hWndNew,tvs)

    sText       = "My"
    tvs.hParent = hTreeWnd(2)
    hTreeWnd(3) = TreeView_InsertItem(hWndNew,tvs)
    sText       = "Shoe"
    hTreeWnd(4) = TreeView_InsertItem(hWndNew,tvs)

    'Expand All
    FOR i = 0 TO 4
      TreeView_Expand hWndNew, hTreeWnd(i), %TVE_EXPAND
    NEXT

  FUNCTION = hWndNew

END FUNCTION

'---------------------------------------------------------------------------
' Create a control based on Control Type and Size with my "Default Styles."
' Entry: <ctype>  type of control to create
' Entry: <rcDraw> contains the position of the new control.
' NOTE: This routine is in no way optimized!
'---------------------------------------------------------------------------
FUNCTION CreateControlEx ( BYVAL hWndParent AS LONG, BYVAL ctype AS LONG, rcDraw AS RECT, sp AS setting_type) AS LONG
    LOCAL hWndNew      AS LONG
    LOCAL szOldWndProc AS ASCIIZ * 11
    LOCAL i            AS LONG
  LOCAL rcWidth           AS LONG
  LOCAL rcHeight          AS LONG

    IF ctype = 0 OR ctype > 18 THEN EXIT FUNCTION          'Make sure we have a valid number

  rcWidth=rcDraw.nRight-rcDraw.nLeft
  rcHeight=rcDraw.nBottom-rcDraw.nTop
    SELECT CASE ctype
  CASE 1  'Static  <%SS_SIMPLE>
      hWndNew =myCreateWindowEx(0,$WC_STATIC, "Static box", _
              %WS_MINIMAL OR %SS_NOTIFY, _
              rcDraw.nLeft, rcDraw.nTop, _
              rcWidth, rcHeight, _
              hWndParent, %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )

  CASE 2    'Button   %IDP_ACT2
      hWndNew =myCreateWindowEx(0,$WC_BUTTON, "&Button", _
             %WS_MINIMAL _
             OR %BS_PUSHBUTTON, _
             rcDraw.nLeft, rcDraw.nTop, _
             rcWidth, rcHeight, _
             hWndParent, BYVAL %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )

  CASE 3  'Edit Box
      hWndNew = myCreateWindowEx(0, $WC_EDIT, $WC_EDIT, _
              %WS_CHILD OR %WS_VISIBLE  OR %WS_BORDER OR %WS_VSCROLL OR %WS_CLIPSIBLINGS, _
              rcDraw.nLeft, rcDraw.nTop, _
              rcWidth, rcHeight, _
              hWndParent, BYVAL %NULL,GetModuleHandle(BYVAL 0), BYVAL %NULL,0)

  CASE 4   'CheckBox <derived from the Button Class>
      hWndNew = myCreateWindowEx(0, $WC_BUTTON, $CW_CHECKBOX, _
             %WS_MINIMAL OR %BS_AUTOCHECKBOX, _
             rcDraw.nLeft, rcDraw.nTop, _
             rcWidth, rcHeight, _
             hWndParent, BYVAL %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0)

  CASE 5  'RadioButton <derived from the Button class>
      hWndNew =myCreateWindowEx(0,$WC_BUTTON, $CW_OPTION, _
             %WS_MINIMAL OR %BS_AUTORADIOBUTTON, _
             rcDraw.nLeft, rcDraw.nTop, _
             rcWidth, rcHeight, _
             hWndParent, BYVAL %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0)

  CASE 6   'GroupBox <derived from Button Class>
      hWndNew =myCreateWindowEx(%WS_EX_TRANSPARENT,$WC_BUTTON, $CW_FRAME, _
              %BS_GROUPBOX OR %WS_MINIMAL , _
              rcDraw.nLeft, rcDraw.nTop, _
              rcWidth, rcHeight, _
              hWndParent, BYVAL %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )

  CASE 7   'Image or Solid color Panel <derived from the Static class>

      'if gfImageAutoSize = ON
      hWndNew = myCreateWindowEx(0,$WC_STATIC , $CW_IMAGE, _
              %WS_MINIMAL OR %SS_BITMAP, _
              rcDraw.nLeft, rcDraw.nTop, _
              rcWidth, rcHeight, _
              hWndParent, %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )

      '---
      IF sp.autoSizeImg = %OFF THEN

     'Let's set the size to the Users specifications instead
     SetWindowPos hWndNew, %HWND_TOP, _
            rcDraw.nLeft, rcDraw.nTop, _
            rcWidth, rcHeight, _
            %SWP_DRAWFRAME
      END IF
      '---

    'Either case, we must show the window
    ShowWindow hWndNew,%SW_SHOW


  CASE 8  'ComboBox      'Todo: test %CBS_SIMPLE style
      LOCAL s AS STRING
      hWndNew  = myCreateWindowEx(0, $WC_COMBOBOX, "", _
              %WS_MINIMAL OR %CBS_DROPDOWNLIST OR %CBS_NOINTEGRALHEIGHT, _
              rcDraw.nLeft, rcDraw.nTop, _
              rcWidth, rcHeight, _
              hWndParent, %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0)


      FOR i = 0 TO 3
    s = "ComboBox " + FORMAT$(i)
    SendMessage hWndNew,%CB_ADDSTRING,i,STRPTR(s$)
      NEXT
      SendMessage hWndNew,%CB_SETCURSEL,0,0
      MoveWindow  hWndNew,rcDraw.nLeft,rcDraw.nTop,rcDraw.nRight-rcDraw.nLeft,400,0
'            ShowWindow  hWndNew,%SW_SHOW



  CASE 9   'ListBox

      hWndNew = myCreateWindowEx(0,$WC_LISTBOX, "", _
               %WS_MINIMAL OR %LBS_NOINTEGRALHEIGHT OR %LBS_STANDARD, _
               rcDraw.nLeft, rcDraw.nTop, _
               rcWidth, rcHeight, _
               hWndParent, %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )

      s = "ListBox"
      SendMessage hWndNew, %LB_ADDSTRING, 0, STRPTR(s)
      SendMessage hWndNew, %LB_ADDSTRING, 1, STRPTR(s)
      SendMessage hWndNew, %LB_ADDSTRING, 2, STRPTR(s)
      SendMessage hWndNew, %LB_ADDSTRING, 3, STRPTR(s)




  CASE 10   'Horizontal Scroll Bar <standard scroll bar>

      hWndNew = myCreateWindowEx(0,$WC_SCROLLBAR, "", _
              %WS_MINIMAL OR %SBS_HORZ, _
              rcDraw.nLeft, rcDraw.nTop, _
              rcWidth, rcHeight, _
              hWndParent, %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )


  '---
  CASE 11   'Vertical Scroll Bar <standard scroll bar>

      hWndNew = myCreateWindowEx(0,$WC_SCROLLBAR, "", _
              %WS_MINIMAL OR %SBS_VERT, _
              rcDraw.nLeft, rcDraw.nTop, _
              rcWidth, rcHeight, _
              hWndParent, %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0)

  CASE 12  : hWndNew = CreateNewTabCtl(hWndParent,rcDraw)  'Tab
  CASE 13  'RichEdit
      hWndNew = myCreateWindowEx(%WS_EX_CLIENTEDGE ,$CW_RICHEDIT,$CW_RICHEDIT, _
               %WS_MINIMAL OR %WS_VSCROLL OR %WS_HSCROLL , _
               rcDraw.nLeft, rcDraw.nTop, _
               rcWidth, rcHeight, _
               hWndParent,BYVAL %NULL,GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )

      SendMessage hWndNew,%EM_SETREADONLY, %TRUE, 0


  CASE 14     'Progress Bar

      hWndNew = myCreateWindowEx(0,$PROGRESS_CLASS,BYVAL %NULL, _
              %WS_MINIMAL, _  'OR %PBS_SMOOTH OR %PBS_VERTICAL
              rcDraw.nLeft, rcDraw.nTop, _
              rcWidth, rcHeight, _
              hWndParent,BYVAL %NULL,GetModuleHandle(BYVAL 0),BYVAL %NULL,0 )

      SendMessage hWndNew, %PBM_SETRANGE,0,MAKDWD(0,100)
      SendMessage hWndNew, %PBM_SETSTEP,30,0
      SendMessage hWndNew, %PBM_STEPIT,0,0

  CASE 15   'Up-Down (aka Spin Control)

      hWndNew = myCreateWindowEx( 0,$UPDOWN_CLASS, "", _
          %WS_BORDER OR %WS_MINIMAL _
          OR %UDS_ARROWKEYS OR %UDS_ALIGNLEFT OR %UDS_WRAP, _
          rcDraw.nLeft, rcDraw.nTop, _
          rcWidth, rcHeight, _
          hWndParent,BYVAL %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )
      SendMessage hWndNew, %UDM_SETRANGE, 0&, MAKLNG(10,1)

  CASE 16  : hWndNew = CreateNewListView(hWndParent,rcDraw)   'ListView
  CASE 17  : hWndNew = CreateNewTreeView(hWndParent,rcDraw) 'TreeView
  CASE 18   'TrackBar

        hWndNew = myCreateWindowEx(0,$TRACKBAR_CLASS, "", _
          %WS_MINIMAL OR %TBS_ENABLESELRANGE OR %TBS_AUTOTICKS _
          OR %TBS_HORZ OR %TBS_BOTTOM, _
          rcDraw.nLeft, rcDraw.nTop, _
          rcWidth, rcHeight, _
          hWndParent,BYVAL %NULL, GetModuleHandle(BYVAL 0), BYVAL %NULL,0 )

        SendMessage hWndNew, %TBM_SETRANGE,    %TRUE,  MAKLNG(0,30)
        SendMessage hWndNew, %TBM_SETPAGESIZE, %TRUE,  2
        SendMessage hWndNew, %TBM_SETSEL,      %FALSE, MAKLNG(0,10)
        SendMessage hWndNew, %TBM_SETPOS,      %TRUE,  2


    END SELECT

    IF ISFALSE  hWndNew THEN EXIT FUNCTION                                    'Make sure we have a valid Handle
    SendMessage hWndNew, %WM_SETFONT, GetStockObject(%DEFAULT_GUI_FONT), 0    'Set the default Font
    szOldWndProc = $OldWndProc                                                'SubClass the Control.
    SetProp hWndNew,szOldWndProc, SetWindowLong (hWndNew, %GWL_WNDPROC, BYVAL CODEPTR(ControlSubClassProc))
    'Make sure this control accepts Double Clicks
    SetClassLong  hWndNew, %GCL_STYLE, GetClassLong(hWndNew,%GCL_STYLE) OR %CS_DBLCLKS
    'add WS_TABSTOP style
    SetWindowLong hWndNew, %GWL_STYLE, GetWindowLong(hWndNew,%GWL_STYLE) OR %WS_TABSTOP

    'Bump up the Array count
    i = UBOUND(gZorder) + 1
    REDIM PRESERVE gZorder(i)

    gZorder(i)   = hWndNew
    SetZorder    gZorder(i)          'Put it on top of the List
    hWndControl = gZorder(0)
    BringToFront hWndControl        'Now Paint it on top

    FUNCTION = hWndNew

END FUNCTION

FUNCTION WndClientProc_LBUTTONUP( BYVAL hWnd AS LONG, BYVAL wParam AS LONG , BYVAL lParam AS LONG , _
                  st AS st_types, sf AS flag_type, sp AS setting_type ) AS LONG

  LOCAL  tmp       AS LONG
  LOCAL  rcControl AS RECT
  LOCAL  rcClient  AS RECT
  LOCAL  rc        AS RECT

  LOCAL  pt        AS POINTAPI
  LOCAL  dnc       AS POINTAPI
  LOCAL  dwc       AS POINTAPI


  '----------------------------------------------------------------
  'Dragging a control MouseUp handler
  '----------------------------------------------------------------
  IF sf.Dragging THEN

    ClipCursor BYVAL %NULL
    ReleaseCapture
    SetCursor LoadCursor(%NULL,BYVAL %IDC_ARROW)       'Restore the cursor, the Client window will do this for us anyway
    DrawRectangle2 hWnd, st.ptDrawStart,st.ptDrawEnd    'Remove Last Rectangle

    IF sp.stickFrames THEN                             'Determine if sp.stickFrames = ON
      tmp = IsGroupBox(hWndControl)                 '1= GroupBox :2= EtchedFrame :3= BlackFrame
      IF tmp = 1 OR tmp = 2 OR tmp = 3 THEN GetStickyControls hWndControl, hWnd, st.rcDrag
    END IF

    LOCAL rcOld AS Rect
    GetWindowRect hWndControl,rcOld                 'Save the old coordinates to see if we moved
    'Position the control to its new home
    SetWindowPos  hWndControl,%HWND_TOP,st.rcDrag.nLeft,st.rcDrag.nTop,st.rcDrag.nRight-st.rcDrag.nLeft,st.rcDrag.nBottom -st.rcDrag.nTop,%SWP_NOSIZE OR %SWP_DRAWFRAME
    sf.Dragging = %OFF                                 'Turn off Dragging operation
    sf.changed  = %true

    'added another optimization, to reduce jumpy handles
    IF  rcOld.nLeft = st.rcDrag.nLeft AND rcOld.nTop = st.rcDrag.nTop THEN
      EXIT IF
    ELSE                                               'Last thing to do is Paint our Handles on this control
      DrawHandles     sp.handles
      InvalidateRect hWndControl, BYVAL %NULL,%TRUE
      UpdateWindow   hWndControl
    END IF

    ptEmpty st.ptDrawEnd : ptEmpty st.ptDrawStart : rcEmpty st.rcDrag  'Clear these before we use them again.

  END IF

  '----------------------------------------------------------------
  'Sizing a control MouseUp handler
  '----------------------------------------------------------------
  IF sf.Sizing THEN

    sf.Sizing  = %OFF
    sf.changed = %true         ' set update flag
    ClipCursor BYVAL %NULL
    ReleaseCapture
    DrawRectangle2 hWnd, st.ptDragStart, st.ptDragLast              'Remove the last Sizing Rectangle
    NormalizeRect st.ptDragStart, st.ptDragEnd               'Normalize the Rectangle

    'Set our control to a default minimum value If Zero.
    IF st.ptDragEnd.x =< st.ptDragStart.x +8 THEN st.ptDragEnd.x = st.ptDragEnd.x +8  'min Width
    IF st.ptDragEnd.y =< st.ptDragStart.y +8 THEN st.ptDragEnd.y = st.ptDragEnd.y +8  'min Height

    'Map our Client points to RECT structure
    rc.nLeft   = st.ptDragStart.x
    rc.nTop    = st.ptDragStart.y
    rc.nRight  = st.ptDragEnd.x - st.ptDragStart.x
    rc.nBottom = st.ptDragEnd.y - st.ptDragStart.y

    'Now lets move the control to its new position
    MoveWindow hWndControl, rc.nLeft, rc.nTop, rc.nRight, rc.nBottom, %TRUE

    'Clear these before we use them again.
    ptEmpty st.ptDragStart
    ptEmpty st.ptDragEnd
    ptEmpty st.ptDragLast
    rcEmpty st.nRect

    '---
    'This repaint is strickly for the Static Image Control
    'I need to skip this call if this control is a GroupBox,
    'otherwise the GroupBox does not repaint correctly.

    tmp& = IsGroupBox(hWndControl)
    IF tmp& =1 OR tmp& = 2  OR tmp& = 3 THEN
       EXIT IF
    ELSE
       InvalidateRect hWndControl,BYVAL %NULL,%TRUE
       UpdateWindow   hWndControl
    END IF
    '---

    'Last thing to do is Paint our Handles on this control
    DrawHandles sp.handles

  END IF


  IF sf.Drawing THEN          'Drawing a New Control MouseUp handler
      IF sp.snapGrid = %ON THEN
        dnc.x = LOWRD ( lParam ) - LOWRD ( lParam ) MOD sp.GridSize
        dnc.y = HIWRD ( lParam ) - HIWRD ( lParam ) MOD sp.GridSize
      ELSE
        'ptFromLng lParam, dnc
        dnc.x=LOWRD(lParam)
        dnc.y=HIWRD(lParam)
      END IF
      ClipCursor BYVAL %NULL
      ReleaseCapture
      IF st.Start_dnc.x  > dnc.x THEN                 'Going left?
        st.rcDraw.nLeft  = dnc.x
        st.rcDraw.nRight = st.Start_dnc.x
      ELSE                                            'Going right
        st.rcDraw.nLeft  = st.Start_dnc.x
        st.rcDraw.nRight = dnc.x
      END IF
      IF st.Start_dnc.y < dnc.y THEN                  'Going down?
        st.rcDraw.nBottom = dnc.y
        st.rcDraw.nTop    = st.Start_dnc.y
      ELSE                                            'Going right
        st.rcDraw.nBottom = st.Start_dnc.y
        st.rcDraw.nTop    = dnc.y
      END IF
      st.ptDrawEnd.x   = st.rcDraw.nLeft              'Convert st.rcDraw to points
      st.ptDrawEnd.y   = st.rcDraw.nTop
      st.ptDrawStart.x = st.rcDraw.nRight
      st.ptDrawStart.y = st.rcDraw.nBottom
      DrawRectangle2 hWnd, st.ptDrawStart, st.ptDrawEnd      'Remove the Last Rectangle
      SetCursor LoadCursor(%NULL,BYVAL %IDC_ARROW)    'Restore the cursor, the Client window will do this for us anyway.
      'Set our control to our default minimum values IF Zero.
      'We do this so we don't loose visibility of our control.
      'use a default size of 24
      IF st.ptDrawStart.x =< st.ptDrawEnd.x + 8 THEN st.rcDraw.nRight = st.ptDrawStart.x + 24  '<- *min width
      IF st.ptDrawStart.y =< st.ptDrawEnd.y + 8 THEN st.rcDraw.nBottom = st.ptDrawEnd.y + 24   '<- *min height
      CreateControlEx hWnd, sf.ctype, st.rcDraw, sp
      DrawHandles     sp.Handles                   'add the Handles to the control
      sf.ctype      = 0
      sf.Drawing    = %OFF                                               'The drawing is over
      sf.changed    = %true                                              'set update flag
      ptEmpty st.ptDrawEnd : ptEmpty st.ptDrawStart : rcEmpty st.rcDraw  'Clear these before we use them again
    SENDMESSAGE hFtoolbar, %TB_CHECKBUTTON, %IDP_ACT0, MAKLNG( %TRUE, 0 )
  END IF
  UpdateSizeStatus hWnd,BYVAL VARPTR(rc)             'Update the status window on the Main Header/menu window
  'UpdateSizeStatus wParam, BYVAL lParam
  IF sf.LostCapture = %ON THEN       'Clean up any  GDI garbage left over if we lost our mouse capture.
     sf.LostCapture = %OFF
     InvalidateRect hWnd, BYVAL %NULL,%TRUE
  END IF

END FUNCTION

FUNCTION WndClientProc( BYVAL hWnd AS LONG, _
          BYVAL Msg AS LONG, BYVAL wParam AS LONG, _
          BYVAL lParam AS LONG ) AS LONG
  STATIC st AS st_types           ' draw & move & size rects
  STATIC sf AS flag_type          ' action & drawing flags
  STATIC sp AS setting_type               ' presets
  LOCAL rc AS RECT
  '       STATIC sh AS hChild_type            ' child handles
  SELECT CASE Msg
    CASE %WM_CREATE : FUNCTION = WndClientProc_Create( hWnd, sf, sp )
      'CASE %WM_DESTROY        : FUNCTION = WndClientProc_DESTROY       ( hWnd, sh )
    CASE %WM_SIZE
      'wParam= handle of Control or Group select marquis
      'lParam= lpRect of Control or Group select marquis
      GetWindowRECT hWnd,rc
      UpdateSizeStatus hWnd, BYVAL VARPTR(rc)
      FUNCTION =0 :EXIT FUNCTION
    CASE %WM_MOVE
      GetWindowRECT hWnd,rc
      UpdateSizeStatus hWnd, BYVAL VARPTR(rc)
      FUNCTION =0 :EXIT FUNCTION
    CASE %WM_CAPTURECHANGED : FUNCTION = WndClientProc_CAPTURECHANGED( hWnd, wParam, lParam, sf )
    CASE %WM_COMMAND : FUNCTION = WndClientProc_COMMAND ( hWnd, wParam )
    CASE %WM_PAINT          : FUNCTION = WndClientProc_PAINT         ( hWnd, wParam ,lParam, sp )
    CASE %WM_USER OR %WM_PRINT      : FUNCTION = WndClientProc_PRINT         ( hWnd, wParam ,lParam ) : EXIT FUNCTION
    CASE %WM_CANCELMODE     : FUNCTION = WndClientProc_CANCELMODE    ( hWnd, st, sf )
    CASE %WM_KEYUP          : IF wParam = %VK_DELETE THEN BEEP : EXIT FUNCTION
    CASE %WM_LBUTTONDBLCLK  : FUNCTION = WndClientProc_LBUTTONDBLCLK ( hWnd )
    CASE %WM_RBUTTONDOWN    : FUNCTION = WndClientProc_RBUTTONDOWN   ( hWnd, sf )
    CASE %WM_CONTEXTMENU    : FUNCTION = WndClientProc_CONTEXTMENU   ( hWnd, lParam, sf )
    CASE %WM_LBUTTONDOWN    : FUNCTION = WndClientProc_LBUTTONDOWN   ( hWnd, wParam ,lParam, st, sf, sp )
    CASE %WM_MOUSEMOVE      : FUNCTION = WndClientProc_MOUSEMOVE     ( hWnd, wParam ,lParam, st, sf, sp )
    CASE %WM_LBUTTONUP      : FUNCTION = WndClientProc_LBUTTONUP     ( hWnd, wParam ,lParam, st, sf, sp )
    '''' user messages
    CASE %WM_USER + 150     : sf.cType = wParam : sf.DrawNewControl = lParam : EXIT FUNCTION   ''' set selection from toolwin
    CASE %WM_USER + 151     : Ini_Get "","",sp                               : EXIT FUNCTION   ''' refresh ini para
    CASE %WM_USER + 152     : sf.fWM_RBUTTONUP= wParam : sf.IDhWndCtl=lParam : EXIT FUNCTION
    CASE %WM_USER + 153     : FUNCTION = WndClientProc_AddControl    ( hWnd, wParam, sp )
    CASE %WM_USER + 154     : FUNCTION = sf.changed : EXIT FUNCTION
    CASE %WM_USER + 155     : sf.changed = %false                : EXIT FUNCTION

  FUNCTION = 0
  END SELECT
  FUNCTION = DEFWINDOWPROC( hWnd, msg, wParam, lParam )
END FUNCTION
FUNCTION WndMDIClientProc( BYVAL hWnd AS LONG, _
          BYVAL Msg AS LONG, BYVAL wParam AS LONG, _
          BYVAL lParam AS LONG ) AS LONG
  '       STATIC st AS st_types               ' draw & move & size rects
  '       STATIC sf AS flag_type              ' action & drawing flags
  '       STATIC sp AS setting_type           ' presets
  '       STATIC sh AS hChild_type            ' child handles
  SELECT CASE Msg
    CASE %WM_CREATE                 ': FUNCTION = WndClientProc_Create        ( hWnd, sf, sp, sh )
      '               CASE %WM_DESTROY        : FUNCTION = WndClientProc_DESTROY       ( hWnd, sh )
    CASE %WM_CAPTURECHANGED                 ': FUNCTION = WndClientProc_CAPTURECHANGED( hWnd, wParam, lParam, sf )
      '               CASE %WM_COMMAND        : FUNCTION = WndClientProc_COMMAND       ( hWnd, wParam )
      '               CASE %WM_PAINT          : FUNCTION = WndClientProc_PAINT         ( hWnd, wParam ,lParam, sp )
      '               CASE %WM_USERPRINT      : FUNCTION = WndClientProc_PRINT         ( hWnd, wParam ,lParam ) : EXIT FUNCTION
      '               CASE %WM_CANCELMODE     : FUNCTION = WndClientProc_CANCELMODE    ( hWnd, st, sf )
      '               CASE %WM_KEYUP          :            WndClientProc_KEYUP  hWnd, wParam ,lParam, st, sf, sp, sh : EXIT FUNCTION
      '               CASE %WM_LBUTTONDBLCLK  : FUNCTION = WndClientProc_LBUTTONDBLCLK ( hWnd, sh )
      '               CASE %WM_RBUTTONDOWN    : FUNCTION = WndClientProc_RBUTTONDOWN   ( hWnd, sf )
      '               CASE %WM_CONTEXTMENU    : FUNCTION = WndClientProc_CONTEXTMENU   ( hWnd, lParam, sf, sh )
      '               CASE %WM_LBUTTONDOWN    : FUNCTION = WndClientProc_LBUTTONDOWN   ( hWnd, wParam ,lParam, st, sf, sp, sh )
      '               CASE %WM_MOUSEMOVE      : FUNCTION = WndClientProc_MOUSEMOVE     ( hWnd, wParam ,lParam, st, sf, sp, sh )
      '               CASE %WM_LBUTTONUP      : FUNCTION = WndClientProc_LBUTTONUP     ( hWnd, wParam ,lParam, st, sf, sp, sh )
      '               '''' user messages
      '               CASE %WM_USER + 150     : sf.cType = wParam : sf.DrawNewControl = lParam : EXIT FUNCTION   ''' set selection from toolwin
      '               CASE %WM_USER + 151     : Ini_Get "","",sp                               : EXIT FUNCTION   ''' refresh ini para
      '               CASE %WM_USER + 152     : sf.fWM_RBUTTONUP= wParam : sf.IDhWndCtl=lParam : EXIT FUNCTION
      '               CASE %WM_USER + 153     : FUNCTION = WndClientProc_AddControl    ( hWnd, wParam, sp, sh )
      '               CASE %WM_USER + 154     : FUNCTION = WndClientProc_isChanged( hWnd, sf ) : EXIT FUNCTION
      '               CASE %WM_USER + 155     : WndClientProc_NoChange hWnd, sf                : EXIT FUNCTION
  END SELECT
  FUNCTION = DEFWINDOWPROC( hWnd, msg, wParam, lParam )
END FUNCTION
FUNCTION NewEditWindow( BYVAL hWndParent AS LONG ) AS LONG
  STATIC nForm AS LONG
  LOCAL zText AS ASCIIZ * 30
  LOCAL hWndNew AS LONG
  INCR nForm : zText = "Form" + STR$( nForm )
  '*Create the default Main Frame Window of our Form
  hWndNew = CREATEWINDOWEX( %WS_EX_LEFT OR %WS_EX_LTRREADING OR %WS_EX_WINDOWEDGE, _
          $MainEdit, zText, _
          %WS_VISIBLE OR %WS_CLIPSIBLINGS OR %WS_CLIPCHILDREN OR _
          %WS_BORDER OR %WS_DLGFRAME OR %WS_OVERLAPPEDWINDOW, _
          GETSYSTEMMETRICS( %SM_CXSCREEN ) \ 8, _
          GETSYSTEMMETRICS( %SM_CYSCREEN ) \ 4, _
          GETSYSTEMMETRICS( %SM_CXSCREEN ) \ 2, _
          GETSYSTEMMETRICS( %SM_CYSCREEN ) \ 2, _
          hWndParent, BYVAL %NULL, GETMODULEHANDLE( BYVAL 0 ), BYVAL %NULL )
  FUNCTION = hWndNew
END FUNCTION
FUNCTION NewEditMDIWindow( BYVAL hWndParent AS LONG ) AS LONG
  LOCAL hWndNew AS LONG
  hWndNew = CreateWindow( $MainMDIEdit, _                 ' window class name窗口类名
          "MDI Window", _                 ' window caption窗口标题
          %WS_OVERLAPPEDWINDOW OR _
          %WS_CLIPCHILDREN, _             ' window style 窗口类型
          GETSYSTEMMETRICS( %SM_CXSCREEN ) \ 8, _
          GETSYSTEMMETRICS( %SM_CYSCREEN ) \ 4, _
          GETSYSTEMMETRICS( %SM_CXSCREEN ) \ 2, _
          GETSYSTEMMETRICS( %SM_CYSCREEN ) \ 2, _
          %NULL, _                ' parent window handle父窗口句柄
          0, _            ' window menu handle窗口菜单句柄
          GETMODULEHANDLE( BYVAL 0 ), _           ' program instance handle窗口实例句柄
          BYVAL %NULL )           ' creation parameters创建参数
  MSGBOX "mdiwindow"
  FUNCTION = hWndNew
END FUNCTION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' file : dialog.bas
' preferences dlg for ide
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' windows dialogs

''''''''''''''''''''''''''''''''''''''''''''''''
' colour
' retval = color
FUNCTION pickColor( hWnd AS LONG ) AS DWORD
    LOCAL ColorSpec      AS CHOOSECOLORAPI
    LOCAL lCounter       AS LONG
    DIM   lCustomCol(15) AS LOCAL LONG
    ' Fill custom colors array with all white
    FOR lCounter = 0 TO UBOUND(lCustomCol)
       lCustomCol(lCounter) = &HFFFFFF
    NEXT
    ColorSpec.lStructSize  = LEN(ColorSpec)
    ColorSpec.hwndOwner    = hWnd
    ColorSpec.lpCustColors = VARPTR(lCustomCol(0))
    ColorSpec.rgbResult    = lCustomCol(0)
    ColorSpec.Flags        = %CC_FULLOPEN OR %CC_RGBINIT
    IF ISFALSE ChooseColor(ColorSpec) THEN EXIT FUNCTION
    FUNCTION = VAL ( "&h" & HEX$(ColorSpec.rgbResult,6) )
END FUNCTION

'$GridSize            = "网格大小 "
%use_HANDLESMALL     = 1
%use_HANDLEMEDIUM    = 2
%use_HANDLELARGE     = 3
%use_backColor       = 1
'''' dialog ID's
%IDC_GRID            = 1001
%IDC_SNAP            = 1002
%IDC_SIZECOMBO       = 1003
%IDC_STICKY          = 1005
%IDC_AUTOIMAGE       = 1006
%IDC_HANDLESMALL     = 1008
%IDC_HANDLEMEDIUM    = 1009
%IDC_HANDLELARGE     = 1010
%IDC_BACKDROP        = 1012
%IDC_BACKCOLOR       = 1013
%IDC_BACKTILED       = 1014
%IDC_DEFAULTEDIT     = 1015
%IDC_IMAGECOLORPICK  = 1016
CALLBACK FUNCTION CBShowFormEditorOptions
    STATIC inis AS settings
    LOCAL  test AS LONG
    LOCAL  sTmp AS STRING
    SELECT CASE CBMSG
       CASE %WM_INITDIALOG
            Ini_Get "", "", inis
            CONTROL SET CHECK  CBHNDL, %IDC_GRID, inis.showGrid
            CONTROL SET CHECK  CBHNDL, %IDC_SNAP, inis.snapGrid
            CONTROL SET TEXT   CBHNDL, %IDC_SIZECOMBO,FORMAT$(inis.gridSize) & " × " & FORMAT$(inis.gridSize)
            CONTROL SET CHECK  CBHNDL, %IDC_BACKDROP, inis.backgrndType
            CONTROL SET CHECK  CBHNDL, %IDC_BACKCOLOR, inis.backColor
            IF inis.backColor  THEN
               sTmp = "&H" + HEX$( inis.backColor )
               CONTROL SET TEXT   CBHNDL, %IDC_DEFAULTEDIT , sTmp
            ELSE
            END IF
            CONTROL SET CHECK  CBHNDL, %IDC_STICKY, inis.stickFrames
            CONTROL SET CHECK  CBHNDL, %IDC_AUTOIMAGE, inis.autoSizeImg
            SELECT CASE inis.handles
                CASE %use_HANDLESMALL  : CONTROL SET CHECK  CBHNDL, %IDC_HANDLESMALL,  %true
                CASE %use_HANDLEMEDIUM : CONTROL SET CHECK  CBHNDL, %IDC_HANDLEMEDIUM, %true
                CASE %use_HANDLELARGE  : CONTROL SET CHECK  CBHNDL, %IDC_HANDLELARGE,  %true
            END SELECT
       CASE %WM_COMMAND
          SELECT CASE LOWRD( CBWPARAM )
              CASE %IDCANCEL           : DIALOG END CBHNDL, %false
              CASE %IDC_IMAGECOLORPICK
                  inis.backColor = pickColor( CBHNDL )
                  sTmp = "&H" + HEX$( inis.backColor )
                  CONTROL SET TEXT  CBHNDL, %IDC_DEFAULTEDIT , sTmp
              CASE %IDOK
                  CONTROL GET CHECK CBHNDL, %IDC_GRID         TO inis.showGrid
                  CONTROL GET CHECK CBHNDL, %IDC_SNAP         TO inis.snapGrid
                  CONTROL GET TEXT  CBHNDL, %IDC_SIZECOMBO    TO sTmp : inis.gridSize = VAL ( RIGHT$(sTmp,2) )
                  CONTROL GET CHECK CBHNDL, %IDC_BACKDROP     TO inis.backgrndType
                  CONTROL GET CHECK CBHNDL, %IDC_HANDLESMALL  TO test : IF test THEN inis.handles = %use_HANDLESMALL
                  CONTROL GET CHECK CBHNDL, %IDC_HANDLEMEDIUM TO test : IF test THEN inis.handles = %use_HANDLEMEDIUM
                  CONTROL GET CHECK CBHNDL, %IDC_HANDLELARGE  TO test : IF test THEN inis.handles = %use_HANDLELARGE
                  CONTROL GET CHECK CBHNDL, %IDC_STICKY       TO inis.stickFrames
                  CONTROL GET CHECK CBHNDL, %IDC_AUTOIMAGE    TO inis.autoSizeImg
                  Ini_Set "", "", inis
                  DIALOG END CBHNDL, %true
          END SELECT
    END SELECT
END FUNCTION
FUNCTION ShowFormEditorOptions(BYVAL hParent AS LONG ) AS LONG
    LOCAL hDlg AS LONG
    DIM optionTxt(1 TO 7) AS LOCAL STRING
    optionTxt(1) = "4 × 4"
    optionTxt(2) = "6 × 6"
    optionTxt(3) = "8 × 8"
    optionTxt(4) = "10 × 10"
    optionTxt(5) = "12 × 12"
    optionTxt(6) = "14 × 14"
    optionTxt(7) = "16 × 16"
    DIALOG  NEW hParent,  "窗口编辑器设置", ,, 352, 208, %WS_CAPTION OR %WS_SYSMENU OR %DS_CENTER, %WS_EX_TOOLWINDOW TO hDlg
    CONTROL ADD LABEL,    hDlg, 0,    "VFE 边框形式", 11, 10, 155, 10
    CONTROL ADD LINE,     hDlg, 0,    "", 5, 25, 343, 2
    CONTROL ADD LABEL,    hDlg, 0,    "", 251, 26, 97, 143, %WS_CHILD OR %WS_VISIBLE , %WS_EX_CLIENTEDGE
    CONTROL ADD IMAGE,    hDlg, 0,    "setup", 257, 39, 91, 123,, %WS_EX_CLIENTEDGE
    CONTROL ADD FRAME,    hDlg, 0, "网格",17,34,97,69
    CONTROL ADD CHECKBOX, hDlg, %IDC_GRID, "显示/隐藏网格", 32,49,80,10, %BS_AUTOCHECKBOX OR %WS_TABSTOP
    CONTROL ADD CHECKBOX, hDlg, %IDC_SNAP, "开/关对齐风格", 32,64,80,10, %BS_AUTOCHECKBOX OR %WS_TABSTOP
    CONTROL ADD LABEL,    hDlg, 0, "网格尺寸:",32,80,50,10
    CONTROL ADD COMBOBOX, hDlg, %IDC_SIZECOMBO,optionTxt(), 67,79,40,50, %CBS_DROPDOWN OR %WS_VSCROLL OR %WS_TABSTOP
    CONTROL ADD FRAME,    hDlg, 0, "背景",17,112,97,82
    CONTROL ADD CHECKBOX, hDlg, %IDC_BACKDROP,  "背景",  32,127,80,11 , %BS_AUTOCHECKBOX OR %WS_TABSTOP
    CONTROL ADD OPTION,   hDlg, %IDC_BACKCOLOR, "颜色", 32,141,38,10, %WS_GROUP OR %WS_TABSTOP
    CONTROL ADD OPTION,   hDlg, %IDC_BACKTILED,  "平铺图象",  32,154,48,10
    CONTROL ADD TEXTBOX,  hDlg, %IDC_DEFAULTEDIT, "", 32,170,58,12, %WS_CHILD OR %WS_VISIBLE OR %ES_READONLY, %WS_EX_CLIENTEDGE
    CONTROL ADD BUTTON,   hDlg, %IDC_IMAGECOLORPICK, "...", 90,170,13,12, %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP
    CONTROL ADD FRAME,    hDlg, 0, "其他",133,34,102,69
    CONTROL ADD CHECKBOX, hDlg, %IDC_STICKY, "粘连边框 开/关", 145,49,80,10, %BS_AUTOCHECKBOX OR %WS_TABSTOP
    CONTROL ADD CHECKBOX, hDlg, %IDC_AUTOIMAGE, "自动调整图象尺寸", 145,66,86,10, %BS_AUTOCHECKBOX OR %WS_TABSTOP
    CONTROL ADD FRAME,    hDlg, 0, "句柄",133,111,102,58
    CONTROL ADD OPTION,   hDlg, %IDC_HANDLESMALL,  "小尺寸",  145,126,33,10 , %WS_GROUP OR %WS_TABSTOP
    CONTROL ADD OPTION,   hDlg, %IDC_HANDLEMEDIUM, "中尺寸", 145,139,40,10
    CONTROL ADD OPTION,   hDlg, %IDC_HANDLELARGE,  "大尺寸",  145,152,34,10
    CONTROL ADD BUTTON,   hDlg, %IDCANCEL, "取消",198,181, 65, 15
    CONTROL ADD BUTTON,   hDlg, %IDOK, "确定", 283,181, 65, 15, %BS_DEFAULT
  DIALOG SET TEXT hDlg,"  窗口编辑器设置"
    DIALOG SHOW MODAL hDlg CALL CBShowFormEditorOptions
END FUNCTION
