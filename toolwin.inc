''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 文件 : toolwin.inc
' 控件窗口：为可视窗口编辑功能选择控件
'
' 控件窗口中的控件ID
%IDP_ACT0               = 40100
%IDP_ACT1               = 40101
%IDP_ACT2               = 40102
%IDP_ACT3               = 40103
%IDP_ACT4               = 40104
%IDP_ACT5               = 40105
%IDP_ACT6               = 40106
%IDP_ACT7               = 40107
%IDP_ACT8               = 40108
%IDP_ACT9               = 40109
%IDP_ACT10              = 40110
%IDP_ACT11              = 40111
%IDP_ACT12              = 40112
%IDP_ACT13              = 40113
%IDP_ACT14              = 40114
%IDP_ACT15              = 40115
%IDP_ACT16              = 40116
%IDP_ACT17              = 40117
%IDP_ACT18              = 40118
%IDP_ACT19              = 40119
%ON               = 1
%OFF              = 0
' 创建控件基本类
'$WC_BUTTON      = "BUTTON"
'$WC_COMBOBOX    = "COMBOBOX"
'$WC_EDIT        = "EDIT"
'$WC_LISTBOX     = "LISTBOX"
$WC_MDICLIENT   = "MDICLIENT"
'$WC_SCROLLBAR   = "SCROLLBAR"
'$WC_STATIC      = "STATIC"
$CW_EDIT        = "TEXTBOX"
$CW_LINE        = "LINE"
$CW_LABEL       = "LABEL"
$CW_IMAGE       = "IMAGE"
$CW_IMAGEX      = "IMAGEX"
$CW_FRAME       = "FRAME"
$CW_OPTION      = "OPTION"
$CW_CHECKBOX    = "CHECKBOX"
'$CW_IMGBUTTON   = "IMGBUTTON"
'$CW_IMGBUTTONX  = "IMGBUTTONX"
'$CW_CHECK3STATE = "CHECK3STATE"
$CW_RICHEDIT    = "RICHEDIT"

%ButtnSize = 24     '按钮大小
'%ID_CONTRILS  = 111   '控件窗口ID，转移到dock.inc
GLOBAL hFtoolbar    AS DWORD
GLOBAL hToolContain AS DWORD
GLOBAL hToolBmp     AS LONG
' 创建控件窗口
FUNCTION CreateToolboxWin(BYVAL gdiI AS LONG) AS LONG
  LOCAL i AS LONG
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  i=gdiI
  gdi(i).ID = %ID_CONTROLS
  IF gdi(i).minH<=0 THEN gdi(i).minH=100
  IF gdi(i).minW<=0 THEN gdi(i).minW=50
'  IF gdi(i).drc.nRight<gdi(i).minW THEN
'    gdi(i).drc.nRight=gdi(i).minW
'  END IF
'  IF gdi(i).drc.nBottom<gdi(i).minH THEN
'    gdi(i).drc.nBottom=gdi(i).minH
'  END IF
'  IF gdi(i).frc.nRight<gdi(i).minW THEN
'    gdi(i).frc.nRight=gdi(i).minW
'  END IF
'  IF gdi(i).frc.nBottom<gdi(i).minH THEN
'    gdi(i).frc.nBottom=gdi(i).minH
'  END IF
  ' 创建控件泊坞窗口
  gdi(i).hWndD = CreateWindowEx( 0 , "DOCKIT", "控件", _
                                  %WS_CHILD, _
                                  gdi(i).drc.nLeft, gdi(i).drc.nTop, _
                                  gdi(i).drc.nRight, gdi(i).drc.nBottom, _  ' width,height
                                  g_hWndMain, %ID_CONTROLS, g_hInst, BYVAL %NULL)
  SetWindowLong gdi(i).hWndD,%GWL_USERDATA,i
  ' 创建控件浮动窗口
  gdi(i).hWndF = CreateWindowEx( 0 , "DOCKIT", "控件", _
                                  %WS_POPUP  OR %WS_BORDER OR %WS_THICKFRAME, _
                                  gdi(i).frc.nLeft, gdi(i).frc.nTop, _
                                  gdi(i).frc.nRight, gdi(i).frc.nBottom, _  ' width,height
                                  g_hWndMain, %NULL, g_hInst, BYVAL %NULL)
  SetWindowLong gdi(i).hWndF,%GWL_USERDATA,i
  IF gdi(i).l="F" OR gdi(i).l="I" THEN
    gdi(i).hWnd=gdi(i).hWndF
  ELSE
    gdi(i).hWnd=gdi(i).hWndD
    IF gdi(i).tid<=0 THEN
      gdi(i).tid=%ID_CLIENTWINDOW
    END IF
  END IF
  IF gdi(i).l<>"I" AND gdi(i).ml<>"I" THEN
    ShowWindow gdi(i).hWnd, %SW_SHOW
    UpdateWindow gdi(i).hWnd
    IF gdi(i).l="F" THEN
      gdi(i).byside=0
    ELSEIF gdi(i).byside<=0 THEN
      SetGdiByside i
    END IF
  END IF
  CreateToolboxContents gdi(i).hWnd
  FUNCTION = gdi(i).hWnd
END FUNCTION
' 创建控件窗口中的控件：工具条，使用静态控件包装，这样就可以随窗口大小调整而对工具条进行调整
FUNCTION CreateToolboxContents(BYVAL hWnd AS DWORD)AS LONG
  LOCAL rc AS RECT
  GetClientRect hWnd,rc
  hToolContain = CreateWindowEx(0,"STATIC","", _          '%WS_EX_TRANSPARENT
                                %WS_CHILD OR %WS_VISIBLE OR %WS_BORDER , _
                                rc.nLeft+4,rc.nTop+23, _
                                rc.nRight-8,rc.nBottom-28, _
                                hWnd,BYVAL %NULL, g_hInst, BYVAL %NULL )
  hFtoolbar = ToolProc_CREATE( hToolContain) ', hToolBmp )
  SetWindowLong hToolContain,%GWL_WNDPROC,CODEPTR(ToolContainProc)
END FUNCTION
' 由ToolContainProc调用，响应按钮通知消息，提示相关的信息
FUNCTION ToolProc_NOTIFY( BYVAL hWnd AS LONG, BYVAL wParam AS LONG, BYVAL lParam AS LONG ) AS LONG
  LOCAL lpToolTip AS TOOLTIPTEXT PTR
  LOCAL zText AS ASCIIZ * 40
  lpToolTip = lParam
  IF @lpToolTip.hdr.code = %TTN_NEEDTEXT THEN
    SELECT CASE @lpToolTip.hdr.idFrom - %IDP_ACT0
      CASE 0 : zText = "选择"
      CASE 1 : zText = $WC_STATIC
      CASE 2 : zText = $WC_BUTTON
      CASE 3 : zText = $WC_EDIT
      CASE 4 : zText = $CW_CHECKBOX
      CASE 5 : zText = $CW_OPTION
      CASE 6 : zText = $CW_FRAME
      CASE 7 : zText = $CW_IMAGE
      CASE 8 : zText = $WC_COMBOBOX
      CASE 9 : zText = $WC_LISTBOX
      CASE 10 : zText = $WC_SCROLLBAR
      CASE 11 : zText = $WC_SCROLLBAR
      CASE 12 : zText = $WC_TABCONTROL
      CASE 13 : zText = $CW_RICHEDIT
      CASE 14 : zText = $PROGRESS_CLASS
      CASE 15 : zText = $UPDOWN_CLASS
      CASE 16 : zText = $WC_LISTVIEW
      CASE 17 : zText = $WC_TREEVIEW
      CASE 18 : zText = $TRACKBAR_CLASS
      CASE ELSE : zText = "不支持"
    END SELECT
    @lpToolTip.lpszText = VARPTR( zText )
  END IF
END FUNCTION
' 工具条中的按钮事件响应，包括通知和按下(单击)
CALLBACK FUNCTION ToolContainProc()AS LONG
  SELECT CASE CB.MSG
'    case %WM_SIZE
'      SENDMESSAGE hFtoolbar, %TB_AUTOSIZE, 0, 0
    CASE %WM_NOTIFY : ToolProc_NOTIFY CB.HNDL, CB.WPARAM, CB.LPARAM
    CASE %WM_COMMAND
      'LOCAL lp_sChld AS childWin_types PTR
      'lp_sChld =
      SELECT CASE LOWRD( CB.WPARAM )
        CASE %IDP_ACT0 TO %IDP_ACT18 :
          sendmessage hWndEdit, %WM_USER+150, LOWRD( CB.WPARAM )-%IDP_ACT0, %ON
        CASE %IDP_ACT19 :               'sendmessage hWndEdit, %WM_USER+150, 0, %OFF
          'SENDMESSAGE hFtoolbar, %TB_CHECKBUTTON, %IDP_ACT0, MAKLNG( %TRUE, 0 )
          CONTROL SET CHECK CB.HNDL,%IDP_ACT0,1
      END SELECT
      EXIT FUNCTION
  END SELECT
  FUNCTION = DEFWINDOWPROC( CBHNDL, CBMSG, CBWPARAM, CBLPARAM )
END FUNCTION
' 创建工具条
FUNCTION ToolProc_CREATE( BYVAL hWnd AS LONG) AS LONG
  LOCAL i AS LONG
  LOCAL hTBar AS LONG
  DIM ftbb( 19 ) AS LOCAL TBBUTTON
  FOR i = LBOUND( ftbb ) TO UBOUND( ftbb )
    ftbb( i ).iBitmap = i
    ftbb( i ).idCommand = %IDP_ACT0 + i
    ftbb( i ).fsState = %TBSTATE_ENABLED
    ftbb( i ).fsStyle = %TBSTYLE_CHECKGROUP OR %TBSTYLE_AUTOSIZE
  NEXT i
  hToolBmp = LOADBITMAP( g_hInst, "PAL" )
  'msgbox "hToolBmp=" & str$(hToolBmp)
  hTBar = CREATETOOLBAREX( hWnd, _
                        %WS_CHILD OR %TBSTYLE_FLAT OR %TBSTYLE_TOOLTIPS OR _            'Create the toolbar for the Floating window
                        %WS_CLIPSIBLINGS OR %TBSTYLE_WRAPABLE OR %TBSTYLE_LIST OR _
                        %WS_VISIBLE OR %CCS_NODIVIDER ,%ID_FTOOLBAR, 21, %NULL, hToolBmp, _
                        ftbb( 0 ), 20, 0, 0, %ButtnSize, %ButtnSize, _
                        LEN( TBBUTTON ))
  SENDMESSAGE hTBar, %TB_AUTOSIZE, 0, 0
  SENDMESSAGE hTBar, %TB_CHECKBUTTON, %IDP_ACT0, MAKLNG( %TRUE, 0 )
  FUNCTION = hTBar
END FUNCTION
' 控件窗口大小调整，同步调整工具条中的控件位置大小
FUNCTION SizeToolboxContents(BYVAL hWnd AS DWORD,BYVAL wParam AS DWORD)AS LONG
  LOCAL rc AS RECT
  LOCAL i AS LONG
  LOCAL gdi1 AS DockInfo
  RESET gdi1
  i=FindGdiIndex(%ID_CONTROLS)
  IF i=-1 THEN
    EXIT FUNCTION
  END IF
  gdi1=gdi(i)
  IF gdi1.id=0 THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  GetClientRect hWnd,rc
  IF (gdi1.ml="B" AND gdi1.byside=(gdi1.byside OR %BYTOP)) OR gdi1.l="B" THEN
    rc.nTop+=5
    rc.nbottom-=5
  END IF
  SetRect rc,rc.nLeft+4,rc.nTop+23,rc.nRight-8,rc.nBottom-28
  MoveWin hToolContain,rc,%TRUE
  'SendMessage hToolContain,%WM_PAINT,0,0  '必须注释掉，放开时，关联窗口调整大小，会有影响(刷掉边框)
  SendMessage hFtoolbar, %TB_AUTOSIZE, 0, 0
  UpdateWindow hToolContain
END FUNCTION
' 控件中按钮事件响应
FUNCTION Command_Toolbox(BYVAL wParam AS DWORD)AS LONG
  SendMessage hToolContain,%WM_COMMAND,wParam,0
END FUNCTION
' 控件窗口销毁事件响应
FUNCTION DestroyToolboxWin() AS LONG
  IF hToolBmp <> 0 THEN DELETEOBJECT hToolBmp
END FUNCTION
' 将控件窗口中的控件移到指定句柄窗口中：切换泊坞与浮动状态时发生
FUNCTION MoveToolboxContents(BYVAL hWnd AS DWORD)AS LONG
  SetParent hToolContain, hWnd
END FUNCTION
