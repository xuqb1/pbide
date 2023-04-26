''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 文件 : findrs.inc
' 查找结果展示窗口
'
GLOBAL hFindrsLB AS DWORD
%IDC_FINDRSLB = 101
' 创建项目窗口
FUNCTION CreateFindrsWin(BYVAL gdiI AS LONG) AS LONG
  LOCAL i AS LONG
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  i=gdiI
  gdi(i).ID = %ID_FINDRS
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
  ' 创建泊坞窗口
  gdi(i).hWndD = CreateWindowEx( 0 , "DOCKIT", "查找结果", _
                                  %WS_CHILD, _
                                  gdi(i).drc.nLeft, gdi(i).drc.nTop, _
                                  gdi(i).drc.nRight, gdi(i).drc.nBottom, _  ' width,height
                                  g_hWndMain, %ID_FINDRS, g_hInst, BYVAL %NULL)
  SetWindowLong gdi(i).hWndD,%GWL_USERDATA,i
  ' 创建浮动窗口
  gdi(i).hWndF = CreateWindowEx( 0 , "DOCKIT", "查找结果", _
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
  CreateFindrsContents gdi(i).hWnd
  FUNCTION = gdi(i).hWnd
END FUNCTION
' 创建工具箱窗口中的控件：工具条，使用静态控件包装，这样就可以随窗口大小调整而对工具条进行调整
FUNCTION CreateFindrsContents(BYVAL hWnd AS DWORD)AS LONG
  LOCAL rc AS RECT
  LOCAL hRoot AS DWORD
  GetClientRect hWnd,rc
  'control add treeview,hWnd,%IDC_PROJECTTV,"qq",rc.nLeft+8,rc.nTop+25,rc.nRight-16,rc.nBottom-30 call ProjectProc
  hFindrsLB = CreateWindowEx(%WS_EX_STATICEDGE,"LISTBOX","Compile Completed", _          '%WS_EX_TRANSPARENT     textbox
                                %WS_CHILD OR %WS_VISIBLE, _
                                rc.nLeft+4,rc.nTop+23, _
                                rc.nRight-8,rc.nBottom-28, _
                                hWnd,%IDC_FINDRSLB, g_hInst, BYVAL %NULL )

  SetWindowLong hFindrsLB,%GWL_WNDPROC,CODEPTR(FindrstbProc)
END FUNCTION
'' 由ProjectProc调用，响应按钮通知消息，提示相关的信息
'FUNCTION Project_NOTIFY( BYVAL hWnd AS LONG, BYVAL wParam AS LONG, BYVAL lParam AS LONG ) AS LONG
'  LOCAL lpToolTip AS TOOLTIPTEXT PTR
'  LOCAL zText AS ASCIIZ * 40
'  lpToolTip = lParam
'  IF @lpToolTip.hdr.code = %TTN_NEEDTEXT THEN
'    SELECT CASE @lpToolTip.hdr.idFrom - %IDP_ACT0
'      CASE 0 : zText = "选择"
'      CASE 1 : zText = $WC_STATIC
'      CASE 2 : zText = $WC_BUTTON
'      CASE 3 : zText = $WC_EDIT
'      CASE 4 : zText = $CW_CHECKBOX
'      CASE 5 : zText = $CW_OPTION
'      CASE 6 : zText = $CW_FRAME
'      CASE 7 : zText = $CW_IMAGE
'      CASE 8 : zText = $WC_COMBOBOX
'      CASE 9 : zText = $WC_LISTBOX
'      CASE 10 : zText = $WC_SCROLLBAR
'      CASE 11 : zText = $WC_SCROLLBAR
'      CASE 12 : zText = $WC_TABCONTROL
'      CASE 13 : zText = $CW_RICHEDIT
'      CASE 14 : zText = $PROGRESS_CLASS
'      CASE 15 : zText = $UPDOWN_CLASS
'      CASE 16 : zText = $WC_LISTVIEW
'      CASE 17 : zText = $WC_TREEVIEW
'      CASE 18 : zText = $TRACKBAR_CLASS
'      CASE ELSE : zText = "不支持"
'    END SELECT
'    @lpToolTip.lpszText = VARPTR( zText )
'  END IF
'END FUNCTION
' 工具条中的按钮事件响应，包括通知和按下(单击)
CALLBACK FUNCTION FindrstbProc()AS LONG
  IF hFindrsLB<=0 THEN
    hFindrsLB=CB.HNDL
    MSGBOX STR$(hFindrsLB)
  END IF
  SELECT CASE CB.MSG
    CASE %WM_INITDIALOG

    CASE %WM_SIZE

'      SENDMESSAGE hFtoolbar, %TB_AUTOSIZE, 0, 0
    'CASE %WM_NOTIFY : ToolProc_NOTIFY CB.HNDL, CB.WPARAM, CB.LPARAM
    CASE %WM_COMMAND
      'LOCAL lp_sChld AS childWin_types PTR
      'lp_sChld =
'      SELECT CASE LOWRD( CB.WPARAM )
'        CASE %IDP_ACT0 TO %IDP_ACT18 :
'          'sendmessage hWndEdit, %WM_USER+150, LOWRD( wParam )-%IDP_ACT0, %ON
'        CASE %IDP_ACT19 :               'sendmessage hWndEdit, %WM_USER+150, 0, %OFF
'          'SENDMESSAGE hFtoolbar, %TB_CHECKBUTTON, %IDP_ACT0, MAKLNG( %TRUE, 0 )
'      END SELECT
      EXIT FUNCTION
  END SELECT
  FUNCTION = DEFWINDOWPROC( CBHNDL, CBMSG, CBWPARAM, CBLPARAM )
END FUNCTION
'' 创建工具条
'FUNCTION ToolProc_CREATE( BYVAL hWnd AS LONG) AS LONG
'  LOCAL i AS LONG
'  LOCAL hTBar AS LONG
'  DIM ftbb( 19 ) AS LOCAL TBBUTTON
'  FOR i = LBOUND( ftbb ) TO UBOUND( ftbb )
'    ftbb( i ).iBitmap = i
'    ftbb( i ).idCommand = %IDP_ACT0 + i
'    ftbb( i ).fsState = %TBSTATE_ENABLED
'    ftbb( i ).fsStyle = %TBSTYLE_CHECKGROUP OR %TBSTYLE_AUTOSIZE
'  NEXT i
'  hToolBmp = LOADBITMAP( g_hInst, "PAL" )
'  'msgbox "hToolBmp=" & str$(hToolBmp)
'  hTBar = CREATETOOLBAREX( hWnd, _
'                        %WS_CHILD OR %TBSTYLE_FLAT OR %TBSTYLE_TOOLTIPS OR _            'Create the toolbar for the Floating window
'                        %WS_CLIPSIBLINGS OR %TBSTYLE_WRAPABLE OR %TBSTYLE_LIST OR _
'                        %WS_VISIBLE OR %CCS_NODIVIDER ,%ID_FTOOLBAR, 21, %NULL, hToolBmp, _
'                        ftbb( 0 ), 20, 0, 0, %ButtnSize, %ButtnSize, _
'                        LEN( TBBUTTON ))
'  SENDMESSAGE hTBar, %TB_AUTOSIZE, 0, 0
'  SENDMESSAGE hTBar, %TB_CHECKBUTTON, %IDP_ACT0, MAKLNG( %TRUE, 0 )
'  FUNCTION = hTBar
'END FUNCTION
' 编译结果展示窗口大小调整，同步调整编译结果展示窗口中的控件位置大小
FUNCTION SizeFindrsContents(BYVAL hWnd AS DWORD,BYVAL wParam AS DWORD)AS LONG
  LOCAL rc AS RECT
  LOCAL i AS LONG
  LOCAL gdi1 AS DockInfo
  LOCAL hCtrl AS DWORD
  RESET gdi1
  'i=FindGdiIndex(%ID_FINDRS)
  i=GetWindowLong(hWnd,%GWL_USERDATA)
  IF i<0 OR i>UBOUND(gdi()) THEN
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
  'canLog=1 : RunLog(gdi1.l & " ml=" & gdi1.ml) : canLog=0
  SetRect rc,rc.nLeft+4,rc.nTop+23,rc.nRight-8,rc.nBottom-28
  MoveWin hFindrsLB,rc,%TRUE
END FUNCTION
' 项目中按钮事件响应
FUNCTION Command_Findrs(BYVAL wParam AS DWORD)AS LONG
  'SendMessage hToolContain,%WM_COMMAND,wParam,0
END FUNCTION
' 窗口销毁事件响应
FUNCTION DestroyFindrsWin() AS LONG
  'IF hToolBmp <> 0 THEN DELETEOBJECT hToolBmp
END FUNCTION
' 将窗口中的控件移到指定句柄窗口中：切换泊坞与浮动状态时发生
FUNCTION MoveFindrsContents(BYVAL hWnd AS DWORD)AS LONG
  SetParent hFindrsLB, hWnd
END FUNCTION
