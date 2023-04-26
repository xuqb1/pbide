''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 文件 : property.inc
' 属性窗口：为可视窗口编辑功能设置控件属性
'
TYPE PROPERTYITEM        '属性自定义类型
  forObj      AS ASCIIZ*255 '确定用于哪种对象，如窗口，按钮，文本框等，值用"|"分隔开，全用时值为"*"，如“名称”属性
  propName    AS ASCIIZ*225 '属性名
  curValue    AS ASCIIZ*255 '当前值
  nItemType   AS INTEGER    '项类型
  cmbItems    AS ASCIIZ*255 '命令项
  tip         AS ASCIIZ*255 '提示
END TYPE
TYPE PROPERTYSET         '把属性列表中设置的属性值与实际编辑程序中对应的参数进行联系，动态修改
  forObje       AS ASCIIZ*255'确定用于何种对象
  propName      AS ASCIIZ*255'确定属性名
  curValue(10)  AS ASCIIZ*255'属性值
  Para(10)      AS ASCIIZ*255 '设置的程序中的对应参数
END TYPE
GLOBAL pItem()        AS PROPERTYITEM '用户设计的自定义控件信息
GLOBAL ghCombo1       AS DWORD   '属性窗口中最上面的下拉列表控件句柄
GLOBAL ghList         AS DWORD   '属性列表控件句柄
GLOBAL ghTip          AS LONG         '属性窗口中的属性信息提示文本框句柄
GLOBAL nDivider       AS INTEGER      '属性列表中的列分隔线位置
GLOBAL bDivIsSet      AS INTEGER
GLOBAL lpOldListProc  AS LONG    '用于绑定属性列表的处理过程而返回的句柄
GLOBAL curSel         AS INTEGER '属性列表中当前选择项
GLOBAL oldSel         AS INTEGER '属性列表中上次选择项
'GLOBAL ghFont   AS LONG          '设定属性列表中字的字体的句柄
GLOBAL ghButton       AS LONG          '属性列表中的按钮句柄
GLOBAL ghCombo        AS LONG          '属性列表中的下拉列表句柄
GLOBAL ghEdit         AS LONG          '属性列表中的文本框句柄
GLOBAL cColor()       AS LONG
GLOBAL prevSel        AS INTEGER
GLOBAL nLastBox       AS INTEGER
%IDC_LISTBOX = 1000
%IDC_COMBO   = 1001
%IDC_BUTTON  = 1002
%IDC_EDITS   = 1003
%IDC_UCC     = 1004
%IDC_LABEL   = 1005
%IDC_COMBO1  = 1006
'PIT = 属性项目类型，按钮为默认
%PIT_COMBO   = 0
%PIT_EDIT    = 1
%PIT_COLOR   = 2
%PIT_FONT    = 3
%PIT_FILE    = 4
' 创建属性窗口
FUNCTION CreatePropertyWin(BYVAL gdiI AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  '设定泊坞窗口的默认状态或从ini文件得到
  REDIM pItem(16)
  nDivider  = 0
  i = gdiI
  gdi(i).ID=%ID_PROPERTY
  '检查并初始窗口最小宽高
  'msgbox "4: " & RcToStr(gdi(i).frc)
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
  ' 创建属性泊坞窗口
  gdi(i).hWndD = CreateWindowEx(  0,"DOCKIT", _
                              "属性", _
                              %WS_CHILD, _
                              gdi(i).drc.nLeft, gdi(i).drc.nTop, _
                              gdi(i).drc.nRight,gdi(i).drc.nBottom, _   ' width,height
                              g_hWndMain, %ID_PROPERTY, g_hInst, BYVAL %NULL )
  SetWindowLong gdi(i).hWndD,%GWL_USERDATA,i
  ' 创建属性浮动窗口
  'rc=gdi(i).frc
  'ScreenToClientRect g_hWndMain,rc
  'msgbox "frc=" & RcToStr(gdi(i).frc) & " rc=" & RcToStr(rc)
  gdi(i).hWndF = CreateWindowEx(0,"DOCKIT", _
                              "属性", _
                              %WS_POPUP  OR %WS_BORDER OR %WS_THICKFRAME, _
                              gdi(i).frc.nLeft, gdi(i).frc.nTop, _
                              gdi(i).frc.nRight,gdi(i).frc.nBottom, _   ' width,height
                              g_hWndMain, %NULL, g_hInst, BYVAL %NULL )
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
    'msgbox gdi(i).l
    ShowWindow gdi(i).hWnd, %SW_SHOW
    'UpdateWindow gdi(i).hWnd
    IF gdi(i).l="F" THEN
      gdi(i).byside=0
    ELSEIF gdi(i).byside=0 THEN
      SetGdiByside i
    END IF
    'SendMessage gdi(i).hWnd,%WM_SIZE,0,0
  END IF
  CreatePropContents gdi(i).hWnd
  FUNCTION = gdi(i).hWnd
END FUNCTION
' 创建属性窗口中的控件
FUNCTION CreatePropContents(BYVAL hWnd AS DWORD)AS LONG
  LOCAL rc AS RECT
  LOCAL x1() AS STRING
  LOCAL a1 AS STRING
  LOCAL i AS INTEGER
  LOCAL gdi1 AS DOCKINFO
  REDIM x1(3)
  i=GetWindowLong(hWnd,%GWL_USERDATA)
  GetClientRect hWnd, rc
'  gdi1=gdi(i)
'  IF gdi(i).l="F" THEN
'    hWnd=gdi(i).hWndF
'    GetClientRect hWnd, rc
'  END IF
'  IF (gdi1.ml="B" AND gdi1.byside=(gdi1.byside OR %BYTOP)) OR gdi1.l="B" THEN
'    rc.nTop+=5
'    rc.nBottom-=5
'  END IF
  SetRect rc,rc.nLeft+4,rc.nTop+23,rc.nRight-8,rc.nBottom-28
  ghCombo1 = CreateWindowEx(  %WS_EX_STATICEDGE,"ComboBox", "", _
                              %WS_CHILD OR %WS_VISIBLE OR %WS_VSCROLL OR _
                              %CBS_DROPDOWNLIST OR %CBS_NOINTEGRALHEIGHT, _
                              rc.nLeft, rc.nTop, _
                              rc.nRight, 100,_
                              hWnd, %IDC_COMBO1, g_hInst, BYVAL %NULL )
  ghList   = CreateWindowEx(  %WS_EX_STATICEDGE,"LISTBOX", "", _
                              %WS_CHILD OR %WS_VISIBLE OR _
                              %WS_TABSTOP OR %WS_VSCROLL OR _
                              %LBS_NOINTEGRALHEIGHT OR %LBS_HASSTRINGS OR _
                              %LBS_OWNERDRAWFIXED OR %LBS_NOTIFY ,_
                              rc.nLeft, rc.nTop + 25, _
                              rc.nRight,rc.nBottom-80,_
                              hWnd, %IDC_LISTBOX, g_hInst, BYVAL %NULL )
  ghTip    =  CreateWindowEx( %WS_EX_STATICEDGE, "EDIT", "", _
                              %WS_CHILD OR %WS_VISIBLE OR _
                              %ES_MULTILINE OR %ES_AUTOVSCROLL OR %ES_READONLY , _
                              rc.nLeft,rc.nTop+rc.nBottom-55,rc.nRight,50, _
                              hWnd, %IDC_LABEL, g_hInst, BYVAL %NULL)
'  GetWindowRect ghCombo1,rc
'  rc.nRight-=rc.nLeft
'  rc.nBottom-=rc.nTop
'  MSGBOX "ghCombo1 rc=" & RcToStr(rc)
  CALL SendMessage( ghCombo1, %WM_SETFONT, ghFont, 0)
  CALL SendMessage( ghTip, %WM_SETFONT, ghFont, 0)
  '对属性列表应用子类，即绑定其处理过程
  lpOldListProc = SetWindowLong( ghList,%GWL_WNDPROC,BYVAL CODEPTR(ListBoxProc) )
  CALL SendMessage( ghList, %WM_SETFONT, ghFont, 0 )
  '创建属性列表中的控件。
  '现在他们都是隐藏的，下一步就会设置大小
  ghButton = CreateWindowEx(  0,"BUTTON","┅" , %WS_CHILD, _
                              10,10,30,10, _
                              ghList, %IDC_BUTTON, g_hInst, BYVAL %NULL)
  CALL SendMessage( ghButton, %WM_SETFONT, ghFont, 0 )
  ghCombo  = CreateWindowEx(  0,"COMBOBOX", "", _
                              %WS_CHILD OR %WS_VISIBLE  OR _
                              %CBS_DROPDOWNLIST OR %CBS_DISABLENOSCROLL, _
                              10,30,30,100, _
                              ghList, %IDC_COMBO, g_hInst, BYVAL %NULL)
  CALL SendMessage( ghCombo, %WM_SETFONT, ghFont, 0 )
  ghEdit   = CreateWindowEx(  0,"EDIT", "True", _
                              %WS_CHILD OR %ES_AUTOHSCROLL , _
                              10,50,30,10, _
                              ghList, %IDC_EDITS, g_hInst, BYVAL %NULL)
  CALL SendMessage( ghEdit, %WM_SETFONT, ghFont, 0 )
  ShowWindow ghButton,%SW_HIDE
  ShowWindow ghCombo,%SW_HIDE
  ShowWindow ghEdit,%SW_HIDE
  '-------------------------------------------------------------------
  ' 这里的CALL调用必须在属性列表创建之后
  CALL InitializePropertyItems()
  CALL InitializeCustomColors()
  CALL SetFocus( ghList )
  x1(0)="United States"
  x1(1)="Afghanistan"
  x1(2)="Albania"
  FOR i=1 TO DATACOUNT
    a1=READ$(i) & x1(i-1)
    SendMessage ghCombo1,%CB_ADDSTRING,i,STRPTR(a1)
  NEXT i
  DATA "1", "2", "3"
END FUNCTION
' 根据属性窗口大小，调整属性窗口中的控件
FUNCTION SizePropContents(BYVAL hWnd AS DWORD,BYVAL wParam AS DWORD)AS LONG
  LOCAL rc AS RECT
  LOCAL i AS LONG
  LOCAL gdi1 AS DockInfo
  RESET gdi1
  'i=FindGdiIndex(%ID_PROPERTY)
  i=GetWindowLong(hWnd,%GWL_USERDATA)
  IF i=-1 THEN
    EXIT FUNCTION
  END IF
  gdi1=gdi(i)
  IF gdi1.id<=0 THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  GetClientRect hWnd,rc
'  IF gdi1.l="F" and gdi1.hWndD=hWnd THEN
'    if gdi1.hWndF>0 and IsWindow(gdi1.hWndF) then
'      GetClientRect gdi1.hWndF,rc
'    end if
''    'msgbox "test"
''    SetRect rc,0,0,gdi1.frc.nRight-14,gdi1.frc.nBottom-14
'  END IF
  IF (gdi1.ml="B" AND gdi1.byside=(gdi1.byside OR %BYTOP)) OR gdi1.l="B" THEN
    rc.nTop+=5
    rc.nBottom-=5
  END IF
  SetRect rc,rc.nLeft+4,rc.nTop+23,rc.nRight-8,rc.nBottom-28
  'canLog=1 : RunLog "ghCombo1 nRight=" & str$(rc.nRight)
  MoveWindow ghCombo1,rc.nLeft,rc.nTop,rc.nRight,100,%TRUE
  MoveWindow ghList,rc.nLeft,rc.nTop+25,rc.nRight,rc.nBottom-80,%TRUE
  MoveWindow ghTip,rc.nLeft,rc.nTop+rc.nBottom-55,rc.nRight,50,%TRUE '90-rc.nTop
'  GetWindowRect ghCombo1,rc
'  rc.nRight-=rc.nLeft
'  rc.nBottom-=rc.nTop
''  canLog=1 : RunLog "ghCombo1 rc=" & RcToStr(rc)
'  UpdateWindow hWnd
  IF rc.nBottom<50 THEN
    ShowWindow ghTip,%SW_HIDE
    ShowWindow ghCombo1,%SW_HIDE
    ShowWindow ghList,%SW_HIDE
  ELSE
    ShowWindow ghTip,%SW_SHOW
    ShowWindow ghCombo1,%SW_SHOW
    ShowWindow ghList,%SW_SHOW
  END IF
  IF wParam <> %SIZE_MINIMIZED THEN
    IF ghCombo  <> 0 THEN CALL ShowWindow( ghCombo,  %SW_HIDE )
    IF ghButton <> 0 THEN CALL ShowWindow( ghButton, %SW_HIDE )
    IF ghEdit   <> 0 THEN CALL ShowWindow( ghEdit,   %SW_HIDE )
  END IF
END FUNCTION
' 绘制属性列表
FUNCTION DrawPropList(BYVAL dis AS DWORD)AS LONG
  LOCAL lpDs      AS DRAWITEMSTRUCT PTR
  LOCAL rectFull  AS RECT
  LOCAL rc        AS RECT
  LOCAL rc2       AS RECT
  LOCAL nIndex    AS INTEGER
  LOCAL hBrush    AS LONG
  LOCAL OrgPen    AS LONG
  LOCAL Lb        AS LOGBRUSH
  LOCAL initClr   AS DWORD
  LOCAL currClr   AS ASCIIZ *255
  lpDs = dis
  IF @lpDs.itemID < 0 THEN 'OR @lpDs.itemID<>%IDC_LISTBOX THEN
    EXIT FUNCTION
  END IF
  rectFull = @lpDs.rcItem
  rc = rectFull
  IF nDivider = 0 THEN nDivider = 80 '(rc.nRight - rc.nLeft)  \ 2
  rc.nleft   = nDivider
  rc2        = rectFull
  rc2.nright = rc.nleft - 1
  nIndex     = @lpDs.itemID
  'IF  nIndex  > -1  THEN
  '画两个矩形，每个是一列
  hBrush=CreatePen(%PS_SOLID,1,RGB(192,192,192))
  OrgPen=SelectObject(@lpDs.hDC,hBrush)
  CALL MoveTo(@lpds.hDC ,0 ,rc2.nbottom )
  CALL LineTo(@lpDs.hDC ,rc.nright,rc2.nbottom)
  CALL MoveTo(@lpDs.hDC,rc2.nright,rc2.ntop)
  CALL LineTo(@lpDs.hDC,rc2.nright,rc2.nbottom)
  SelectObject @lpDs.hDC,OrgPen
  DeleteObject hBrush
  CALL SetBkMode( @lpDs.hDC , %TRANSPARENT )
  rc2.ntop    = rc2.ntop    + 1
  '在第一列靠左写上属性名，第二列靠左写上对应的属性值，检测是否处于选中状态，如果是，则高亮显示
  '并把选中的属性对应的提示显示在ghTip文本框中。
  SELECT CASE @lpDs.itemAction
    CASE %ODA_DRAWENTIRE ,%ODA_SELECT
      'CLEAR BACKGROUND
      'canLog=1 : RunLog("@lpDs.itemState=" & str$(@lpDs.itemState) & " %ODS_SELECTED=" & STR$(%ODS_SELECTED))
      IF @lpDs.itemState  <> 769 THEN      'AND %ODS_SELECTED)
        FillRect(@lpDs.hDC,rc2 , GetSysColorBrush(%COLOR_WINDOW))
        CALL SetBkColor(@lpds.hDC, GetSysColor(%COLOR_WINDOW))             'text background
        CALL SetTextColor(@lpds.hDC, GetSysColor(%COLOR_WINDOWTEXT))       'text color
      ELSE
        FillRect(@lpds.hDC, rc2, GetSysColorBrush(%COLOR_HIGHLIGHT)) 'clear background
        CALL SetBkColor(@lpds.hDC, GetSysColor(%COLOR_HIGHLIGHT))             'text background
        CALL SetTextColor(@lpds.hDC, GetSysColor(%COLOR_HIGHLIGHTTEXT))       'text color
        CALL SendMessage(ghTip,%WM_SETTEXT,1,VARPTR(pItem(nIndex).tip))
      END IF

      rc2.nleft   = rc2.nleft   + 3
      CALL DrawText(@lpDs.hDC , pItem(nIndex).propName, _
                      LEN(pItem(nIndex).propName), _
                      rc2, %DT_LEFT OR %DT_SINGLELINE )
  END SELECT
  rc.nleft   = rc.nleft   + 3
  rc.ntop    = rc.ntop    + 1
  IF pItem(nIndex).nItemType = %PIT_COMBO  THEN
    rc.nright=rc.nright-19
    FillRect(@lpDs.hDC,rc , GetSysColorBrush(%COLOR_WINDOW))
  END IF
  '如果属性为颜色，则用以下代码画出表示颜色的小方框。
  IF pItem(nIndex).nItemType = %PIT_COLOR THEN
    currClr = pItem(nIndex).curValue
    IF  currClr > "" THEN
      '分离属性值的当前颜色值
      currClr = LTRIM$(currClr,"&&")
      currClr="&" + currClr
      currClr = RTRIM$(currClr,"&&")
      initClr = VAL(currClr)
    ELSE
      initClr=VAL("&H8080FF")
    END IF
    rc.nTop=rc.nTop+1
    rc.nLeft=rc.nLeft-2
    rc.nRight=rc.nLeft+13
    rc.nBottom=rc.nBottom-1
    Lb.lbStyle = %BS_SOLID
    Lb.lbColor = initClr
    hBrush& = CreateBrushIndirect(lb)
    FillRect @lpds.hDC, rc, hBrush&
    DeleteObject hBrush&
    CALL FrameRect(@lpds.hDc,rc,GetStockObject(%BLACK_BRUSH))
    SetBkMode @lpds.hDC,%TRANSPARENT
    rc.nLeft=rc.nRight+1
    rc.nRight=rc.nLeft+100
  END IF
  '第二列中写出对应的属性值
  CALL SetTextColor(@lpds.hDC, GetSysColor(%COLOR_WINDOWTEXT))
  CALL DrawText(@lpDs.hDC , pItem(nIndex).curValue, _
                  LEN(pItem(nIndex).curValue), _
                  rc, %DT_LEFT OR %DT_SINGLELINE )
END FUNCTION
' 属性命令执行，控件基本消息响应
FUNCTION Command_Prop(BYVAL wParam AS DWORD)AS LONG
  SELECT CASE LOWRD(wParam)
    CASE %IDC_LISTBOX
      SELECT CASE HIWRD(wParam)
        CASE %LBN_SELCHANGE  '选择有改变
          CALL OnSelChange()
      END SELECT
  END SELECT
END FUNCTION
'-------------------------------------------------------------------------
' 属性列表处理过程子类
'-------------------------------------------------------------------------
FUNCTION ListBoxProc(   BYVAL hWnd AS LONG, BYVAL wMsg AS LONG, _
                        BYVAL wParam AS LONG, BYVAL lParam AS LONG) AS LONG
  LOCAL pt        AS POINTAPI
  LOCAL hBrush    AS LONG
  LOCAL idx&,tLen&
  SELECT CASE wMsg
    '捕捉Combo,Edit,Button子控件的消息响应
    CASE %WM_COMMAND
      SELECT CASE LOWRD(wParam)
        CASE %IDC_COMBO
          SELECT CASE HIWRD(wParam)
            CASE %CBN_SELCHANGE
              LOCAL selStr AS ASCIIZ*255
              IF ghCombo <> 0  THEN
                idx& = SendMessage( ghCombo, %CB_GETCURSEL ,0 ,0 )
                CALL SendMessage( ghCombo, %CB_GETLBTEXT, idx& , VARPTR(selStr) )
                pItem(curSel).curValue = selStr
              END IF
          END SELECT
        CASE %IDC_EDITS
          SELECT CASE HIWRD(wParam)
            CASE %EN_CHANGE
              LOCAL newStr AS ASCIIZ*255
              tLen&  = SendMessage(ghEdit, %WM_GETTEXTLENGTH, 0, 0) + 1
              newStr = SPACE$(tLen&)
              CALL SendMessage(ghEdit, %WM_GETTEXT, tLen&, VARPTR(newStr))
              pItem(curSel).curValue = newStr
          END SELECT
        CASE %IDC_BUTTON
          SELECT CASE HIWRD(wParam)
            CASE %BN_CLICKED
              CALL OnButtonClicked()
          END SELECT
      END SELECT
    CASE %WM_MOUSEWHEEL
      IF ghCombo  <> 0 THEN CALL ShowWindow( ghCombo,  %SW_HIDE )
      IF ghButton <> 0 THEN CALL ShowWindow( ghButton, %SW_HIDE )
      IF ghEdit   <> 0 THEN CALL ShowWindow( ghEdit,   %SW_HIDE )
      CALL InvalidateRect(hWnd ,BYVAL %NULL, 1)
      CALL UpdateWindow(hWnd)
      'CALL UpdateWindow(ghProp)
    CASE %WM_VSCROLL
      IF ghCombo  <> 0 THEN CALL ShowWindow( ghCombo,  %SW_HIDE )
      IF ghButton <> 0 THEN CALL ShowWindow( ghButton, %SW_HIDE )
      IF ghEdit   <> 0 THEN CALL ShowWindow( ghEdit,   %SW_HIDE )
      CALL InvalidateRect(hWnd ,BYVAL %NULL, 1)
      CALL UpdateWindow(hWnd)
      'CALL UpdateWindow(ghProp)
    CASE %WM_LBUTTONDOWN '鼠标左键按下消息
      IF ghCombo  <> 0 THEN CALL ShowWindow( ghCombo,  %SW_HIDE )
      IF ghEdit   <> 0 THEN CALL ShowWindow( ghEdit,   %SW_HIDE )
      IF ghButton <> 0 THEN CALL ShowWindow( ghButton, %SW_HIDE )
      '-----------------------[ 分隔 ]-----------------------
      pt.x = LOWRD(lParam)
      pt.y = HIWRD(lParam)
      IF ((pt.x >= nDivider-5) AND (pt.x <= nDivider+5)) THEN
        '如果鼠标在分隔条上，则开始调节
        CALL SetCursor(hHsplit)
        LOCAL rc AS RECT
        CALL GetWindowRect( hWnd ,rc )
        rc.nleft  = rc.nleft  +10
        rc.nright = rc.nright -20
        '限制鼠标移动范围在列表框内
        CALL ClipCursor( rc )
        CALL GetClientRect( hWnd ,rc )
        bTracking = %TRUE
        nDivTop   = rc.ntop
        nDivBtm   = rc.nbottom
        IF nDivBtm>17*15 THEN
          nDivBtm=17*15
        END IF
        nOldDivX  = pt.x
        CALL InvertLine( nOldDivX ,nDivTop ,nOldDivX ,nDivBtm )
        '捕捉鼠标
        CALL SetCapture(hWnd)
      END IF
    CASE %WM_MOUSEMOVE
      pt.x = LOWRD(lParam)
      pt.y = HIWRD(lParam)
      '移动分隔线到鼠标位置
      '调节列大小
      IF bTracking = %TRUE  THEN
        '删除旧的分隔线
        CALL InvertLine( nOldDivX ,nDivTop ,nOldDivX ,nDivBtm )
        '画新分隔线
        CALL InvertLine( pt.x ,nDivTop ,pt.x ,nDivBtm )
        nOldDivX = pt.x
      ELSE
        '如果鼠标在行分隔线上，设置鼠标为可调节样式
        IF (pt.x >= nDivider-5) AND (pt.x <= nDivider+5) THEN
          CALL SetCursor(hHsplit)
        END IF
      END IF
    CASE %WM_LBUTTONUP
      pt.x = LOWRD(lParam)
      pt.y = HIWRD(lParam)
      '如果列被调整了,鼠标弹起后指示调节完成.需要重绘列以显示他们的新的宽度
      IF bTracking = %TRUE THEN
        bTracking = %FALSE
        CALL ReleaseCapture()
        CALL ClipCursor( BYVAL %NULL )
        CALL InvertLine( pt.x ,nDivTop ,pt.x ,nDivBtm )
        '设置调节线位置为新值
        nDivider = pt.x
        '刷新属性列表控件
        CALL InvalidateRect(hWnd ,BYVAL %NULL, 1)
        CALL UpdateWindow(hWnd)
      END IF
    CASE %WM_CAPTURECHANGED
      IF lParam <> ghList THEN
        IF bTracking = %TRUE THEN
          bTracking = %FALSE
          CALL ReleaseCapture()
          CALL ClipCursor( BYVAL %NULL )
          CALL InvalidateRect(ghList ,BYVAL %NULL, 1)
          CALL UpdateWindow(ghList)
        END IF
      END IF
  END SELECT
  FUNCTION = CallWindowProc( BYVAL lpOldListProc,hWnd,wMsg,wParam,lParam )
END FUNCTION
' 属性窗口销毁事件
FUNCTION DestroyPropertyWin() AS LONG
  IF lpOldListProc <> 0 THEN CALL SetWindowLong( ghList,%GWL_WNDPROC,lpOldListProc )
END FUNCTION
' 移动属性窗口控件到句柄指定窗口：在泊坞和浮动窗口之间切换
FUNCTION MovePropContents(BYVAL hWnd AS DWORD)AS LONG
  SetParent ghCombo1, hWnd
  SetParent ghList, hWnd
  SetParent ghTip,hWnd
  'SizePropContents hWnd,0
END FUNCTION
' *********************************************************************************************
'----------------------------------------------------------------------------
' 下拉控件选择内容有变化
'----------------------------------------------------------------------------
SUB OnSelChange()
  LOCAL rc AS RECT
  LOCAL lBoxSelText AS STRING
  LOCAL sel AS INTEGER
  LOCAL hBrush AS LONG
  LOCAL rc3 AS RECT
  LOCAL hDC AS LONG
  LOCAL cmbItems$,delimiter$,nCount&,i&,j&,dataItem$
  LOCAL nWidth&
  CALL UpdateWindow ( ghList )
  curSel = SendMessage( ghList, %LB_GETCURSEL, 0, 0 )
  CALL SendMessage( ghList, %LB_GETITEMRECT, curSel,VARPTR(rc) )
  rc.nleft = nDivider
  IF ghCombo  <> 0 THEN CALL ShowWindow( ghCombo,  %SW_HIDE )
  IF ghButton <> 0 THEN CALL ShowWindow( ghButton, %SW_HIDE )
  IF ghEdit   <> 0 THEN CALL ShowWindow( ghEdit,   %SW_HIDE )
  IF pItem(curSel).nItemType = %PIT_COMBO  THEN
    '显示下拉列表框,并移动它到新的位置
    nLastBox = 0
    IF ghCombo <> 0 THEN
      CALL MoveWindow( ghCombo,rc.nleft-1,rc.ntop+1,rc.nright+2-rc.nleft,rc.nbottom+100,1 )
    END IF
    '为这个属性增加选择项
    cmbItems$   = pItem(curSel).cmbItems
    lBoxSelText = pItem(curSel).curValue
    CALL SendMessage( ghCombo,%CB_RESETCONTENT, 0,0 )
    delimiter$ = "|"
    nCount& = TALLY(cmbItems$,delimiter$)
    FOR i& = 1 TO nCount&
      dataItem$ = PARSE$(cmbItems$,delimiter$,i&)
      '把字符串增加到ComboBox
      CALL SendMessage(ghCombo, %CB_ADDSTRING, i&-1, STRPTR(dataItem$) )
    NEXT
    '跳到下拉框里的属性当前值
    j& = SendMessage(ghCombo,%CB_FINDSTRINGEXACT,0,STRPTR(lBoxSelText) )
    IF j& <> %CB_ERR  THEN
      CALL SendMessage( ghCombo,%CB_SETCURSEL,j&,0 )
    ELSE
      '如果没有当前值,设置列表中的第一项为默认值
      CALL SendMessage( ghCombo,%CB_SETCURSEL,0,0 )
    END IF
    CALL ShowWindow( ghCombo,%SW_SHOW )
    CALL setFocus(ghCombo)
  ELSEIF pItem(curSel).nItemType = %PIT_EDIT  THEN
    '显示编辑框
    nLastBox   = 1
    prevSel    = curSel
    rc.nbottom = rc.nbottom - 3
    IF ghEdit <> 0 THEN
      CALL MoveWindow( ghEdit,rc.nleft+2,rc.ntop+3,rc.nright-rc.nleft,rc.nbottom-rc.ntop,1 )
    END IF
    lBoxSelText = pItem(curSel).curValue
    CALL ShowWindow( ghEdit, %SW_SHOW )
    CALL SetFocus( ghEdit )
    CALL SendMessage(ghEdit, %WM_SETTEXT, 0, STRPTR(lBoxSelText) )
    CALL SendMessage(ghEdit,%EM_SETSEL,0,255)
    '设置编辑框中的文本为属性当前值
  ELSE
    '当属性为颜色/文件/字体选择时显示按钮
    nLastBox = 2
    prevSel  = curSel
    nWidth& = rc.nright - rc.nleft
    IF  nWidth& > 25  THEN
      rc.nleft   = rc.nright  - 25
    END IF
    IF ghButton <> 0  THEN
        CALL MoveWindow( ghButton,rc.nRight-22,rc.ntop+1,22, 20, 1 )  'rc.nRight -rc.nleft
    END IF
    CALL ShowWindow( ghButton, %SW_SHOW )
  END IF
  IF oldSel=0 AND curSel=0 THEN
    oldSel=1
  END IF
  oldSel=curSel
  CALL UpdateWindow ( ghList )
  'CALL UpdateWindow(ghProp)
END SUB
'----------------------------------------------------------------------------
' 属性列表框中选择字体，颜色，文件按钮按下事件
'----------------------------------------------------------------------------
SUB OnButtonClicked()
  LOCAL initClr        AS DWORD
  LOCAL currClr        AS ASCIIZ*255
  LOCAL clrStr         AS ASCIIZ*255
  LOCAL ColorSpec      AS CHOOSECOLORAPI
  LOCAL lResult        AS LONG
  LOCAL lCounter       AS LONG
  LOCAL lCustomColor() AS LONG
  LOCAL rc             AS RECT
  LOCAL currPath$,faceName$,fName$,tmp&
  '根据类型显示相应的通用对话框
  '根据相关属性显示相应的普通对话框。
  '为选择颜色对话框进行检查
  IF pItem(curSel).nItemType = %PIT_COLOR  THEN
    currClr = pItem(curSel).curValue
    IF  currClr > "" THEN
      '分析属性的当前颜色值
      'd$  = ","
      currClr = LTRIM$(currClr,"&&")
      currClr="&" + currClr
      currClr = RTRIM$(currClr,"&&")
      initClr = VAL(currClr)
    ELSE
      '设置默认值
      initClr = VAL("&H8080FF")'RGB( 255,128,128 )
    END IF
    ColorSpec.lStructSize  = LEN(ColorSpec)
    ColorSpec.hwndOwner    = ghList
    ColorSpec.lpCustColors = VARPTR(cColor(0))
    ColorSpec.rgbResult    = initClr
    ColorSpec.Flags        = ColorSpec.Flags OR %CC_RGBINIT
    lResult = ChooseColor(ColorSpec)
    IF lResult = 0 THEN '检查用户是否取消了对话框
      EXIT SUB
    ELSE
      'selClr = ColorSpec.rgbResult
      clrStr = HEX$(ColorSpec.rgbResult,7)
      clrStr="&&H" + clrStr
      clrStr=clrStr+"&&"
      pItem(curSel).curValue = clrStr
      CALL InvalidateRect( ghList,BYVAL %NULL,1)
      CALL UpdateWindow ( ghList )
    END IF
  '下一步检查打开文件对话框
  ELSEIF  pItem(curSel).nItemType = %PIT_FILE THEN
    LOCAL SelectedFile AS ASCIIZ*255
    LOCAL zTxt  AS ASCIIZ * 255
    zTxt = "所有图片(*.bmp,*.ico)|*.BMP;*.ICO|"
    zTxt = zTxt & "位图(*.bmp)|*.BMP|"
    ZTxt = zTxt & "图标(*.ico)|*.ICO|"
    currPath$ = pItem(curSel).curValue
    IF currPath$ = "none" THEN
      fName$ = ""
    ELSE
      fName$ = currPath$
    END IF
    '使用简单的打开文件对话框
    tmp& = OpenFileDialog(GetParent(ghList),"选择文件:",fName$,CURDIR$,zTxt,"BMP", _   'ghProp
                            %OFN_FILEMUSTEXIST OR %OFN_HIDEREADONLY OR %OFN_LONGNAMES)
    IF tmp& THEN
      pItem(curSel).curValue = fName$
      CALL InvalidateRect( ghList,BYVAL %NULL,1)
      CALL UpdateWindow ( ghList )
    ELSE
      pItem(curSel).curValue = "NONE"
      CALL InvalidateRect( ghList,BYVAL %NULL,1)
      CALL UpdateWindow ( ghList )
    END IF
  '最后,检查选择字体对话框
  ELSEIF pItem(curSel).nItemType = %PIT_FONT  THEN
    DIM cf AS CHOOSEFONTAPI
    DIM lfFont AS LOGFONT
    faceName$ = pItem(curSel).curValue
    IF faceName$= "none" OR faceName$="NONE" OR faceName$ = "" THEN
      pItem(curSel).curValue="Arial"
    END IF
    lfFont.lfFaceName=pItem(curSel).curValue
    cf.iPointSize=100
    cf.lStructSize = SIZEOF(cf)
    cf.hWndOwner   = %NULL'ghProp
    cf.lpLogFont   = VARPTR(lfFont)
    cf.Flags       = %CF_EFFECTS OR %CF_BOTH OR %CF_INITTOLOGFONTSTRUCT
    lResult = ChooseFont(cf)
    IF lResult = 0 THEN '检查用户是否取消取对话框
      EXIT SUB
    ELSE
      faceName$   = lfFont.lfFaceName
      IF faceName$ = "" THEN
        EXIT SUB
      END IF
      '这里使用格式.
      pItem(curSel).curValue = faceName$  '& "-" & rv$ & gv$ & bv$ & "-" & TRIM$(STR$(cf.iPointSize/10)) & "]"
      CALL InvalidateRect( ghList,BYVAL %NULL,1)
      CALL UpdateWindow ( ghList )
    END IF
  END IF
END SUB
'--------------------------------------------------------------------------
' 画属性列表框中的调整指示分隔线
'--------------------------------------------------------------------------
SUB InvertLine( BYVAL xleft1 AS INTEGER ,BYVAL yFrom AS INTEGER, _
                BYVAL xleft2 AS INTEGER ,BYVAL yTo AS INTEGER )
  LOCAL hDC       AS LONG
  LOCAL nDrawMode AS LONG
  LOCAL oldpt     AS POINTAPI
  '*取得属性列表的DC
  hDC = GetDC(ghList)
  '*使用默认的Pen颜色及类型
  '*设置GDI前景为混合模式
  nDrawMode = SetROP2(hDC, %R2_NOT)
  '*画线
  CALL MoveToEx( hDC ,xLeft1 ,yFrom ,oldpt)
  CALL LineTo( hDC ,xLeft2 ,yTo)
  '*重载DC
  CALL SetROP2(hDC, nDrawMode)
  CALL ReleaseDC(ghList, hDC)
END SUB
'--------------------------------------------------------------------------
' 初始属性列表
'--------------------------------------------------------------------------
SUB InitializePropertyItems()
  pItem(0).propName  =  "工具提示文本"
  pItem(0).curValue  =  "Litte Red Riding Hood..."
  pItem(0).nItemType =  %PIT_EDIT
  pItem(0).cmbItems  =  ""
  pItem(0).tip       =  "这个是工具提示文本，显示提示"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(0).propName))

  pItem(1).propName  =  "Enabled"
  pItem(1).curValue  =  "True"
  pItem(1).nItemType =  %PIT_COMBO
  pItem(1).cmbItems  =  "True|False|"
  pItem(1).tip       =  "显示为是否可用"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(1).propName))

  pItem(2).propName  =  "Visible"
  pItem(2).curValue  =  "True"
  pItem(2).nItemType =  %PIT_COMBO
  pItem(2).cmbItems  =  "True|False|"
  pItem(2).tip       =  "指示是否可见"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(2).propName))

  pItem(3).propName  =  "ForeColor"
  pItem(3).curValue  =  "&&H000FFFF&&"
  pItem(3).nItemType =  %PIT_COLOR
  pItem(3).cmbItems  =  ""
  pItem(3).tip       =  "设置前景色"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(3).propName))

  pItem(4).propName  =  "BackColor"
  pItem(4).curValue  =  "&&H08000FF&&"
  pItem(4).nItemType =  %PIT_COLOR
  pItem(4).cmbItems  =  ""
  pItem(4).tip       =  "设置背景色"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(4).propName))

  pItem(5).propName  =  "Opaque"
  pItem(5).curValue  =  "False"
  pItem(5).nItemType =  %PIT_COMBO
  pItem(5).cmbItems  =  "True|False|"
  pItem(5).tip       =  "设置是否以透明状态显示"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(5).propName))

  pItem(6).propName  =  "AutoScroll"
  pItem(6).curValue  =  "True"
  pItem(6).nItemType =  %PIT_COMBO
  pItem(6).cmbItems  =  "True|False|"
  pItem(6).tip       =  "设置是否自动滚动"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(6).propName))

 pItem(7).propName  =  "Double Buffered"
 pItem(7).curValue  =  "True"
 pItem(7).nItemType =  %PIT_COMBO
 pItem(7).cmbItems  =  "True|False|"
 pItem(7).tip       =  "这个意思是设置双重缓冲区，我也不知道是什么意思"
 CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(7).propName))

 pItem(8).propName  =  "Font"
 pItem(8).curValue  =  "Arial"
 pItem(8).nItemType =  %PIT_FONT
 pItem(8).cmbItems  =  ""
 pItem(8).tip       =  "没什么说的，这就是字体设置"
 CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(8).propName))

  pItem(9).propName  = "Text"
  pItem(9).curValue  =  "Big Bad Wolf!"
  pItem(9).nItemType =  %PIT_EDIT
  pItem(9).cmbItems  =  ""
  pItem(9).tip       =  "这当然就是设置文本了"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(9).propName))

  pItem(10).propName  =  "HorizAlign"
  pItem(10).curValue  =  "CENTER"
  pItem(10).nItemType =  %PIT_COMBO
  pItem(10).cmbItems  =  "CENTER|LEFT|RIGHT|"
  pItem(10).tip       =  "这里设置水平对齐方式"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(10).propName))

  pItem(11).propName  =  "VertAlign"
  pItem(11).curValue  =  "CENTER"
  pItem(11).nItemType =  %PIT_COMBO
  pItem(11).cmbItems  =  "CENTER|TOP|BOTTOM|"
  pItem(11).tip       =  "这里设置垂直对齐方式"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(11).propName))

  pItem(12).propName  =  "Bmp ,Icon"
  pItem(12).curValue  =  "NONE"
  pItem(12).nItemType =  %PIT_FILE
  pItem(12).cmbItems  =  ""
  pItem(12).tip       =  "选择位图或图标文件"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(12).propName))

  pItem(13).propName  =  "Border Painted"
  pItem(13).curValue  =  "False"
  pItem(13).nItemType =  %PIT_COMBO
  pItem(13).cmbItems  =  "True|False|"
  pItem(13).tip       =  "决定是否显示一个黑边"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(13).propName))

  pItem(14).propName  =  "Fill Content Area"
  pItem(14).curValue  =  "True"
  pItem(14).nItemType =  %PIT_COMBO
  pItem(14).cmbItems  =  "True|False|"
  pItem(14).tip       =  "填充内容区域，表面意思是这样的，具体指什么，我也不清楚"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(14).propName))

  pItem(15).propName  = "Focus Painted"
  pItem(15).curValue  =  "True"
  pItem(15).nItemType =  %PIT_COMBO
  pItem(15).cmbItems  =  "True|False|"
  pItem(15).tip       =  "确定是否画出一个聚焦框，就是比控件稍小点的虚线框"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(15).propName))

  pItem(16).propName  =  "Horiz. Text Pos."
  pItem(16).curValue  =  "RIGHT"
  pItem(16).nItemType =  %PIT_COMBO
  pItem(16).cmbItems  =  "RIGHT|LEFT|CENTER|LEADING|TRAILING|"
  pItem(16).tip       =  "这个用来确定文本的水平方向位置"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(16).propName))
END SUB                                                                                                                                                                                        _
'--------------------------------------------------------------------------
' 为选择颜色对话框加载16个自定义颜色
'--------------------------------------------------------------------------
SUB InitializeCustomColors()
  LOCAL lCounter AS LONG
  DIM cColor(0:30)
  '加载自定义的颜色
  RANDOMIZE TIMER
  FOR lCounter = 0 TO 15
    '示例颜色
    cColor(lCounter) = RND(0,16777215) 'GetSysColor(lCounter)
  NEXT
END SUB
