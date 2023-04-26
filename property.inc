''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' �ļ� : property.inc
' ���Դ��ڣ�Ϊ���Ӵ��ڱ༭�������ÿؼ�����
'
TYPE PROPERTYITEM        '�����Զ�������
  forObj      AS ASCIIZ*255 'ȷ���������ֶ����細�ڣ���ť���ı���ȣ�ֵ��"|"�ָ�����ȫ��ʱֵΪ"*"���硰���ơ�����
  propName    AS ASCIIZ*225 '������
  curValue    AS ASCIIZ*255 '��ǰֵ
  nItemType   AS INTEGER    '������
  cmbItems    AS ASCIIZ*255 '������
  tip         AS ASCIIZ*255 '��ʾ
END TYPE
TYPE PROPERTYSET         '�������б������õ�����ֵ��ʵ�ʱ༭�����ж�Ӧ�Ĳ���������ϵ����̬�޸�
  forObje       AS ASCIIZ*255'ȷ�����ں��ֶ���
  propName      AS ASCIIZ*255'ȷ��������
  curValue(10)  AS ASCIIZ*255'����ֵ
  Para(10)      AS ASCIIZ*255 '���õĳ����еĶ�Ӧ����
END TYPE
GLOBAL pItem()        AS PROPERTYITEM '�û���Ƶ��Զ���ؼ���Ϣ
GLOBAL ghCombo1       AS DWORD   '���Դ�����������������б��ؼ����
GLOBAL ghList         AS DWORD   '�����б��ؼ����
GLOBAL ghTip          AS LONG         '���Դ����е�������Ϣ��ʾ�ı�����
GLOBAL nDivider       AS INTEGER      '�����б��е��зָ���λ��
GLOBAL bDivIsSet      AS INTEGER
GLOBAL lpOldListProc  AS LONG    '���ڰ������б��Ĵ������̶����صľ��
GLOBAL curSel         AS INTEGER '�����б��е�ǰѡ����
GLOBAL oldSel         AS INTEGER '�����б����ϴ�ѡ����
'GLOBAL ghFont   AS LONG          '�趨�����б����ֵ�����ľ��
GLOBAL ghButton       AS LONG          '�����б��еİ�ť���
GLOBAL ghCombo        AS LONG          '�����б��е������б����
GLOBAL ghEdit         AS LONG          '�����б��е��ı�����
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
'PIT = ������Ŀ���ͣ���ťΪĬ��
%PIT_COMBO   = 0
%PIT_EDIT    = 1
%PIT_COLOR   = 2
%PIT_FONT    = 3
%PIT_FILE    = 4
' �������Դ���
FUNCTION CreatePropertyWin(BYVAL gdiI AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  '�趨���봰�ڵ�Ĭ��״̬���ini�ļ��õ�
  REDIM pItem(16)
  nDivider  = 0
  i = gdiI
  gdi(i).ID=%ID_PROPERTY
  '��鲢��ʼ������С����
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
  ' �������Բ��봰��
  gdi(i).hWndD = CreateWindowEx(  0,"DOCKIT", _
                              "����", _
                              %WS_CHILD, _
                              gdi(i).drc.nLeft, gdi(i).drc.nTop, _
                              gdi(i).drc.nRight,gdi(i).drc.nBottom, _   ' width,height
                              g_hWndMain, %ID_PROPERTY, g_hInst, BYVAL %NULL )
  SetWindowLong gdi(i).hWndD,%GWL_USERDATA,i
  ' �������Ը�������
  'rc=gdi(i).frc
  'ScreenToClientRect g_hWndMain,rc
  'msgbox "frc=" & RcToStr(gdi(i).frc) & " rc=" & RcToStr(rc)
  gdi(i).hWndF = CreateWindowEx(0,"DOCKIT", _
                              "����", _
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
' �������Դ����еĿؼ�
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
  '�������б�Ӧ�����࣬�����䴦������
  lpOldListProc = SetWindowLong( ghList,%GWL_WNDPROC,BYVAL CODEPTR(ListBoxProc) )
  CALL SendMessage( ghList, %WM_SETFONT, ghFont, 0 )
  '���������б��еĿؼ���
  '�������Ƕ������صģ���һ���ͻ����ô�С
  ghButton = CreateWindowEx(  0,"BUTTON","��" , %WS_CHILD, _
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
  ' �����CALL���ñ����������б�����֮��
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
' �������Դ��ڴ�С���������Դ����еĿؼ�
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
' ���������б�
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
  '���������Σ�ÿ����һ��
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
  '�ڵ�һ�п���д�����������ڶ��п���д�϶�Ӧ������ֵ������Ƿ���ѡ��״̬������ǣ��������ʾ
  '����ѡ�е����Զ�Ӧ����ʾ��ʾ��ghTip�ı����С�
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
  '�������Ϊ��ɫ���������´��뻭����ʾ��ɫ��С����
  IF pItem(nIndex).nItemType = %PIT_COLOR THEN
    currClr = pItem(nIndex).curValue
    IF  currClr > "" THEN
      '��������ֵ�ĵ�ǰ��ɫֵ
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
  '�ڶ�����д����Ӧ������ֵ
  CALL SetTextColor(@lpds.hDC, GetSysColor(%COLOR_WINDOWTEXT))
  CALL DrawText(@lpDs.hDC , pItem(nIndex).curValue, _
                  LEN(pItem(nIndex).curValue), _
                  rc, %DT_LEFT OR %DT_SINGLELINE )
END FUNCTION
' ��������ִ�У��ؼ�������Ϣ��Ӧ
FUNCTION Command_Prop(BYVAL wParam AS DWORD)AS LONG
  SELECT CASE LOWRD(wParam)
    CASE %IDC_LISTBOX
      SELECT CASE HIWRD(wParam)
        CASE %LBN_SELCHANGE  'ѡ���иı�
          CALL OnSelChange()
      END SELECT
  END SELECT
END FUNCTION
'-------------------------------------------------------------------------
' �����б�������������
'-------------------------------------------------------------------------
FUNCTION ListBoxProc(   BYVAL hWnd AS LONG, BYVAL wMsg AS LONG, _
                        BYVAL wParam AS LONG, BYVAL lParam AS LONG) AS LONG
  LOCAL pt        AS POINTAPI
  LOCAL hBrush    AS LONG
  LOCAL idx&,tLen&
  SELECT CASE wMsg
    '��׽Combo,Edit,Button�ӿؼ�����Ϣ��Ӧ
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
    CASE %WM_LBUTTONDOWN '������������Ϣ
      IF ghCombo  <> 0 THEN CALL ShowWindow( ghCombo,  %SW_HIDE )
      IF ghEdit   <> 0 THEN CALL ShowWindow( ghEdit,   %SW_HIDE )
      IF ghButton <> 0 THEN CALL ShowWindow( ghButton, %SW_HIDE )
      '-----------------------[ �ָ� ]-----------------------
      pt.x = LOWRD(lParam)
      pt.y = HIWRD(lParam)
      IF ((pt.x >= nDivider-5) AND (pt.x <= nDivider+5)) THEN
        '�������ڷָ����ϣ���ʼ����
        CALL SetCursor(hHsplit)
        LOCAL rc AS RECT
        CALL GetWindowRect( hWnd ,rc )
        rc.nleft  = rc.nleft  +10
        rc.nright = rc.nright -20
        '��������ƶ���Χ���б�����
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
        '��׽���
        CALL SetCapture(hWnd)
      END IF
    CASE %WM_MOUSEMOVE
      pt.x = LOWRD(lParam)
      pt.y = HIWRD(lParam)
      '�ƶ��ָ��ߵ����λ��
      '�����д�С
      IF bTracking = %TRUE  THEN
        'ɾ���ɵķָ���
        CALL InvertLine( nOldDivX ,nDivTop ,nOldDivX ,nDivBtm )
        '���·ָ���
        CALL InvertLine( pt.x ,nDivTop ,pt.x ,nDivBtm )
        nOldDivX = pt.x
      ELSE
        '���������зָ����ϣ��������Ϊ�ɵ�����ʽ
        IF (pt.x >= nDivider-5) AND (pt.x <= nDivider+5) THEN
          CALL SetCursor(hHsplit)
        END IF
      END IF
    CASE %WM_LBUTTONUP
      pt.x = LOWRD(lParam)
      pt.y = HIWRD(lParam)
      '����б�������,��굯���ָʾ�������.��Ҫ�ػ�������ʾ���ǵ��µĿ���
      IF bTracking = %TRUE THEN
        bTracking = %FALSE
        CALL ReleaseCapture()
        CALL ClipCursor( BYVAL %NULL )
        CALL InvertLine( pt.x ,nDivTop ,pt.x ,nDivBtm )
        '���õ�����λ��Ϊ��ֵ
        nDivider = pt.x
        'ˢ�������б��ؼ�
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
' ���Դ��������¼�
FUNCTION DestroyPropertyWin() AS LONG
  IF lpOldListProc <> 0 THEN CALL SetWindowLong( ghList,%GWL_WNDPROC,lpOldListProc )
END FUNCTION
' �ƶ����Դ��ڿؼ������ָ�����ڣ��ڲ���͸�������֮���л�
FUNCTION MovePropContents(BYVAL hWnd AS DWORD)AS LONG
  SetParent ghCombo1, hWnd
  SetParent ghList, hWnd
  SetParent ghTip,hWnd
  'SizePropContents hWnd,0
END FUNCTION
' *********************************************************************************************
'----------------------------------------------------------------------------
' �����ؼ�ѡ�������б仯
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
    '��ʾ�����б���,���ƶ������µ�λ��
    nLastBox = 0
    IF ghCombo <> 0 THEN
      CALL MoveWindow( ghCombo,rc.nleft-1,rc.ntop+1,rc.nright+2-rc.nleft,rc.nbottom+100,1 )
    END IF
    'Ϊ�����������ѡ����
    cmbItems$   = pItem(curSel).cmbItems
    lBoxSelText = pItem(curSel).curValue
    CALL SendMessage( ghCombo,%CB_RESETCONTENT, 0,0 )
    delimiter$ = "|"
    nCount& = TALLY(cmbItems$,delimiter$)
    FOR i& = 1 TO nCount&
      dataItem$ = PARSE$(cmbItems$,delimiter$,i&)
      '���ַ������ӵ�ComboBox
      CALL SendMessage(ghCombo, %CB_ADDSTRING, i&-1, STRPTR(dataItem$) )
    NEXT
    '����������������Ե�ǰֵ
    j& = SendMessage(ghCombo,%CB_FINDSTRINGEXACT,0,STRPTR(lBoxSelText) )
    IF j& <> %CB_ERR  THEN
      CALL SendMessage( ghCombo,%CB_SETCURSEL,j&,0 )
    ELSE
      '���û�е�ǰֵ,�����б��еĵ�һ��ΪĬ��ֵ
      CALL SendMessage( ghCombo,%CB_SETCURSEL,0,0 )
    END IF
    CALL ShowWindow( ghCombo,%SW_SHOW )
    CALL setFocus(ghCombo)
  ELSEIF pItem(curSel).nItemType = %PIT_EDIT  THEN
    '��ʾ�༭��
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
    '���ñ༭���е��ı�Ϊ���Ե�ǰֵ
  ELSE
    '������Ϊ��ɫ/�ļ�/����ѡ��ʱ��ʾ��ť
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
' �����б�����ѡ�����壬��ɫ���ļ���ť�����¼�
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
  '����������ʾ��Ӧ��ͨ�öԻ���
  '�������������ʾ��Ӧ����ͨ�Ի���
  'Ϊѡ����ɫ�Ի�����м��
  IF pItem(curSel).nItemType = %PIT_COLOR  THEN
    currClr = pItem(curSel).curValue
    IF  currClr > "" THEN
      '�������Եĵ�ǰ��ɫֵ
      'd$  = ","
      currClr = LTRIM$(currClr,"&&")
      currClr="&" + currClr
      currClr = RTRIM$(currClr,"&&")
      initClr = VAL(currClr)
    ELSE
      '����Ĭ��ֵ
      initClr = VAL("&H8080FF")'RGB( 255,128,128 )
    END IF
    ColorSpec.lStructSize  = LEN(ColorSpec)
    ColorSpec.hwndOwner    = ghList
    ColorSpec.lpCustColors = VARPTR(cColor(0))
    ColorSpec.rgbResult    = initClr
    ColorSpec.Flags        = ColorSpec.Flags OR %CC_RGBINIT
    lResult = ChooseColor(ColorSpec)
    IF lResult = 0 THEN '����û��Ƿ�ȡ���˶Ի���
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
  '��һ�������ļ��Ի���
  ELSEIF  pItem(curSel).nItemType = %PIT_FILE THEN
    LOCAL SelectedFile AS ASCIIZ*255
    LOCAL zTxt  AS ASCIIZ * 255
    zTxt = "����ͼƬ(*.bmp,*.ico)|*.BMP;*.ICO|"
    zTxt = zTxt & "λͼ(*.bmp)|*.BMP|"
    ZTxt = zTxt & "ͼ��(*.ico)|*.ICO|"
    currPath$ = pItem(curSel).curValue
    IF currPath$ = "none" THEN
      fName$ = ""
    ELSE
      fName$ = currPath$
    END IF
    'ʹ�ü򵥵Ĵ��ļ��Ի���
    tmp& = OpenFileDialog(GetParent(ghList),"ѡ���ļ�:",fName$,CURDIR$,zTxt,"BMP", _   'ghProp
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
  '���,���ѡ������Ի���
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
    IF lResult = 0 THEN '����û��Ƿ�ȡ��ȡ�Ի���
      EXIT SUB
    ELSE
      faceName$   = lfFont.lfFaceName
      IF faceName$ = "" THEN
        EXIT SUB
      END IF
      '����ʹ�ø�ʽ.
      pItem(curSel).curValue = faceName$  '& "-" & rv$ & gv$ & bv$ & "-" & TRIM$(STR$(cf.iPointSize/10)) & "]"
      CALL InvalidateRect( ghList,BYVAL %NULL,1)
      CALL UpdateWindow ( ghList )
    END IF
  END IF
END SUB
'--------------------------------------------------------------------------
' �������б����еĵ���ָʾ�ָ���
'--------------------------------------------------------------------------
SUB InvertLine( BYVAL xleft1 AS INTEGER ,BYVAL yFrom AS INTEGER, _
                BYVAL xleft2 AS INTEGER ,BYVAL yTo AS INTEGER )
  LOCAL hDC       AS LONG
  LOCAL nDrawMode AS LONG
  LOCAL oldpt     AS POINTAPI
  '*ȡ�������б���DC
  hDC = GetDC(ghList)
  '*ʹ��Ĭ�ϵ�Pen��ɫ������
  '*����GDIǰ��Ϊ���ģʽ
  nDrawMode = SetROP2(hDC, %R2_NOT)
  '*����
  CALL MoveToEx( hDC ,xLeft1 ,yFrom ,oldpt)
  CALL LineTo( hDC ,xLeft2 ,yTo)
  '*����DC
  CALL SetROP2(hDC, nDrawMode)
  CALL ReleaseDC(ghList, hDC)
END SUB
'--------------------------------------------------------------------------
' ��ʼ�����б�
'--------------------------------------------------------------------------
SUB InitializePropertyItems()
  pItem(0).propName  =  "������ʾ�ı�"
  pItem(0).curValue  =  "Litte Red Riding Hood..."
  pItem(0).nItemType =  %PIT_EDIT
  pItem(0).cmbItems  =  ""
  pItem(0).tip       =  "����ǹ�����ʾ�ı�����ʾ��ʾ"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(0).propName))

  pItem(1).propName  =  "Enabled"
  pItem(1).curValue  =  "True"
  pItem(1).nItemType =  %PIT_COMBO
  pItem(1).cmbItems  =  "True|False|"
  pItem(1).tip       =  "��ʾΪ�Ƿ����"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(1).propName))

  pItem(2).propName  =  "Visible"
  pItem(2).curValue  =  "True"
  pItem(2).nItemType =  %PIT_COMBO
  pItem(2).cmbItems  =  "True|False|"
  pItem(2).tip       =  "ָʾ�Ƿ�ɼ�"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(2).propName))

  pItem(3).propName  =  "ForeColor"
  pItem(3).curValue  =  "&&H000FFFF&&"
  pItem(3).nItemType =  %PIT_COLOR
  pItem(3).cmbItems  =  ""
  pItem(3).tip       =  "����ǰ��ɫ"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(3).propName))

  pItem(4).propName  =  "BackColor"
  pItem(4).curValue  =  "&&H08000FF&&"
  pItem(4).nItemType =  %PIT_COLOR
  pItem(4).cmbItems  =  ""
  pItem(4).tip       =  "���ñ���ɫ"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(4).propName))

  pItem(5).propName  =  "Opaque"
  pItem(5).curValue  =  "False"
  pItem(5).nItemType =  %PIT_COMBO
  pItem(5).cmbItems  =  "True|False|"
  pItem(5).tip       =  "�����Ƿ���͸��״̬��ʾ"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(5).propName))

  pItem(6).propName  =  "AutoScroll"
  pItem(6).curValue  =  "True"
  pItem(6).nItemType =  %PIT_COMBO
  pItem(6).cmbItems  =  "True|False|"
  pItem(6).tip       =  "�����Ƿ��Զ�����"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(6).propName))

 pItem(7).propName  =  "Double Buffered"
 pItem(7).curValue  =  "True"
 pItem(7).nItemType =  %PIT_COMBO
 pItem(7).cmbItems  =  "True|False|"
 pItem(7).tip       =  "�����˼������˫�ػ���������Ҳ��֪����ʲô��˼"
 CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(7).propName))

 pItem(8).propName  =  "Font"
 pItem(8).curValue  =  "Arial"
 pItem(8).nItemType =  %PIT_FONT
 pItem(8).cmbItems  =  ""
 pItem(8).tip       =  "ûʲô˵�ģ��������������"
 CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(8).propName))

  pItem(9).propName  = "Text"
  pItem(9).curValue  =  "Big Bad Wolf!"
  pItem(9).nItemType =  %PIT_EDIT
  pItem(9).cmbItems  =  ""
  pItem(9).tip       =  "�⵱Ȼ���������ı���"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(9).propName))

  pItem(10).propName  =  "HorizAlign"
  pItem(10).curValue  =  "CENTER"
  pItem(10).nItemType =  %PIT_COMBO
  pItem(10).cmbItems  =  "CENTER|LEFT|RIGHT|"
  pItem(10).tip       =  "��������ˮƽ���뷽ʽ"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(10).propName))

  pItem(11).propName  =  "VertAlign"
  pItem(11).curValue  =  "CENTER"
  pItem(11).nItemType =  %PIT_COMBO
  pItem(11).cmbItems  =  "CENTER|TOP|BOTTOM|"
  pItem(11).tip       =  "�������ô�ֱ���뷽ʽ"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(11).propName))

  pItem(12).propName  =  "Bmp ,Icon"
  pItem(12).curValue  =  "NONE"
  pItem(12).nItemType =  %PIT_FILE
  pItem(12).cmbItems  =  ""
  pItem(12).tip       =  "ѡ��λͼ��ͼ���ļ�"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(12).propName))

  pItem(13).propName  =  "Border Painted"
  pItem(13).curValue  =  "False"
  pItem(13).nItemType =  %PIT_COMBO
  pItem(13).cmbItems  =  "True|False|"
  pItem(13).tip       =  "�����Ƿ���ʾһ���ڱ�"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(13).propName))

  pItem(14).propName  =  "Fill Content Area"
  pItem(14).curValue  =  "True"
  pItem(14).nItemType =  %PIT_COMBO
  pItem(14).cmbItems  =  "True|False|"
  pItem(14).tip       =  "����������򣬱�����˼�������ģ�����ָʲô����Ҳ�����"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(14).propName))

  pItem(15).propName  = "Focus Painted"
  pItem(15).curValue  =  "True"
  pItem(15).nItemType =  %PIT_COMBO
  pItem(15).cmbItems  =  "True|False|"
  pItem(15).tip       =  "ȷ���Ƿ񻭳�һ���۽��򣬾��Ǳȿؼ���С������߿�"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(15).propName))

  pItem(16).propName  =  "Horiz. Text Pos."
  pItem(16).curValue  =  "RIGHT"
  pItem(16).nItemType =  %PIT_COMBO
  pItem(16).cmbItems  =  "RIGHT|LEFT|CENTER|LEADING|TRAILING|"
  pItem(16).tip       =  "�������ȷ���ı���ˮƽ����λ��"
  CALL SendMessage(ghList,%LB_ADDSTRING,0,VARPTR(pItem(16).propName))
END SUB                                                                                                                                                                                        _
'--------------------------------------------------------------------------
' Ϊѡ����ɫ�Ի������16���Զ�����ɫ
'--------------------------------------------------------------------------
SUB InitializeCustomColors()
  LOCAL lCounter AS LONG
  DIM cColor(0:30)
  '�����Զ������ɫ
  RANDOMIZE TIMER
  FOR lCounter = 0 TO 15
    'ʾ����ɫ
    cColor(lCounter) = RND(0,16777215) 'GetSysColor(lCounter)
  NEXT
END SUB