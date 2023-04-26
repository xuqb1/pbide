''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 文件 : dock.inc
' 实现窗口泊坞功能
'
TYPE DockInfo      '泊坞信息自定义类型
  hWnd          AS DWORD      '泊坞时与泊坞句柄一致，浮动时与浮动句柄一致
  hWndD         AS DWORD      '泊坞时的句柄 hDock
  hWndF         AS DWORD      '浮动时的句柄 hFloat
  ID            AS DWORD      '窗口ID，如%ID_PROPERTY
  tid           AS DWORD      '指示窗口泊坞到的窗口ID，如%ID_TOOLBOX，%ID_CLIENTWINDOW
  ml            AS STRING * 1 '泊坞到主窗口的位置，继承自tid的ml,如果tid是客户区的话，则与l一致
  l             AS STRING * 1 '泊坞位置, "T" "B" "L" "R" "F"(上下左右浮动) "I"(隐藏)
  drc           AS RECT       '泊坞矩形：左上角，宽和高 dockrc
  frc           AS RECT       '浮动矩形：左上角，宽和高 floatrc
  byside        AS BYTE       '对泊坞区的靠边情况：左上右下，靠就是1,不靠就是0， 初始或继承自tid，并根据泊坞位置有所变化
  minH          AS LONG       '最小高度
  minW          AS LONG       '最小宽度
END TYPE
GLOBAL gdi()      AS DockInfo '管理所有泊坞窗口的泊坞信息
GLOBAL tA         AS BYTE     '客户区的上面泊坞区靠边情况：010,表示左右都不到边;0，表示泊坞区不存在
GLOBAL tARc       AS RECT     '上泊坞区位置大小（左上宽高）
GLOBAL lA         AS BYTE     '左泊坞区靠边情况指示
GLOBAL lARc       AS RECT     '左泊坞区位置大小：左上宽高
GLOBAL rA         AS BYTE     '右泊坞区靠边情况指示
GLOBAL rARc       AS RECT     '右泊坞区位置大小：左上宽高
GLOBAL bA         AS BYTE     '下泊坞区靠边情况指示
GLOBAL bARc       AS RECT     '下泊坞区位置大小：左上宽高

GLOBAL gSizer           AS LONG     '当前可调尺寸的位置
GLOBAL fakeXButtonRect  AS RECT     '假关闭按钮大小
GLOBAL hWndDockBtn1     AS DWORD    '指示按钮：集合上
GLOBAL hWndDockBtn2     AS DWORD    '指示按钮：集合下
GLOBAL hWndDockBtn3     AS DWORD    '指示按钮：集合左
GLOBAL hWndDockBtn4     AS DWORD    '指示按钮：集合右
GLOBAL hWndDockBtn5     AS DWORD    '指示按钮：分离左
GLOBAL hWndDockBtn6     AS DWORD    '指示按钮：分离右
GLOBAL hWndDockBtn7     AS DWORD    '指示按钮：分离上
GLOBAL hWndDockBtn8     AS DWORD    '指示按钮：分离下
GLOBAL desthWnd         AS DWORD    '泊坞到目标窗口句柄
GLOBAL destWinRc        AS RECT     '泊坞到目标窗口位置大小：左上宽高
GLOBAL hHsplit          AS LONG     '调节尺寸样式的鼠标句柄:横向调节
GLOBAL hVsplit          AS LONG     '调节尺寸样式的鼠标句柄:纵向调节
GLOBAL nDivTop          AS INTEGER  '大小调整时的分隔条指示：开始点，横向调节是顶值，纵向调节是左值
GLOBAL nDivBtm          AS INTEGER  '大小调整时的分隔条指示：结束点，横向调节是底值，纵向调节是右值
GLOBAL nOldDivX         AS INTEGER  '大小调整时的分隔条指示：横向调节是鼠标x值，纵向调节是鼠标y值
GLOBAL bTracking        AS LONG
GLOBAL activehWnd       AS DWORD
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
'泊坞窗口ID
%ID_CLIENTWINDOW  = %WM_USER + 1100  '%WM_USER值是1024，主窗口客户区实体窗口ID
%ID_PROPERTY  = 110   '属性窗口ID
%ID_CONTROLS  = 111   '控件窗口ID
%ID_PROJECT   = 112   '项目窗口ID
%ID_COMPILERS = 113   '编译结果窗口ID
%ID_FINDRS    = 114   '查找结果窗口ID
%ID_FTOOLBAR      = %WM_USER + 1024&  ' 工具箱窗口中的工具栏
' 注册泊坞窗口类
FUNCTION RegisterDockClass()AS LONG
  LOCAL wClass        AS WndClassEx
  LOCAL szClassName   AS ASCIIZ*50
  szClassName          = "DOCKIT"
  wClass.cbSize        = SIZEOF(wClass)
  wClass.style         = %CS_HREDRAW OR %CS_VREDRAW
  wClass.lpfnWndProc   = CODEPTR( DockitProc )    '绑定相应的处理过程
  wClass.cbClsExtra    = 0
  wClass.cbWndExtra    = 0
  wClass.hInstance     = g_hInst
  wClass.hIcon         = LoadIcon  ( %NULL, BYVAL %IDI_APPLICATION )
  wClass.hCursor       = LoadCursor( %NULL, BYVAL %IDC_ARROW )
  wClass.hbrBackground = GetSysColorBrush(%COLOR_3DFACE) 'GetStockObject(%LTGRAY_BRUSH)'GetStockObject(%COLOR_WINDOW )'%LTGRAY_BRUSH )%COLOR_APPWORKSPACE
  wClass.lpszMenuName  = %NULL
  wClass.lpszClassName = VARPTR(szClassName )
  wClass.hIConSm       = LoadIcon(%NULL, BYVAL %IDI_APPLICATION )
  CALL RegisterClassEx (wClass)
END FUNCTION
' 获取指定点所在的的窗口句柄, hWnd是要避开的当前窗口句柄，主窗口与各泊坞状态的泊坞窗口为可选项
FUNCTION GetWinPtIn(BYVAL hWnd AS DWORD ,BYVAL pt AS POINTAPI) AS DWORD
  LOCAL rc AS RECT
  LOCAL i AS LONG
  IF ISTRUE IsWindowVisible(g_hWndMain) AND hWnd<>g_hWndMain THEN '主窗口可见，但检测的不是主窗口
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).l<>"F" AND gdi(i).l<>"I" AND gdi(i).hWnd<>hWnd _  '目标窗口必须是泊坞状态
          AND gdi(i).hWndD<>hWnd AND gdi(i).hWndF<>hWnd AND IsWindowVisible(gdi(i).hWndD) THEN     '避开本窗口本身
        GetWindowRect gdi(i).hWndD,rc
        IF PtInRect(rc,pt.x,pt.y) THEN
          FUNCTION=gdi(i).hWndD
          EXIT FUNCTION
        END IF
      END IF
    NEXT i
    GetWindowRECT g_hWndMain,rc
    IF PtInRect(rc,pt.x,pt.y) THEN
      FUNCTION = g_hWndMain
      EXIT FUNCTION
    END IF
  END IF
  FUNCTION = 0
END FUNCTION
FUNCTION GetMainClientRect(BYREF rc AS RECT)AS LONG
  LOCAL toolHeight AS LONG
  LOCAL StatHeight AS LONG
  LOCAL tmpRc AS RECT
  'GetWindowRect g_hRebar,tmpRc
  GetClientRect g_hRebar,tmpRc
  toolHeight=tmpRc.nBottom '-tmpRc.nTop
  'GetWindowRect g_hStatus,tmpRc
  GetClientRect g_hStatus,tmpRc
  StatHeight=tmpRc.nBottom '-tmpRc.nTop
  GetClientRect g_hWndMain,rc
  rc.nTop=toolHeight
  rc.nBottom = rc.nBottom-toolHeight-StatHeight
END FUNCTION
' 创建泊坞窗口
FUNCTION CreateDockWin()AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  IF UBOUND(gdi())<0 THEN   '数组中无元素，则增加一个，然后默认将属性窗口放到主窗口右侧
    REDIM gdi(0)
    rA = %TOTOP OR %TOCENTER OR %TOBOTTOM '右区通上通下
    'GetClientRect g_hWndMain,rc
    GetMainClientRect rc
    rc.nRight  -= 200
    MoveWin g_hWndClient,rc,%TRUE
    SetRect rc,rc.nLeft+rc.nRight,rc.nTop,200,rc.nBottom '转换为宽度和高度
    rARc=rc
    gdi(0).ID   = %ID_PROPERTY
    gdi(0).tid  = %ID_CLIENTWINDOW
    gdi(0).ml   = "R"
    gdi(0).l    = "R"
    gdi(0).drc  = rc
    StrToRc("200,200,200,300",gdi(i).frc)
    gdi(0).minW   = 50
    gdi(0).minH   = 100
    gdi(0).byside = 15
    IF CreatePropertyWin(0)>0 THEN  '创建属性窗口
      SendMessage gdi(0).hWnd,%WM_SIZE,0,0
    END IF
    EXIT FUNCTION
  END IF
  FOR i=0 TO UBOUND(gdi())  '根据数组中的每个元素，创建相应的窗口
    SELECT CASE gdi(i).ID
      CASE %ID_CONTROLS
        CreateToolboxWin i
        'SendMessage gdi(i).hWnd,%WM_SIZE,0,0
        IF gdi(i).l<>"I" THEN
          MENU SET STATE g_hMenuView,BYCMD %IDM_CONTROLS,%MF_CHECKED OR %MF_ENABLED
        END IF
      CASE %ID_PROPERTY
        CreatePropertyWin i
        'SendMessage gdi(i).hWnd,%WM_SIZE,0,0
        IF gdi(i).l<>"I" THEN
          MENU SET STATE g_hMenuView,BYCMD %IDM_PROPERTY,%MF_CHECKED OR %MF_ENABLED
        END IF
      CASE %ID_PROJECT
        CreateProjectWin i
        'SendMessage gdi(i).hWnd,%WM_SIZE,0,0
        IF gdi(i).l<>"I" THEN
          MENU SET STATE g_hMenuView,BYCMD %IDM_PROJECT,%MF_CHECKED OR %MF_ENABLED
        END IF
      CASE %ID_COMPILERS
        CreateCompilersWin i
        'SendMessage gdi(i).hWnd,%WM_SIZE,0,0
        IF gdi(i).l<>"I" THEN
          MENU SET STATE g_hMenuView,BYCMD %IDM_COMPILERS,%MF_CHECKED OR %MF_ENABLED
        END IF
      CASE %ID_FINDRS
        CreateFindrsWin i
        'SendMessage gdi(i).hWnd,%WM_SIZE,0,0
        IF gdi(i).l<>"I" THEN
          MENU SET STATE g_hMenuView,BYCMD %IDM_FINDRS,%MF_CHECKED OR %MF_ENABLED
        END IF
    END SELECT
  NEXT i
END FUNCTION
' 初始gdi结构数组中指定元素，默认浮动
FUNCTION InitGdiF(BYVAL gdiI AS LONG,BYVAL wID AS DWORD)AS LONG
  LOCAL i AS LONG
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  i=gdiI
  RESET gdi(i)
  gdi(i).ID=wID
  StrToRc("200,200,200,300",gdi(i).frc)
  gdi(i).ml="F"
  gdi(i).l="F"
  gdi(i).byside=0
END FUNCTION
' 创建用于指示泊坞的分布式按钮
FUNCTION CreateDockButton() AS LONG
  LOCAL hDlg          AS DWORD
  LOCAL wcex          AS WNDCLASSEX
  LOCAL szClassName   AS ASCIIZ * 80
  LOCAL rc            AS RECT
  LOCAL szCaption     AS ASCIIZ * 255
  LOCAL nLeft         AS LONG
  LOCAL nTop          AS LONG
  LOCAL nWidth        AS LONG
  LOCAL nHeight       AS LONG
  LOCAL destPos       AS POINT
  LOCAL destWid       AS DWORD
  LOCAL destHei       AS DWORD
  LOCAL mainClientPos AS POINT
  LOCAL mainClientWid AS DWORD
  LOCAL mainClientHei AS DWORD
  LOCAL centerBorder  AS RECT
  LOCAL x             AS LONG
  LOCAL y             AS LONG
  LOCAL w             AS LONG
  LOCAL h             AS LONG
  szClassName        = "DockButClass"
  wcex.cbSize        = SIZEOF(wcex)
  wcex.style         = %CS_HREDRAW OR %CS_VREDRAW
  wcex.lpfnWndProc   = CODEPTR(DockButsProc)
  wcex.cbClsExtra    = 0
  wcex.cbWndExtra    = 0
  wcex.hInstance     = g_hInst
  wcex.hCursor       = LoadCursor (%NULL, BYVAL %IDC_ARROW)
  wcex.hbrBackground = %COLOR_3DFACE + 1
  wcex.lpszMenuName  = %NULL
  wcex.lpszClassName = VARPTR(szClassName)
  wcex.hIcon         = LoadIcon (%NULL, BYVAL %IDI_APPLICATION)
  wcex.hIconSm       = LoadIcon (%NULL, BYVAL %IDI_APPLICATION)
  RegisterClassEx wcex
  szCaption="泊坞按钮"
  'dialog get client g_hWndMain to rc.nRight,rc.nBottom
  GetClientRect g_hWndMain,rc
  'GetMainClientRect rc
  mainClientWid=rc.nRight
  mainClientHei=rc.nBottom
  ClientToScreenRECT g_hWndMain,rc
  mainClientPos.x=rc.nLeft
  mainClientPos.y=rc.nTop

  centerBorder.nLeft    = rc.nLeft+24+20
  centerBorder.nTop     = rc.nTop+24+20+24
  centerBorder.nRight   = rc.nRight-24-20-72
  centerBorder.nBottom  = rc.nBottom-24-20-21-25
  w=24
  h=21
  GetClientRect desthWnd,destWinRc
  IF desthWnd<>g_hWndMain THEN
    GetWindowRect desthWnd,destWinRc
    ScreenToClientRect g_hWndMain,destWinRc
    destWinRc.nRight-=destWinRc.nLeft
    destWinRc.nBottom-=destWinRc.nTop
  END IF
  x=destWinRc.nLeft+(destWinRc.nRight-72)/2  '限制被泊窗口最小宽度和高度为120
  y=destWinRc.nTop+(destWinRc.nBottom-h)/2
'  IF x<centerBorder.nLeft THEN
'    x=centerBorder.nLeft
'  ELSEIF x>centerBorder.nRight THEN
'    x=centerBorder.nRight
'  END IF
'  IF y<centerBorder.nTop THEN
'    y=centerBorder.nTop
'  ELSEIF y>centerBorder.nBottom THEN
'    y=centerBorder.nBottom
'  END IF
  hDlg=CreateWindowEx(%WS_EX_TOOLWINDOW OR %WS_EX_TOPMOST , _           ' extended style
                       szClassName, _       ' 窗口类名
                       szCaption, _         ' 窗口标题
                       %WS_POPUP, _         ' 窗口样式 %WS_CAPTION OR %WS_SYSMENU
                       mainClientPos.x, _   ' 初始 x 位置
                       mainClientPos.y, _   ' 初始 y 位置
                       mainClientWid, _     ' 初始宽度
                       mainClientHei, _     ' 初始调试
                       %NULL, _             ' 父窗口
                       0, _                 ' 窗口菜单句柄
                       g_hInst, _            ' 程序实例句柄
                       BYVAL %NULL)
  '集合上按钮
  hWndDockBtn1=CreateWindowEx(0,"STATIC","",_
                  %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS _
                  OR %SS_BITMAP, _
                  x+w+1,y-w-1,21,24, _
                  hDlg,BYVAL %NULL,g_hInst,BYVAL %NULL)
  CALL SendMessage(hWndDockBtn1,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_TOP"))
  '集合下按钮
  hWndDockBtn2=CreateWindowEx(0 ,"STATIC","",_
                  %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS _
                  OR %SS_BITMAP, _
                  x+w+1,y+h+1,21,24, _
                  hDlg,BYVAL %NULL,g_hInst,BYVAL %NULL)
  CALL SendMessage(hWndDockBtn2,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_BOTTOM"))
  '集合左按钮
  hWndDockBtn3=CreateWindowEx(0,"STATIC","",_
                  %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS _
                  OR %SS_BITMAP, _
                  x,y,24,21, _
                  hDlg,BYVAL %NULL,g_hInst,BYVAL %NULL)
  CALL SendMessage(hWndDockBtn3,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_LEFT"))
  '集合右按钮
  hWndDockBtn4=CreateWindowEx(0,"STATIC","",_
                  %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS _
                  OR %SS_BITMAP, _
                  x+w+h+2,y,24,21, _
                  hDlg,BYVAL %NULL,g_hInst,BYVAL %NULL)
  CALL SendMessage(hWndDockBtn4,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_RIGHT"))
  '分离左按钮
  hWndDockBtn5=CreateWindowEx(0 ,"STATIC","",_
                  %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS _
                  OR %SS_BITMAP, _
                  10,(mainClientHei-21)/2,24,21, _
                  hDlg,BYVAL %NULL,g_hInst,BYVAL %NULL)
  CALL SendMessage(hWndDockBtn5,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_LEFT"))
  '分离右按钮
  hWndDockBtn6=CreateWindowEx(0,"STATIC","",_
                  %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS _
                  OR %SS_BITMAP, _ OR %BS_BITMAP OR %BS_CENTER OR %BS_VCENTER, _
                  mainClientWid-24-10,(mainClientHei-21)/2,24,21, _
                  hDlg,BYVAL %NULL,g_hInst,BYVAL %NULL)
  CALL SendMessage(hWndDockBtn6,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_RIGHT"))
  '分离上按钮
  hWndDockBtn7=CreateWindowEx(0,"STATIC","",_
                  %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS _
                  OR %SS_BITMAP, _
                  (mainClientWid-21)/2,10,21,24, _
                  hDlg,BYVAL %NULL,g_hInst,BYVAL %NULL)
  CALL SendMessage(hWndDockBtn7,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_TOP"))
  '分离下按钮
  hWndDockBtn8=CreateWindowEx(0,"STATIC","",_
                  %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS _
                  OR %SS_BITMAP, _
                  (mainClientWid-21)/2,mainClientHei-24-10,21,24, _
                  hDlg,BYVAL %NULL,g_hInst,BYVAL %NULL)
  CALL SendMessage(hWndDockBtn8,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_BOTTOM"))
  ShowWindow hDlg, %SW_SHOW 'nCmdShow
  UpdateWindow hDlg
  LOCAL uMsg AS tagMsg
  WHILE GetMessage(uMsg, %NULL, 0, 0)
    IF ISFALSE IsDialogMessage(hDlg, uMsg) THEN
      TranslateMessage uMsg
      DispatchMessage uMsg
    END IF
  WEND
  FUNCTION=uMsg.wParam
END FUNCTION
'泊坞按钮窗口回调，消息处理
FUNCTION DockButsProc (BYVAL hWnd AS DWORD, BYVAL wMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG
  SELECT CASE wMsg
    CASE %WM_CREATE
      SetRegion hWnd
    CASE %WM_DESTROY
      PostQuitMessage 0
      EXIT FUNCTION
  END SELECT
  FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam)
END FUNCTION
'对泊坞窗口进行裁剪，只保留有按钮的部分
SUB SetRegion(BYVAL hWnd AS LONG)
  LOCAL hRgn              AS DWORD
  LOCAL x,y               AS LONG
  LOCAL w,h               AS LONG
  LOCAL rc                AS RECT
  LOCAL rgPoints()        AS POINTAPI
  LOCAL mainClientPos     AS POINTAPI
  LOCAL mainClientWid     AS DWORD
  LOCAL mainClientHei     AS DWORD
  LOCAL centerBorder      AS RECT '中间面板边界
  LOCAL startPos          AS POINT'左上角启点
  LOCAL arr()             AS LONG
  REDIM arr(7)
  REDIM rgPoints(1 TO 32)
  ARRAY ASSIGN arr()=4,4,4,4,4,4,4,4
  GetClientRECT hWnd,rc
  IF desthWnd<>g_hWndMain THEN
    GetWindowRect hWnd,rc
    ScreenToClientRect g_hWndMain,rc
    rc.nRight-=rc.nLeft
    rc.nBottom-=rc.nTop
  END IF
  centerBorder.nLeft    = rc.nLeft+24+20
  centerBorder.nTop     = rc.nTop+24+20+24
  centerBorder.nRight   = rc.nRight-24-20-72
  centerBorder.nBottom  = rc.nBottom-24-20-21-25
  ' 左侧按钮
  w=24
  h=21
  x=10
  y=(rc.nBottom-h)/2
  rgPoints(1).x=x     : rgPoints(1).y=y
  rgPoints(2).x=x+w   : rgPoints(2).y=y
  rgPoints(3).x=x+w   : rgPoints(3).y=y+h
  rgPoints(4).x=x     : rgPoints(4).y=y+h
  '右侧按钮
  x=rc.nRight-w-10
  y=(rc.nBottom-h)/2
  rgPoints(5).x=x     : rgPoints(5).y=y
  rgPoints(6).x=x+w   : rgPoints(6).y=y
  rgPoints(7).x=x+w   : rgPoints(7).y=y+h
  rgPoints(8).x=x     : rgPoints(8).y=y+h
  '上面按钮
  w=21
  h=24
  x=(rc.nRight-rc.nLeft-w)/2
  y=10
  rgPoints(9).x=x     : rgPoints(9).y=y
  rgPoints(10).x=x+w   : rgPoints(10).y=y
  rgPoints(11).x=x+w   : rgPoints(11).y=y+h
  rgPoints(12).x=x     : rgPoints(12).y=y+h
  '下面按钮
  x=(rc.nRight-rc.nLeft-w)/2
  y=rc.nBottom-h-10
  rgPoints(13).x=x     : rgPoints(13).y=y
  rgPoints(14).x=x+w   : rgPoints(14).y=y
  rgPoints(15).x=x+w   : rgPoints(15).y=y+h
  rgPoints(16).x=x     : rgPoints(16).y=y+h
  '中间面板
  w=24
  h=21
  IF desthWnd<>g_hWndMain THEN
    GetWindowRect desthWnd,destWinRc
    ScreenToClientRect g_hWndMain,destWinRc
    destWinRc.nRight-=destWinRc.nLeft
    destWinRc.nBottom-=destWinRc.nTop
  END IF
  x=destWinRc.nLeft+(destWinRc.nRight-72)/2
'  IF x<centerBorder.nLeft THEN
'    x=centerBorder.nLeft
'  ELSEIF x>centerBorder.nRight THEN
'    x=centerBorder.nRight
'  END IF
  y=destWinRc.nTop+(destWinRc.nBottom-h)/2
'  IF y<centerBorder.nTop THEN
'    y=centerBorder.nTop
'  ELSEIF y>centerBorder.nBottom THEN
'    y=centerBorder.nBottom
'  END IF
  '左按钮
  w=24
  h=21
  startPos.x=x         : startPos.y=y
  rgPoints(17).x=x     : rgPoints(17).y=y
  rgPoints(18).x=x+w   : rgPoints(18).y=y
  rgPoints(19).x=x+w   : rgPoints(19).y=y+h
  rgPoints(20).x=x     : rgPoints(20).y=y+h
  '上按钮
  startPos.x=x+w+1      : startPos.y=y-w-1
  rgPoints(21).x=startPos.x     : rgPoints(21).y=startPos.y
  rgPoints(22).x=startPos.x+h   : rgPoints(22).y=startPos.y
  rgPoints(23).x=startPos.x+h   : rgPoints(23).y=startPos.y+w
  rgPoints(24).x=startPos.x     : rgPoints(24).y=startPos.y+w
  '右按钮
  startPos.x=x+w+1+h+1      : startPos.y=y
  rgPoints(25).x=startPos.x     : rgPoints(25).y=startPos.y
  rgPoints(26).x=startPos.x+w   : rgPoints(26).y=startPos.y
  rgPoints(27).x=startPos.x+w   : rgPoints(27).y=startPos.y+h
  rgPoints(28).x=startPos.x     : rgPoints(28).y=startPos.y+h
  '下按钮
  startPos.x=x+w+1      : startPos.y=y+h+1
  rgPoints(29).x=startPos.x     : rgPoints(29).y=startPos.y
  rgPoints(30).x=startPos.x+h   : rgPoints(30).y=startPos.y
  rgPoints(31).x=startPos.x+h   : rgPoints(31).y=startPos.y+w
  rgPoints(32).x=startPos.x     : rgPoints(32).y=startPos.y+w
  hRgn =  CreatePolyPolygonRgn (rgPoints(1), arr(0),8, %ALTERNATE)
  SetWindowRgn hWnd, hRgn, %TRUE
END SUB
'----------------------------------------------------------------------
' 泊坞窗口回调，消息处理过程
'----------------------------------------------------------------------
FUNCTION DockitProc(BYVAL hWnd    AS DWORD, BYVAL wMsg    AS DWORD,_
                        BYVAL wParam  AS DWORD, BYVAL lParam  AS DWORD) AS LONG
  LOCAL   hDC             AS LONG
  LOCAL   ps              AS PAINTSTRUCT
  LOCAL   rc              AS RECT
  LOCAL   pt              AS POINTAPI
  STATIC  rcDrag          AS RECT
  STATIC  rcOldPopup      AS RECT
  STATIC  rcOldChild      AS RECT
  STATIC  xStart          AS INTEGER
  STATIC  yStart          AS INTEGER
  STATIC  ptDrawStart     AS POINTAPI
  STATIC  ptDrawEnd       AS POINTAPI
  LOCAL   x               AS INTEGER
  LOCAL   y               AS INTEGER
  STATIC  bDragging       AS LONG
  STATIC  bLostCapture    AS LONG
  STATIC  bTracking       AS LONG
  STATIC  hadDrag         AS LONG
  STATIC  hadTrack        AS LONG
  STATIC  nFirstHit       AS INTEGER  '左键按下的第一位置
  STATIC  nLastHit        AS INTEGER  '左键放开时的最后位置
  STATIC  nDivTop         AS INTEGER
  STATIC  nDivBtm         AS INTEGER
  STATIC  nOldDivX        AS INTEGER
  LOCAL   lpMs            AS MEASUREITEMSTRUCT PTR
  LOCAL   Lb              AS LOGBRUSH
  STATIC  temphWnd        AS LONG
  LOCAL   pt1             AS POINTAPI
  LOCAL   gdi1            AS DockInfo
  LOCAL   gdiI            AS LONG
  LOCAL   gdiId           AS LONG
  STATIC  lastl           AS STRING
  STATIC  lastml          AS STRING
  STATIC  resizeDA        AS STRING
  LOCAL   tmpPt           AS POINTAPI
  LOCAL   mrc             AS RECT '主窗口客户区位置/大小
  STATIC  olddsthWnd      AS LONG
  LOCAL   ptCur           AS POINTAPI
  LOCAL   tmpRc           AS RECT
  LOCAL   i               AS LONG
  LOCAL   tmpAsc          AS ASCIIZ * 18
  LOCAL   tmpStr          AS STRING
  LOCAL   allWids         AS LONG
  LOCAL   lefts           AS LONG
  LOCAL   allHeit         AS LONG
  LOCAL   top             AS LONG
  LOCAL   hMove&
  LOCAL   nDir&
  LOCAL   hBrush&
  LOCAL   nRight&
  LOCAL   j&
  SELECT CASE ( wMsg )
    CASE %WM_CREATE
      bDragging = %FALSE
      bTracking = %FALSE
      IF ghFont<=0 THEN
        ghFont = MakeStdFont("Arial", 9)
      END IF
      '设置默认尺寸，或从ini文件读取
      GetWindowRect hWnd, rcOldPopup
      '设置可调区域默认位置或从ini文件得到
      gSizer = 180
      '加载可调区域的鼠标形状，任意使用一个或从资源文件中自定义一个
      IF hVsplit<=0 THEN
        hVsplit = LoadCursor( %NULL,BYVAL %IDC_SIZENS )
        hHsplit = LoadCursor( %NULL,BYVAL %IDC_SIZEWE )
      END IF
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_SIZE
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
      '如果是弹出窗口，则复制窗口的扩展方式
      IF gdi(gdiI).l="F" THEN

        'IF gdi(gdiI).hWndF=hWnd THEN
        GetWindowRect hWnd, rcOldPopup
        IF gdi(gdiI).hWndF=hWnd THEN
          gdi(gdiI).frc=rcOldPopup
          gdi(gdiI).frc.nRight-=gdi(gdiI).frc.nLeft
          gdi(gdiI).frc.nBottom-=gdi(gdiI).frc.nTop
        END IF
'          MoveDockWinContents gdiI
'          msgbox "tt"
        'END IF
      END IF
      '窗口位置
      On_SizeDockWin gdi(gdiI).ID,hWnd,wParam '泊坞窗口大小调整事件
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_MEASUREITEM
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
      IF gdi(gdiI).ID=%ID_PROPERTY THEN
        lpMs = lParam
        '设定属性列表的每行高度
        @lpMs.itemHeight = 15
      END IF
      FUNCTION = %TRUE
      EXIT FUNCTION
    CASE %WM_DRAWITEM
      IF bTracking = %TRUE THEN EXIT FUNCTION
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
      IF gdi(gdiI).ID=%ID_PROPERTY THEN
        DrawPropList lParam
      END IF
      FUNCTION = %TRUE : EXIT FUNCTION
    CASE %WM_COMMAND
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
      On_CommandDockWin gdi(gdiI).ID,wParam '泊坞窗口中控件命令触发事件
    CASE %WM_DESTROY
      IF ghFont <> 0 THEN CALL DeleteObject( ghFont )
      'DIALOG GET USER hWnd,1 TO gdiI
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
      On_DestroyDockWin gdi(gdiI).ID '泊坞窗口销毁事件
      CALL PostQuitMessage( 0 )
    CASE %WM_CAPTURECHANGED '针对控件的鼠标获取与失去发生变化时
      '-------------------------------------------------------
      '丢失鼠标捕捉句柄
      '用于Windows v4及以上
      '-------------------------------------------------------
      '如果释放了鼠标捕捉，则不做它
      IF bDragging = %TRUE THEN
        FUNCTION = 0
        EXIT IF
      ELSE
        IF bLostCapture = %TRUE THEN
          bLostCapture = %FALSE
          '清除桌面上的GDI垃圾
          InvalidateRect %NULL,BYVAL %NULL,%TRUE
          UpdateWindow %NULL
          FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam )
          EXIT FUNCTION
        END IF
      END IF
    CASE %WM_CANCELMODE
      '----------------------------------------------------------
      '用于Windows 95/98及NT4
      '在系统“偷取”鼠标捕捉前发送此消息
      '---------------------------------------------------------
      '取消拖动操作
      IF bDragging = %TRUE THEN
        bDragging = %FALSE
        '如果在拖动，再次清除之前我们用到的。
        rcDrag.nLeft=0:rcDrag.nTop=0:rcDrag.nRight=0:rcDrag.nBottom=0
        ptDrawEnd.x=0:ptDrawEnd.y=0:ptDrawStart.x=0:ptDrawStart.y=0
        InvalidateRect %NULL,BYVAL %NULL,%TRUE
        UpdateWindow %NULL
      END IF
      '*ReleaseCapture() 定义的窗口处理过程将用到它
      '清除留下的GDI垃圾
      bLostCapture = %TRUE
      FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam )
      EXIT FUNCTION
    CASE %WM_MOUSEACTIVATE
      activehWnd = hWnd
      FOR i=0 TO UBOUND(gdi())
        IF gdi(i).l<>"I" AND gdi(i).hWnd>0 AND IsWindow(gdi(i).hWnd) THEN
          DIALOG REDRAW gdi(i).hWnd
        END IF
      NEXT i
    CASE %WM_LBUTTONDOWN
      'canLog=1 : RunLog("mousedown activehWnd=" & str$(activehWnd)) : canLog=0
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
      gdi1=gdi(gdiI)
      lastl=gdi1.l
      IF gdi1.l = "F" THEN
        '强制主窗口标题栏为激活状态
        SendMessage g_hWndMain, %WM_NCACTIVATE, %TRUE, 0
        GetWindowRect hWnd, rcOldPopup
      ELSE
        rcOldPopup=gdi1.frc
        rcOldPopup.nRight+=rcOldPopup.nLeft
        rcOldPopup.nBottom+=rcOldPopup.nTop
      END IF
      '----------------------------------------------------------------
      '测试标题区域的拖放与泊坞
      '----------------------------------------------------------------
      GetClientRect hWnd, rc
      IF gdi1.l="B" THEN  '泊坞到下面时，留出一点作为分隔条位置
        rc.nTop=rc.nTop+3
      END IF
      IF gdi1.ml="B" AND gdi1.byside=(gdi1.byside OR %BYTOP) THEN
        rc.nTop=rc.nTop+3
      END IF
      rc.nLeft+=3
      rc.nRight-=3
      rc.nBottom = rc.nTop+17   '设置测试区为顶端部分
      GetCursorPos pt           '取得鼠标位置（屏幕坐标）
      ScreenToClient hWnd, pt   '转换到客户区坐标
      '--测试是否在假的关闭按钮区域内
      SetRect fakeXButtonRect,rc.nRight-10,rc.nTop,rc.nRight-1,rc.nBottom
      IF PtInRect(fakeXButtonRect,pt.x,pt.y) THEN
        IF activehWnd=hWnd THEN
          activehWnd=0
        END IF
        DockWinHide gdiI
        EXIT FUNCTION
      END IF
      activehWnd = hWnd
      FOR i=0 TO UBOUND(gdi())
        IF gdi(i).l<>"I" AND gdi(i).hWnd>0 AND IsWindow(gdi(i).hWnd) THEN
          DIALOG REDRAW gdi(i).hWnd
        END IF
      NEXT i
      '--如果在假的标题栏区域内
      IF PtInRect( rc, pt.x, pt.y ) THEN
        '捕获鼠标
        SetCapture hWnd
        '得到当前窗口的矩形RECT信息
        GetWindowRect hWnd,rc
        '这是将要维护的虚拟窗口
        '将使用最终的rcDraw RECT来定位控件
        rcDrag.nLeft   = rc.nLeft
        rcDrag.nTop    = rc.nTop
        rcDrag.nRight  = rc.nRight
        rcDrag.nBottom = rc.nBottom
        '设置开始点
        xStart = LOWRD ( lParam )
        yStart = HIWRD ( lParam )
        '由首个矩形开始
        IF gdi(gdiI).l="F" THEN
          DrawRectangle ptDrawStart, ptDrawEnd,3
        ELSE
          DrawRectangle ptDrawStart, ptDrawEnd,1
        END IF
        '开始拖动窗口
        bDragging = %TRUE
        FUNCTION = 0 : EXIT FUNCTION
      END IF
      '-------------------------------------------------------
      '测试每个可调区域，并设置我们的鼠标
      '-------------------------------------------------------
      pt.x = LO(WORD,lParam)
      pt.y = HI(WORD,lParam)
      'canLog=1 : RunLog("lbuttondown pt=" & PtToStr(pt)) : canLog=0
      '客户区也需要调整大小
      tmpPt=pt
      ClientToScreen hWnd,pt
      ScreenToClient g_hWndMain,pt
      IF lA>0 THEN
        SetRect rc,lARc.nLeft+lARc.nRight-5,lARc.nTop,lARc.nLeft+lARc.nRight+2,lARc.nTop+lARc.nBottom
        IF PtInRect(rc,pt.x,pt.y) THEN
          nDivTop   = rc.ntop
          nDivBtm   = rc.nbottom
          nOldDivX  = tmpPt.x 'pt.x
          '保存可调区域的初始位置
          nFirstHit = nOldDivx
          '需要确定SplitBarInvertLine在这里能不能用 ，不能用，分隔条位置错误
          'SplitBarInvertLine g_hWndClient, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,0
          '捕获鼠标
          SetCapture hWnd
          bTracking = %TRUE
          SetCursor hHsplit
          resizeDA="L"
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      IF tA>0 THEN
        SetRect rc,tARc.nLeft,tARc.nTop+tARc.nBottom-5,tARc.nLeft+tARc.nRight,tARc.nTop+tARc.nBottom+2
        IF PtInRect(rc,pt.x,pt.y) THEN
          nDivTop   = rc.nLeft 'ntop
          nDivBtm   = rc.nRight 'nbottom
          nOldDivX  = tmpPt.y'pt.y
          '保存可调区域的初始位置
          nFirstHit = nOldDivx
          '需要确定SplitBarInvertLine在这里能不能用
          'SplitBarInvertLine g_hWndClient, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,1
          '捕获鼠标
          SetCapture hWnd
          bTracking = %TRUE
          SetCursor hVsplit
          resizeDA="T"
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      IF rA>0 THEN
        SetRect rc,rARc.nLeft-2,rARc.nTop,rARc.nLeft+5,rARc.nTop+rARc.nBottom
        IF PtInRect(rc,pt.x,pt.y) THEN
          nDivTop   = rc.ntop
          nDivBtm   = rc.nbottom
          nOldDivX  = tmpPt.x 'pt.x
          '保存可调区域的初始位置
          nFirstHit = nOldDivx
          '需要确定SplitBarInvertLine在这里能不能用
          'SplitBarInvertLine g_hWndClient, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,0
          '捕获鼠标
          SetCapture hWnd
          bTracking = %TRUE
          SetCursor hHsplit
          resizeDA="R"
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      IF bA>0 THEN
        SetRect rc,bARc.nLeft,bARc.nTop-2,bARc.nLeft+bARc.nRight,bARc.nTop+5
        IF PtInRect(rc,pt.x,pt.y) THEN
          nDivTop   = rc.nLeft'ntop
          nDivBtm   = rc.nRight'nbottom
          nOldDivX  = tmpPt.y 'pt.y
          '保存可调区域的初始位置
          nFirstHit = nOldDivx
          '需要确定SplitBarInvertLine在这里能不能用
          'SplitBarInvertLine g_hWndClient, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,1
          '捕获鼠标
          SetCapture hWnd
          bTracking = %TRUE
          SetCursor hVsplit
          resizeDA="B"
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      resizeDA=""
      '在排除了客户区边缘调整之外，泊坞窗口的调节就只能是相互之间的调整了
      '只需要注意右侧和下方
      IF gdi1.l="F" OR gdi1.l="I" THEN '非泊坞窗口，不需要调整
        EXIT FUNCTION
      END IF
      '主窗口右下角泊坞窗口，不需要调整
      IF gdi1.byside=(gdi1.byside OR %BYRIGHT) AND gdi1.byside=(gdi1.byside OR %BYBOTTOM) THEN
        EXIT FUNCTION
      END IF
      IF gdi1.byside<>(gdi1.byside OR %BYRIGHT) THEN '不靠右边，则右侧可调整
        rc=gdi1.drc
        top=getLastLTopVal(gdiI,rc.nLeft+rc.nRight)'rc.nTop
        allHeit=GetLastLBottomVal(gdiI,gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight)
        SetRect rc,rc.nLeft+rc.nRight-5,top,rc.nLeft+rc.nRight+2,allHeit     '右侧调整
        IF PtInRect(rc,pt.x,pt.y) THEN
          nDivTop = rc.nTop
          nDivBtm = rc.nBottom
          nOldDivX = tmpPt.x'pt.x     'tmpPt.x
          nFirstHit = nOldDivx
          'SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,0
          '捕获鼠标
          SetCapture hWnd
          bTracking = %TRUE
          SetCursor hHsplit
          resizeDA = "hr"
          EXIT FUNCTION
        END IF
      END IF
      IF gdi1.byside<>(gdi1.byside OR %BYBOTTOM) THEN '不靠下边，则下边可调整
        rc=gdi1.drc
        lefts=GetLastTLeftVal(gdiI,gdi1.drc.nTop+gdi1.drc.nBottom) '找最左的左边
        allWids=GetLastTRightVal(gdiI,gdi1.drc.nTop+gdi1.drc.nBottom) '找最右的右边，用于画线
        SetRect rc,lefts,rc.nTop+rc.nBottom-5,allWids,rc.nTop+rc.nBottom+2   '下边调整
        'canLog=1 : RunLog("lbuttondown rc=" & RcToStr(rc) & " pt=" & PtToStr(pt))
        'canLog=0
        IF PtInRect(rc,pt.x,pt.y) THEN
          nDivTop = rc.nLeft
          nDivBtm = rc.nRight'Bottom
          nOldDivX = tmpPt.y 'pt.y
          nFirstHit = nOldDivx
          'SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,1
          '捕获鼠标
          SetCapture hWnd
          bTracking = %TRUE
          SetCursor hHsplit
          resizeDA="hb"
          EXIT FUNCTION
        END IF
      END IF
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_MOUSEMOVE
      '-------------------------------------------------------------------
      '用聚焦矩形来拖动控件
      '-------------------------------------------------------------------
      '如果我们拖动了控件，那么就更新窗口
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
'      canLog=1 : RunLog("GetIndex=" & str$(GetIndex(hWnd)) & " gdiI=" & str$(gdiI))
'      RunLog("lastl=" & lastl & " gdi(" & str$(gdiI) & ").id=" & str$(gdi(gdiI).id))
'      canLog=0
      gdi1=gdi(gdiI)
      IF bDragging = %TRUE THEN
        hadDrag=%TRUE
        IF lastl<>"F" AND lastl<>"" AND lastl<>"I" THEN '之前是泊坞状态，拖动前，先把原来泊坞状态窗口隐藏，并调整泊坞区
          ShowWindow gdi1.hWndD,%SW_HIDE
          gdi1.hWnd=gdi1.hWndF
          MoveDockWinContents gdiI '鼠标移动并拖动泊坞窗口时转移其中的控件到当前窗口中
          AdjustForMove gdiI
          gdi1.l="F"
          gdi1.ml="F"
          gdi(gdiI)=gdi1
          lastl="F"
        END IF
        GetCursorPos pt1           '取得鼠标位置（屏幕坐标）
        temphWnd=GetWinPtIn(hWnd,pt1)
        IF hWndDockBtn1=0 AND temphWnd>0 THEN
          desthWnd=temphWnd
          CreateDockButton
          olddsthWnd=temphWnd
        END IF
        'canLog=1 : RunLog() : canLog=0
        'canLog=1 : RunLog("mousemove temphWnd=" & str$(temphWnd) & " olddsthWnd=" & str$(olddsthWnd)) : canLog=0
        IF temphWnd<>olddsthWnd AND temphWnd>0 AND hWndDockBtn1>0 THEN
          DestroyWindow GetParent(hWndDockBtn1)
          olddsthWnd=temphWnd
          hWndDockBtn1=0
        END IF
        x = LOWRD ( lParam )
        y = HIWRD ( lParam )
        '移动鼠标即移动矩形
        IF (( x <> xStart ) OR ( y <> yStart )) THEN
          '擦除以前的矩形
          IF gdi(gdiI).l="F" THEN
            DrawRectangle  ptDrawStart,ptDrawEnd,3
          ELSE
            DrawRectangle  ptDrawStart,ptDrawEnd,1
          END IF
          IF bDragging = %FALSE THEN
            FUNCTION=0
            EXIT FUNCTION
          END IF
          '为新位置得到新点
          rcDrag.nLeft   = rcDrag.nLeft   + x - xStart
          rcDrag.nRight  = rcDrag.nRight  + x - xStart
          rcDrag.nTop    = rcDrag.nTop    + y - yStart
          rcDrag.nBottom = rcDrag.nBottom + y - yStart
          ' 确定鼠标在RECT内部，如果不是，则调整RECT来包含鼠标
          GetCursorPos ptCur
          IF ISFALSE(PtInRect( rcDrag,ptCur.x, ptCur.y )) THEN
            '调节虚拟RECT包含鼠标
            rcDrag.nLeft =   ptCur.x - ((rcDrag.nRight-rcDrag.nLeft)/2)
          END IF
          '*决定要靠到哪一边：L,T,R,B,ML,MT,MR,MB 或浮动窗口。改变虚拟矩形来示意用户窗口将会绑定到哪一边或浮动
          LOCAL lRes AS LONG
          'GetClientRect g_hWndMain,mrc
          GetMainClientRect mrc
          mrc.nRight+=mrc.nLeft
          mrc.nBottom+=mrc.nTop
          ClientToScreenRECT g_hWndMain,mrc
          lRes = HitTestMain1( temphWnd, ptCur )
          IF temphWnd=g_hWndMain THEN
            'GetClientRect temphWnd,rc
            GetMainClientRect rc
            rc.nRight+=rc.nLeft
            rc.nBottom+=rc.nTop
            ClientToScreenRECT temphWnd,rc
          ELSE
            GetWindowRect temphWnd,rc
          END IF
          IF lRes = %MLEFT THEN '泊坞到了主窗口左侧
            '为新位置得到新点
            rcDrag.nLeft=mrc.nLeft
            rcDrag.nRight=rcDrag.nLeft+150
            rcDrag.nTop=mrc.nTop
            rcDrag.nBottom=mrc.nBottom
            IF lA>0 THEN
              tmpRc=lARc
              ClientToScreenRect g_hWndMain,lARc
              rcDrag.nLeft=lARc.nLeft
              rcDrag.nRight=rcDrag.nLeft+tmpRc.nRight/2
              rcDrag.nTop=lARc.nTop
              rcDrag.nBottom=tmpRc.nTop+lARc.nBottom
              ScreenToClientRect g_hWndMain,lARc
            ELSE
              IF tA>0 AND tA=(tA OR %TOLEFT) THEN
                rcDrag.nTop+=tARc.nBottom
              END IF
              IF bA>0 AND bA=(bA OR %TOLEFT) THEN
                rcDrag.nBottom-=bARc.nBottom
              END IF
            END IF
            gdi1.tid=%ID_CLIENTWINDOW
            gdi1.l = "L"
            gdi1.ml = "L"
          ELSEIF lRes = %MTOP THEN  '泊坞到了主窗口的上方
            '为新位置得到新点
            rcDrag.nTop=mrc.nTop
            rcDrag.nLeft=mrc.nLeft
            rcDrag.nRight=mrc.nRight
            rcDrag.nBottom=rcDrag.nTop+150
            IF tA>0 THEN
              tmpRc=tARc
              ClientToScreenRect g_hWndMain,tARc
              rcDrag.nTop     =tARc.nTop
              rcDrag.nLeft    =tARc.nLeft
              rcDrag.nRight   =tmpRc.nLeft+tARc.nRight
              rcDrag.nBottom  =rcDrag.nTop+tmpRc.nBottom/2
              ScreenToClientRect g_hWndMain,tARc
            ELSE
              IF lA>0 AND lA=(lA OR %TOTOP) THEN
                rcDrag.nLeft+=lARc.nRight
              END IF
              IF rA>0 AND rA=(rA OR %TOTOP) THEN
                rcDrag.nRight-=rARc.nRight
              END IF
            END IF
            gdi1.tid=%ID_CLIENTWINDOW
            gdi1.l = "T"
            gdi1.ml = "T"
          ELSEIF lRes = %MRIGHT THEN  '泊坞到了主窗口的右侧
            '为新位置得到新点
            rcDrag.nRight   =mrc.nRight
            rcDrag.nLeft    =rcDrag.nRight-150
            rcDrag.nTop     =mrc.nTop
            rcDrag.nBottom  =mrc.nBottom
            IF rA>0 THEN
              tmpRc=rARc
              ClientToScreenRect g_hWndMain,rARc
              rcDrag.nTop     =rARc.nTop
              rcDrag.nRight   =rARc.nLeft
              rcDrag.nLeft    =rcDrag.nRight-tmpRc.nRight/2
              rcDrag.nBottom  =tmpRc.nTop+rARc.nBottom
              ScreenToClientRect g_hWndMain,rARc
            ELSE
              IF tA>0 AND tA=(tA OR %TORIGHT) THEN
                rcDrag.nTop+=tARc.nBottom
              END IF
              IF bA>0 AND bA=(bA OR %TORIGHT) THEN
                rcDrag.nBottom-=bARc.nBottom
              END IF
            END IF
            gdi1.tid=%ID_CLIENTWINDOW
            gdi1.l = "R"
            gdi1.ml= "R"
          ELSEIF lRes = %MBOTTOM THEN '泊坞到了主窗口的下方
            '为新位置得到新点
            rcDrag.nBottom=mrc.nBottom
            rcDrag.nTop=rcDrag.nBottom-150
            rcDrag.nLeft=mrc.nLeft
            rcDrag.nRight=mrc.nRight
            IF bA>0 THEN
              tmpRc=bARc
              ClientToScreenRect g_hWndMain,bARc
              rcDrag.nLeft=bARc.nLeft
              rcDrag.nRight=tmpRc.nLeft+bARc.nRight
              rcDrag.nBottom=bARc.nTop
              rcDrag.nTop=bARc.nTop-tmpRc.nBottom/2'rcDrag.nBottom-bARc.nBottom/2
              ScreenToClientRect g_hWndMain,bARc
            ELSE
              IF lA>0 AND lA=(lA OR %TOBOTTOM) THEN
                rcDrag.nLeft+=lARc.nRight
              END IF
              IF rA>0 AND rA=(rA OR %TOBOTTOM) THEN
                rcDrag.nRight-=rARc.nRight
              END IF
            END IF
            gdi1.tid=%ID_CLIENTWINDOW
            gdi1.l = "B"
            gdi1.ml = "B"
          ELSEIF lRes = %DLEFT    THEN  '泊坞到了某个泊坞窗口的左侧
            rcDrag.nLeft=rc.nLeft
            rcDrag.nTop=rc.nTop
            rcDrag.nRight=rc.nLeft+(rc.nRight-rc.nLeft)/2
            rcDrag.nBottom=rc.nBottom
            gdi1.l="L"
            gdi1.ml=gdi(GetIndex(temphWnd)).ml
            gdi1.tid=gdi(GetIndex(temphWnd)).ID
          ELSEIF lRes = %DTOP     THEN  '泊坞到了某个泊坞窗口的上方
            rcDrag.nLeft=rc.nLeft
            rcDrag.nTop=rc.nTop
            rcDrag.nRight=rc.nRight
            rcDrag.nBottom=rc.nTop+(rc.nBottom-rc.nTop)/2
            gdi1.l="T"
            gdi1.ml=gdi(GetIndex(temphWnd)).ml
            gdi1.tid=gdi(GetIndex(temphWnd)).ID
          ELSEIF lRes = %DRIGHT   THEN  '泊坞到了某个泊坞窗口的右侧
            rcDrag.nLeft=rc.nLeft+(rc.nRight-rc.nLeft)/2
            rcDrag.nTop=rc.nTop
            rcDrag.nRight=rc.nRight
            rcDrag.nBottom=rc.nBottom
            gdi1.l="R"
            gdi1.ml=gdi(GetIndex(temphWnd)).ml
            gdi1.tid=gdi(GetIndex(temphWnd)).ID
          ELSEIF lRes = %DBOTTOM  THEN  '泊坞到了某个泊坞窗口的下方
            rcDrag.nTop=rc.nTop+(rc.nBottom-rc.nTop)/2
            rcDrag.nLeft=rc.nLeft
            rcDrag.nRight=rc.nRight
            rcDrag.nBottom=rc.nBottom
            gdi1.l="B"
            gdi1.ml=gdi(GetIndex(temphWnd)).ml
            gdi1.tid=gdi(GetIndex(temphWnd)).ID
          ELSE
            '如果浮动，重载rcOld RECT cx & cy
            ''为新位置得到新点，保持相同尺寸
            GetCursorPos ptCur
            rcDrag.nTop = ptCur.y-8
            rcDrag.nRight  = rcDrag.nLeft + (rcOldPopup.nRight-rcOldPopup.nLeft)
            rcDrag.nBottom = rcDrag.nTop  + (rcOldPopup.nBottom-rcOldPopup.nTop)
            gdi1.l = "F"
            gdi1.ml = "F"
            gdi1.tid=0
          END IF
          gdi(gdiI)=gdi1
          '转换rcDraw为新点
          ptDrawEnd.x   = rcDrag.nLeft
          ptDrawEnd.y   = rcDrag.nTop
          ptDrawStart.x = rcDrag.nRight
          ptDrawStart.y = rcDrag.nBottom
          '移动矩形到新位置
          IF gdi(gdiI).l="F" THEN
            ptDrawStart.x = rcDrag.nLeft+gdi(gdiI).frc.nRight
            ptDrawStart.y = rcDrag.nTop+gdi(gdiI).frc.nBottom
            DrawRectangle ptDrawStart, ptDrawEnd,3
          ELSE
            DrawRectangle ptDrawStart, ptDrawEnd,1
          END IF
          '记住上次的位置
          xStart = x
          yStart = y
        END IF
      END IF
      '---------------------------------------------------------
      '如果我们在跟踪它，取得调节区域句柄
      '---------------------------------------------------------
      pt.x = LOWRD(lParam)
      pt.y = HIWRD(lParam)
      'canLog=1
      '移动分隔线到鼠标位置
      IF bTracking = %TRUE  THEN
        hadTrack=%TRUE
        SELECT CASE resizeDA
          CASE "L"
            SetCursor hHsplit
            '删除旧的分隔线
            SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,0
            '画新的分隔线
            SplitBarInvertLine hWnd, pt.x ,nDivTop ,pt.x ,nDivBtm,0
            nOldDivX=pt.x
            'RunLog "mousemove " & resizeDA & " lARc=" & RcToStr(lARc) & " top/Btm:" & STR$(nDivTop) & "/" & STR$(nDivBtm)
            EXIT FUNCTION
          CASE "T"
            SetCursor hVsplit
            '删除旧分隔线
            SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,1
            '画新的分隔线
            SplitBarInvertLine hWnd, pt.y ,nDivTop ,pt.y ,nDivBtm,1
            nOldDivX = pt.y
            'RunLog "mousemove " & resizeDA & " tARc=" & RcToStr(tARc) & " top/Btm:" & STR$(nDivTop) & "/" & STR$(nDivBtm)
            EXIT FUNCTION
          CASE "R"
            SetCursor hHsplit
            '删除旧的分隔线
            SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,0
            '画新的分隔线
            SplitBarInvertLine hWnd, pt.x ,nDivTop ,pt.x ,nDivBtm,0
            nOldDivX = pt.x
            'RunLog "mousemove " & resizeDA & " rARc=" & RcToStr(rARc) & " top/Btm:" & STR$(nDivTop) & "/" & STR$(nDivBtm)
            EXIT FUNCTION
          CASE "B"
            SetCursor hVsplit
            '删除旧分隔线
            SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,1
            '画新的分隔线
            SplitBarInvertLine hWnd, pt.y ,nDivTop ,pt.y ,nDivBtm,1
            nOldDivX = pt.y
            'RunLog "mousemove " & resizeDA & " bARc=" & RcToStr(bARc) & " top/Btm:" & STR$(nDivTop) & "/" & STR$(nDivBtm)
            EXIT FUNCTION
          CASE "hr"
            SetCursor hHsplit
            '删除旧的分隔线
            SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,0
           '画新的分隔线
            SplitBarInvertLine hWnd, pt.x ,nDivTop ,pt.x ,nDivBtm,0
            nOldDivX = pt.x
            'RunLog "mousemove " & resizeDA
            EXIT FUNCTION
          CASE "hb"
            SetCursor hVsplit
            '删除旧分隔线
            SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm,1
            '画新的分隔线
            SplitBarInvertLine hWnd, pt.y ,nDivTop ,pt.y ,nDivBtm,1
            nOldDivX = pt.y
            'RunLog "mousemove " & resizeDA
            EXIT FUNCTION
        END SELECT
      END IF
      '点位测试，以确定是否在分隔条附近，然后显示可调节的鼠标
      '首先测试各泊坞区边缘
      'canLog=1 : RunLog("mousemove 0 pt=" & PtToStr(pt))
      'GetCursorPos pt
      'RunLog("1 pt=" & PtToStr(pt))
      'canLog=0
      '将鼠标坐标转成主窗口相对坐标
      ClientToScreen hWnd,pt
      ScreenToClient g_hWndMain,pt
      '检查泊坞区可调情况:判断各泊坞区是否存在，检查鼠标是否在泊坞区与客户区邻接的边缘
      IF lA>0 THEN  '左区
        SetRect rc,lARc.nLeft+lARc.nRight-5,lARc.nTop,lARc.nLeft+lARc.nRight+2,lARc.nTop+lARc.nBottom
        IF PtInRect(rc,pt.x,pt.y) THEN
          SetCursor hHsplit
          rc=fakeXButtonRect
          SetRect rc,rc.nLeft-2,rc.nTop-2,rc.nRight+4,rc.nBottom+6
          InvalidateRect( hwnd,rc,%TRUE)
          'DIALOG REDRAW hWnd
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      IF tA>0 THEN  '上区
        SetRect rc,tARc.nLeft,tARc.nTop+tARc.nBottom-5,tARc.nLeft+tARc.nRight,tARc.nTop+tARc.nBottom+2
        IF PtInRect(rc,pt.x,pt.y) THEN
          SetCursor hVsplit
          rc=fakeXButtonRect
          SetRect rc,rc.nLeft-2,rc.nTop-2,rc.nRight+4,rc.nBottom+6
          InvalidateRect( hwnd,rc,%TRUE)
          'DIALOG REDRAW hWnd
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      IF rA>0 THEN  '右区
        SetRect rc,rARc.nLeft-2,rARc.nTop,rARc.nLeft+5,rARc.nTop+rARc.nBottom
        IF PtInRect(rc,pt.x,pt.y) THEN
          SetCursor hHsplit
          rc=fakeXButtonRect
          SetRect rc,rc.nLeft-2,rc.nTop-2,rc.nRight+4,rc.nBottom+6
          InvalidateRect( hwnd,rc,%TRUE)
          'DIALOG REDRAW hWnd
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      IF bA>0 THEN  '下区
        SetRect rc,bARc.nLeft,bARc.nTop-2,bARc.nLeft+bARc.nRight,bARc.nTop+5
        IF PtInRect(rc,pt.x,pt.y) THEN
          SetCursor hVsplit
          rc=fakeXButtonRect
          SetRect rc,rc.nLeft-2,rc.nTop-2,rc.nRight+4,rc.nBottom+6
          InvalidateRect( hwnd,rc,%TRUE)
          'DIALOG REDRAW hWnd
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      '测试当前窗口边缘
      'rc=gdi1.drc
      '先测试水平分隔条
      IF gdi1.byside<>(gdi1.byside OR %BYRIGHT) THEN '没有靠到右侧，则右侧可调
        rc=gdi1.drc
        SetRect  rc,rc.nLeft+rc.nRight-3,rc.nTop,rc.nLeft+rc.nRight,rc.nTop+rc.nBottom
        IF PtInRect(rc,pt.x,pt.y) THEN
          SetCursor hHsplit
          FUNCTION=0:EXIT FUNCTION
        END IF
        rc=gdi1.drc
        top=getLastLTopVal(gdiI,rc.nLeft+rc.nRight)'rc.nTop
        allHeit=GetLastLBottomVal(gdiI,gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight)
        SetRect rc,rc.nLeft+rc.nRight-5,top,rc.nLeft+rc.nRight+2,allHeit     '右侧调整
        IF PtInRect(rc,pt.x,pt.y) THEN
          SetCursor hHsplit
          rc=fakeXButtonRect
          SetRect rc,rc.nLeft-2,rc.nTop-2,rc.nRight+4,rc.nBottom+6
          InvalidateRect( hwnd,rc,%TRUE)
          'DIALOG REDRAW hWnd
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      '再测试垂直分隔条...
      IF gdi1.byside<>(gdi1.byside OR %BYBOTTOM) THEN '没有靠到底边，则底边可调
        rc=gdi1.drc
        SetRect  rc,rc.nLeft,rc.nTop+rc.nBottom-3,rc.nLeft+rc.nRight,rc.nTop+rc.nBottom
        'canLog=1 : RunLog("rc=" & RcToStr(rc) & " pt=" & PtToStr(pt))
        IF PtInRect(rc,pt.x,pt.y) THEN
          SetCursor hVsplit
          rc=fakeXButtonRect
          SetRect rc,rc.nLeft-2,rc.nTop-2,rc.nRight+4,rc.nBottom+6
          InvalidateRect( hwnd,rc,%TRUE)
          FUNCTION=0:EXIT FUNCTION
        END IF
        rc=gdi1.drc
        lefts=GetLastTLeftVal(gdiI,gdi1.drc.nTop+gdi1.drc.nBottom) '找最左的左边
        allWids=GetLastTRightVal(gdiI,gdi1.drc.nTop+gdi1.drc.nBottom) '找最右的右边，用于画线
        SetRect rc,lefts,rc.nTop+rc.nBottom-5,allWids,rc.nTop+rc.nBottom+2   '下边调整
        'canLog=1 : RunLog("lbuttondown rc=" & RcToStr(rc) & " pt=" & PtToStr(pt))
        'canLog=0
        IF PtInRect(rc,pt.x,pt.y) THEN
          SetCursor hVsplit
          rc=fakeXButtonRect
          SetRect rc,rc.nLeft-2,rc.nTop-2,rc.nRight+4,rc.nBottom+6
          InvalidateRect( hwnd,rc,%TRUE)
          'DIALOG REDRAW hWnd
          FUNCTION=0:EXIT FUNCTION
        END IF
      END IF
      '测试鼠标是否在关闭按钮范围内
      GetClientRect hWnd, rc
      IF gdi1.l="B" OR gdi1.l="F" THEN  '泊坞到下面时，留出一点作为分隔条位置
        rc.nTop=rc.nTop+3
      END IF
      IF gdi1.ml="B" AND gdi1.byside=(gdi1.byside OR %BYTOP) THEN
        rc.nTop=rc.nTop+3
      END IF
      IF gdi1.byside<>(gdi1.byside OR %BYTOP) THEN
        rc.nTop+=3
      END IF
      rc.nLeft+=3
      rc.nRight-=3
      rc.nBottom = rc.nTop+17   '设置测试区为顶端部分
      GetCursorPos pt           '取得鼠标位置（屏幕坐标）
      ScreenToClient hWnd, pt   '转换到客户区坐标
      '--测试是否在假的关闭按钮区域内
      SetRect rc,rc.nRight-10,rc.nTop,rc.nRight-2,rc.nBottom-2
      IF PtInRect(rc,pt.x,pt.y) THEN
        'SendMessage hWnd,%WM_PAINT,0,0
        GetClientRect hWnd, rc
        IF gdi1.l="F" THEN
'          rc.nTop+=3
'          rc.nLeft+=5
          SetRect rc,rc.nRight-8,rc.nTop+11,rc.nRight+6,rc.nTop+25
        ELSE
          IF gdi1.byside<>(gdi1.byside OR %BYTOP) THEN
            SetRect rc,rc.nRight-12,rc.nTop+10,rc.nRight,rc.nTop+22
          ELSE
            IF gdi1.ml="B" THEN
              SetRect rc,rc.nRight-12,rc.nTop+10,rc.nRight,rc.nTop+22
            ELSE
              SetRect rc,rc.nRight-12,rc.nTop+4,rc.nRight,rc.nTop+16
            END IF
          END IF
        END IF
        IF gdi1.byside=(gdi1.byside OR %BYRIGHT) THEN
          rc.nLeft-=1
          rc.nRight-=1
        END IF
        IF gdi1.l<>"F" AND gdi1.byside<>(gdi1.byside OR %BYTOP) AND gdi1.byside<>(gdi1.byside OR %BYRIGHT) THEN
          rc.nLeft-=1
          rc.nRight-=1
        END IF
        'SetRect rc,rc.nLeft,rc.nTop+5,rc.nLeft+13,rc.nTop+15
        DrawXButBorder GetWindowDC(hWnd),rc
      ELSE
        SetRect rc,rc.nLeft-4,rc.nTop-4,rc.nRight+8,rc.nBottom+8
        InvalidateRect( hwnd,rc,%TRUE)
        'dialog redraw hWnd
        'SendMessage hWnd,%wm_paint,0,0
      END IF
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_LBUTTONUP
      '----------------------------------------------------------------
      '拖动窗口，然后鼠标左键弹起
      '----------------------------------------------------------------
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
      gdi1=gdi(gdiI)
      IF bDragging = %TRUE AND hadDrag = %FALSE THEN
        bDragging=%FALSE
        'DIALOG REDRAW hWnd
        ReleaseCapture
        EXIT FUNCTION
      END IF
      IF bDragging = %TRUE  THEN
        DestroyWindow GetParent(hWndDockBtn1)
        hWndDockBtn1=0
        'DIALOG REDRAW hWnd
        ReleaseCapture
        '删除最后的矩形
        IF gdi(gdiI).l="F" THEN
          DrawRectangle  ptDrawStart, ptDrawEnd,3
        ELSE
          DrawRectangle ptDrawStart,ptDrawEnd,1
        END IF
        '我们需要改变它为弹出窗口吗？
        IF gdi(gdiI).l = "F" THEN
          ShowWindow gdi1.hWndD, %SW_HIDE
          ShowWindow gdi1.hWndF, %SW_SHOW
          hMove& = gdi1.hWndF
          gdi(gdiI).hWnd=hMove&
          '移动控件到弹出窗口
          'MoveDockWinContents gdi(gdiI).ID
          SELECT CASE gdi(gdiI).ID
            CASE %ID_PROPERTY
              MovePropContents gdi(gdiI).hWnd
            CASE %ID_CONTROLS
              MoveToolboxContents gdi(gdiI).hWnd
            CASE %ID_PROJECT
              MoveProjectContents gdi(gdiI).hWnd
            CASE %ID_COMPILERS
              MoveCompilersContents gdi(gdiI).hWnd
            CASE %ID_FINDRS
              MoveFindrsContents gdi(gdiI).hWnd
          END SELECT
          SendMessage gdi1.hWndF, %WM_SIZE, 0, 0
          '强制主窗口标题栏为激活状态
          SendMessage g_hWndMain, %WM_NCACTIVATE, %TRUE, 0
        ELSE
          ShowWindow gdi1.hWndF, %SW_HIDE
          ShowWindow gdi1.hWndD, %SW_SHOW
          hMove& = gdi1.hWndD
          gdi(gdiI).hWnd=hMove&
          '移动控件到泊坞子窗口
          'MoveDockWinContents gdi(gdiI).ID
          SELECT CASE gdi(gdiI).ID
            CASE %ID_PROPERTY
              MovePropContents gdi(gdiI).hWnd
            CASE %ID_CONTROLS
              MoveToolboxContents gdi(gdiI).hWnd
            CASE %ID_PROJECT
              MoveProjectContents gdi(gdiI).hWnd
            CASE %ID_COMPILERS
              MoveCompilersContents gdi(gdiI).hWnd
            CASE %ID_FINDRS
              MoveFindrsContents gdi(gdiI).hWnd
          END SELECT
        END IF
        gdi(gdiI).hWnd=hMove&
        rcDrag.nRight  = rcDrag.nRight  -rcDrag.nLeft
        rcDrag.nBottom = rcDrag.nBottom -rcDrag.nTop
        IF gdi1.l <>"F" THEN
          '重绘点位为主窗口的相对点位
          pt.x = rcDrag.nLeft
          pt.y = rcDrag.nTop
          ScreenToClient g_hWndMain, pt
          rcDrag.nLeft = pt.x
          rcDrag.nTop  = pt.y
        END IF
        '定位泊坞窗口到它的新位置
        MoveWin hMove&,rcDrag,%TRUE
        IF gdi(gdiI).l="F" THEN
          gdi(gdiI).frc=rcDrag
        ELSE
          gdi(gdiI).drc=rcDrag
          AdjustForInsert gdiI
        END IF
        '停止拖动操作
        bDragging = %FALSE
        hadDrag = %FALSE
        '在我们再次使用前清除这些变量
        rcDrag.nLeft=0:rcDrag.nTop=0:rcDrag.nRight=0:rcDrag.nBottom=0
        ptDrawEnd.x=0:ptDrawEnd.y=0:ptDrawStart.x=0:ptDrawStart.y=0
        '发送一个调整大小消息给主窗口
        SendMessage g_hWndMain,%WM_SIZE,0,0
        'AdjustRDockArea
        activehWnd = hMove&
        FOR i=0 TO UBOUND(gdi())
          IF gdi(i).l<>"I" AND gdi(i).hWnd>0 AND IsWindow(gdi(i).hWnd) THEN
            DIALOG REDRAW gdi(i).hWnd
            'canLog=1 : RunLog("enddrag redraw " & str$(gdi(i).id)) : canLog=0
          END IF
        NEXT i
      END IF
      '-------------------------------------------------------------
      ' 处理调整大小操作
      '-------------------------------------------------------------
      IF bTracking = %TRUE AND hadTrack = %FALSE THEN
        bTracking = %FALSE
        'DIALOG REDRAW hWnd
        ReleaseCapture
      END IF
      IF bTracking = %TRUE THEN
        bTracking = %FALSE
        hadTrack = %FALSE
        'DIALOG REDRAW hWnd
        ReleaseCapture
        gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
        gdi1=gdi(gdiI)
        IF resizeDA="L" OR resizeDA="R" OR resizeDA="hr"  THEN nDir& = 0
        IF resizeDA="T" OR resizeDA="B" OR resizeDA="hb"  THEN nDir& = 1
        SplitBarInvertLine hWnd, nOldDivX ,nDivTop ,nOldDivX ,nDivBtm, nDir&
        '保存位置及设置最小最大范围
        'msgbox resizeDA
        'GetClientRect g_hWndMain,rc
        GetMainClientRect rc
        SELECT CASE resizeDA
          CASE "L"
            pt.x=nOldDivX : pt.y=0
            ClientToScreen hWnd,pt
            ScreenToClient g_hWndMain,pt
            nOldDivX=pt.x
            gSizer=nOldDivX
            lARc.nRight=gSizer+1
          CASE "R"
            pt.x=nOldDivX : pt.y=0
            ClientToScreen hWnd,pt
            ScreenToClient g_hWndMain,pt
            nOldDivX=pt.x
            gSizer=rc.nRight-nOldDivX
            rARc.nLeft=nOldDivX
            rARc.nRight=gSizer
          CASE "T"
            pt.x=0 : pt.y=nOldDivX
            ClientToScreen hWnd,pt
            ScreenToClient g_hWndMain,pt
            nOldDivX=pt.y
            gSizer=nOldDivX
            GetClientRect g_hReBar,rc
            tARc.nBottom=gSizer+1-rc.nbottom
          CASE "B"
            pt.x=0 : pt.y=nOldDivX
            ClientToScreen hWnd,pt
            ScreenToClient g_hWndMain,pt
            nOldDivX=pt.y
            gSizer=rc.nTop+rc.nBottom-nOldDivX
            'GetClientRect g_hStatus,tmpRc
            bARc.nTop=rc.nTop+rc.nBottom-gSizer '+tmpRc.nBottom
            bARc.nBottom=gSizer '-tmpRc.nBottom
          CASE "hr"
            pt.x=nOldDivX : pt.y=0
            ClientToScreen hWnd,pt
            ScreenToClient g_hWndMain,pt
            nOldDivX=pt.x
            gSizer=nOldDivX+1-gdi(gdiI).drc.nLeft '新宽度
            gdi(gdiI).drc.nRight=gSizer
            AdjustRWin gdiI
          CASE "hb"   '不靠下边，则下方可调
            pt.x=0  : pt.y=nOldDivX
            ClientToScreen hWnd,pt
            ScreenToClient g_hWndMain,pt
            nOldDivX=pt.y
            gSizer=nOldDivX+1-gdi(gdiI).drc.nTop '新高度
            gdi(gdiI).drc.nBottom=gSizer
            AdjustBWin gdiI
        END SELECT
        '设置泊坞子窗口到新的尺寸
        IF resizeDA="R" THEN
          AdjustRDockArea
        ELSEIF resizeDA="B" THEN
          AdjustBDockArea
        END IF
        SendMessage g_hWndMain, %WM_SIZE, 0, 0
        'END IF
        resizeDA=""
        activehWnd = hWnd
        FOR i=0 TO UBOUND(gdi())
          IF gdi(i).l<>"I" AND gdi(i).hWnd>0 AND IsWindow(gdi(i).hWnd) THEN
            DIALOG REDRAW gdi(i).hWnd
          END IF
        NEXT i
      END IF
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_PAINT
      gdiI=GetWindowLong(hWnd,%GWL_USERDATA)
      gdi1=gdi(gdiI)
      hDC = BeginPaint( hWnd, ps)
      GetClientRect hWnd, rc
      Lb.lbStyle = %BS_SOLID
      'canLog=1 : RunLog("paint activehWnd=" & STR$(activehWnd) & " hWnd=" & str$(hWnd))
      'canLog=0
      IF hWnd=activehWnd THEN
        Lb.lbColor = GetSysColor(%COLOR_ACTIVECAPTION)
        SetTextColor hDC, GetSysColor(%COLOR_CAPTIONTEXT)
      ELSE
        Lb.lbColor = GetSysColor(%COLOR_INACTIVECAPTION) '%GRAY '
        SetTextColor hDC, GetSysColor(%COLOR_INACTIVECAPTIONTEXT)
      END IF
      rc.nBottom = rc.nTop + 19
      hBrush& = CreateBrushIndirect(lb)
      SELECT CASE gdi1.l
        CASE "F"
          '画标题栏
          FillRect hDC, rc, hBrush&
          DeleteObject hBrush&
          '输出标题文本
          SetBkColor (hDC,Lb.lbColor)
          SelectObject(hDC,ghFont)
          rc.nLeft=rc.nLeft+4
          rc.nRight-=4
          GetWindowText hWnd,tmpAsc,18
          tmpStr=tmpAsc
          DrawText hDC,tmpAsc,LEN(tmpStr),rc,%DT_LEFT OR %DT_VCENTER
          '画关闭按钮
          nRight& = rc.nRight - 10
          SetRect  rc,nRight&,5,nRight&+9,14
          DrawXBut hDC,rc
          fakeXButtonRect = rc
        CASE "L"
          '画上边的假标题栏
          rc.nLeft=rc.nLeft+2
          rc.nRight-=1
          IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN '靠右，宽度左让一个
            rc.nRight=rc.nRight-1
          END IF
          FillRect hDC, rc, hBrush&
          DeleteObject hBrush&
          '输出标题文本
          SetBkColor (hDC,Lb.lbColor)
          SelectObject(hDC,ghFont)
          rc.nLeft=rc.nLeft+4
          rc.nTop+=2
          GetWindowText hWnd,tmpAsc,18
          tmpStr=tmpAsc
          DrawText hDC,tmpAsc,LEN(tmpStr),rc,%DT_LEFT OR %DT_VCENTER
          '画假的关闭按钮
          nRight& = rc.nRight - 10
          SetRect  rc,nRight&,5,nRight&+9,14
          DrawXBut hDC,rc
          fakeXButtonRect = rc
          '画右边的可调区为高亮
          GetClientRect  hWnd, rc
          SetRect  rc,rc.nRight-2,rc.nTop,rc.nRight-1,rc.nBottom
          IF gdi(gdiI).ml="L" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN '左泊，在左区，靠右边时，画右侧可调线
            DrawResizeBorder hDC,rc
          END IF
        CASE "R"
          '画上边的假标题栏
          IF (gdi1.ml="B" AND gdi1.byside=(gdi1.byside OR %BYTOP)) OR gdi1.byside<>(gdi1.byside OR %BYTOP) THEN '(在下区，靠上边，)或不靠上边，标题要向下让6个
            rc.nTop=rc.nTop+6
            rc.nBottom = rc.nTop + 19
          END IF
          rc.nLeft=rc.nLeft+4
          rc.nRight-=1
          IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN '靠右边，则宽度左让一个
            rc.nRight=rc.nRight-1
          END IF
          FillRect hDC, rc, hBrush&
          DeleteObject hBrush&
          '输出标题文本
          SetBkColor (hDC,Lb.lbColor)
          SelectObject(hDC,ghFont)
          rc.nLeft=rc.nLeft+6
          rc.nTop+=2
          GetWindowText hWnd,tmpAsc,18
          tmpStr=tmpAsc
          DrawText hDC,tmpAsc,LEN(tmpStr),rc,%DT_LEFT OR %DT_VCENTER
          '画假的关闭按钮
          nRight& = rc.nRight - 10
          SetRect  rc,nRight&,5,nRight&+9,14
          IF (gdi1.ml="B" AND gdi1.byside=(gdi1.byside OR %BYTOP)) OR gdi1.byside<>(gdi1.byside OR %BYTOP) THEN '(在下区，靠上边，)或不靠上边，则上下位置加6
            SetRect  rc,nRight&,5+6,nRight&+9,14+6
          END IF
          DrawXBut hDC,rc
          fakeXButtonRect = rc
          '画左边的可调区为高亮
          GetClientRect  hWnd, rc
          SetRect  rc,rc.nLeft+1,rc.nTop+1,rc.nLeft+2,rc.nBottom+1
          IF gdi(gdiI).ml="T" THEN
            rc.nBottom-=2
          END IF
          IF gdi(gdiI).ml="B" THEN
            rc.nTop+=2
          END IF
          IF gdi(gdiI).byside<>(gdi(gdiI).byside OR %BYBOTTOM) THEN
            rc.nBottom+=2
          END IF
          DrawResizeBorder hDC,rc '画左侧可调线
          IF gdi(gdiI).byside<>(gdi(gdiI).byside OR %BYTOP) THEN '不靠上,画上面可调线
            GetClientRect hWnd,rc
            IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN
              rc.nRight -= 1
            END IF
            SetRect rc,rc.nLeft,rc.nTop+1,rc.nRight,rc.nTop+2
            DrawResizeBorder hDC,rc,1
          END IF
          IF gdi(gdiI).ml="L" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN '左区，靠右时，画右侧可调线
            GetClientRect  hWnd, rc
            SetRect  rc,rc.nLeft+rc.nRight-2,rc.nTop,rc.nLeft+rc.nRight-1,rc.nBottom
            DrawResizeBorder hDC,rc
          END IF
          IF gdi(gdiI).ml="T" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYBOTTOM) THEN '上区，靠下时，画下面可调线
            GetClientRect hWnd,rc
            SetRect rc,rc.nLeft,rc.nBottom-2,rc.nRight,rc.nBottom-1
            DrawResizeBorder hDC,rc,1
          END IF
          IF gdi(gdiI).ml="B" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYTOP) THEN '下区，靠上时，画上面可调线
            GetClientRect hWnd,rc
            SetRect rc,rc.nLeft,rc.nTop+1,rc.nRight,rc.nTop+2
            DrawResizeBorder hDC,rc,1
          END IF
        CASE "T"
          '画上边的假标题栏
          rc.nLeft=rc.nLeft+2
          rc.nRight-=1
          IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN '靠右边，宽度左让一个
            rc.nRight=rc.nRight-1
          END IF
          FillRect hDC, rc, hBrush&
          DeleteObject hBrush&
          '输出标题文本
          SetBkColor (hDC,Lb.lbColor)
          SelectObject(hDC,ghFont)
          rc.nLeft=rc.nLeft+4
          rc.nTop+=2
          GetWindowText hWnd,tmpAsc,18
          tmpStr=tmpAsc
          DrawText hDC,tmpAsc,LEN(tmpStr),rc,%DT_LEFT OR %DT_VCENTER
          '画假的关闭按钮
          nRight& = rc.nRight - 10
          SetRect  rc,nRight&,5,nRight&+9,14
          DrawXBut hDC,rc
          fakeXButtonRect = rc
          '画下边的可调区为高亮
          GetClientRect  hWnd, rc
          SetRect rc,rc.nLeft,rc.nBottom-2,rc.nRight,rc.nBottom-1
          IF gdi(gdiI).ml="T" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYBOTTOM) THEN '上泊，在上区，靠下边，画下面线
            DrawResizeBorder hDC,rc,1
          END IF
        CASE "B"
          '画上边的假标题栏
          rc.nTop=rc.nTop+6
          rc.nBottom = rc.nTop + 19
          rc.nLeft=rc.nLeft+2
          rc.nRight-=1
          IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN
            rc.nRight=rc.nRight-1
          END IF
          IF gdi(gdiI).ml="R" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYLEFT) THEN
            rc.nLeft+=2
            IF gdi(gdiI).byside<>(gdi(gdiI).byside OR %BYRIGHT) THEN
              rc.nRight-=1
            END IF
          END IF
          FillRect hDC, rc, hBrush&
          DeleteObject hBrush&
          '输出标题文本
          SetBkColor (hDC,Lb.lbColor)
          SelectObject(hDC,ghFont)
          rc.nLeft=rc.nLeft+4
          rc.nTop+=2
          GetWindowText hWnd,tmpAsc,18
          tmpStr=tmpAsc
          DrawText hDC,tmpAsc,LEN(tmpStr),rc,%DT_LEFT OR %DT_VCENTER
          '画假的关闭按钮
          nRight& = rc.nRight - 10
          SetRect  rc,nRight&,11,nRight&+9,20
          DrawXBut hDC,rc
          fakeXButtonRect = rc
          '画上边的可调区为高亮
          GetClientRect  hWnd, rc
          SetRect rc,rc.nLeft-1,rc.nTop+1,rc.nRight+1,rc.nTop+2
          IF gdi(gdiI).ml="L" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN
            rc.nRight-=2
          END IF
          IF gdi(gdiI).ml="R" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYLEFT) THEN
            rc.nLeft+=2
          END IF
          IF gdi(gdiI).byside<>(gdi(gdiI).byside OR %BYLEFT) THEN
            rc.nLeft+=2
          END IF
          DrawResizeBorder hDC,rc,1 '1表示是横线
          IF gdi(gdiI).byside<>(gdi(gdiI).byside OR %BYLEFT) THEN  '不靠左时，画左侧可调线
            GetClientRect  hWnd, rc
            SetRect rc,rc.nLeft+1,rc.nTop,rc.nLeft+2,rc.nBottom
            IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYBOTTOM) THEN
              rc.nBottom-=2
            END IF
            DrawResizeBorder hDC,rc
          END IF
          IF gdi(gdiI).ml="T" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYBOTTOM) THEN '在上区，且靠下，则画下面可调线
            GetClientRect  hWnd, rc
            SetRect rc,rc.nLeft-1,rc.nBottom-2,rc.nRight,rc.nBottom-1
            DrawResizeBorder hDC,rc,1
          END IF
          IF gdi(gdiI).ml="L" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN '在左区，且靠右，则画右侧可调线
            GetClientRect  hWnd, rc
            SetRect rc,rc.nRight-2,rc.nTop,rc.nRight-1,rc.nBottom
            DrawResizeBorder hDC,rc
          END IF
          IF gdi(gdiI).ml="R" AND gdi(gdiI).byside=(gdi(gdiI).byside OR %BYLEFT) THEN '在右区，且靠左，则画左侧可调线
            GetClientRect  hWnd, rc
            SetRect rc,rc.nLeft+1,rc.nTop,rc.nLeft+2,rc.nBottom
            DrawResizeBorder hDC,rc
          END IF
      END SELECT
      EndPaint hWnd, ps
      FUNCTION = 0 : EXIT FUNCTION
    CASE %WM_DESTROY
  END SELECT
  FUNCTION = DefWindowProc( hWnd, wMsg, wParam, lParam )
END FUNCTION
FUNCTION GetActiveDocker()AS DWORD
  LOCAL hWndTest AS DWORD
  LOCAL hWndParent AS DWORD
  LOCAL rs          AS DWORD
  LOCAL tmpLng      AS LONG
  LOCAL i           AS LONG
  hWndTest=GetFocus()
  WHILE hWndTest>0
    hWndParent=GetParent(hWndTest)
    IF hWndParent=hWndTest THEN
      EXIT LOOP
    END IF
    hWndTest=hWndParent
    tmpLng=GetIndex(hWndTest)
    IF tmpLng>=0 THEN
      EXIT LOOP
    END IF
  WEND
  IF gdi(tmpLng).l="I" THEN
    rs=-1
  ELSE
    IF gdi(tmpLng).l="F" THEN
      rs=gdi(tmpLng).hWndF
    ELSE
      rs=gdi(tmpLng).hWndD
    END IF
  END IF
  FUNCTION=rs
END FUNCTION
' 绘制可调区示意边,用到左上下三个值,dir为1表示画横线，默认是画竖线
FUNCTION DrawResizeBorder(BYVAL hDC AS DWORD,BYVAL rc AS RECT,OPT dir AS INTEGER)AS LONG
  LOCAL hBrush&
  hBrush& = CreatePen(%PS_SOLID,1,%GRAY)
  SelectObject hDC,hBrush&
  MoveToEx hDC,rc.nLeft,rc.nTop,BYVAL 0
  IF ISFALSE(ISMISSING(dir)) AND dir=1 THEN
    LineTo hDC,rc.nRight,rc.nTop
    DeleteObject hBrush&
    EXIT FUNCTION
  END IF
  LineTo hDC,rc.nLeft,rc.nBottom
  DeleteObject hBrush&
END FUNCTION
' 检测鼠标放在哪个泊坞窗口内，返回泊坞方向
FUNCTION HitTestMain1( BYVAL hWnd AS LONG, pt AS POINTAPI ) AS LONG
  LOCAL rc    AS RECT  '主窗口矩形
  LOCAL rc1   AS RECT
  LOCAL rc2   AS RECT
  LOCAL rc3   AS RECT
  LOCAL rc4   AS RECT
  LOCAL rc5   AS RECT
  LOCAL rc6   AS RECT
  LOCAL rc7   AS RECT
  LOCAL rc8   AS RECT
  LOCAL nZone AS LONG
  LOCAL rs    AS LONG
  '定义点测试区域
  GetWindowRect hWnd,rc
  GetWindowRect hWndDockBtn1,rc1   '上 desthWnd
  GetWindowRect hWndDockBtn2,rc2   '下
  GetWindowRect hWndDockBtn3,rc3   '左
  GetWindowRect hWndDockBtn4,rc4   '右
  GetWindowRect hWndDockBtn5,rc5   '左 g_hWndMain
  GetWindowRect hWndDockBtn6,rc6   '右
  GetWindowRect hWndDockBtn7,rc7   '上
  GetWindowRect hWndDockBtn8,rc8   '下
  rs = 0
  IF PtInRect(rc1,pt.x,pt.y) THEN
    CALL SendMessage(hWndDockBtn1,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_HTOP"))
    rs = %DTOP
  ELSEIF PtInRect(rc7,pt.x,pt.y) THEN
    CALL SendMessage(hWndDockBtn7,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_HTOP"))
    rs = %MTOP
  ELSEIF PtInRect(rc2,pt.x,pt.y) THEN
    CALL SendMessage(hWndDockBtn2,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_HBOTTOM"))
    rs = %DBOTTOM
  ELSEIF PtInRect(rc8,pt.x,pt.y) THEN
    CALL SendMessage(hWndDockBtn8,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_HBOTTOM"))
    rs = %MBOTTOM
  ELSEIF PtInRect(rc3,pt.x,pt.y) THEN
    CALL SendMessage(hWndDockBtn3,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_HLEFT"))
    rs = %DLEFT
  ELSEIF PtInRect(rc5,pt.x,pt.y) THEN
    CALL SendMessage(hWndDockBtn5,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_HLEFT"))
    rs = %MLEFT
  ELSEIF PtInRect(rc4,pt.x,pt.y) THEN
    CALL SendMessage(hWndDockBtn4,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_HRIGHT"))
    rs = %DRIGHT
  ELSEIF PtInRect(rc6,pt.x,pt.y) THEN
    CALL SendMessage(hWndDockBtn6,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_HRIGHT"))
    rs = %MRIGHT
  ELSE
    CALL SendMessage(hWndDockBtn1,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_TOP"))
    CALL SendMessage(hWndDockBtn2,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_BOTTOM"))
    CALL SendMessage(hWndDockBtn3,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_LEFT"))
    CALL SendMessage(hWndDockBtn4,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_RIGHT"))
    CALL SendMessage(hWndDockBtn5,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_LEFT"))
    CALL SendMessage(hWndDockBtn6,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_RIGHT"))
    CALL SendMessage(hWndDockBtn7,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_TOP"))
    CALL SendMessage(hWndDockBtn8,%STM_SETIMAGE,%IMAGE_BITMAP,LoadBitmap(g_hInst,"BMP_BOTTOM"))
    FUNCTION=0
  END IF
  IF g_hWndMain=hWnd AND rs>4 THEN
    FUNCTION=rs-4
  ELSE
    FUNCTION=rs
  END IF
END FUNCTION
'------------------------------------------------------------
' 画拖动、泊坞示意矩形
'------------------------------------------------------------
SUB DrawRectangle(ptStart AS POINTAPI,ptEnd AS POINTAPI,BYVAL penwidth AS INTEGER)
  LOCAL hPen      AS LONG
  LOCAL hOldPen   AS LONG
  LOCAL hBrush    AS LONG
  LOCAL hOldBrush AS LONG
  LOCAL hDC       AS LONG
  LOCAL nDrawMode AS LONG
  LOCAL Lb AS LOGBRUSH
  LOCAL rc        AS RECT
  '得到整个屏幕的DC,用来绘制所有控件到顶层
  '(稍后,需要增加到我们的MDI应用程序,因此需要绘制整个应用程序)
  hDC = GetDC(%HWND_DESKTOP)
  hPen = CreatePen(%PS_SOLID , penWidth, 0)
  hOldPen   = SelectObject(hDC, hPen)
  hBrush    = GetStockObject(%NULL_BRUSH)
  hOldBrush = SelectObject(hDC, hBrush)
  nDrawMode = SetROP2(hDC, %R2_NOT )
  '画矩形
  CALL Rectangle(hDC, ptStart.x, ptStart.y, ptEnd.x, ptEnd.y)
  '重载DC
  CALL SetROP2(hDC, nDrawMode)
  CALL SelectObject(hDC, hOldBrush)
  CALL SelectObject(hDC, hOldPen)
  CALL ReleaseDC(%HWND_DESKTOP, hDC)
  '删除GDI对象
  CALL DeleteObject(hPen)
END SUB
' 画关闭按钮外框
FUNCTION DrawXButBorder(BYVAL hDC AS LONG,BYVAL rc AS RECT)AS LONG
  LOCAL hBrush  AS LONG
  LOCAL hPen    AS LONG
  LOCAL hWnd    AS DWORD
  hWnd=WindowFromDC(hDC)
  IF hWnd=activehWnd THEN
    hPen = CreatePen(%PS_SOLID,1,GetSysColor(%COLOR_CAPTIONTEXT))
  ELSE
    hPen = CreatePen(%PS_SOLID,1,GetSysColor(%COLOR_INACTIVECAPTIONTEXT))
  END IF
  'hPen = CreatePen(%PS_SOLID,1,%BLACK)
  SelectObject hDC,hPen
  hBrush    = GetStockObject(%NULL_BRUSH)
  SelectObject(hDC, hBrush)
  Rectangle(hDC, rc.nLeft, rc.nTop, rc.nRight, rc.nBottom)
  DeleteObject  hBrush
  DeleteObject  hPen
  'ReleaseDC(%HWND_DESKTOP, hDC)
END FUNCTION
' 画关闭按钮
FUNCTION DrawXBut(BYVAL hDC AS LONG,BYVAL rc AS RECT)AS LONG
  LOCAL hBrush  AS LONG
  LOCAL hPen    AS LONG
  LOCAL pt      AS POINTAPI
  LOCAL hWnd    AS DWORD
  LOCAL testRc  AS RECT
  LOCAL i       AS LONG
  hWnd=WindowFromDC(hDC)
  IF hWnd=activehWnd THEN
    hBrush = CreatePen(%PS_SOLID,1,GetSysColor(%COLOR_CAPTIONTEXT))
  ELSE
    hBrush = CreatePen(%PS_SOLID,1,GetSysColor(%COLOR_INACTIVECAPTIONTEXT))
  END IF

  SelectObject hDC,hBrush&
  MoveToEx hDC,rc.nLeft,rc.nTop,BYVAL 0
  LineTo hDC,rc.nLeft+9,rc.nTop+9
  MoveToEx hDC,rc.nLeft+9,rc.nTop,BYVAL 0
  LineTo hDC,rc.nLeft,rc.nTop+9
  DeleteObject  hBrush
'  '测试鼠标是否在关闭按钮范围内
'  hWnd=WindowFromDC(hDC)
'  i=GetIndex(hWnd)
'  GetClientRect hWnd, rc
'  IF gdi(i).l="B" or gdi(i).l="F" THEN  '泊坞到下面时，留出一点作为分隔条位置
'    rc.nTop=rc.nTop+3
'  END IF
'  IF gdi(i).ml="B" AND gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
'    rc.nTop=rc.nTop+3
'  END IF
'  'rc.nLeft+=3
'  'rc.nRight-=3
'  rc.nBottom = rc.nTop+17   '设置测试区为顶端部分
'  GetCursorPos pt           '取得鼠标位置（屏幕坐标）
'  ScreenToClient hWnd, pt   '转换到客户区坐标
'  '--测试是否在假的关闭按钮区域内
'  SetRect rc,rc.nRight-13,rc.nTop,rc.nRight-1,rc.nBottom
'  IF PtInRect(rc,pt.x,pt.y) THEN
'    'DrawXButBorder GetWindowDC(hWnd),rc
'    hPen = CreatePen(%PS_SOLID,1,%BLACK)
'    SelectObject hDC,hPen
'    hBrush    = GetStockObject(%NULL_BRUSH)
'    SelectObject(hDC, hBrush)
'    Rectangle(hDC, rc.nLeft, rc.nTop, rc.nRight, rc.nBottom)
'    DeleteObject  hBrush
'    DeleteObject  hPen
'  END IF
  'ReleaseDC(%HWND_DESKTOP, hDC)
END FUNCTION
'-----------------------------------------------------------------------------
' 在鼠标拖动分隔条操作期间,绘制分隔条
' nDir=> 绘制水平=0 or 垂直=1
'-----------------------------------------------------------------------------
SUB SplitBarInvertLine( BYVAL hWnd AS LONG,BYVAL xleft1 AS INTEGER, _
                        BYVAL yFrom AS INTEGER,BYVAL xleft2 AS INTEGER , _
                        BYVAL yTo AS INTEGER, BYVAL nDir AS LONG )
  LOCAL hDC       AS LONG
  LOCAL nDrawMode AS LONG
  LOCAL rc        AS RECT
  LOCAL hPen      AS LONG
  LOCAL pt        AS POINTAPI
  LOCAL tmprc     AS RECT
  LOCAL memDC     AS LONG
  LOCAL hbitmap   AS LONG
  LOCAL hBrush    AS LONG
  LOCAL hBrush2   AS LONG
  LOCAL oldBrush  AS LONG
  LOCAL oldBmp    AS LONG
  LOCAL rcb       AS RECT
  LOCAL i&,j&
  '得到桌面的DC
  hDC = GetDC( %HWND_DESKTOP )
'  '得到分隔条的窗口扩展,并重测绘它
  GetWindowRect hWnd,rc
  IF nDir = 0 THEN
    SetRect tmprc,0,yFrom,0,yTo
    ClientToScreenRect g_hWndMain,tmprc
    yFrom=tmprc.nTop+2
    yTo=tmprc.nBottom-tmprc.nTop-4
    xLeft1 = rc.nLeft +xLeft1
    xleft2 = rc.nLeft +xLeft2
  ELSEIF nDir = 1 THEN
    SetRect tmprc,yFrom,0,yTo,0
    ClientToScreenRect g_hWndMain,tmprc
    yFrom=tmprc.nLeft+2
    yTo=tmprc.nRight-tmprc.nLeft-4
    xLeft1 = rc.nTop +xLeft1
    xleft2 = rc.nTop +xLeft2
  END IF
  '在内存中创建位图
  memDC    = CreateCompatibleDC(hdc)
  hBitmap  = CreateCompatibleBitmap(hDC, 8, 8)
  oldBmp   = SelectObject( memDC, hBitmap )
  rcb.ntop = 0 :rcb.nleft = 0 :rcb.nright = 8 :rcb.nbottom = 8
  hBrush   = GetStockObject( %BLACK_BRUSH )
  FillRect  memDC,rcb,hBrush
  '支持网蓝图案(0x55 & 0xAA)
  FOR j& = 0 TO 7 STEP 1
    IF j&=0 OR j&=2 OR j&=4 OR j&=6 THEN
      FOR i& = 0 TO 7 STEP 2
        SetPixel memDC,i&,j&,RGB(255,255,255)
      NEXT
    ELSE
      FOR i& = 1 TO 7 STEP 2
        SetPixel memDC,i&,j&,RGB(255,255,255)
      NEXT
    END IF
  NEXT
  '创建图案刷
  hBrush2 = CreatePatternBrush( hBitmap )
  oldBrush = SelectObject(hDC,hBrush2)
  LOCAL rcc AS RECT
  GetClientRect hWnd, rcc
  IF nDir = 0 THEN
    PatBlt hDC, xLeft1, yFrom,5 , yTo , %PATINVERT
  ELSEIF nDir =1 THEN
    PatBlt hDc, yFrom, xLeft1, yTo, 5, %PATINVERT
  END IF
  '清除...
  SelectObject  memDC, oldbmp
  DeleteObject  hBitmap
  DeleteDC      memDC
  SelectObject  hDC, oldBrush
  DeleteObject  hBrush2
  SetROP2       hDC, nDrawMode
  ReleaseDC     hWnd, hDC
END SUB
' 根据周边情况，设置泊坞窗口的泊坞结构中的byside值
FUNCTION SetGdiByside(BYVAL gdiI AS LONG)AS LONG
  LOCAL i,j AS LONG
  LOCAL rc AS RECT
  LOCAL bydock AS BYTE
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  IF gdi(gdiI).l="F" OR gdi(gdiI).l="I" THEN '浮动或隐藏时，设置byside为0
    gdi(gdiI).ml=gdi(i).l
    gdi(gdiI).byside=0
    FUNCTION=1
    EXIT FUNCTION
  END IF
  gdi(gdiI).byside=15 '初始值
  '检查它泊坞到的情况
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).ID=gdi(gdiI).tid AND i<>gdiI THEN
      IF gdi(i).tid=%ID_CLIENTWINDOW THEN
        'gdi(gdiI).l=gdi(gdiI).ml
        gdi(gdiI).byside=gdi(i).byside
        IF gdi(gdiI).l="R"  THEN
          gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYLEFT)
          IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
            gdi(i).byside=gdi(i).byside AND (NOT %BYRIGHT)
          END IF
        END IF
        IF gdi(gdiI).l="B" THEN
          gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYTOP)
          IF gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
            gdi(i).byside=gdi(i).byside AND (NOT %BYBOTTOM)
          END IF
        END IF
      ELSEIF gdi(i).l="F" OR gdi(i).l="I" THEN '它泊坞到的窗口已经浮动或隐藏
        gdi(gdiI).tid=gdi(i).tid
        IF gdi(i).tid=%ID_CLIENTWINDOW OR gdi(i).tid=0 THEN
          EXIT FOR
        END IF
        gdi(gdiI).l=gdi(i).l
        FOR j=0 TO UBOUND(gdi())
          IF gdi(j).ID=gdi(i).tid AND gdi(j).l<>"F" AND gdi(j).l<>"I" THEN
            gdi(gdiI).byside=gdi(j).byside
            IF gdi(gdiI).l="R"  THEN
              gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYLEFT)
              IF gdi(j).byside=(gdi(j).byside OR %BYRIGHT) THEN
                gdi(j).byside=gdi(j).byside AND (NOT %BYRIGHT)
              END IF
            END IF
            IF gdi(gdiI).l="B" THEN
              gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYTOP)
              IF gdi(j).byside=(gdi(j).byside OR %BYBOTTOM) THEN
                gdi(j).byside=gdi(j).byside AND (NOT %BYBOTTOM)
              END IF
            END IF
            EXIT FOR
          END IF
        NEXT j
      ELSE
        gdi(gdiI).byside=gdi(i).byside
'        IF gdi(gdiI).l="L" AND gdi(i).byside=(gdi(i).byside OR %BYLEFT)   THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYLEFT)
'        IF gdi(gdiI).l="T" AND gdi(i).byside=(gdi(i).byside OR %BYTOP)    THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYTOP)
'        IF gdi(gdiI).l="R" AND gdi(i).byside=(gdi(i).byside OR %BYRIGHT)  THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYRIGHT)
'        IF gdi(gdiI).l="B" AND gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYBOTTOM)
        IF gdi(gdiI).l="R"  THEN
          gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYLEFT)
          IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
            gdi(i).byside=gdi(i).byside AND (NOT %BYRIGHT)
          END IF
        END IF
        IF gdi(gdiI).l="B" THEN
          gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYTOP)
          IF gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
            gdi(i).byside=gdi(i).byside AND (NOT %BYBOTTOM)
          END IF
        END IF
      END IF
    END IF
  NEXT i
  '泊坞到了客户区，且是泊坞状态，判定是否是独占一个泊坞区(即，没有被泊)
  bydock=0
  IF gdi(gdiI).tid=%ID_CLIENTWINDOW AND gdi(gdiI).l<>"F" AND gdi(gdiI).l<>"I" THEN
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).tid=gdi(gdiI).ID AND gdi(i).l<>"F" AND gdi(i).l<>"I" THEN
        bydock=1
        EXIT FOR
      END IF
    NEXT i
    IF bydock=0 THEN
      gdi(gdiI).byside=15 '即独占
    ELSE
      FOR i=0 TO UBOUND(gdi())
        IF gdi(i).tid=gdi(gdiI).ID AND gdi(i).l<>"F" AND gdi(i).l<>"I" THEN '被泊坞
          IF gdi(i).l="L" THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYLEFT)
          IF gdi(i).l="T" THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYTOP)
          IF gdi(i).l="R" THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYRIGHT)
          IF gdi(i).l="B" THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYBOTTOM)
        END IF
      NEXT i
    END IF
    FUNCTION=1
    EXIT FUNCTION
  END IF
  '被泊坞了
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).tid=gdi(gdiI).ID AND gdi(i).l<>"F" AND gdi(i).l<>"I" THEN '被泊坞
      IF gdi(i).l="L" THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYLEFT)
      IF gdi(i).l="T" THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYTOP)
      IF gdi(i).l="R" THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYRIGHT)
      IF gdi(i).l="B" THEN gdi(gdiI).byside=gdi(gdiI).byside AND (NOT %BYBOTTOM)
    END IF
  NEXT i
  FUNCTION=1
END FUNCTION
'通过id查dock结构数组索引值
FUNCTION FindGdiIndex(BYVAL wID AS DWORD)AS LONG
  LOCAL i AS LONG
  IF UBOUND(gdi())<0 THEN
    FUNCTION=-1
  END IF
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).ID=wID THEN
      FUNCTION=i
      EXIT FUNCTION
    END IF
  NEXT i
  FUNCTION=-1
END FUNCTION
' *********************************************************************************************
' 隐藏窗口
' *********************************************************************************************
FUNCTION DockWinHide(BYVAL gdiI AS LONG) AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    EXIT FUNCTION
  END IF
  IF gdi(i).l = "F" THEN
    gdi(i).ml="I"
    gdi(i).l="I"
    gdi(i).hWnd=gdi(i).hWndF
    GetWindowRect gdi(i).hWndF,rc
    rc.nRight-=rc.nLeft
    rc.nBottom-=rc.nTop
    gdi(i).frc=rc         '浮动窗口的位置是相对于屏幕的，存储左上宽高
    ShowWindow gdi(i).hWndF, %SW_HIDE
    SELECT CASE gdi(gdiI).ID
      CASE %ID_PROPERTY
        MENU SET STATE g_hMenuView,BYCMD %IDM_PROPERTY,%MF_UNCHECKED OR %MF_ENABLED
      CASE %ID_CONTROLS
        MENU SET STATE g_hMenuView,BYCMD %IDM_CONTROLS,%MF_UNCHECKED OR %MF_ENABLED
      CASE %ID_PROJECT
        MENU SET STATE g_hMenuView,BYCMD %IDM_PROJECT,%MF_UNCHECKED OR %MF_ENABLED
      CASE %ID_COMPILERS
        MENU SET STATE g_hMenuView,BYCMD %IDM_COMPILERS,%MF_UNCHECKED OR %MF_ENABLED
      CASE %ID_FINDRS
        MENU SET STATE g_hMenuView,BYCMD %IDM_FINDRS,%MF_UNCHECKED OR %MF_ENABLED
    END SELECT
    EXIT FUNCTION
  END IF
  ShowWindow gdi(i).hWndD, %SW_HIDE
  AdjustForMove i
  gdi(i).ml="I"
  gdi(i).l="I"
  gdi(i).hWnd=gdi(i).hWndF
  gdi(i).tid=0
  gdi(i).byside=0
  SELECT CASE gdi(gdiI).ID
    CASE %ID_PROPERTY
      MENU SET STATE g_hMenuView,BYCMD %IDM_PROPERTY,%MF_UNCHECKED OR %MF_ENABLED
    CASE %ID_CONTROLS
      MENU SET STATE g_hMenuView,BYCMD %IDM_CONTROLS,%MF_UNCHECKED OR %MF_ENABLED
    CASE %ID_PROJECT
      MENU SET STATE g_hMenuView,BYCMD %IDM_PROJECT,%MF_UNCHECKED OR %MF_ENABLED
    CASE %ID_COMPILERS
      MENU SET STATE g_hMenuView,BYCMD %IDM_COMPILERS,%MF_UNCHECKED OR %MF_ENABLED
    CASE %ID_FINDRS
      MENU SET STATE g_hMenuView,BYCMD %IDM_FINDRS,%MF_UNCHECKED OR %MF_ENABLED
  END SELECT
END FUNCTION
''获得指定窗口关联的最右边的泊坞窗口gdi数组索引
FUNCTION GetLastRight(BYVAL gdiI AS LONG)AS LONG
  LOCAL i AS LONG
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).tid=gdi(gdiI).ID AND gdi(i).l="R" THEN
      FUNCTION=GetLastRight(i)
      EXIT FUNCTION
    END IF
  NEXT i
  FUNCTION=gdiI
END FUNCTION
''获得指定窗口关联的最下边的泊坞窗口gdi数组索引
FUNCTION GetLastBottom(BYVAL gdiI AS LONG)AS LONG
  LOCAL i AS LONG
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).tid=gdi(gdiI).ID AND gdi(i).l="B" THEN
      FUNCTION=GetLastBottom(i)
      EXIT FUNCTION
    END IF
  NEXT i
  FUNCTION=gdiI
END FUNCTION
' *********************************************************************************************
' 显示窗口 仅以浮动形式显示
' *********************************************************************************************
FUNCTION DockWinShow(BYVAL gdiI AS LONG) AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  i= gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    EXIT FUNCTION
  END IF
  ' I表示隐藏状态，如果不是隐藏状态，即显示状态，则退出（原来就是显示的，不需要再显示）
  IF gdi(i).l <> "I" THEN EXIT FUNCTION
  IF ISFALSE gdi(i).hWndD OR ISFALSE gdi(i).hWndF THEN EXIT FUNCTION '窗口句柄异常，则退出
  ShowWindow gdi(i).hWnd,%SW_SHOW
  'GetClientRect g_hWndMain,rc
  GetMainClientRect rc
  ClientToScreenRect g_hWndMain,rc
  IF gdi(i).frc.nLeft<rc.nLeft THEN
    gdi(i).frc.nLeft=rc.nLeft
    MoveWin gdi(i).hWnd,gdi(i).frc,%TRUE
  END IF
  IF gdi(i).frc.nTop<rc.nTop THEN
    gdi(i).frc.nTop=rc.nTop
    MoveWin gdi(i).hWnd,gdi(i).frc,%TRUE
  END IF
  gdi(i).hWnd=gdi(i).hWndF
  'MoveDockWinContents i
  SELECT CASE gdi(i).ID
    CASE %ID_PROPERTY
      MovePropContents gdi(i).hWnd
    CASE %ID_CONTROLS
      MoveToolboxContents gdi(i).hWnd
    CASE %ID_PROJECT
      MoveProjectContents gdi(i).hWnd
    CASE %ID_COMPILERS
      MoveCompilersContents gdi(i).hWnd
    CASE %ID_FINDRS
      MoveFindrsContents gdi(i).hWnd
  END SELECT
  activehWnd = gdi(i).hWnd
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).l<>"I" AND gdi(i).hWnd>0 AND IsWindow(gdi(i).hWnd) THEN
      DIALOG REDRAW gdi(i).hWnd
    END IF
  NEXT i
END FUNCTION
'积累宽度，并捕捉最左侧的值
'FUNCTION AccumulateWids(BYVAL gdiI AS LONG,BYVAL lastB AS LONG,BYREF lefts AS LONG)AS LONG
'  LOCAL i AS LONG
'  LOCAL rs AS LONG
'  i=gdiI
'  IF i<0 OR i>UBOUND(gdi()) OR gdi(i).l="" OR gdi(i).l="I" OR gdi(i).l="F" THEN
'    FUNCTION=0
'    EXIT FUNCTION
'  END IF
'  rs=AccumulateRWids(i,lastB) + AccumulateLWids(i,lastB)
'  lefts=GetLastLeftVal(i,lastB)
'  FUNCTION=rs
'END FUNCTION
'获得准线lastB上方最左的泊坞窗口nLeft
FUNCTION GetLastTLeftVal(BYVAL gdiI AS LONG,BYVAL lastB AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt  AS POINTAPI
  LOCAL tmpLng AS LONG
  LOCAL tmpI   AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '泊坞在左侧只有一种情况，就是左泊坞区的左上角，或者已经顶左面了，返回左侧值
  IF gdi(i).l="L" OR gdi(i).byside=(gdi(i).byside OR %BYLEFT) OR HasBottom(i)>=0 THEN
    rs=gdi(i).drc.nLeft
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口左下角离边一像素为用于测试的点位
  pt.x=gdi(gdiI).drc.nLeft-1
  pt.y=lastB-1 'gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往左就是主窗口或出了可测试范围，返回当前窗口左侧值
    rs=gdi(gdiI).drc.nLeft
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口左侧值
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs=gdi(gdiI).drc.nLeft
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口下方值与限定值相等，检查不是还有泊坞到其下方的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nTop+gdi(tmpI).drc.nBottom=lastB THEN
        IF HasBottom(tmpI)>=0 THEN   '还有泊坞到其下方的窗口，则返回当前窗口左侧值
          rs=gdi(tmpI).drc.nLeft
        ELSE                        '没有，则再左一个，递归检查
          rs=GetLastTLeftVal(tmpI,lastB)
        END IF
      ELSE                          '不在同一个泊坞区或下方值不与限定值相等，则返回当前窗口左侧值
        rs=gdi(gdiI).drc.nLeft
      END IF
    END IF
  END IF
  FUNCTION = rs
END FUNCTION
'获得准线lastT下方最左的泊坞窗口nLeft
FUNCTION GetLastBLeftVal(BYVAL gdiI AS LONG,BYVAL lastT AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt  AS POINTAPI
  LOCAL tmpLng AS LONG
  LOCAL tmpI   AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '泊坞在左侧只有一种情况，就是左泊坞区的左上角，或者已经顶左面了，返回左侧值
  IF gdi(i).byside=(gdi(i).byside OR %BYLEFT) THEN 'OR gdi(i).l="B"
    rs=gdi(i).drc.nLeft
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口左上角离边一像素为用于测试的点位
  pt.x=gdi(gdiI).drc.nLeft-1
  pt.y=lastT+1 'gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往左就是主窗口或出了可测试范围，返回当前窗口左侧值
    rs=gdi(gdiI).drc.nLeft
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口左侧值
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs=gdi(gdiI).drc.nLeft
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口下方值与限定值相等，检查不是还有泊坞到其下方的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nTop=lastT THEN
        IF gdi(tmpI).l="B" THEN   '还有泊坞到其下方的窗口，则返回当前窗口左侧值
          rs=gdi(gdiI).drc.nLeft
        ELSE                        '没有，则再左一个，递归检查
          rs=GetLastBLeftVal(tmpI,lastT)
        END IF
      ELSE                          '不在同一个泊坞区或下方值不与限定值相等，则返回当前窗口左侧值
        rs=gdi(gdiI).drc.nLeft
      END IF
    END IF
  END IF
  FUNCTION = rs
END FUNCTION
'获得准线lastB上方最右的关联泊坞窗口nLeft+nRight
FUNCTION GetLastTRightVal(BYVAL gdiI AS LONG,BYVAL lastB AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt  AS POINTAPI
  LOCAL tmpLng AS LONG
  LOCAL tmpI   AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '泊坞在左侧只有一种情况，就是左泊坞区的左上角，或者已经顶左面了，返回左侧值
  IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
    rs=gdi(i).drc.nLeft+gdi(i).drc.nRight
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口右下角离边一像素为用于测试的点位
  pt.x=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight+1
  pt.y=lastB-1 'gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往左就是主窗口或出了可测试范围，返回当前窗口左侧值
    rs=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口左侧值
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口下方值与限定值相等，检查不是还有泊坞到其下方的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nTop+gdi(tmpI).drc.nBottom=lastB THEN
        IF HasBottom(tmpI)>=0 THEN   '还有泊坞到其下方的窗口，则返回当前窗口左侧值
          rs=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
        ELSE                        '没有，则再左一个，递归检查
          rs=GetLastTRightVal(tmpI,lastB)
        END IF
      ELSE                          '不在同一个泊坞区或下方值不与限定值相等，则返回当前窗口左侧值
        rs=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
      END IF
    END IF
  END IF
  FUNCTION = rs
END FUNCTION
'获得准线lastT下方最右的关联泊坞窗口nLeft+nRight
FUNCTION GetLastBRightVal(BYVAL gdiI AS LONG,BYVAL lastT AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt  AS POINTAPI
  LOCAL tmpLng AS LONG
  LOCAL tmpI   AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '已经顶右面了，返回右侧值
  IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
    rs=gdi(i).drc.nLeft+gdi(i).drc.nRight
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口右下角离边一像素为用于测试的点位
  pt.x=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight+1
  pt.y=lastT+1 'gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往左就是主窗口或出了可测试范围，返回当前窗口左侧值
    rs=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口左侧值
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口下方值与限定值相等，检查不是还有泊坞到其下方的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nTop=lastT THEN
        IF gdi(tmpI).l="B" THEN   '还有泊坞到其下方的窗口，则返回当前窗口左侧值
          rs=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
        ELSE                        '没有，则再左一个，递归检查
          rs=GetLastBRightVal(tmpI,lastT)
        END IF
      ELSE                          '不在同一个泊坞区或下方值不与限定值相等，则返回当前窗口左侧值
        rs=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
      END IF
    END IF
  END IF
  FUNCTION = rs
END FUNCTION
' 观察指定窗口是否有从下面泊坞到它上面的窗口，返回该数组索引值
FUNCTION HasBottom(BYVAL gdiI AS LONG)AS LONG
  LOCAL i AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).ml=gdi(gdiI).ml AND gdi(i).tid=gdi(gdiI).ID AND gdi(i).l="B" THEN
      FUNCTION=i
      EXIT FUNCTION
    END IF
  NEXT i
  FUNCTION=-1
END FUNCTION
'积累泊坞窗口右侧泊坞窗口宽度，包括本身宽度
FUNCTION AccumulateRWids(BYVAL gdiI AS LONG,BYVAL lastB AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt  AS POINTAPI
  LOCAL tmpLng AS LONG
  LOCAL tmpI   AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '已经靠右侧了，返回0
  IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
    rs=gdi(i).drc.nRight
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口右下角离边一像素为用于测试的点位
  pt.x=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight+1
  pt.y=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往右就是主窗口或出了可测试范围，返回当前窗口宽度+
    rs+=gdi(gdiI).drc.nRight
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口宽度+
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs+=gdi(gdiI).drc.nRight
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口下方值与限定值相等，检查不是还有泊坞到其下方的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nTop+gdi(tmpI).drc.nBottom=lastB THEN
        IF HasBottom(tmpI)>=0 THEN   '还有泊坞到其下方的窗口，则返回当前窗口宽度+
          rs+=gdi(gdiI).drc.nRight
        ELSE                        '没有，则再右一个，递归检查
          rs+=AccumulateRWids(tmpI,lastB)
        END IF
      ELSE                          '不在同一个泊坞区或下方值不与限定值相等，则返回当前窗口宽度+
        rs+=gdi(gdiI).drc.nRight
      END IF
    END IF
  END IF
  FUNCTION = rs
END FUNCTION
'积累泊坞窗口左侧泊坞窗口宽度，不包括自身
FUNCTION AccumulateLWids(BYVAL gdiI AS LONG,BYVAL lastB AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt  AS POINTAPI
  LOCAL tmpLng AS LONG
  LOCAL tmpI   AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '泊坞在左侧只有一种情况，就是左泊坞区的左上角，或者已经顶左面了，返回0
  IF gdi(i).l="L" OR gdi(i).byside=(gdi(i).byside OR %BYLEFT) THEN
    rs=0 'gdi(i).drc.nBottom
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口左下角离边一像素为用于测试的点位
  pt.x=gdi(gdiI).drc.nLeft-1
  pt.y=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往左就是主窗口或出了可测试范围，返回当前窗口宽度+
    rs+=gdi(gdiI).drc.nRight
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口宽度+
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs+=gdi(gdiI).drc.nRight
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口下方值与限定值相等，检查不是还有泊坞到其下方的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nTop+gdi(tmpI).drc.nBottom=lastB THEN
        IF HasBottom(tmpI)>=0 THEN   '还有泊坞到其下方的窗口，则返回当前窗口宽度+
          rs+=gdi(gdiI).drc.nRight
        ELSE                        '没有，则再左一个，递归检查
          rs+=AccumulateLWids(tmpI,lastB)
        END IF
      ELSE                          '不在同一个泊坞区或下方值不与限定值相等，则返回当前窗口宽度+
        rs+=gdi(gdiI).drc.nRight
      END IF
    END IF
  END IF
  FUNCTION = rs
END FUNCTION
'积累高度，并捕捉最上面的值
'FUNCTION AccumulateHeit(BYVAL gdiI AS LONG,BYVAL lastR AS LONG,BYREF top AS LONG)AS LONG
'  LOCAL i AS LONG
'  LOCAL tmpRc AS RECT
'  LOCAL j AS LONG
'  LOCAL rs AS LONG
'  i=gdiI
'  IF i<0 OR i>UBOUND(gdi()) OR gdi(i).l="" OR gdi(i).l="I" OR gdi(i).l="F" THEN
'    FUNCTION=0
'    EXIT FUNCTION
'  END IF
'  rs=AccumulateBHeit(i,lastR) + AccumulateTHeit(i,lastR)
'  top=getLastTopVal(i,lastR)
'  FUNCTION=rs
'END FUNCTION
'获得准线lastR左侧最上面的泊坞窗口nTop
FUNCTION GetLastLTopVal(BYVAL gdiI AS LONG,BYVAL lastR AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt AS POINTAPI
  LOCAL tmpLng AS DWORD
  LOCAL tmpI  AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '泊坞在左侧只有一种情况，就是左泊坞区的左上角，或者已经顶上面了，返回该顶值
  IF gdi(i).l="L" OR gdi(i).byside=(gdi(i).byside OR %BYTOP) OR HasRight(i)>=0 THEN
    rs=gdi(i).drc.nTop
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口右上角离边一像素为用于测试的点位
  pt.x=lastR-1
  pt.y=gdi(gdiI).drc.nTop-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往上就是主窗口或出了可测试范围，返回当前窗口顶值
    rs=gdi(gdiI).drc.nTop
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口顶值
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs=gdi(gdiI).drc.nTop
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口右侧值与限定值相等，检查不是还有泊坞到其右侧的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nLeft+gdi(tmpI).drc.nRight=lastR THEN
        IF HasRight(tmpI)>=0 THEN   '还有泊坞到其右侧的窗口，则返回该窗口顶值
          rs=gdi(tmpI).drc.nTop
        ELSE                        '没有，则再上一层，递归检查
          rs=getLastLTopVal(tmpI,lastR)
        END IF
      ELSE                          '不在同一个泊坞区或右侧值不与限定值相等，则返回当前窗口顶值
        rs=gdi(gdiI).drc.nTop
      END IF
    END IF
  END IF
  FUNCTION=rs
END FUNCTION
'获得准线lastL右侧最上面的泊坞窗口nTop
FUNCTION GetLastRTopVal(BYVAL gdiI AS LONG,BYVAL lastL AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt AS POINTAPI
  LOCAL tmpLng AS DWORD
  LOCAL tmpI  AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '泊坞在左侧只有一种情况，就是左泊坞区的左上角，或者已经顶上面了，返回该顶值
  IF gdi(i).l="L" OR gdi(i).byside=(gdi(i).byside OR %BYTOP) OR HasRight(i)>=0 THEN
    rs=gdi(i).drc.nTop
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口右上角离边一像素为用于测试的点位
  pt.x=lastL+1
  pt.y=gdi(gdiI).drc.nTop-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往上就是主窗口或出了可测试范围，返回当前窗口顶值
    rs=gdi(gdiI).drc.nTop
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口顶值
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs=gdi(gdiI).drc.nTop
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口右侧值与限定值相等，检查不是还有泊坞到其右侧的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nLeft=lastL THEN
        IF gdi(tmpI).l="R" THEN   '泊坞到其它窗口的右侧，则返回该窗口顶值
          rs=gdi(tmpI).drc.nTop
        ELSE                        '没有，则再上一层，递归检查
          rs=getLastRTopVal(tmpI,lastL)
        END IF
      ELSE                          '不在同一个泊坞区或右侧值不与限定值相等，则返回当前窗口顶值
        rs=gdi(gdiI).drc.nTop
      END IF
    END IF
  END IF
  FUNCTION=rs
END FUNCTION
'获得准线lastR左侧最下面的关联泊坞窗口nTop+nBottom
FUNCTION GetLastLBottomVal(BYVAL gdiI AS LONG,BYVAL lastR AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt AS POINTAPI
  LOCAL tmpLng AS DWORD
  LOCAL tmpI  AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '已经靠底了，返回该底值
  IF gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
    rs=gdi(i).drc.nTop+gdi(i).drc.nBottom
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口右下角离边一像素为用于测试的点位
  pt.x=lastR-1 'gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight-1
  pt.y=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom+1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往下就是主窗口或出了可测试范围，返回当前窗口底值
    rs=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口顶值
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口右侧值与限定值相等，检查不是还有泊坞到其右侧的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nLeft+gdi(tmpI).drc.nRight=lastR THEN
        IF HasRight(tmpI)>=0 THEN   '还有泊坞到其右侧的窗口，则返回当前窗口顶值
          rs=gdi(gdiI).drc.nTop+gdi(i).drc.nBottom
        ELSE
          rs=getLastLBottomVal(tmpI,lastR)
        END IF
      ELSE                          '不在同一个泊坞区或右侧值不与限定值相等，则返回当前窗口底值
        rs=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom
      END IF
    END IF
  END IF
  FUNCTION=rs
END FUNCTION
'获得准线lastL右侧最下面的关联泊坞窗口nTop+nBottom
FUNCTION GetLastRBottomVal(BYVAL gdiI AS LONG,BYVAL lastL AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  LOCAL pt AS POINTAPI
  LOCAL tmpLng AS DWORD
  LOCAL tmpI  AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '已经靠底了，返回该底值
  IF gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
    rs=gdi(i).drc.nTop+gdi(i).drc.nBottom
    FUNCTION=rs
    EXIT FUNCTION
  END IF
  '取当前窗口右下角离边一像素为用于测试的点位
  pt.x=lastL+1 'gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight-1
  pt.y=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom+1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往下就是主窗口或出了可测试范围，返回当前窗口底值
    rs=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口顶值
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rs=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口右侧值与限定值相等，检查不是还有泊坞到其右侧的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nLeft=lastL THEN
        IF gdi(tmpI).l="R" THEN   '还有泊坞到其右侧的窗口，则返回当前窗口顶值
          rs=gdi(gdiI).drc.nTop+gdi(i).drc.nBottom
        ELSE
          rs=getLastRBottomVal(tmpI,lastL)
        END IF
      ELSE                          '不在同一个泊坞区或右侧值不与限定值相等，则返回当前窗口底值
        rs=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom
      END IF
    END IF
  END IF
  FUNCTION=rs
END FUNCTION
' 观察指定窗口是否被右侧泊坞了，返回右侧泊坞窗口的数组索引值
FUNCTION HasRight(BYVAL gdiI AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  rs=-1
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).ml=gdi(gdiI).ml AND gdi(i).l="R" AND gdi(i).tid=gdi(gdiI).ID THEN
      rs=i
      EXIT FOR
    END IF
  NEXT i
  FUNCTION=rs
END FUNCTION
'积累泊坞窗口下方泊坞窗口高度，包括当前窗口高度
FUNCTION AccumulateBHeit(BYVAL gdiI AS LONG,BYVAL lastR AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rstop AS LONG
  LOCAL pt  AS POINTAPI
  LOCAL tmpLng AS LONG
  LOCAL tmpI   AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  'rstop=gdi(gdiI).drc.nBottom
  '已经顶下面了，返回当前窗口高度
  IF gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
    rstop=gdi(i).drc.nBottom
    FUNCTION=rstop
    EXIT FUNCTION
  END IF
  '取当前窗口右下角离边一像素为用于测试的点位
  pt.x=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight-1
  pt.y=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom+1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往下就是主窗口或出了可测试范围，返回当前窗口高度+
    rstop+=gdi(gdiI).drc.nBottom
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口高度+
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rstop+=gdi(gdiI).drc.nBottom
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口右侧值与限定值相等，检查不是还有泊坞到其右侧的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nLeft+gdi(tmpI).drc.nRight=lastR THEN
        IF HasRight(tmpI)>=0 THEN   '还有泊坞到其右侧的窗口，则返回当前窗口高度+
          rstop+=gdi(gdiI).drc.nBottom
        ELSE                        '没有，则再上一层，递归检查
          rstop+=AccumulateBHeit(tmpI,lastR)
        END IF
      ELSE                          '不在同一个泊坞区或右侧值不与限定值相等，则返回当前窗口高度+
        rstop+=gdi(gdiI).drc.nBottom
      END IF
    END IF
  END IF
  FUNCTION=rstop
END FUNCTION
'积累泊坞窗口上方泊坞窗口高度, 不包括当前窗口高度
FUNCTION AccumulateTHeit(BYVAL gdiI AS LONG,BYVAL lastR AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rstop AS LONG
  LOCAL pt  AS POINTAPI
  LOCAL tmpLng AS LONG
  LOCAL tmpI   AS LONG
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) THEN
    FUNCTION=-1
    EXIT FUNCTION
  END IF
  '泊坞在左侧只有一种情况，就是左泊坞区的左上角，或者已经顶上面了，返回0
  IF gdi(i).l="L" OR gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
    rstop=0 'gdi(i).drc.nBottom
    FUNCTION=rstop
    EXIT FUNCTION
  END IF
  '取当前窗口右上角离边一像素为用于测试的点位
  pt.x=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight-1
  pt.y=gdi(gdiI).drc.nTop-1
  ClientToScreen g_hWndMain,pt
  tmpLng=GetWinPtIn(gdi(gdiI).hWndD,pt) '获取包含点位的窗口句柄
  IF tmpLng=g_hWndMain OR tmpLng<=0 THEN '再往上就是主窗口或出了可测试范围，返回当前窗口高度+
    rstop+=gdi(gdiI).drc.nBottom
  ELSE
    tmpI=GetIndex(tmpLng) '根据测得的句柄获取数组索引值，值小于0或大于数组上标值,则返回当前窗口高度+
    IF tmpI<0 OR tmpI>UBOUND(gdi()) THEN
      rstop+=gdi(gdiI).drc.nBottom
    ELSE '数组索引值合适
      '被测窗口与当前窗口在同一泊坞区，且泊坞窗口右侧值与限定值相等，检查不是还有泊坞到其右侧的窗口
      IF gdi(tmpI).ml=gdi(gdiI).ml AND gdi(tmpI).drc.nLeft+gdi(tmpI).drc.nRight=lastR THEN
        IF HasRight(tmpI)>=0 THEN   '还有泊坞到其右侧的窗口，则返回当前窗口高度+
          rstop+=gdi(gdiI).drc.nBottom
        ELSE                        '没有，则再上一层，递归检查
          rstop+=AccumulateTHeit(tmpI,lastR)
        END IF
      ELSE                          '不在同一个泊坞区或右侧值不与限定值相等，则返回当前窗口高度+
        rstop+=gdi(gdiI).drc.nBottom
      END IF
    END IF
  END IF
  FUNCTION=rstop
END FUNCTION
' 调节右区中的泊坞窗口,用于泊坞区边缘手动调节
FUNCTION AdjustRDockArea()AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  LOCAL clientRc AS RECT
  IF rA<=0 THEN EXIT FUNCTION
  GetWindowRect g_hWndClient,clientRc
  ScreenToClientRect g_hWndMain,clientRc
  clientRc.nRight -= clientRc.nLeft
  clientRc.nBottom -= clientRc.nTop
  IF rARc.nRight<=0   THEN rARc.nRight=20 '给定泊坞区最小宽度
  IF rARc.nBottom<=0  THEN rARc.nBottom=20 '给定泊坞区最小高度
  clientRc.nRight=rARc.nLeft-clientRc.nLeft
  MoveWin g_hWndClient,clientRc,%TRUE
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).ml="R" THEN
      IF gdi(i).byside=(gdi(i).byside OR %BYLEFT) THEN
        'IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
          gdi(i).drc.nRight=gdi(i).drc.nLeft+gdi(i).drc.nRight-rARc.nLeft 'gdi(i).drc.nLeft
        'END IF
        gdi(i).drc.nLeft=rARc.nLeft
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
        'SendMessage gdi(i).hWndD,%WM_SIZE,0,0
      END IF
    END IF
  NEXT i
END FUNCTION
' 调节下区中的泊坞窗口,用于泊坞区边缘手动调节
FUNCTION AdjustBDockArea()AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  LOCAL clientRc AS RECT
  IF bA<=0 THEN EXIT FUNCTION
  GetWindowRect g_hWndClient,clientRc
  ScreenToClientRect g_hWndMain,clientRc
  clientRc.nRight -= clientRc.nLeft
  clientRc.nBottom -= clientRc.nTop
  IF rARc.nRight<=0   THEN rARc.nRight=20 '给定泊坞区最小宽度
  IF rARc.nBottom<=0  THEN rARc.nBottom=20 '给定泊坞区最小高度
  clientRc.nBottom=bARc.nTop-clientRc.nTop
  MoveWin g_hWndClient,clientRc,%TRUE
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).ml="B" THEN
      IF gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
        'IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
          gdi(i).drc.nBottom=gdi(i).drc.nTop+gdi(i).drc.nBottom-bARc.nTop 'gdi(i).drc.nLeft
        'END IF
        gdi(i).drc.nTop=bARc.nTop
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
        'SendMessage gdi(i).hWndD,%WM_SIZE,0,0
      END IF
    END IF
  NEXT i
END FUNCTION
' 因为主窗口尺寸调整，导致泊坞区尺寸相应调整
FUNCTION AdjustDockArea()AS LONG
  LOCAL i AS LONG
  LOCAL tmpLng AS LONG
  IF lA=0 AND tA=0 AND rA=0 AND bA=0 THEN
    RESET lARc  : RESET tARc : RESET rARc : RESET bARc
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml<>"F" THEN
        gdi(i).ml="I"
        gdi(i).l="I"
        gdi(i).tid=0
        ShowWindow gdi(i).hWnd,%SW_HIDE
        gdi(i).hWnd=gdi(i).hWndF
'      else
'        'SendMessage gdi(i).hWnd,%WM_SIZE,0,0
'        msgbox "hWnd=" & str$(gdi(i).hWnd) & " hWndF=" & str$(gdi(i).hWndF)
'        MoveWin gdi(i).hWnd,gdi(i).frc,%TRUE
      END IF
    NEXT i
    FUNCTION=0
    EXIT FUNCTION
  END IF
'  FOR i=0 TO UBOUND(gdi())
'    IF gdi(i).l<>"I" THEN
'      IF gdi(i).l="F" THEN
'        SendMessage gdi(i).hWndF,%WM_SIZE,0,0
'      ELSE
'        SendMessage gdi(i).hWndD,%WM_SIZE,0,0
'      END IF
'    END IF
'  NEXT i
  IF lA=0 THEN
    RESET lARc
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="L" THEN
        gdi(i).ml="I"
        gdi(i).l="I"
        ShowWindow gdi(i).hWnd,%SW_HIDE
      END IF
    NEXT i
  END IF
  IF tA=0 THEN
    RESET tARc
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="T" THEN
        gdi(i).ml="I"
        gdi(i).l="I"
        ShowWindow gdi(i).hWnd,%SW_HIDE
      END IF
    NEXT i
  END IF
  IF rA=0 THEN
    RESET rARc
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="R" THEN
        gdi(i).ml="I"
        gdi(i).l="I"
        ShowWindow gdi(i).hWnd,%SW_HIDE
      END IF
    NEXT i
  END IF
  IF bA=0 THEN
    RESET bARc
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="B" THEN
        gdi(i).ml="I"
        gdi(i).l="I"
        ShowWindow gdi(i).hWnd,%SW_HIDE
      END IF
    NEXT i
  END IF
  IF lA>0 THEN
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="L" THEN
        IF gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
          gdi(i).drc.nBottom=gdi(i).drc.nTop+gdi(i).drc.nBottom-lARc.nTop'gdi(i).drc.nTop
          gdi(i).drc.nTop=lARc.nTop
          'MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          'SendMessage gdi(i).hWndD,%WM_SIZE,0,0 '调整窗口中的内容
        END IF
        IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
          gdi(i).drc.nRight=lARc.nRight -gdi(i).drc.nLeft
          'MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          'SendMessage gdi(i).hWndD,%WM_SIZE,0,0
        END IF
        IF gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
          gdi(i).drc.nBottom=lARc.nTop+lARc.nBottom-gdi(i).drc.nTop
          'MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          'SendMessage gdi(i).hWndD,%WM_SIZE,0,0
        END IF
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
      END IF
    NEXT i
  END IF
  IF tA>0 THEN
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="T" THEN
        IF gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
          gdi(i).drc.nBottom=gdi(i).drc.nTop+gdi(i).drc.nBottom-tARc.nTop
          gdi(i).drc.nTop=tARc.nTop
        END IF
        IF gdi(i).byside=(gdi(i).byside OR %BYLEFT) THEN
          gdi(i).drc.nRight=gdi(i).drc.nLeft+gdi(i).drc.nRight-tARc.nLeft
          gdi(i).drc.nLeft=tARc.nLeft
          'MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          'SendMessage gdi(i).hWndD,%WM_SIZE,0,0
        END IF
        IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
          gdi(i).drc.nRight=tARc.nLeft+tARc.nRight-gdi(i).drc.nLeft
          'MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          'SendMessage gdi(i).hWndD,%WM_SIZE,0,0
        END IF
        IF gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
          gdi(i).drc.nBottom=tARc.nTop+tARc.nBottom-gdi(i).drc.nTop
          'MoveWindow gdi(i).hWndD,gdi(i).drc.nLeft,gdi(i).drc.nTop,gdi(i).drc.nRight,gdi(i).drc.nBottom,%TRUE
          'SendMessage gdi(i).hWndD,%WM_SIZE,0,0
        END IF
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
      END IF
    NEXT i
  END IF
  IF rA>0 THEN
    tmpLng=0
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="R" THEN
        IF gdi(i).byside=(gdi(i).byside OR %BYLEFT) THEN
          IF tmpLng=0 THEN
            tmpLng=rARc.nLeft-gdi(i).drc.nLeft
            EXIT FOR
          END IF
        END IF
      END IF
    NEXT i
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="R" THEN
        gdi(i).drc.nLeft+=tmpLng
        IF gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
          gdi(i).drc.nBottom=gdi(i).drc.nTop+gdi(i).drc.nBottom-rARc.nTop
          gdi(i).drc.nTop=rARc.nTop
        END IF
        IF gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
          gdi(i).drc.nBottom=rARc.nTop+rARc.nBottom-gdi(i).drc.nTop
        END IF
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
      END IF
    NEXT i
  END IF
  IF bA>0 THEN
    tmpLng=0
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="B" THEN
        IF gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
          IF tmpLng=0 THEN
            tmpLng=bARc.nTop-gdi(i).drc.nTop
            EXIT FOR
          END IF
        END IF
      END IF
    NEXT i
    FOR i=0 TO UBOUND(gdi())
      IF gdi(i).ml="B" THEN
        gdi(i).drc.nTop+=tmpLng
        IF gdi(i).byside=(gdi(i).byside OR %BYLEFT) THEN
          gdi(i).drc.nRight=gdi(i).drc.nLeft+gdi(i).drc.nRight-bARc.nLeft
          gdi(i).drc.nLeft=bARc.nLeft
        END IF
        IF gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
          gdi(i).drc.nRight=bARc.nLeft+bARc.nRight-gdi(i).drc.nLeft
        END IF
        IF gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
          gdi(i).drc.nBottom=bARc.nTop+bARc.nBottom-gdi(i).drc.nTop
        END IF
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
      END IF
    NEXT i
  END IF
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).l<>"I" THEN
      IF gdi(i).l="F" THEN
        SendMessage gdi(i).hWndF,%WM_SIZE,0,0
      ELSE
        SendMessage gdi(i).hWndD,%WM_SIZE,0,0
      END IF
      DIALOG REDRAW gdi(i).hWnd
    END IF
  NEXT i
  FUNCTION = 1
END FUNCTION
' 根据窗口ID获取该窗口在数组中对应元素的索引值
FUNCTION GetIndexById(BYVAL wId AS DWORD)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  rs=-1
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).ID=wId THEN
      rs=i
      EXIT FOR
    END IF
  NEXT i
  FUNCTION=rs
END FUNCTION
' 根据窗口句柄获取该窗口在数组中对应元素的索引值
FUNCTION GetIndex(BYVAL hWnd AS DWORD)AS LONG
  LOCAL i AS LONG
  LOCAL rs AS LONG
  rs=-1
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).hWnd=hWnd OR gdi(i).hWndD=hWnd OR gdi(i).hWndF=hWnd THEN
      rs=i
      EXIT FOR
    END IF
  NEXT i
  FUNCTION=rs
END FUNCTION
' 因为拖动泊坞窗口，导致泊坞区需要调整
FUNCTION AdjustForMove(BYVAL gdiI AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL itsR  AS LONG
  LOCAL itsB  AS LONG
  LOCAL toR   AS LONG
  LOCAL toB   AS LONG
  LOCAL pt    AS POINTAPI
  LOCAL byleft AS INTEGER
  LOCAL bytop AS INTEGER
  LOCAL byright AS INTEGER
  LOCAL bybottom AS INTEGER
  LOCAL cltRc AS RECT '客户区RC
  LOCAL tmpRc AS RECT
  LOCAL lastBottom AS LONG
  LOCAL lastRight AS LONG
  itsR  = -1
  itsB  = -1
  toR   = -1
  toB   = -1
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).tid=gdi(gdiI).ID THEN
      IF gdi(i).l="R" THEN itsR=i
      IF gdi(i).l="B" THEN itsB=i
    END IF
    IF gdi(i).ID=gdi(gdiI).tid THEN
      IF gdi(gdiI).l="R" THEN toR=i
      IF gdi(gdiI).l="B" THEN toB=i
    END IF
  NEXT i
  SetRect tmpRc,itsR,itsB,toR,toB
  'canLog=1 : RunLog("itsR,B,toR,B=" & RcToStr(tmpRc))
  GetWindowRect g_hWndClient,cltRc
  ScreenToClientRect g_hWndMain,cltRc
  cltRc.nRight-=cltRc.nLeft
  cltRc.nBottom-=cltRc.nTop
  IF itsR<0 AND itsB<0 AND toR<0 AND toB<0 THEN  '没有泊坞到它上的窗口，它也没泊坞到其它窗口
    SELECT CASE gdi(gdiI).ml   '根据泊坞到主窗口的哪边，来进行相应的调整
      CASE "L"
        'canLog=1 : RunLog("tA=" & STR$(tA) & " bA=" & STR$(bA))
        'canLog=0
        IF tA>0 AND tA<>(tA OR %TOLEFT) THEN
          tA=(tA OR %TOLEFT)
          tARc.nLeft=gdi(gdiI).drc.nLeft
          tARc.nRight+=gdi(gdiI).drc.nRight
        END IF
        IF bA>0 AND bA<>(bA OR %TOLEFT) THEN
          bA=(bA OR %TOLEFT)
          bARc.nLeft=gdi(gdiI).drc.nLeft
          bARc.nRight+=gdi(gdiI).drc.nRight
        END IF
        lA=0 : RESET lARc
        cltRc.nLeft=gdi(gdiI).drc.nLeft
        cltRc.nRight+=gdi(gdiI).drc.nRight
        MoveWin g_hWndClient,cltRc,%TRUE
      CASE "T"
        IF lA>0 AND lA<>(lA OR %TOTOP) THEN
          lA=(lA OR %TOTOP)
          lARc.nTop=gdi(gdiI).drc.nTop
          lARc.nBottom+=gdi(gdiI).drc.nbottom
        END IF
        IF rA>0 AND rA<>(rA OR %TOTOP) THEN
          rA=(rA OR %TOTOP)
          rARc.nTop=gdi(gdiI).drc.nTop
          rARc.nBottom+=gdi(gdiI).drc.nBottom
        END IF
        tA=0 : RESET tARc
        cltRc.nTop=gdi(gdiI).drc.nTop
        cltRc.nBottom+=gdi(gdiI).drc.nBottom
        MoveWin g_hWndClient,cltRc,%TRUE
      CASE "R"
        IF tA>0 AND tA<>(tA OR %TORIGHT) THEN
          tA=(tA OR %TORIGHT)
          tARc.nRight+=gdi(gdiI).drc.nRight
        END IF
        IF bA>0 AND bA<>(bA OR %TORIGHT) THEN
          bA=(bA OR %TORIGHT)
          bARc.nRight+=gdi(gdiI).drc.nRight
        END IF
        rA=0 : RESET rARc
        cltRc.nRight+=gdi(gdiI).drc.nRight
        MoveWin g_hWndClient,cltRc,%TRUE
      CASE "B"
        IF lA>0 AND lA<>(lA OR %TOBOTTOM) THEN
          lA=(lA OR %TOBOTTOM)
          lARc.nBottom+=gdi(gdiI).drc.nbottom
        END IF
        IF rA>0 AND rA<>(rA OR %TOBOTTOM) THEN
          rA=(rA OR %TOTOP)
          rARc.nBottom+=gdi(gdiI).drc.nBottom
        END IF
        bA=0 : RESET bARc
        cltRc.nBottom+=gdi(gdiI).drc.nBottom
        MoveWin g_hWndClient,cltRc,%TRUE
    END SELECT
    AdjustDockArea
    EXIT FUNCTION
  END IF
  tmpRc=gdi(gdiI).drc
  tmpRc.nRight+=tmpRc.nLeft
  tmpRc.nBottom+=tmpRc.nTop
  IF itsR>=0 OR itsB>=0 THEN '首先考虑被泊坞的情况，
    IF itsR>=0 THEN
      lastBottom=GetLastRBottomVal(itsR,gdi(itsR).drc.nLeft) '获取右泊坞窗口关联下泊坞窗口底值
      IF lastBottom=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom THEN 'itsR>itsB THEN  '有右泊坞窗口，且其底值与要撤出的窗口的底值相等，则
        IF itsB>=0 THEN                                 '要撤出窗口有下泊坞窗口，则使下泊坞窗口的目标泊坞窗口为撤出窗口的右泊坞窗口的关联最下泊坞窗口
          gdi(itsB).tid=gdi(GetLastBottom(itsR)).ID
        END IF
        gdi(itsR).tid=gdi(gdiI).tid
        gdi(itsR).l=gdi(gdiI).l
        IF gdi(itsR).tid=%ID_CLIENTWINDOW OR gdi(gdiI).byside=(gdi(gdiI).byside OR %BYLEFT) THEN
          byleft=1
          gdi(itsR).byside=(gdi(itsR).byside OR %BYLEFT)
        END IF
        FOR i=0 TO UBOUND(gdi())
          pt.x=gdi(i).drc.nLeft-1
          pt.y=gdi(i).drc.nTop+1
          IF i<>gdiI AND gdi(i).ml=gdi(gdiI).ml AND PtInRect(tmpRc,pt.x,pt.y) THEN
            IF byleft=1 THEN
              gdi(i).byside=(gdi(i).byside OR %BYLEFT)
            END IF
            gdi(i).drc.nLeft=gdi(gdiI).drc.nLeft
            gdi(i).drc.nRight+=gdi(gdiI).drc.nRight
            MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          END IF
        NEXT i
        EXIT FUNCTION
      END IF
    END IF
    IF itsB>=0 THEN
      lastRight=GetLastBRightVal(itsB,gdi(itsB).drc.nTop) '+gdi(itsB).drc.nBottom)
      IF lastRight=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight THEN 'itsR<itsB THEN
        IF itsR>=0 THEN
          gdi(itsR).tid=gdi(itsB).ID
        END IF
        gdi(itsB).tid=gdi(gdiI).tid
        gdi(itsB).l=gdi(gdiI).l
        IF gdi(itsB).tid=%ID_CLIENTWINDOW OR gdi(gdiI).byside=(gdi(gdiI).byside OR %BYTOP) THEN
          bytop=1
          gdi(itsB).byside=(gdi(itsB).byside OR %BYTOP)
        END IF
        FOR i=0 TO UBOUND(gdi())
          pt.x=gdi(i).drc.nLeft+1
          pt.y=gdi(i).drc.nTop-1
          IF i<>gdiI AND gdi(i).ml=gdi(gdiI).ml AND gdi(i).l<>"F" AND gdi(i).l<>"I" AND PtInRect(tmpRc,pt.x,pt.y) THEN
            IF bytop=1 THEN
              gdi(i).byside=(gdi(i).byside OR %BYTOP)
            END IF
            gdi(i).drc.nTop=gdi(gdiI).drc.nTop
            gdi(i).drc.nBottom+=gdi(gdiI).drc.nBottom
            MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
            'SendMessage gdi(i).hWndD,%WM_SIZE,MAK(WORD,gdi(i).drc.nRight,gdi(i).drc.nBottom),0
          END IF
        NEXT i
        EXIT FUNCTION
      END IF
    END IF
  END IF
  IF toR>=0 THEN
    lastBottom=GetLastLBottomVal(toR,gdi(toR).drc.nLeft+gdi(toR).drc.nRight)
    IF lastBottom=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom THEN
      IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN
        byright=1
        gdi(toR).byside=(gdi(toR).byside OR %BYRIGHT)
      END IF
      FOR i=0 TO UBOUND(gdi())
        pt.x=gdi(i).drc.nLeft+gdi(i).drc.nRight+1
        pt.y=gdi(i).drc.nTop+1
        IF i<>gdiI AND gdi(i).ml=gdi(gdiI).ml AND gdi(i).l<>"F" AND gdi(i).l<>"I" AND IsWindow(gdi(i).hWndD) AND PtInRect(tmpRc,pt.x,pt.y) THEN
          IF byright=1 THEN
            gdi(i).byside=(gdi(i).byside OR %BYRIGHT)
          END IF
          gdi(i).drc.nRight+=gdi(gdiI).drc.nRight
          MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          'SendMessage gdi(i).hWndD,%WM_SIZE,MAK(WORD,gdi(i).drc.nRight,gdi(i).drc.nBottom),0
        END IF
      NEXT i
      EXIT FUNCTION
    END IF
  END IF
  IF toB>=0 THEN
    lastRight=GetLastTRightVal(toB,gdi(toB).drc.nTop+gdi(toB).drc.nBottom)
    IF lastRight=gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight THEN
      IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYBOTTOM) THEN
        bybottom=1
        gdi(toR).byside=(gdi(toR).byside OR %BYBOTTOM)
      END IF
      FOR i=0 TO UBOUND(gdi())
        pt.x=gdi(i).drc.nLeft+1
        pt.y=gdi(i).drc.nTop+gdi(i).drc.nBottom+1
        'canLog=1 : RunLog("toB>0 tmpRc=" & RcToStr(tmpRc) & " pt=" & PtToStr(pt) & " gdi(i).ml=" & gdi(i).ml)
        IF i<>gdiI AND gdi(i).ml=gdi(gdiI).ml AND gdi(i).l<>"F" AND gdi(i).l<>"I" AND IsWindow(gdi(i).hWndD) AND PtInRect(tmpRc,pt.x,pt.y) THEN
          RunLog "in"
          IF bybottom=1 THEN
            gdi(i).byside=(gdi(i).byside OR %BYBOTTOM)
          END IF
          gdi(i).drc.nBottom+=gdi(gdiI).drc.nBottom
          MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          'SendMessage gdi(i).hWndD,%WM_SIZE,MAK(WORD,gdi(i).drc.nRight,gdi(i).drc.nBottom),0
        END IF
      NEXT i
      EXIT FUNCTION
    END IF
  END IF
END FUNCTION
' 根据gdi数组，重排dock窗口，客户区，各泊坞区，主窗口WM_SIZE的主体
FUNCTION SizeDockWin(BYVAL lParam AS DWORD,BYVAL wParam AS DWORD)AS LONG 'BYVAL hParent AS DWORD)AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  LOCAL hadDock() AS LONG
  LOCAL clientRc AS RECT
  LOCAL tRect     AS RECT
  LOCAL ToolHeight AS LONG
  LOCAL StatHeight AS LONG
  LOCAL TabHeight AS LONG
  LOCAL tmpRc   AS RECT
  GetClientRect g_hWndMain,rc
  IF rc.nRight<600 THEN
    DIALOG SET SIZE g_hWndMain,650,rc.nBottom
  END IF
  IF rc.nBottom<500 THEN
    DIALOG SET SIZE g_hWndMain,rc.nRight,560 '要加上rebar及状态条高度
  END IF
  MoveWindow g_hRebar, 0, 0, LOWRD(lParam), HIWRD(lParam), %TRUE
  GetWindowRect g_hRebar, tRect
  ToolHeight = tRect.nBottom - tRect.nTop         '工具栏高度
  GetWindowRect g_hStatus, tRect
  StatHeight = tRect.nBottom - tRect.nTop         '状态栏高度 '
  SendMessage g_hStatus, %WM_SIZE, wParam, lParam
  GetMainClientRect rc
  clientRc=rc
  IF g_hTabMdi>0 AND IsWindowVisible(g_hTabMdi) THEN
    GetClientRect g_hTabMdi,tmpRc
    IF tmpRc.nBottom>0 THEN
      TabHeight=26
    END IF
  END IF
  '调整泊坞区
  IF UBOUND(gdi())<0 THEN '泊坞结构数组为空，则泊坞区清空，重置客户区
    lA=0    : RESET lARc
    tA=0    : RESET tARc
    rA=0    : RESET rARc
    bA=0    : RESET bARc
    IF g_hWndClient>0 AND IsWindow(g_hWndClient) THEN
      IF g_hTabMdi>0 AND IsWindowVisible(g_hTabMdi) AND TabHeight=26 THEN
        MoveWindow g_hTabMdi,rc.nLeft,rc.nTop,rc.nRight,TabHeight,%TRUE
      END IF
      MoveWin g_hWndClient,rc,%TRUE
      EXIT FUNCTION
    END IF
  END IF
  '确定泊坞区有无
  REDIM hadDock(3)
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).l<>"I" AND gdi(i).l<>"F" THEN
      SELECT CASE gdi(i).ml
        CASE "L"  : hadDock(0)=1
        CASE "T"  : hadDock(1)=1
        CASE "R"  : hadDock(2)=1
        CASE "B"  : hadDock(3)=1
      END SELECT
    END IF
  NEXT i
  IF hadDock(0)=0 THEN lA=0   : RESET lARc
  IF hadDock(1)=0 THEN tA=0   : RESET tARc
  IF hadDock(2)=0 THEN rA=0   : RESET rARc
  IF hadDock(3)=0 THEN bA=0   : RESET bARc
  '泊坞区不存在，重置客户区
  IF lA=0 AND tA=0 AND rA=0 AND bA=0 THEN
    IF g_hWndClient>0 AND IsWindow(g_hWndClient) THEN
      IF g_hTabMdi>0 AND IsWindowVisible(g_hTabMdi) THEN
        MoveWindow g_hTabMdi,rc.nLeft,rc.nTop,rc.nRight,TabHeight,%TRUE
      END IF
      MoveWin g_hWndClient,rc,%TRUE
      EXIT FUNCTION
    END IF
  END IF
  '泊坞区要包含其中的最大（宽/高）泊坞窗口，并确定客户区尺寸
  '四个区域矩形为左上角和右下角坐标（相对主窗口）
  '在主窗口调整大小过程中，左右泊坞区宽度不变，上下泊坞区高度不变，如宽度或高度超出范围，则调整到范围内
  '这样，要求主客户区大小最小是500x500
  '四个泊坞区rc定义为：左上角坐标（相对g_hWndMain），宽高
  IF lA>0 THEN
    IF lARc.nRight<=0   THEN lARc.nRight=20 '给定泊坞区最小宽度
    IF lARc.nBottom<=0  THEN lARc.nBottom=20 '给定泊坞区最小高度
  END IF
  IF rA>0 THEN
    IF rARc.nRight<=0   THEN rARc.nRight=20 '给定泊坞区最小宽度
    IF rARc.nBottom<=0  THEN rARc.nBottom=20 '给定泊坞区最小高度
  END IF
  IF tA>0 THEN
    IF tARc.nRight<=0   THEN tARc.nRight=20 '给定泊坞区最小宽度
    IF tARc.nBottom<=0  THEN tARc.nBottom=20 '给定泊坞区最小高度
  END IF
  IF bA>0 THEN
    IF bARc.nRight<=0   THEN bARc.nRight=20 '给定泊坞区最小宽度
    IF bARc.nBottom<=0  THEN bARc.nBottom=20 '给定泊坞区最小高度
  END IF
  IF lA>0 THEN '重新确定左泊坞区大小，初始客户区左侧位置
    lARc.nLeft=rc.nLeft '左区的左侧开始位置为当前主客户区的左侧
    'lARc.nRight+=lARc.nLeft
    IF tA=0 OR lA=(lA OR %TOTOP) THEN  '上区域不存在或左区域通顶
      lA=(lA OR %TOTOP)                       '则左区域上部为主客户区上部
      lARc.nTop=rc.nTop
    ELSE  '否则（上区域存在，且左区域未通顶），设置左区域上部为上区域下部
      lARc.nTop=tARc.nBottom+tARc.nTop
      'lARc.nBottom-=tARc.nBottom
    END IF
    IF bA=0 OR lA=(lA OR %TOBOTTOM) THEN '下区域不存在或左区域通底
      lA=(lA OR %TOBOTTOM)                         '则区域下部为主客户区下部
      lARc.nBottom=rc.nBottom+rc.nTop-lARc.nTop
    ELSE  '否则（下区域存在，且左区域未通底），设置左区域下部为主客户区下部减下区域高度
      lARc.nBottom=rc.nBottom+rc.nTop-bARc.nBottom-lARc.nTop
    END IF
    IF lARc.nRight<=0   THEN lARc.nRight=20 '给定泊坞区最小宽度
    IF lARc.nBottom<=0  THEN lARc.nBottom=20 '给定泊坞区最小高度
    clientRc.nLeft=lARc.nRight+lARc.nLeft '设置客户区的左侧为左区域的右侧
    clientRc.nRight=rc.nRight-lARc.nRight
  END IF
  IF tA>0 THEN   '上方泊坞区
    tARc.nTop=rc.nTop
    IF lA=0 OR tA=(tA OR %TOLEFT) THEN '左区不存在，或上区直通左侧边
      tA=(tA OR %TOLEFT)  : tARc.nLeft=rc.nLeft
    ELSE
      tARc.nLeft=lARc.nRight+lARc.nLeft
    END IF
    IF rA=0 OR tA=(tA OR %TORIGHT) THEN
      tA=(tA OR %TORIGHT)
      tARc.nRight=rc.nLeft+rc.nRight-tARc.nLeft
    ELSE
      tARc.nRight=rc.nLeft+rc.nRight-rARc.nRight-tARc.nLeft
    END IF
    IF tARc.nRight<=0   THEN tARc.nRight=20 '给定泊坞区最小宽度
    IF tARc.nBottom<=0  THEN tARc.nBottom=20 '给定泊坞区最小高度
    clientRc.nTop= tARc.nTop+tARc.nBottom
    clientRc.nBottom = rc.nBottom-tARc.nBottom
  END IF
  IF rA>0 THEN  '右侧泊坞区
    rARc.nLeft=rc.nLeft+rc.nRight-rARc.nRight
    IF tA=0 OR rA=(rA OR %TOTOP) THEN
      rA=(rA OR %TOTOP)
      rARc.nTop=rc.nTop
    ELSE
      rARc.nTop=tARc.nTop+tARc.nBottom
    END IF
    IF bA=0 OR rA=(rA OR %TOBOTTOM) THEN
      rA=rA OR %TOBOTTOM
      rARc.nBottom=rc.nTop+rc.nBottom-rARc.nTop
    ELSE
      rARc.nBottom=rc.nTop+rc.nBottom-bARc.nBottom-rARc.nTop
    END IF
    IF rARc.nRight<=0   THEN rARc.nRight=20 '给定泊坞区最小宽度
    IF rARc.nBottom<=0  THEN rARc.nBottom=20 '给定泊坞区最小高度
    clientRc.nRight=rc.nLeft+rc.nRight-rARc.nRight-clientRc.nLeft
  END IF
  IF bA>0 THEN  '下方泊坞区
    bARc.nTop=rc.nTop+rc.nBottom-bARc.nBottom
    IF lA=0 OR bA=(bA OR %TOLEFT) THEN
      bA=(bA OR %TOLEFT)
      bARc.nLeft=rc.nLeft
    ELSE
      bARc.nLeft=lARc.nLeft+lARc.nRight
    END IF
    IF rA=0 OR bA=(bA OR %TORIGHT) THEN
      bA=bA OR %TORIGHT
      bARc.nRight=rc.nLeft+rc.nRight-bARc.nLeft
    ELSE
      bARc.nRight=rc.nLeft+rc.nRight-rARc.nRight-bARc.nLeft
    END IF
    IF bARc.nRight<=0   THEN bARc.nRight=20 '给定泊坞区最小宽度
    IF bARc.nBottom<=0  THEN bARc.nBottom=20 '给定泊坞区最小高度
    clientRc.nBottom=rc.nTop+rc.nBottom-bARc.nBottom-clientRc.nTop
  END IF
  IF g_hTabMdi>0 AND IsWindowVisible(g_hTabMdi) THEN
    MoveWindow g_hTabMdi,clientRc.nLeft,clientRc.nTop,clientRc.nRight,TabHeight,%TRUE
    clientRc.nTop+=TabHeight
    clientRc.nBottom-=TabHeight
  END IF
  MoveWin g_hWndClient,clientRc,%TRUE
  AdjustDockArea '四周泊坞区都确定的情况下，调整泊坞区内的泊坞窗口
END FUNCTION
' 指定窗口切换到泊坞状态，泊坞到泊坞区或某个泊坞窗口后，对相应的泊坞区及目标泊坞窗口进行调整
FUNCTION AdjustForInsert(BYVAL gdiI AS LONG)AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  LOCAL cltRc AS RECT
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) OR gdi(i).l="F" OR gdi(i).l="I" THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  'GetClientRect g_hWndMain,rc
  GetMainClientRect rc
  GetWindowRect g_hWndClient,cltRc
  ScreenToClientRect g_hWndMain,cltRc
  cltRc.nRight-=cltRc.nLeft
  cltRc.nBottom-=cltRc.nTop
  '独占某个泊坞区
  SELECT CASE gdi(gdiI).ml
    CASE "L"
      IF lA=0 THEN
        RESET lARc
        lARc.nLeft=rc.nLeft
        lARc.nRight=gdi(gdiI).drc.nRight
        lA=%TOCENTER
        IF tA=0 THEN
          lA=(lA OR %TOTOP)
          lARc.nTop=rc.nTop
        ELSE
          lARc.nTop=tARc.nTop+tARc.nBottom
        END IF
        IF bA=0 THEN
          lA=(lA OR %TOBOTTOM)
          lARc.nBottom=rc.nTop+rc.nBottom-lARc.nTop
        ELSE
          lARc.nBottom=rc.nTop+rc.nBottom-lARc.nTop-bARc.nBottom
        END IF
        cltRc.nLeft=lARc.nLeft+lARc.nRight
        cltRc.nRight-=lARc.nRight
        MoveWin g_hWndClient,cltRc,%TRUE
        gdi(gdiI).byside=15
        gdi(gdiI).hWnd=gdi(gdiI).hWndD
        gdi(gdiI).tid=%ID_CLIENTWINDOW
        'canLog=1 : RunLog("gdi(" & FORMAT$(gdiI) & ").drc=" & RcToStr(gdi(gdiI).drc))
        SendMessage gdi(gdiI).hWndD,%WM_SIZE,0,0
        'RunLog("1 Insert lARc=" & RcToStr(lARc) & " mainclientrc=" & RcToStr(rc))
        AdjustDockArea
        'RunLog("2 Insert lARc=" & RcToStr(lARc) & " mainclientrc=" & RcToStr(rc)) : canLog=0
        EXIT FUNCTION
      END IF
    CASE "T"
      IF tA=0 THEN
        RESET tARc
        tARc.nTop=rc.nTop
        tARc.nBottom=gdi(gdiI).drc.nBottom
        tA=%TOCENTER
        IF lA=0 THEN
          tA=(tA OR %TOLEFT)
          tARc.nLeft=rc.nLeft
        ELSE
          tARc.nLeft=lARc.nLeft+lARc.nRight
        END IF
        IF rA=0 THEN
          tA=(tA OR %TORIGHT)
          tARc.nRight=rc.nLeft+rc.nRight-tARc.nLeft
        ELSE
          tARc.nRight=rc.nLeft+rc.nRight-tARc.nLeft-rARc.nRight
        END IF
        cltRc.nTop=tARc.nTop+tARc.nBottom
        cltRc.nBottom-=tARc.nBottom
        MoveWin g_hWndClient,cltRc,%TRUE
        gdi(gdiI).byside=15
        gdi(gdiI).hWnd=gdi(gdiI).hWndD
        gdi(gdiI).tid=%ID_CLIENTWINDOW
        SendMessage gdi(gdiI).hWndD,%WM_SIZE,0,0
        AdjustDockArea
        EXIT FUNCTION
      END IF
    CASE "R"
      IF rA=0 THEN
        rARc.nLeft=rc.nLeft+rc.nRight-gdi(gdiI).drc.nRight
        rARc.nRight=gdi(gdiI).drc.nRight
        rA=%TOCENTER
        IF tA=0 THEN
          rA=(rA OR %TOTOP)
          rARc.nTop=rc.nTop
        ELSE
          rARc.nTop=tARc.nTop+tARc.nBottom
        END IF
        IF bA=0 THEN
          rA=(rA OR %TOBOTTOM)
          rARc.nBottom=rc.nTop+rc.nBottom-rARc.nTop
        ELSE
          rARc.nBottom=rc.nTop+rc.nBottom-rARc.nTop-bARc.nBottom
        END IF
        cltRc.nRight=rc.nLeft+rc.nRight-rARc.nRight-cltRc.nLeft
        MoveWin g_hWndClient,cltRc,%TRUE
        gdi(gdiI).byside=15
        gdi(gdiI).hWnd=gdi(gdiI).hWndD
        gdi(gdiI).tid=%ID_CLIENTWINDOW
        SendMessage gdi(gdiI).hWndD,%WM_SIZE,0,0
        AdjustDockArea
        EXIT FUNCTION
      END IF
    CASE "B"
      IF bA=0 THEN
        bARc.nTop=rc.nTop+rc.nBottom-gdi(gdiI).drc.nBottom
        bARc.nBottom=gdi(gdiI).drc.nBottom
        bA=%TOCENTER
        IF lA=0 THEN
          bA=bA OR %TOLEFT
          bARc.nLeft=rc.nLeft
        ELSE
          bARc.nLeft=lARc.nLeft+lARc.nRight
        END IF
        IF rA=0 THEN
          bA=bA OR %TORIGHT
          bARc.nRight=rc.nLeft+rc.nRight-bARc.nLeft
        ELSE
          bARc.nRight=rc.nRight-bARc.nLeft-rARc.nRight
        END IF
        'canLog=1 : RunLog("bARc=" & RcToStr(bARc) & " rARc=" & RcToStr(rARc))
        cltRc.nBottom=rc.nTop+rc.nBottom-bARc.nBottom-cltRc.nTop
        MoveWin g_hWndClient,cltRc,%TRUE
        gdi(gdiI).byside=15
        gdi(gdiI).hWnd=gdi(gdiI).hWndD
        gdi(gdiI).tid=%ID_CLIENTWINDOW
        SendMessage gdi(gdiI).hWndD,%WM_SIZE,0,0
        AdjustDockArea
        EXIT FUNCTION
      END IF
  END SELECT
  '直接泊坞到某个泊坞区，该泊坞区已经有了泊坞窗口
  IF gdi(gdiI).tid=%ID_CLIENTWINDOW THEN
    gdi(gdiI).ml=gdi(gdiI).l
    SELECT CASE gdi(gdiI).l
      CASE "L"                       '泊坞到泊坞区的左侧，则挤到最左边
        lARc.nRight+=gdi(gdiI).drc.nRight
        cltRc.nRight-=gdi(gdiI).drc.nRight
        cltRc.nLeft+=gdi(gdiI).drc.nRight
        MoveWin g_hWndClient,cltRc,%TRUE
        FOR i=0 TO UBOUND(gdi())
          IF i<>gdiI AND gdi(i).ml="L" THEN
            gdi(i).drc.nLeft += gdi(gdiI).drc.nRight
            IF gdi(i).byside=(gdi(i).byside OR %BYLEFT) THEN
              gdi(i).byside=(gdi(i).byside AND (NOT %BYLEFT))
            END IF
            IF gdi(i).l="L" AND gdi(i).tid=%ID_CLIENTWINDOW THEN
              gdi(i).l="R"
              gdi(i).tid=gdi(gdiI).ID
            END IF
            MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
            'SendMessage gdi(i).hWndD,%WM_SIZE,0,0
          END IF
        NEXT i
        gdi(gdiI).byside=(%BYLEFT OR %BYTOP OR %BYBOTTOM)
        IF tA>0 AND lA=(lA OR %TOTOP) THEN
          tARc.nLeft+=gdi(gdiI).drc.nRight
          tARc.nRight-=gdi(gdiI).drc.nRight
          'SendMessage g_hWndMain,%WM_SIZE,0,0
          AdjustDockArea
        END IF
        IF bA>0 AND lA=(lA OR %TOBOTTOM) THEN
          bARc.nLeft+=gdi(gdiI).drc.nRight
          bARc.nRight-=gdi(gdiI).drc.nRight
          'SendMessage g_hWndMain,%WM_SIZE,0,0
          AdjustDockArea
        END IF
      CASE "T" '泊坞到上泊坞区，挤到上面
        tARc.nBottom+=gdi(gdiI).drc.nBottom
        cltRc.nTop+=gdi(gdiI).drc.nBottom
        cltRc.nBottom-=gdi(gdiI).drc.nBottom
        MoveWin g_hWndClient,cltRc,%TRUE
        FOR i=0 TO UBOUND(gdi())
          IF i<>gdiI AND gdi(i).ml="T" THEN
            gdi(i).drc.nTop+=gdi(gdiI).drc.nBottom
            IF gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
              gdi(i).byside=(gdi(i).byside AND (NOT %BYTOP))
            END IF
            IF gdi(i).l="T" AND gdi(i).tid=%ID_CLIENTWINDOW THEN
              gdi(i).l="B"
              gdi(i).tid=gdi(gdiI).ID
            END IF
            MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
          END IF
        NEXT i
        gdi(gdiI).byside=(%BYLEFT OR %BYTOP OR %BYRIGHT)
        IF lA>0 AND tA=(tA OR %TOLEFT) THEN
          lARc.nTop+=gdi(gdiI).drc.nBottom
          lARc.nBottom-=gdi(gdiI).drc.nBottom
          AdjustDockArea
        END IF
        IF rA>0 AND tA=(tA OR %TORIGHT) THEN
          rARc.nTop+=gdi(gdiI).drc.nBottom
          rARc.nBottom-=gdi(gdiI).drc.nBottom
          AdjustDockArea
        END IF
      CASE "R"
        rARc.nLeft-=gdi(gdiI).drc.nRight
        rARc.nRight+=gdi(gdiI).drc.nRight
        cltRc.nRight-=gdi(gdiI).drc.nRight
        MoveWin g_hWndClient,cltRc,%TRUE
        FOR i=0 TO UBOUND(gdi())
          IF i<>gdiI AND gdi(i).ml="R" THEN
            IF gdi(i).byside=(gdi(i).byside OR %BYLEFT) THEN
              gdi(i).byside=(gdi(i).byside AND (NOT %BYLEFT))
            END IF
            IF gdi(i).l="R" AND gdi(i).tid=%ID_CLIENTWINDOW THEN
              gdi(i).tid=gdi(gdiI).ID
            END IF
          END IF
        NEXT i
        gdi(gdiI).byside=(%BYLEFT OR %BYTOP OR %BYBOTTOM)
        IF tA>0 AND rA=(rA OR %TOTOP) THEN
          tARc.nRight-=gdi(gdiI).drc.nRight
          AdjustDockArea
        END IF
        IF bA>0 AND rA=(rA OR %TOBOTTOM) THEN
          bARc.nRight-=gdi(gdiI).drc.nRight
          AdjustDockArea
        END IF
      CASE "B"
        bARc.nTop-=gdi(gdiI).drc.nBottom
        bARc.nBottom+=gdi(gdiI).drc.nBottom
        cltRc.nBottom-=gdi(gdiI).drc.nBottom
        MoveWin g_hWndClient,cltRc,%TRUE
        FOR i=0 TO UBOUND(gdi())
          IF i<>gdiI AND gdi(i).ml="B" THEN
            IF gdi(i).byside=(gdi(i).byside OR %BYTOP) THEN
              gdi(i).byside=(gdi(i).byside AND (NOT %BYTOP))
            END IF
            IF gdi(i).l="B" AND gdi(i).tid=%ID_CLIENTWINDOW THEN
              gdi(i).tid=gdi(gdiI).ID
            END IF
          END IF
        NEXT i
        gdi(gdiI).ml="B"
        gdi(gdiI).l="B"
        gdi(gdiI).byside=(%BYLEFT OR %BYTOP OR %BYRIGHT)
        IF lA>0 AND tA=(tA OR %TOLEFT) THEN
          lARc.nBottom-=gdi(gdiI).drc.nBottom
          AdjustDockArea
        END IF
        IF rA>0 AND tA=(tA OR %TORIGHT) THEN
          rARc.nBottom-=gdi(gdiI).drc.nBottom
          AdjustDockArea
        END IF
    END SELECT
    EXIT FUNCTION
  END IF
  '泊坞对象是一个泊坞窗口
  LOCAL dstI AS LONG
  dstI=FindGdiIndex(gdi(gdiI).tid)
  SELECT CASE gdi(gdiI).l
    CASE "L"
      gdi(gdiI).l=gdi(dstI).l
      gdi(gdiI).tid=gdi(dstI).tid
      gdi(gdiI).byside=gdi(dstI).byside
      IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYRIGHT) THEN
        gdi(gdiI).byside=(gdi(gdiI).byside AND (NOT %BYRIGHT))
      END IF
      gdi(dstI).l="R"
      gdi(dstI).tid=gdi(gdiI).ID
      IF gdi(dstI).byside=(gdi(dstI).byside OR %BYLEFT) THEN
        gdi(dstI).byside=(gdi(dstI).byside AND (NOT %BYLEFT))
      END IF
      'gdi(dstI).drc.nLeft+=gdi(dstI).drc.nRight/2
      gdi(dstI).drc.nRight += gdi(dstI).drc.nLeft
      gdi(dstI).drc.nLeft = gdi(gdiI).drc.nLeft+gdi(gdiI).drc.nRight
      gdi(dstI).drc.nRight -= gdi(dstI).drc.nLeft
      MoveWin gdi(dstI).hWndD,gdi(dstI).drc,%TRUE
      FOR i=0 TO UBOUND(gdi())
        IF gdi(i).ml=gdi(dstI).ml AND gdi(i).tid=gdi(dstI).ID AND gdi(i).l="B" THEN
          gdi(i).tid=gdi(gdiI).ID
          EXIT FOR
        END IF
      NEXT i
    CASE "T"
      gdi(gdiI).l=gdi(dstI).l
      gdi(gdiI).tid=gdi(dstI).tid
      gdi(gdiI).byside=gdi(dstI).byside
      IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYBOTTOM) THEN
        gdi(gdiI).byside=(gdi(gdiI).byside AND (NOT %BYBOTTOM))
      END IF
      gdi(dstI).l="B"
      gdi(dstI).tid=gdi(gdiI).ID
      IF gdi(dstI).byside=(gdi(dstI).byside OR %BYTOP) THEN
        gdi(dstI).byside=(gdi(dstI).byside AND (NOT %BYTOP))
      END IF
      'gdi(dstI).drc.nTop+=gdi(dstI).drc.nBottom/2
      gdi(dstI).drc.nBottom += gdi(dstI).drc.nTop
      gdi(dstI).drc.nTop=gdi(gdiI).drc.nTop+gdi(gdiI).drc.nBottom
      gdi(dstI).drc.nBottom -= gdi(dstI).drc.nTop 'gdi(dstI).drc.nBottom/2
      MoveWin gdi(dstI).hWndD,gdi(dstI).drc,%TRUE
      FOR i=0 TO UBOUND(gdi())
        IF gdi(i).ml=gdi(dstI).ml AND gdi(i).tid=gdi(dstI).ID AND gdi(i).l="R" THEN
          gdi(i).tid=gdi(gdiI).ID
          EXIT FOR
        END IF
      NEXT i
    CASE "R"
      gdi(gdiI).byside=gdi(dstI).byside
      IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYLEFT) THEN
        gdi(gdiI).byside=(gdi(gdiI).byside AND (NOT %BYLEFT))
      END IF
      IF gdi(dstI).byside=(gdi(dstI).byside OR %BYRIGHT) THEN
        gdi(dstI).byside=(gdi(dstI).byside AND (NOT %BYRIGHT))
      END IF
      gdi(dstI).drc.nRight=gdi(gdiI).drc.nLeft-gdi(dstI).drc.nLeft'gdi(dstI).drc.nRight/2
      MoveWin gdi(dstI).hWndD,gdi(dstI).drc,%TRUE
      FOR i=0 TO UBOUND(gdi())
        IF i<>gdiI AND gdi(i).ml=gdi(dstI).ml AND gdi(i).tid=gdi(dstI).ID AND gdi(i).l="R" THEN
          gdi(i).tid=gdi(gdiI).ID
          EXIT FOR
        END IF
      NEXT i
    CASE "B"
      gdi(gdiI).byside=gdi(dstI).byside
      IF gdi(gdiI).byside=(gdi(gdiI).byside OR %BYTOP) THEN
        gdi(gdiI).byside=(gdi(gdiI).byside AND (NOT %BYTOP))
      END IF
      IF gdi(dstI).byside=(gdi(dstI).byside OR %BYBOTTOM) THEN
        gdi(dstI).byside=(gdi(dstI).byside AND (NOT %BYBOTTOM))
      END IF
      'gdi(dstI).drc.nBottom=gdi(dstI).drc.nBottom/2
      'gdi(dstI).drc.nBottom=gdi(dstI).drc.nTop+gdi(dstI).drc.nBottom-gdi(gdiI).drc.nTop
      'IF gdi(dstI).ml="L" OR gdi(dstI).ml="R" THEN '在左右泊坞区时，泊坞到目标窗口下方时，不明原因高度小了1
      '  gdi(dstI).drc.nBottom+=1
      'END IF
      gdi(dstI).drc.nBottom=gdi(gdiI).drc.nTop-gdi(dstI).drc.nTop
      MoveWin gdi(dstI).hWndD,gdi(dstI).drc,%TRUE
      FOR i=0 TO UBOUND(gdi())
        IF i<>gdiI AND gdi(i).ml=gdi(dstI).ml AND gdi(i).tid=gdi(dstI).ID AND gdi(i).l="B" THEN
          gdi(i).tid=gdi(gdiI).ID
          EXIT FOR
        END IF
      NEXT i
  END SELECT
END FUNCTION
' 窗口右侧调整大小后，调整周边泊坞窗口及它自身
' 调整后的尺寸已存为gdi(gdiI).drc.nRight
FUNCTION AdjustRWin(BYVAL gdiI AS LONG)AS LONG
  LOCAL lastTop AS LONG
  LOCAL lastBottom AS LONG
  LOCAL lastR   AS LONG
  LOCAL newR   AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) OR gdi(i).l="F" OR gdi(i).l="I" OR gdi(i).byside=(gdi(i).byside OR %BYRIGHT) THEN
    'msgbox "test"
    EXIT FUNCTION
  END IF
  GetWindowRect gdi(i).hWndD,rc
  ScreenToClientRect g_hWndMain,rc
  rc.nRight -= rc.nLeft
  rc.nBottom -= rc.nTop
  lastR       = rc.nLeft+rc.nRight                  '调整之前的右侧位置
  newR        = gdi(i).drc.nLeft+gdi(i).drc.nRight  '调整之后的右侧位置
  'canLog=1 : RunLog("AdjustRWin before getLastTopVal")
  lastTop     = getLastLTopVal(gdiI,lastR)           '关联的最上面的顶值
  'canLog=1 : RunLog("AdjustRWin begin...")
  lastBottom  = GetLastLBottomVal(gdiI,lastR)        '关联的最下面的底值
  'canLog=1 : RunLog("AdjustRWin lastR=" & format$(lastR) & " newR=" & format$(newR) _
  '            & " lastTop=" & format$(lastTop) & " lastBottom=" & format$(lastbottom))
  FOR i=0 TO UBOUND(gdi())
    '同一泊坞区中，底值小于等于最下底值，顶值大于等于最上顶值
'    RunLog "gdiI=" & format$(gdiI) & " gdi(" & format$(i) & ").ml=" & gdi(i).ml & " gdi(gdiI).ml=" & gdi(gdiI).ml
'    RunLog "gdi(i).nTop+nBottom=" & format$(gdi(i).drc.nTop+gdi(i).drc.nBottom) & " lastBottom=" & format$(lastBottom)
'    RunLog "gdi(i).drc.nTop=" & format$(gdi(i).drc.nTop) & " lastTop=" & format$(lastTop)
    IF gdi(i).ml=gdi(gdiI).ml AND i<>gdiI AND IsWindow(gdi(i).hWndD) AND gdi(i).drc.nTop+gdi(i).drc.nBottom<=lastBottom _
          AND gdi(i).drc.nTop>=lastTop THEN
'      RunLog "gdi(" & str$(i) & ").nLeft+nRight=" & str$(gdi(i).drc.nLeft+gdi(i).drc.nRight) & $crlf _
'              & "lastR=" & str$(lastR)
      IF gdi(i).drc.nLeft+gdi(i).drc.nRight=lastR THEN
        gdi(i).drc.nRight=newR-gdi(i).drc.nLeft
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
        SendMessage gdi(i).hWndD,%WM_SIZE,0,0
      ELSEIF gdi(i).drc.nLeft=lastR THEN
        gdi(i).drc.nRight=gdi(i).drc.nLeft+gdi(i).drc.nRight-newR
        gdi(i).drc.nLeft=newR
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
        SendMessage gdi(i).hWndD,%WM_SIZE,0,0
      END IF
    END IF
  NEXT i
  MoveWin gdi(gdiI).hWndD,gdi(gdiI).drc,%TRUE
  'SendMessage gdi(gdiI).hWndD,%WM_SIZE,0,0
END FUNCTION
' 窗口下方调整大小后，调整周边泊坞窗口及它自身
' 调整后的尺寸已存入gdi(gdiI).drc.nBottom
FUNCTION AdjustBWin(BYVAL gdiI AS LONG)AS LONG
  LOCAL lastLeft AS LONG
  LOCAL lastRight AS LONG
  LOCAL lastB   AS LONG
  LOCAL newB   AS LONG
  LOCAL i AS LONG
  LOCAL rc AS RECT
  i=gdiI
  IF i<0 OR i>UBOUND(gdi()) OR gdi(i).l="F" OR gdi(i).l="I" OR gdi(i).byside=(gdi(i).byside OR %BYBOTTOM) THEN
    EXIT FUNCTION
  END IF
  GetWindowRect gdi(i).hWndD,rc
  ScreenToClientRect g_hWndMain,rc
  rc.nRight -= rc.nLeft
  rc.nBottom -= rc.nTop
  lastB       = rc.nTop+rc.nBottom
  newB        = gdi(i).drc.nTop +gdi(i).drc.nBottom
  lastLeft    = GetLastTLeftVal(gdiI,lastB)
  lastRight   = GetLastTRightVal(gdiI,lastB) 'lastLeft+AccumulateWids(gdiI,lastB,lastRight)'GetLastBottom(gdiI)
  FOR i=0 TO UBOUND(gdi())
    IF gdi(i).ml=gdi(gdiI).ml AND gdi(i).drc.nLeft+gdi(i).drc.nRight<=lastRight AND gdi(i).drc.nLeft>=lastLeft THEN
      IF gdi(i).drc.nTop+gdi(i).drc.nBottom=lastB THEN
        gdi(i).drc.nBottom=newB-gdi(i).drc.nTop
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
        SendMessage gdi(i).hWndD,%WM_SIZE,0,0
      ELSEIF gdi(i).drc.nTop=lastB THEN
        gdi(i).drc.nBottom=gdi(i).drc.nTop+gdi(i).drc.nBottom-newB
        gdi(i).drc.nTop=newB
        MoveWin gdi(i).hWndD,gdi(i).drc,%TRUE
        SendMessage gdi(i).hWndD,%WM_SIZE,0,0
      END IF
    END IF
  NEXT i
  MoveWin gdi(gdiI).hWndD,gdi(gdiI).drc,%TRUE
  'SendMessage gdi(gdiI).hWndD,%WM_SIZE,0,0
END FUNCTION
' 泊坞窗口大小调整事件响应
FUNCTION On_SizeDockWin(BYVAL wID AS DWORD,BYVAL hWnd AS DWORD,BYVAL wParam AS DWORD)AS LONG
  LOCAL i AS LONG
  IF hWnd<=0 THEN
    EXIT FUNCTION
  END IF
  IF ISFALSE IsWindowVisible(hWnd) THEN
    EXIT FUNCTION
  END IF
  i=GetWindowLong(hWnd,%GWL_USERDATA)
  IF i<0 OR i>UBOUND(gdi()) THEN
    EXIT FUNCTION
  END IF
  wID=gdi(i).ID
  SELECT CASE wID
    CASE %ID_PROPERTY
      SizePropContents hWnd,wParam ',lParam
    CASE %ID_CONTROLS
      SizeToolboxContents hWnd,wParam
    CASE %ID_PROJECT
      SizeProjectContents hWnd,wParam
    CASE %ID_COMPILERS
      SizeCompilersContents hWnd,wParam
    CASE %ID_FINDRS
      SizeFindrsContents hWnd,wParam
  END SELECT
END FUNCTION
' 泊坞窗口中控件命令触发事件
FUNCTION On_CommandDockWin(BYVAL wID AS DWORD,BYVAL wParam AS DWORD)AS LONG
  SELECT CASE wID
    CASE %ID_PROPERTY
      '属性列表控件消息
      Command_Prop wParam
    CASE %ID_CONTROLS
      Command_Toolbox wParam
    CASE %ID_PROJECT
      Command_Project wParam
    CASE %ID_COMPILERS
      Command_Compilers wParam
    CASE %ID_FINDRS
      Command_Findrs wParam
  END SELECT
END FUNCTION
' 泊坞窗口销毁事件
FUNCTION On_DestroyDockWin(BYVAL wID AS DWORD)AS LONG
  SELECT CASE wID
    CASE %ID_PROPERTY
      DestroyPropertyWin
    CASE %ID_CONTROLS
      DestroyToolboxWin
    CASE %ID_PROJECT
      DestroyProjectWin
    CASE %ID_COMPILERS
      DestroyCompilersWin
    CASE %ID_FINDRS
      DestroyFindrsWin
  END SELECT
END FUNCTION
' 转移泊坞中的控件到当前窗口中  (被鼠标移动，左键弹起及DockWinShow调用)
FUNCTION MoveDockWinContents(BYVAL gdiI AS LONG) AS LONG
  SELECT CASE gdi(gdiI).ID 'gdi1.id
    CASE %ID_PROPERTY
      IF gdi(gdiI).l="F" THEN
        MovePropContents gdi(gdiI).hWndF
      ELSE
        MovePropContents gdi(gdiI).hWndD
      END IF
    CASE %ID_CONTROLS
      IF gdi(gdiI).l="F" THEN
        MoveToolboxContents gdi(gdiI).hWndF
      ELSE
        MoveToolboxContents gdi(gdiI).hWndD
      END IF
    CASE %ID_PROJECT
      IF gdi(gdiI).l="F" THEN
        MoveProjectContents gdi(gdiI).hWndF
      ELSE
        MoveProjectContents gdi(gdiI).hWndD
      END IF
    CASE %ID_COMPILERS
      IF gdi(gdiI).l="F" THEN
        MoveCompilersContents gdi(gdiI).hWndF
      ELSE
        MoveCompilersContents gdi(gdiI).hWndD
      END IF
    CASE %ID_FINDRS
      IF gdi(gdiI).l="F" THEN
        MoveFindrsContents gdi(gdiI).hWndF
      ELSE
        MoveFindrsContents gdi(gdiI).hWndD
      END IF
  END SELECT
END FUNCTION
