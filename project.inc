''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 文件 : project.inc
' 项目窗口：加载并展示当前项目中的源文件
'
' 创建项目窗口
%IDC_PROJECTTV = 102
FUNCTION CreateProjectWin(BYVAL gdiI AS LONG) AS DWORD
  LOCAL i AS LONG
  IF gdiI<0 OR gdiI>UBOUND(gdi()) THEN
    FUNCTION=0
    EXIT FUNCTION
  END IF
  i=gdiI
  gdi(i).ID = %ID_PROJECT
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
  ' 创建工具箱泊坞窗口
  gdi(i).hWndD = CreateWindowEx( 0 , "DOCKIT", "项目", _
                                  %WS_CHILD, _
                                  gdi(i).drc.nLeft, gdi(i).drc.nTop, _
                                  gdi(i).drc.nRight, gdi(i).drc.nBottom, _  ' width,height
                                  g_hWndMain, %ID_PROJECT, g_hInst, BYVAL %NULL)
  SetWindowLong gdi(i).hWndD,%GWL_USERDATA,i
  ' 创建工具箱浮动窗口
  gdi(i).hWndF = CreateWindowEx( 0 , "DOCKIT", "项目", _
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
  CreateProjectContents gdi(i).hWnd
  FUNCTION = gdi(i).hWnd
END FUNCTION
' 创建工具箱窗口中的控件：工具条，使用静态控件包装，这样就可以随窗口大小调整而对工具条进行调整
FUNCTION CreateProjectContents(BYVAL hWnd AS DWORD)AS LONG
  LOCAL rc AS RECT
  LOCAL hRoot AS DWORD
  LOCAL hFile1 AS DWORD
  LOCAL hDialog1 AS DWORD
  LOCAL hDlg1 AS DWORD
  LOCAL hProc1 AS DWORD
  LOCAL hDlgProc1 AS DWORD
  GetClientRect hWnd,rc
  'control add treeview,hWnd,%IDC_PROJECTTV,"qq",rc.nLeft+8,rc.nTop+25,rc.nRight-16,rc.nBottom-30 call ProjectProc
  hProjectTV = CreateWindowEx(0,"SYSTREEVIEW32","", _          '%WS_EX_TRANSPARENT
                                %TVS_LINESATROOT  OR %TVS_HASBUTTONS OR _  '%TVS_HASLINESATROOT OR OR %TVS_TRACKSELECT
                                %TVS_HASLINES OR %TVS_SHOWSELALWAYS OR %TVS_FULLROWSELECT OR _
                                %WS_CHILD OR %WS_VISIBLE OR %WS_BORDER , _
                                rc.nLeft+4,rc.nTop+23, _
                                rc.nRight-8,rc.nBottom-28, _
                                hWnd,%IDC_PROJECTTV, g_hInst, BYVAL %NULL )
'  TREEVIEW INSERT ITEM hWnd,%IDC_PROJECTTV,0,%TVI_FIRST,0,0,"Dockit04-5" TO hRoot
'  hPTRoot = hRoot
'  TREEVIEW INSERT ITEM hWnd,%IDC_PROJECTTV,hRoot,%TVI_LAST,0,0,"Dockit04.bas" TO hFile1
'  TREEVIEW INSERT ITEM hWnd,%IDC_PROJECTTV,hFile1,%TVI_LAST,0,0,"窗口" TO hDialog1
'  TREEVIEW INSERT ITEM hWnd,%IDC_PROJECTTV,hDialog1,%TVI_LAST,0,0,"WINMAIN" TO hDlg1
'  TREEVIEW INSERT ITEM hWnd,%IDC_PROJECTTV,hDlg1,%TVI_LAST,0,0,"回调" TO hProc1
'  TREEVIEW INSERT ITEM hWnd,%IDC_PROJECTTV,hProc1,%TVI_LAST,0,0,"MainWndProc" TO hDlgProc1
END FUNCTION
' *********************************************************************************************
' 加载项目树
' *********************************************************************************************
FUNCTION ProjectLoadTree(fName AS STRING) AS LONG
   LOCAL hItem         AS LONG
   LOCAL sName         AS STRING
   LOCAL i             AS LONG
   LOCAL sPrimary      AS STRING
   LOCAL szSection     AS STRING 'ASCIIZ * 255
   LOCAL szPath        AS STRING 'ASCIIZ * %MAX_PATH
   LOCAL szKey         AS STRING 'ASCIIZ * 255
   LOCAL szInString    AS STRING 'ASCIIZ * %MAX_PATH
   LOCAL szDefault     AS STRING 'ASCIIZ * 255
   LOCAL szProjectFile AS STRING 'ASCIIZ * %MAX_PATH
   LOCAL lRes          AS LONG
   LOCAL nFiles        AS LONG
   sProjectName = fName
   szProjectFile = fName
   szSection = "BAS Files"
   szKey = "Primary Source File"
   'lRes = GetPrivateProfileString(szSection, szKey, szDefault, szInString, %MAX_PATH, szProjectFile)
   sPrimary = IniRead(szProjectFile,szSection,szKey,"") '主文件名
   'IF lRes THEN sPrimary = TRIM$(EXTRACT$(szInString, CHR$(0)))
   szKey = "Files"
   'lRes = GetPrivateProfileString(szSection, szKey, szDefault, szInString, %MAX_PATH, szProjectFile)
   nFiles = VAL(IniRead(szProjectFile,szSection,szKey,"0")) 'bas文件数量
   'IF lRes THEN nFiles = VAL(TRIM$(EXTRACT$(szInString, CHR$(0))))
   FOR i = 1 TO nFiles
      szKey = "File " & FORMAT$(i)
      'lRes = GetPrivateProfileString(szSection, szKey, szDefault, szInString, %MAX_PATH, szProjectFile)
      szPath = IniRead(szProjectFile,szSection,szKey,"")
      szPath = TRIM$(szPath)
      IF szPath="" THEN
        ITERATE FOR
      END IF
      'IF lRes THEN
      '   szPath = TRIM$(EXTRACT$(szInString, CHR$(0)))
         hItem = TreeView_FindItem(ghProjectTV, GetFileName(szPath), %FALSE, %Null, %TRUE, %TRUE, hItem)
         IF hItem = 0 THEN
            REDIM PRESERVE udtProjectData(UBOUND(udtProjectData) + 1) AS PROJECT_PATH_NAMES
            udtProjectData(UBOUND(udtProjectData)).sFName = GetFileName(szPath)
            udtProjectData(UBOUND(udtProjectData)).sFPathName = GetFilePath(szPath)
         END IF
      'END IF
   NEXT

   szSection = "INC Files"
   szKey = "Files"
   'lRes = GetPrivateProfileString(szSection, szKey, szDefault, szInString, %MAX_PATH, szProjectFile)
   'IF lRes THEN nFiles = VAL(TRIM$(EXTRACT$(szInString, CHR$(0))))
   nFiles = VAL(IniRead(szProjectFile,szSection,szKey,"0"))
   FOR i = 1 TO nFiles
      szKey = "File " & FORMAT$(i)
      szPath = IniRead(szProjectFile,szSection,szKey,"")
      szPath=TRIM$(szPath)
      IF szPath="" THEN
        ITERATE FOR
      END IF
      'lRes = GetPrivateProfileString(szSection, szKey, szDefault, szInString, %MAX_PATH, szProjectFile)
      'IF lRes THEN
      '   szPath = TRIM$(EXTRACT$(szInString, CHR$(0)))
         hItem = TreeView_FindItem(ghProjectTV, GetFileName(szPath), %FALSE, %Null, %TRUE, %TRUE, hItem)
         IF hItem = 0 THEN
            REDIM PRESERVE udtProjectData(UBOUND(udtProjectData) + 1) AS PROJECT_PATH_NAMES
            udtProjectData(UBOUND(udtProjectData)).sFName = GetFileName(szPath)
            udtProjectData(UBOUND(udtProjectData)).sFPathName = GetFilePath(szPath)
         END IF
      'END IF
   NEXT

   szSection = "RC Files"
   szKey = "Files"
   'lRes = GetPrivateProfileString(szSection, szKey, szDefault, szInString, %MAX_PATH, szProjectFile)
   nFiles = VAL(IniRead(szProjectFile,szSection,szKey,"0"))
   'IF lRes THEN nFiles = VAL(TRIM$(EXTRACT$(szInString, CHR$(0))))
   FOR i = 1 TO nFiles
      szKey = "File " & FORMAT$(i)
      szPath= IniRead(szProjectFile,szSection,szKey,"")
      szPath=TRIM$(szPath)
      IF szPath="" THEN
        ITERATE FOR
      END IF
      'lRes = GetPrivateProfileString(szSection, szKey, szDefault, szInString, %MAX_PATH, szProjectFile)
      'IF lRes THEN
         'szPath = TRIM$(EXTRACT$(szInString, CHR$(0)))
         hItem = TreeView_FindItem(ghProjectTV, GetFileName(szPath), %FALSE, %Null, %TRUE, %TRUE, hItem)
         IF hItem = 0 THEN
            REDIM PRESERVE udtProjectData(UBOUND(udtProjectData) + 1) AS PROJECT_PATH_NAMES
            udtProjectData(UBOUND(udtProjectData)).sFName = GetFileName(szPath)
            udtProjectData(UBOUND(udtProjectData)).sFPathName = GetFilePath(szPath)
         END IF
      'END IF
   NEXT

   IF LEN(sPrimary) THEN
      IF GetFilePath(sPrimary) = ".\">0 THEN
         PrimarySourceFile = GetFilePath(sProjectName) & GetFileName(sPrimary)
      ELSE
         PrimarySourceFile =  sPrimary
      END IF
      sProjectPrimary = PrimarySourceFile
   END IF

   ARRAY SORT udtProjectData()
   FOR i = LBOUND(udtProjectData)  TO UBOUND(udtProjectData)
      sName = TRIM$(udtProjectData(i).sFName)
      TreeViewInsertFile(sName)
   NEXT

   TreeView_Expand(ghProjectTV, hPTRoot, %TVE_EXPAND)

'   IF gdi.l = "I" THEN gdi.l = gdi.sl
'   IF gdi.l <> "E" AND gdi.l <> "W" AND gdi.l <> "N" AND gdi.l <> "S" THEN gdi.l = "E"
'   SetParent ghTreeView, ghDockC
'   ShowWindow ghDockC, %SW_SHOW
'   SendMessage ghDockC, %WM_SIZE, 0, 0
'   SendMessage hWndMain, %WM_SIZE, 0, 0

   IF LEN(PrimarySourceFile) THEN
      szPath = PrimarySourceFile
      szPath = GetFileName(szPath)
'      SendMessage hStatusbar, %SB_SETTEXT, 4, VARPTR(szPath)
   END IF
   szProjectDir = GetFilePath(fName)
   WriteRecentProjects fName
   ChangeButtonsState

   TV_SetItemText ghProjectTV, hPTRoot, GetFileName(fName)

   FUNCTION = %TRUE

END FUNCTION
' *********************************************************************************************
' 保存最近打开工程列表
' *********************************************************************************************
SUB WriteRecentProjects (BYVAL OpenFName AS STRING)
  LOCAL Ac        AS LONG
  LOCAL szText    AS STRING 'ASCIIZ * %MAX_PATH
  LOCAL szSection AS STRING 'ASCIIZ * 30
  LOCAL szKey     AS STRING 'ASCIIZ * 30

  IF INSTR(OpenFName, ANY ":\/") <= 0 THEN   ' 路径不可用
    IF LEFT$(UCASE$(GetFileName(OpenFName)), 8) = "UNTITLED" THEN EXIT SUB
  END IF
  szSection = "Reopen projects"
  IF LEN(OpenFName) THEN
    ARRAY SCAN RecentProjects(), COLLATE UCASE, = UCASE$(OpenFName), TO Ac
    IF Ac THEN ARRAY DELETE RecentProjects(Ac)
    ARRAY INSERT RecentProjects(), OpenFName
  END IF
  FOR Ac = 1 TO 4
    szText = RecentProjects(Ac)
    IF ISFALSE FileExist(szText) THEN szText = ""
    szKey   = "Project " & FORMAT$(Ac)
    IniWrite g_zIni,szSection,szKey,szText
    'WritePrivateProfileString szSection, szKey, szText, IniFile
  NEXT
  GetRecentProjects   ' Update MRU menu
END SUB
' *********************************************************************************************
' 设置选中的树显示项文本Set the selected treview item text
' *********************************************************************************************
FUNCTION TV_SetItemText (BYVAL hTreeView AS DWORD, BYVAL hTvItem AS DWORD, szText AS ASCIIZ) AS LONG
  LOCAL udt_tv_item AS TV_ITEM              ' // 树显示项的属性
  IF hTvItem = 0 THEN EXIT FUNCTION
  udt_tv_item.hitem = hTvItem
  udt_tv_item.mask = %TVIF_TEXT
  udt_tv_item.psztext = VARPTR(szText)
  udt_tv_item.cchtextmax = SIZEOF(szText)
  FUNCTION = TreeView_SetItem(hTreeView, udt_tv_item)
END FUNCTION
' *********************************************************************************************
' 读取最近打开项目文件列表
' *********************************************************************************************
SUB GetRecentProjects
  LOCAL Bc        AS LONG
  LOCAL mii       AS MENUITEMINFO
  LOCAL InString  AS ASCIIZ * 300
  LOCAL i         AS LONG
  LOCAL fInserted AS LONG
  REDIM RecentProjects(1 TO 4) AS STRING

  FOR i = 1 TO 4  ' 清除重打开菜单
    IF RemoveMenu(g_hMenuProject, %IDM_RECENTPROJECTS + i, %MF_BYCOMMAND) = 0 THEN EXIT FOR
  NEXT
  RemoveMenu(g_hMenuProject, %IDM_RECENTPROJECTSSEPARATOR, %MF_BYCOMMAND)
  DrawMenuBar g_hWndMain
  Bc = 1
  FOR i = 1 TO 4
    InString = GetRecentProjectName(i)
    IF LEN(InString) THEN
      IF ISTRUE FileExist(InString) THEN
        RecentProjects(Bc) = InString
        InString = "&" & FORMAT$(Bc) & " " & InString
        AppendMenu g_hMenuProject, %MF_ENABLED, %IDM_RECENTPROJECTS + Bc, InString
        ModifyMenu g_hMenuProject, %IDM_RECENTPROJECTS + Bc, %MF_OWNERDRAW, %IDM_RECENTPROJECTS + Bc, BYVAL %NULL
        fInserted = %TRUE
        INCR Bc
      END IF
    END IF
  NEXT

  IF ISTRUE fInserted THEN
    mii.cbSize = SIZEOF(mii)
    mii.fMask  = %MIIM_TYPE OR %MIIM_ID OR %MIIM_SUBMENU
    mii.fType  = %MFT_STRING
    mii.wID = %IDM_RECENTPROJECTSSEPARATOR
    mii.fType = %MFT_SEPARATOR OR %MFT_OWNERDRAW
    InsertMenuItem g_hMenuProject, %IDM_RECENTPROJECTS + 1, %FALSE, mii
  END IF
END SUB
' *********************************************************************************************
' 得到最近项目文件名
' *********************************************************************************************
FUNCTION GetRecentProjectName(BYVAL idx AS LONG) AS STRING
  LOCAL szSection AS STRING 'ASCIIZ * 30
  LOCAL szKey     AS STRING 'ASCIIZ * 30
  LOCAL szDefault AS STRING 'ASCIIZ * 30
  LOCAL lRes      AS LONG
  LOCAL InString  AS STRING 'ASCIIZ * 300
  szSection = "Reopen projects"
  szKey = "Project " & FORMAT$(idx)
  'lRes = GetPrivateProfileString(szSection, szKey, szDefault, InString, %MAX_PATH, IniFile)
  InString = IniRead(g_zIni,szSection,szKey,"")
  InString = TRIM$(InString)
  'IF lRes THEN FUNCTION = TRIM$(EXTRACT$(InString, CHR$(0)))
  FUNCTION = InString
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
CALLBACK FUNCTION ProjectProc()AS LONG
  IF hProjectTV<=0 THEN
    hProjectTV=CB.HNDL
    'MSGBOX STR$(hProjectTV)
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
' 项目窗口大小调整，同步调整项目窗口中的控件位置大小
FUNCTION SizeProjectContents(BYVAL hWnd AS DWORD,BYVAL wParam AS DWORD)AS LONG
  LOCAL rc AS RECT
  LOCAL i AS LONG
  LOCAL gdi1 AS DockInfo
  LOCAL hCtrl AS DWORD
  RESET gdi1
  i=FindGdiIndex(%ID_PROJECT)
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
  MoveWin hProjectTV,rc,%TRUE
END FUNCTION
' 项目中按钮事件响应
FUNCTION Command_Project(BYVAL wParam AS DWORD)AS LONG
  'SendMessage hToolContain,%WM_COMMAND,wParam,0
END FUNCTION
' 工具箱窗口销毁事件响应
FUNCTION DestroyProjectWin() AS LONG
  'IF hToolBmp <> 0 THEN DELETEOBJECT hToolBmp
END FUNCTION
' 将工具箱窗口中的控件移到指定句柄窗口中：切换泊坞与浮动状态时发生
FUNCTION MoveProjectContents(BYVAL hWnd AS DWORD)AS LONG
  SetParent hProjectTV, hWnd
END FUNCTION
' 判断文件是否在当前项目里
FUNCTION FileIsInProject(BYVAL fname AS STRING)AS LONG
  LOCAL strFullName AS STRING
  LOCAL sFPathName  AS STRING
  LOCAL i           AS LONG
  IF UBOUND(udtProjectData())=-1 THEN
    REDIM udtProjectData()
    EXIT FUNCTION
  END IF
  FOR i = LBOUND(udtProjectData) TO UBOUND(udtProjectData)
    sFPathName = TRIM$(udtProjectData(i).sFPathName, ANY CHR$(32,0))
    IF sFPathName = ".\" THEN sFPathName = szProjectDir
    IF RIGHT$(sFPathName, 1) <> "\" THEN sFPathName = sFPathName & "\"
    strFullName = sFPathName & TRIM$(udtProjectData(i).sFName, ANY CHR$(32,0))
    IF UCASE$(strFullName) = UCASE$(fname) THEN
      FUNCTION = %TRUE
      EXIT FOR
    END IF
  NEXT
END FUNCTION
' *****************************************************************************
' 根据指定字符串查树表中的项，返回该项句柄
' *****************************************************************************
FUNCTION TreeView_FindItem( BYVAL hTreeView AS DWORD, BYVAL sQuery AS STRING, _
                            BYVAL nCaseSens AS LONG, _
                            BYVAL hStartItem AS DWORD, _
                            BYVAL nSearchChildren AS LONG, _
                            BYVAL nSelectItem AS LONG, _
                            OPTIONAL BYVAL hItem AS DWORD) AS DWORD
  LOCAL tvi AS TV_ITEM, zText AS ASCIIZ * 1024, hTmp AS DWORD
  tvi.mask       = %TVIF_TEXT OR %TVIF_CHILDREN
  tvi.pszText    = VARPTR(zText)
  tvi.cchTextMax = SIZEOF(zText)
  IF nCaseSens = %FALSE THEN sQuery = UCASE$(sQuery)
  ' 从 hStartItem 指定的项开始往下查所有子项
  tvi.hItem = SendMessage(hTreeView, %TVM_GETNEXTITEM, _
          IIF&(hItem, %TVGN_NEXT, %TVGN_CHILD), IIF&(hItem, hItem, hStartItem))
  DO UNTIL (tvi.hItem = %NULL)
    SendMessage hTreeView, %TVM_GETITEM, 0, VARPTR(tvi)
    IF nCaseSens = %FALSE THEN
      IF UCASE$(zText) = sQuery THEN EXIT DO
    ELSE
      IF zText = sQuery THEN EXIT DO
    END IF
    ' 查子项(we use this function again in a loop)...
    IF (nSearchChildren <> 0) AND (tvi.cChildren <> 0) THEN
      hTmp = TreeView_FindItem(hTreeView, sQuery, nCaseSens, tvi.hItem, %TRUE, nSelectItem, %Null)
      IF hTmp <> %Null THEN
        ' 找到了，返回句柄，退出
        FUNCTION = hTmp
        EXIT FUNCTION
     END IF
    END IF
    tvi.hItem = SendMessage(hTreeView, %TVM_GETNEXTITEM, %TVGN_NEXT, tvi.hItem)
  LOOP
  IF (tvi.hItem <> %NULL) AND (nSelectItem <> %FALSE) THEN
    ' 如果有要求的话，选中该项
    SendMessage(hTreeView, %TVM_SELECTITEM, %TVGN_CARET, tvi.hItem)
  END IF
  ' 返回项句柄
  FUNCTION = tvi.hItem
END FUNCTION
' *********************************************************************************************
' 插入文件到项目中
' *********************************************************************************************
FUNCTION ProjectFileInsert (fName AS STRING) AS DWORD
   LOCAL sPath  AS STRING
   LOCAL sName  AS STRING
   LOCAL sExt   AS STRING
   LOCAL lIndex AS LONG
   LOCAL hRet   AS LONG
   LOCAL hItem  AS LONG

   sPath = GetFilePath(fName)
   sName = GetFileName(fName)
   lIndex = INSTR(sName, ".")
   sExt = UCASE$(MID$(sName,lIndex + 1))
   IF sExt <> "BAS" AND _
      sExt <> "INC" AND _
      sExt <> "RC" THEN EXIT FUNCTION
   IF sExt = "RC" AND hPTRcParent <> 0 THEN
      IF SED_MsgBox(g_hWndMain, "替换当前资源文件?  ", _
             %MB_YESNO, "替换资源") = %IDNO THEN
         EXIT FUNCTION
      ELSE
         TreeView_DeleteItem(ghProjectTV, hPTRcParent)
         hPTRcParent = 0
      END IF
   END IF
   IF sPath = GetFilePath(sProjectName) THEN sPath = ".\"
   hItem = TreeView_FindItem(ghProjectTV, sName, %FALSE, %Null, %TRUE, %FALSE, hItem)
   IF hItem  = 0 THEN
      hRet = TreeViewInsertFile(sName)
      IF hRet THEN
         REDIM PRESERVE udtProjectData(UBOUND(udtProjectData) + 1)
         udtProjectData(UBOUND(udtProjectData)-1).sFName = sName
         udtProjectData(UBOUND(udtProjectData)-1).sFPathName = sPath
         SELECT CASE sExt
            CASE "BAS"
               TreeView_SortChildren(ghProjectTV, hPTBasParent,0)
            CASE "INC"
               TreeView_SortChildren(ghProjectTV, hPTIncParent,0)
            CASE "RC"
               TreeView_SortChildren(ghProjectTV,hPTRcParent,0)
         END SELECT
         TreeView_Select(ghProjectTV, hRet, %tvgn_caret)
      END IF
   END IF
   FUNCTION = hRet
END FUNCTION
' *********************************************************************************************
' 将文件插入项目树视图
' *********************************************************************************************
FUNCTION TreeViewInsertFile(sItem AS STRING) AS LONG
   LOCAL hItem AS LONG
   LOCAL sExt AS STRING
   LOCAL lIndex AS LONG

   lIndex = INSTR(sItem, ".")
   sExt = MID$(sItem, lIndex + 1)
   SELECT CASE UCASE$(sExt)
      CASE "BAS"
         IF hPTBasParent = 0 THEN
            hPTBasParent = TreeViewInsertItem(ghProjectTV, hPTRoot, "BAS", %IMAGE_FOLD)
            CALL TreeView_SortChildren(ghProjectTV, hPTROOT, 0)
         END IF

         IF GetFileName(PrimarySourceFile) = sItem THEN
            hItem = TreeViewInsertItem(ghProjectTV, hPTBasParent, sItem, %IMAGE_PRI)
         ELSE
            hItem = TreeViewInsertItem(ghProjectTV, hPTBasParent, sItem, %IMAGE_BAS)
         END IF
         TreeView_Expand(ghProjectTV, hPTBasParent, %TVE_EXPAND)
         FUNCTION = hItem

      CASE "INC"
         IF hPTIncParent = 0 THEN
            hPTIncParent = TreeViewInsertItem(ghProjectTV, hPTRoot, "INC", %IMAGE_FOLD)
            CALL TreeView_SortChildren(ghProjectTV, hPTROOT, 0)
         END IF
         hItem =  TreeViewInsertItem(ghProjectTV, hPTIncParent, sItem, %IMAGE_INC)
         TreeView_Expand(ghProjectTV, hPTIncParent, %TVE_EXPAND)
         FUNCTION = hItem

      CASE "RC"
         IF hPTRcParent = 0 THEN
            hPTRcParent = TreeViewInsertItem(ghProjectTV, hPTRoot, "RC", %IMAGE_FOLD)
            CALL TreeView_SortChildren(ghProjectTV, hPTROOT, 0)
         END IF
         hItem =  TreeViewInsertItem(ghProjectTV, hPTRcParent, sItem, %IMAGE_RC)
         TreeView_Expand(ghProjectTV, hPTRcParent, %TVE_EXPAND)
         FUNCTION = hItem
      CASE ELSE
         hItem = %FALSE
   END SELECT
   FUNCTION = hItem
END FUNCTION
' *********************************************************************************************
' 插入一项到树视图
' *********************************************************************************************
FUNCTION TreeViewInsertItem(BYVAL hTree AS DWORD, BYVAL hParent AS DWORD, sItem AS STRING, lImage AS LONG) AS LONG

   LOCAL tTVInsert   AS TV_INSERTSTRUCT
   LOCAL tTVItem     AS TV_ITEM

   IF hParent THEN
      tTVItem.mask        = %TVIF_CHILDREN OR %TVIF_HANDLE
      tTVItem.hItem       = hParent
      tTVItem.cchildren   = 1
      TreeView_SetItem hTree, tTVItem
   END IF

   tTVInsert.hParent                   = hParent
   tTVInsert.Item.mask            = %TVIF_TEXT OR %TVIF_IMAGE OR %TVIF_SELECTEDIMAGE
   tTVInsert.Item.pszText         = STRPTR(sItem)
   tTVInsert.Item.cchTextMax      = LEN(sItem)
   tTVInsert.Item.iImage          = lImage
   tTVInsert.Item.iSelectedImage  = lImage

   FUNCTION = TreeView_InsertItem(hTree, tTVInsert)

END FUNCTION
' *********************************************************************************************
' 保存项目树
' *********************************************************************************************
FUNCTION ProjectSaveTree(SaveNew AS LONG) AS LONG

   LOCAL hItem         AS LONG
   LOCAL tvi           AS TV_ITEM
   LOCAL zText         AS ASCIIZ  * 1024
   LOCAL sText         AS STRING
   LOCAL fOptions      AS STRING
   LOCAL PATH          AS STRING
   LOCAL STYLE         AS DWORD
   LOCAL f             AS STRING

   LOCAL szSection     AS ASCIIZ  * 255
   LOCAL szPath        AS ASCIIZ  * %MAX_PATH
   LOCAL szKey         AS ASCIIZ  * 255
   LOCAL szProjectFile AS ASCIIZ  * %MAX_PATH
   LOCAL idx           AS LONG

   IF LEN(sProjectName) = 0 THEN EXIT FUNCTION ' A project is not active
   IF UCASE$(sProjectName) = "NONAME.SPF" AND SaveNew = 0 THEN EXIT FUNCTION

   IF ISTRUE(SaveNew) THEN
      fOptions =          "PB项目文件 (*.PBP)|*.PBP"
      fOptions = fOptions & "所有文件 (*.*)|*.*"
      STYLE = %OFN_EXPLORER OR %OFN_FILEMUSTEXIST OR %OFN_HIDEREADONLY OR %OFN_OVERWRITEPROMPT
      IF ISFALSE(SaveFileDialog(g_hWndMain, "", f, PATH, fOptions, "SPF", STYLE)) THEN EXIT FUNCTION
      PATH = GetFileName(f)
      IF RIGHT$(UCASE$(PATH),4) <> ".SPF" THEN
         PATH = f & ".spf"
      ELSE
         PATH = f
      END IF
   ELSE
      PATH = sProjectName
   END IF

   tvi.mask       = %TVIF_TEXT OR %TVIF_CHILDREN
   tvi.pszText    = VARPTR(zText)
   tvi.cchTextMax = SIZEOF(zText)

   IF FileExist(PATH) THEN KILL PATH
   szProjectFile = PATH

   szSection = "BAS Files"
   idx = 0

   szKey = "Primary Source File"

   IF GetFilePath(PrimarySourceFile) = GetFilePath(sProjectName) THEN
     szPath = ".\" & GetFileName(PrimarySourceFile)
   ELSE
     szPath = PrimarySourceFile
   END IF

   WritePrivateProfileString szSection, szKey, szPath, szProjectFile

   ' Get .BAS files
   IF hPTBasParent <> 0 THEN
      tvi.hItem  = TreeView_GetNextItem(ghProjectTV, hPTBasParent, %tvgn_child)
      DO UNTIL (tvi.hItem = %NULL)
         TreeView_GetItem(ghProjectTV, tvi)
         sText =  zText
         IF ProjectArrayScan(sText) = GetFilePath(sProjectName) THEN
            sText = ".\" & zText
         ELSE
            sText = ProjectArrayScan(sText) & zText
         END IF
         INCR idx
         szKey = "File " & FORMAT$(idx)
         szPath = sText
         WritePrivateProfileString szSection, szKey, szPath, szProjectFile
         tvi.hItem  = TreeView_GetNextItem(ghProjectTV, tvi.hItem, %tvgn_next)
      LOOP
   END IF

   szKey = "Files"
   szPath = TRIM$(STR$(idx))
   WritePrivateProfileString szSection, szKey, szPath, szProjectFile

   szSection = "INC Files"
   idx = 0

   ' Get .INC files
   IF hPTIncParent <> 0 THEN
      tvi.hItem  = TreeView_GetNextItem(ghProjectTV, hPTIncParent, %tvgn_child)
      DO UNTIL (tvi.hItem = %NULL)
         TreeView_GetItem(ghProjectTV, tvi)
         sText =  zText
         IF ProjectArrayScan(sText) = GetFilePath(sProjectName) THEN
            sText = ".\" & zText
         ELSE
            sText = ProjectArrayScan(sText) & zText
         END IF
         INCR idx
         szKey = "File " & FORMAT$(idx)
         szPath = sText
         WritePrivateProfileString szSection, szKey, szPath, szProjectFile
         tvi.hItem  = TreeView_GetNextItem(ghProjectTV, tvi.hItem, %tvgn_next)
      LOOP
    END IF

   szKey = "Files"
   szPath = TRIM$(STR$(idx))
   WritePrivateProfileString szSection, szKey, szPath, szProjectFile

   szSection = "RC Files"
   idx = 0

   ' Get .RC file
   IF  hPTRcParent <> 0 THEN
      tvi.hItem  = TreeView_GetNextItem(ghProjectTV,hPTRcParent,%tvgn_child)
      DO UNTIL (tvi.hItem = %NULL)
         TreeView_GetItem(ghProjectTV, tvi)
         sText =  zText
         IF ProjectArrayScan(sText) = GetFilePath(sProjectName) THEN
            sText = ".\" & zText
         ELSE
            sText = ProjectArrayScan(sText) & zText
         END IF
         INCR idx
         szKey = "File " & FORMAT$(idx)
         szPath = sText
         WritePrivateProfileString szSection, szKey, szPath, szProjectFile
         tvi.hItem  = TreeView_GetNextItem(ghProjectTV,tvi.hItem,%tvgn_next)
      LOOP
    END IF

   szKey = "Files"
   szPath = TRIM$(STR$(idx))
   WritePrivateProfileString szSection, szKey, szPath, szProjectFile

'   ' Save file
'   IF ISTRUE(SaveNew) THEN
'      gdi.sl = gdi.l
'      ProjectFileClose
'      ProjectLoadTree PATH
'   END IF

   FUNCTION = %TRUE

END FUNCTION
' *********************************************************************************************
' Scan udt project Array
' *********************************************************************************************
FUNCTION ProjectArrayScan(sFind AS STRING) AS STRING

   LOCAL lLooper AS LONG

   FOR lLooper = LBOUND(udtProjectData) TO UBOUND(udtProjectData)
      IF TRIM$(udtProjectData(lLooper).sFName) = TRIM$(sFind) THEN
         FUNCTION = TRIM$(udtProjectData(lLooper).sFPathName)
         EXIT FUNCTION
      END IF
   NEXT

   FUNCTION = ""

END FUNCTION
