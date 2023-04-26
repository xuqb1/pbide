' *********************************************************************************************
' 包含作为MDI子窗口Tab控件相关
' Some of the functions are based on public domain code by Jules Marchildon
' http://www.powerbasic.com/support/forums/Forum7/HTML/001320.html
' *********************************************************************************************
DECLARE FUNCTION CreateTabMdiCtl( BYVAL hWnd AS LONG, rc AS RECT ) AS DWORD
DECLARE FUNCTION InsertTabMdiItem( BYVAL hTab AS DWORD, BYVAL ITEM AS LONG, szTabText AS ASCIIZ ) AS LONG
'DECLARE SUB DrawTabItem (BYVAL lParam AS LONG)
DECLARE SUB SetTabName( BYVAL nTab AS LONG, BYVAL strName AS STRING )
'DECLARE FUNCTION GetTabName (BYVAL nTab AS LONG) AS STRING
DECLARE SUB SED_ActivateMdiWindow( BYVAL strFilePath AS STRING )
' *********************************************************************************************
' 创建tab 控件
' *********************************************************************************************
FUNCTION CreateTabMdiCtl( BYVAL hWnd AS LONG, rc AS RECT ) AS DWORD
  LOCAL hWndTab     AS DWORD
  LOCAL hIcon       AS DWORD
  LOCAL ToolHeight  AS LONG
  LOCAL rctb        AS RECT
  GETWINDOWRECT hToolbar, rctb
  ToolHeight = rctb.nBottom - rctb.nTop
  hWndTab = CREATEWINDOWEX( %WS_EX_STATICEDGE, "SysTabControl32", BYVAL %NULL, _
                            %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %TCS_SINGLELINE OR %TCS_RIGHTJUSTIFY _
                            OR %TCS_TABS OR %TCS_FOCUSNEVER OR %TCS_TOOLTIPS, _
                            0, 0, 0, 0, _
                            hWnd, %IDC_TABMDI, GETMODULEHANDLE( "" ), BYVAL %NULL )
  hTabImageList = IMAGELIST_CREATE( 16, 16, %ILC_MASK, 2, 1 )
  hIcon = LOADICON( GETMODULEHANDLE( "" ), "SEDICON_EDIT" )
  ImageList_AddIcon hTabImageList, hIcon
  DESTROYICON hIcon
  TabCtrl_SetImageList hWndTab, hTabImageList
  FUNCTION = hWndTab
END FUNCTION
' *********************************************************************************************
' 根据给定的文件的路径激活窗口并加载文件
' *********************************************************************************************
SUB SED_ActivateMdiWindow( BYVAL strFilePath AS STRING )
  LOCAL hMdi    AS DWORD
  LOCAL szPath  AS ASCIIZ * %MAX_PATH
  LOCAL nTab    AS LONG
  GETWINDOWTEXT MdiGetActive( hWndClient ), szPath, SIZEOF( szPath )
  IF szPath = strFilePath THEN EXIT SUB
  ' 获取第一个句柄
  hMdi = GETWINDOW( hWndClient, %GW_CHILD )
  ' 循环，遍历所有已打开窗口
  WHILE hMdi
    ' 获取标题文本
    GETWINDOWTEXT( hMdi, szPath, %MAX_PATH )
    IF szPath = strFilePath THEN     '如果有已打开的文件与输入文件相同，则激活该Tab和编辑器
      ARRAY SCAN g_TabFilePaths( ), = szPath, TO nTab
      IF nTab THEN
        SENDMESSAGE ghTabMdi, %TCM_SETCURSEL, nTab - 1, 0       ' 激活相关的Tab标签
        IF ISICONIC( hMdi ) THEN                                '最小化状态，则恢复
          SENDMESSAGE hWndClient, %WM_MDIRESTORE, hMdi, 0
        ELSE
          SENDMESSAGE hWndClient, %WM_MDIACTIVATE, hMdi, 0
        END IF
        EXIT LOOP
      END IF
    END IF
    hMdi = GETWINDOW( hMdi, %GW_HWNDNEXT )
  WEND
END SUB
' *********************************************************************************************
' 插入Tab新标签
' *********************************************************************************************
FUNCTION InsertTabMdiItem( BYVAL hTab AS DWORD, BYVAL ITEM AS LONG, szTabText AS ASCIIZ ) AS LONG
  LOCAL ttc_item AS TC_ITEM
  IF SENDMESSAGE( ghTabMdi, %TCM_GETITEMCOUNT, 0, 0 ) = 0 THEN SHOWWINDOW ghTabMdi, %SW_SHOW
  ' 插入到Tab
  ttc_item.mask = %TCIF_TEXT OR _
                  %TCIF_IMAGE OR _
                  %TCIF_RTLREADING OR _
                  %TCIF_PARAM OR _
                  %TCIF_STATE OR _
                  %TCIS_BUTTONPRESSED OR _
                  %TCIS_HIGHLIGHTED
  ttc_item.pszText = VARPTR( szTabText )
  ttc_item.cchTextMax = LEN( szTabText )
  '   ttc_item.iImage     = -1
  ttc_item.iImage = 0
  ttc_item.lParam = 0
  FUNCTION = SENDMESSAGE( hTab, %TCM_INSERTITEM, ITEM, BYVAL VARPTR( ttc_item ))
END FUNCTION
' *********************************************************************************************
' 修改Tab标签名字
' *********************************************************************************************
SUB SetTabName( BYVAL nTab AS LONG, BYVAL strName AS STRING )
    LOCAL ttc_item AS TC_ITEM
    ttc_item.mask = %TCIF_TEXT
    ttc_item.pszText = STRPTR( strName )
    ttc_item.cchTextMax = 256
    SENDMESSAGE ghTabMdi, %TCM_SETITEM, nTab, BYVAL VARPTR( ttc_item )
END SUB
' *********************************************************************************************
' 再次重绘所有移除的和插入的Tab标签
' Workaround for a problem with the Tab control not redrawing properly when there are many
' tabs and you delete one tab, causing the deletion to scroll the tabs.
' *********************************************************************************************
SUB UpdateSysTab
    LOCAL ttc_item AS TC_ITEM
    LOCAL i AS LONG
    LOCAL szFileName AS ASCIIZ * 256        ' %MAX_PATH
    '   SendMessage ghTabMdi, %WM_SETREDRAW, 0, 0
    SENDMESSAGE ghTabMdi, %TCM_DELETEALLITEMS, 0, 0
    ttc_item.Mask = %TCIF_TEXT OR %TCIF_IMAGE
    ttc_item.cchTextMax = SIZEOF( szFileName )      ' %MAX_PATH
    ttc_item.iImage = 0
    FOR i = LBOUND( g_TabFilePaths ) TO UBOUND( g_TabFilePaths )
        szFileName = GetFileName( g_TabFilePaths( i ))
        ttc_item.pszText = VARPTR( szFilename )
        SENDMESSAGE ghTabMdi, %TCM_INSERTITEM, i, BYVAL VARPTR( ttc_item )
    NEXT
    '   SendMessage ghTabMdi, %WM_SETREDRAW, 1, 0
    '   InvalidateRect ghTabMdi, BYVAL 0, 1
    '   UpdateWindow ghTabMdi
END SUB
' *********************************************************************************************
