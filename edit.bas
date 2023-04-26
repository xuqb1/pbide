'edit.bas
'����ʵ�ִ���༭����

GLOBAL EdOpt  AS EditorOptionsType
GLOBAL CpOpt  AS CompilerOptionsType
GLOBAL SciColorsAndFonts AS SciColorsAndFontsType
'GLOBAL sAutoCompletionTypes AS STRING
'GLOBAL gTabFilePaths() AS STRING
'GLOBAL cUntitledFiles AS LONG

FUNCTION  initEditClass() AS LONG
  LOCAL szClassName AS ASCIIZ * 80
  LOCAL wce AS WNDCLASSEX
  LOCAL hResult AS LONG

  '----------------------------------------------------------------------------
  ' ע����봰����
  szClassName       = $EDITCLASSNAME
  wce.cbSize        = SIZEOF(wce)
  wce.STYLE         = 0 '%CS_DBLCLKS OR %CS_HREDRAW OR %CS_VREDRAW
  wce.lpfnWndProc   = CODEPTR(CodeProc)
  wce.cbClsExtra    = 0
  wce.cbWndExtra    = 4
  wce.hInstance     = g_hInst
  wce.hIcon         = LoadIcon(g_hInst, "NOTEICON")
  IF wce.hIcon = 0 THEN '���û����Դͼ�꣬��ʹ��ϵͳ��Ӧ�ó���ͼ��
     wce.hIcon = LoadIcon(0, BYVAL %IDI_APPLICATION)
  END IF
  wce.hCursor       = %NULL 'LoadCursor(%NULL, BYVAL %IDC_IBEAM)
  wce.hbrBackground = %NULL
  wce.lpszMenuName  = %NULL
  wce.lpszClassName = VARPTR(szClassName)
  wce.hIconSm       = LoadIcon(g_hInst, BYVAL %IDI_APPLICATION)
  hResult= RegisterClassEx(wce)
  IF ISFALSE hResult THEN
      hResult=RegisterClass(BYVAL (VARPTR(wce) + 4))
  END IF
  FUNCTION=hResult
END FUNCTION
' *****************************************************************************
' MDI����༭���Ӵ��ڴ������
' �þ����Ϣ����MDI�Ӵ���.
' *****************************************************************************
FUNCTION CodeProc( BYVAL hWnd AS DWORD, BYVAL wMsg AS LONG,_
        BYVAL wParam AS DWORD, BYVAL lParam AS LONG ) AS LONG
  ' ���洰�ڱ���ı���,�����ļ���·��������
  LOCAL szText            AS ASCIIZ * %MAX_PATH
  LOCAL RetVal            AS LONG               ' ����ֵ
  LOCAL dwRes             AS DWORD              ' GetWindow����ֵ
  LOCAL fTime             AS DWORD              ' �ļ�ʱ��
  LOCAL fSavedTime        AS DWORD              ' �����ļ�ʱ��
  LOCAL hPopupMenu        AS DWORD              ' �����˵����
  LOCAL x                 AS LONG               ' �ֿ�ʼλ��
  LOCAL y                 AS LONG               ' �ֽ���λ��
  LOCAL txtrg             AS TEXTRANGE          ' �ı���Χ�ṹ
  LOCAL p                 AS LONG               ' λ��
  LOCAL curPos            AS LONG               ' ��ǰλ��
  LOCAL strTxt            AS STRING             ' һ��Ŀ�����
  LOCAL i                 AS LONG               ' ѭ������
  LOCAL WFD               AS WIN32_FIND_DATA    ' �����ļ��ṹ
  LOCAL CpOpt             AS CompilerOptionsType' ������ѡ��
  STATIC buffer           AS STRING             ' һ��Ŀ�껺��
  LOCAL nTabs             AS DWORD              ' tab����
  LOCAL nTab              AS DWORD              ' tab��
  LOCAL Pt                AS POINTAPI           ' POINTAPI�ṹ
  LOCAL fHasItems         AS LONG               ' ��ʶ
  LOCAL bitnum            AS LONG               ' λ��
  LOCAL cp                AS LONG               ' ���λ��
  LOCAL selStart          AS LONG               ' ѡ��ʼλ��
  LOCAL selEnd            AS LONG               ' ѡ�����λ��
  LOCAL tmpPROC           AS PROC               ' PROC�ṹ
  STATIC rc               AS RECT
  LOCAL tmpRc             AS RECT
  LOCAL tmpHeight         AS LONG
  LOCAL zText             AS ASCIIZ * 255
  LOCAL wi AS WININFOSTRUC PTR
  LOCAL tmpStr            AS STRING
  ' �õ��Ӵ������ݽṹ�����ڳ�ʼ��
  wi = GETWINDOWLONG( hWnd, %GWL_USERDATA )
  SELECT CASE wMsg
    CASE %WM_CREATE
      ' ����һ��Scintilla�ؼ���ʵ��
      ' �����һ���򿪶����������Ӧ���ļ�
      ' ��ָ�뵽�Ӵ��ڵķָ���Ϣ�ṹ��.
      wi = HEAPALLOC( GETPROCESSHEAP, %HEAP_ZERO_MEMORY, SIZEOF( @wi ))
      IF wi THEN
        SETWINDOWLONG hWnd, %GWL_USERDATA, wi   ' ��������õ���ָ�뵽�û�������
      ELSE
        FUNCTION = ( - 1 ) : EXIT FUNCTION      ' �����ڴ�ʧ�ܣ������˳�
      END IF
      CreateSciControl hWnd, @wi
      ' �����µ�tab
      nTabs = SENDMESSAGE( g_hTabMdi, %TCM_GETITEMCOUNT, 0, 0 )
      DIALOG GET TEXT hWnd TO tmpStr
      IF LEFT$(tmpStr,8)="Untitled" THEN
        IF g_FileType=1 THEN
          tmpStr=LEFT$(tmpStr,LEN(tmpStr)-3) & "INC"
        ELSEIF g_FileType=2 THEN
          tmpStr=LEFT$(tmpStr,LEN(tmpStr)-3) & "RC"
        ELSEIF g_FileType=3 THEN
          tmpStr=LEFT$(tmpStr,LEN(tmpStr)-3) & "TXT"
        END IF
        IF g_FileType>=1 THEN
          DIALOG SET TEXT hWnd,tmpStr
        END IF
      END IF
      szText=tmpStr

      'msgInfobox szText
      REDIM PRESERVE gTabFilePaths( UBOUND( gTabFilePaths ) + 1 )
      gTabFilePaths( UBOUND( gTabFilePaths )) = szText
      szText = GetFileName( szText )
      InsertTabMdiItem( g_hTabMdi, nTabs, szText )
      ' ��tab�ؼ���������Ϊ��ǰѡ��
      SENDMESSAGE( g_hTabMdi, %TCM_SETCURSEL, nTabs, 0 )
      ShowWindow g_hTabMdi,%SW_SHOW
      '������׸��ĵ�����ʾ��ذ�ť,0��ʾ��ʾ,1��ʾ����
      IF MdiGetActive(g_hWndClient) = 0 THEN HideButtons 0
      EXIT FUNCTION
    CASE %WM_MDIACTIVATE
      IF lParam = hWnd THEN
        ' ˢ�´������combobox�������б�
        ' ���������ⲿ�ִ���,���Ե������ĵ���ѡ��ʱ,
        ' �������combobox�½��б��������ɵ�ǰ�ļ�����������.
        ' �������ļ��н�������(����Ҫˢ�������б�Ϳ���֪���ҵ�����ЩSUB�ͺ���)
        ' ��״̬����ʾ���̵�����.
        ResetCodeFinder
        IF ISTRUE ShowProcedureName THEN
          SED_WithinProc( tmpPROC )
          szText = tmpPROC.ProcName
          SENDMESSAGE g_hStatus, %SB_SETTEXT, 4, VARPTR( szText )
        END IF
        EXIT FUNCTION
      END IF
    CASE %WM_MOUSEMOVE
      STATIC Ysplit AS LONG, rc2 AS RECT
      pt.x = LOWRD( lParam )
      pt.y = HIWRD( lParam )
      IF GETCAPTURE = hWnd THEN       ' �ƶ��ָ���
        CLIPCURSOR rc2
        IF Ysplit THEN                ' �ָ����Ǻ���ģ��������ƶ�
          @wi.SplitY = pt.y           ' ������ĵ�ǰY����
          IF pt.y < rc.nBottom - %SPLITSIZE AND CINT( pt.y ) = > 0 THEN
          ELSE                        ' ����Y�������ӿͻ���������
            @wi.SplitY = 0
            RELEASECAPTURE
            CLIPCURSOR BYVAL %NULL
            SETCURSOR hCurNorm
            AssignFocus @wi
          END IF
        ELSE                          ' �ָ���������ģ��������ƶ�
          IF pt.y > @wi.SplitY THEN   ' ����ں���ָ���������ʱ
            @wi.SplitXB = pt.x        ' �������X����
          ELSE
            @wi.SplitXT = pt.x
          END IF
          IF pt.x < rc.nRight - %SPLITSIZE AND CINT( pt.x ) = > 0 THEN
          ELSE                        ' ����X�������ӿͻ�������
            IF pt.y > @wi.SplitY + %SPLITSIZE THEN
              @wi.SplitXB = 0
            ELSE
              @wi.SplitXT = 0
            END IF
            RELEASECAPTURE
            CLIPCURSOR BYVAL %NULL
            SETCURSOR hCurNorm
            AssignFocus @wi
          END IF
        END IF
        DrawWindow hWnd, %TRUE
      ELSE
        ' ������û�б��ػ�,������ڷָ�������ʱ���ǽ����õ�һ��WM_MOUSEMOVE��Ϣ.
        ' ������ʾ��ȷ�ķָ����
        ' �������ƿ��ָ���λ��ʱ,���ڻ��Զ��ı��굽����״̬
        GETCLIENTRECT hWnd, rc
        SETRECT rc2, rc.nLeft, @wi.SplitY, rc.nRight, @wi.SplitY + %SPLITSIZE
        IF PTINRECT( rc2, pt.x, pt.y ) THEN
          SETCURSOR hCurSplit
          Ysplit = %TRUE
        ELSE
          SETCURSOR hCurSplitV
          Ysplit = %FALSE
        END IF
      END IF
      EXIT FUNCTION
    CASE %WM_LBUTTONDOWN
      IF GETCURSOR = hCurSplit OR GETCURSOR = hCurSplitV THEN
        SETCAPTURE hWnd
        IF Ysplit THEN
          pt.x = LOWRD( lParam ) : pt.y = @wi.SplitY
          ' "����" ��굽����ָ����ĵ�ǰXλ��.
          SETRECT rc2, pt.x, - 1, pt.x + 1, rc.nBottom
        ELSE
          pt.y = HIWRD( lParam )
          ' ȷ��ʹ���ĸ�����ָ�����λ��
          pt.x = IIF& ( pt.y > @wi.SplitY, @wi.SplitXB, @wi.SplitXT )
          ' "����"��굽��ѡ������ָ����ĵ�ǰYλ��.
          SETRECT rc2, - 1, pt.y, rc.nRight, pt.y + 1
        END IF
        MAPWINDOWPOINTS hWnd, BYVAL %NULL, rc2, 2 ' ת��rc2����Ļ����ϵ
        CLIENTTOSCREEN hWnd, pt
        SETCURSORPOS pt.x, pt.y                   ' ���ù�굽��ѡ�ָ���
        EXIT FUNCTION
      END IF
    CASE %WM_LBUTTONUP
      IF GETCAPTURE = hWnd THEN
        RELEASECAPTURE
        CLIPCURSOR BYVAL %NULL
        AssignFocus @wi
        EXIT FUNCTION
      END IF
    CASE %WM_PAINT
      ' ֻ��Ҫ���Ʒָ���
      DrawWindow hWnd, %FALSE
    CASE %WM_SIZE
      IF wParam <> %SIZE_MINIMIZED THEN
        IF @wi.SplitY THEN      ' �������Ŵ��ڵ��Ӵ��ڳߴ�ĳ���ֵ
            @wi.SplitY = MULDIV( HIWRD( lParam ), @wi.SplitY, @wi.OldY )
          IF @wi.SplitXT THEN
            @wi.SplitXT = MULDIV( LOWRD( lParam ), @wi.SplitXT, @wi.OldX )
          END IF
        END IF
        IF @wi.SplitXB THEN
          @wi.SplitXB = MULDIV( LOWRD( lParam ), @wi.SplitXB, @wi.OldX )
        END IF
        @wi.OldY = HIWRD( lParam ) : @wi.OldX = LOWRD( lParam )
        SETRECT rc, 0, 0, @wi.OldX, @wi.OldY
        DrawWindow hWnd, %TRUE
      END IF
    CASE %WM_SETFOCUS
      ' Ϊ�Ժ����ؽ��㷢��һ����Ϣ
      ' (����������ھ�������,ĳЩ������͵�߽���)
      POSTMESSAGE hWnd, %WM_USER + 999, GetEdit, 0
      EXIT FUNCTION
    CASE %WM_USER + 999
      ' �Ӵ��ڱ���õ�����
      DIALOG GET TEXT hWnd TO tmpStr
      szText=tmpStr
      ' �õ��ļ�����޸�ʱ��
      fTime = SED_GetFileTime( tmpStr )
      ' �õ��洢ʱ��
      fSavedTime = GETPROP( hWnd, "FTIME" )
      ' �������޸�ʱ����ļ��洢ʱ�䲻ͬ,��ѯ���Ƿ�����,�ǣ��򱣴���ʱ��
      IF fSavedTime THEN
        IF fSavedTime <> fTime THEN
          IF MsgQueBox(tmpStr & $CRLF & _
                  "�ļ��Ѿ��������������޸ġ�Ҫ��������") = %IDYES THEN
            OpenThisFile tmpStr
          ELSE
            SETPROP hWnd, "FTIME", fTime
          END IF
        END IF
      END IF
      IF ISTRUE wParam THEN SETFOCUS wParam   ' ���ý���
      ShowLinCol                              ' ��ʾ�к���
      ' ���������tab
      ARRAY SCAN gTabFilePaths( ), = szText, TO nTab
      IF nTab THEN SENDMESSAGE g_hTabMdi, %TCM_SETCURSEL, nTab - 1, 0
      ' ��Treeview��ѡ���ļ�
      IF FileIsInProject( tmpStr ) THEN
        TreeView_FindItem( hProjectTV, GetFileName(tmpStr), %FALSE, _ '2017/11/23 21:10
                hPTROOT, %TRUE, %TRUE )
      END IF
      EXIT FUNCTION
    CASE %WM_DRAWITEM
      ' �Ի��Ʋ˵���Ҫ�ػ��ơ�
      IF wParam = 0 THEN        ' ������Ϊ0����ʾ��Ϣ�ɲ˵�����
        DrawMenu lParam         ' ���Ʋ˵�
        FUNCTION = %TRUE
        EXIT FUNCTION
      END IF
    CASE %WM_MEASUREITEM          ' �õ����С
      IF wParam = 0 THEN          ' �ɲ˵�����
        MeasureMenu hWnd, lParam  ' �ڶ�����sub�д������й���
        FUNCTION = %TRUE
        EXIT FUNCTION
      END IF
    CASE %WM_CONTEXTMENU '�Ҽ��˵�
      ' ---------------------------------------------------------------
      ' ���û��ѡ���ı���ı佹�㼰���ַ�λ��
      ' ---------------------------------------------------------------
      SETFOCUS wParam
      pt.x = LOWRD( lParam )
      pt.y = HIWRD( lParam )
      SCREENTOCLIENT wParam, pt
      cp = SENDMESSAGE( wParam, %SCI_POSITIONFROMPOINT, pt.x, pt.y )
      selStart = SENDMESSAGE( wParam, %SCI_GETSELECTIONSTART, 0, 0 )
      selEnd = SENDMESSAGE( wParam, %SCI_GETSELECTIONEND, 0, 0 )
      IF selStart = selEnd THEN SENDMESSAGE wParam, %SCI_SETSEL, cp, cp
      ' ---------------------------------------------------------------
      x = 0
      y = 0
      buffer = ""
      ' �õ���ǰλ��
      curPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )
      ' �õ����ʵĿ�ʼλ��
      x = SENDMESSAGE( GetEdit, %SCI_WORDSTARTPOSITION, curPos, %TRUE )
      ' �õ����ʵĽ���λ��
      y = SENDMESSAGE( GetEdit, %SCI_WORDENDPOSITION, curPos, %FALSE )
      IF y > x THEN
        buffer = SPACE$( y - x + 1 )
        ' �ı���Χ
        txtrg.chrg.cpMin = x
        txtrg.chrg.cpMax = y
        txtrg.lpstrText = STRPTR( buffer )
        SENDMESSAGE GetEdit, %SCI_GETTEXTRANGE, 0, BYVAL VARPTR( txtrg )
        ' ɾ��$NUL
        p = INSTR( buffer, CHR$( 0 ))
        IF p THEN buffer = LEFT$( buffer, p - 1 )
      END IF
      buffer = REMOVE$( buffer, ANY CHR$( 13, 10 ))
      buffer = REMOVE$( buffer, ANY "()%," & $DQ )
      hPopupMenu = CREATEPOPUPMENU
      fHasItems = %FALSE      '���봰�ڴ���
      GETWINDOWTEXT hWnd, szText, SIZEOF( szText )
      szText = GetFileName( szText )
      'sProjectName="test"
      IF ISFALSE TreeView_FindItem( ghTreeView, szText, %FALSE, %NULL, %TRUE, _
              %FALSE ) AND LEN( sProjectName )>0 AND _
              LEFT$( UCASE$( szText ), 8 ) <> "UNTITLED" THEN
        APPENDMENU hPopupMenu,%MF_OWNERDRAW OR %MF_ENABLED,%IDM_INSERTINPROJECT,""
        fHasItems = %TRUE
      END IF
      IF ( GETMENUSTATE( g_hMenu, %IDM_INITASNEW, %MF_BYCOMMAND ) AND %MF_GRAYED ) <> %MF_GRAYED THEN
        APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_INITASNEW,""
      END IF
      IF fHasItems THEN
        APPENDMENU hPopupMenu, %MF_SEPARATOR OR %MF_OWNERDRAW, 0, ""
        fHasItems = 0
      END IF
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_UNDO,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_REDO,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_CUT,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_COPY,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_PASTE,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_PASTEIE,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_CLEAR,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_CLEARALL,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_LINEDELETE,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_SELALL,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_INDENT,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_OUTDENT,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_COMMENT,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_UNCOMMENT,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_FORMATREGION,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_TABULATEREGION,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_SELTOUPPERCASE,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_SELTOLOWERCASE,""
      APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, %IDM_SELTOMIXEDCASE,""
      IF SENDMESSAGE(GetEdit,%SCI_GETSELECTIONSTART,0,0)= _
              SENDMESSAGE( GetEdit,%SCI_GETSELECTIONEND,0,0) THEN
        IF UCASE$( RIGHT$( buffer, 4 )) = ".BAS" OR _
                UCASE$( RIGHT$( buffer, 4 )) = ".INC" THEN
          APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_SEPARATOR , 0, ""
          strContextMenu = buffer
          APPENDMENU hPopupMenu,%MF_OWNERDRAW OR %MF_ENABLED,%IDM_GOTOSELFILE, _
                  BYCOPY strContextMenu
        ELSE
          IF TRIM$( buffer ) <> "" THEN
            APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_SEPARATOR, 0, ""
            strContextMenu = buffer
            IF LEN( strContextMenu ) > 64 THEN
              strContextMenu = LEFT$( strContextMenu, 64 )
            END IF
            APPENDMENU hPopupMenu,%MF_OWNERDRAW OR %MF_ENABLED,%IDM_GOTOSELPROC,_
                    BYCOPY strContextMenu
          END IF
        END IF
        ' δѡ������ʱ�����Ҽ��˵�ɾ�����õĲ˵����ʹ�˵����׹���
        REMOVEMENU hPopupMenu, %IDM_CUT, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_COPY, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_CLEAR, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_INDENT, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_OUTDENT, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_COMMENT, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_UNCOMMENT, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_FORMATREGION, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_TABULATEREGION, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_SELTOUPPERCASE, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_SELTOLOWERCASE, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_SELTOMIXEDCASE, %MF_BYCOMMAND
      END IF
      IF SED_LastPosition.Position <> - 1 THEN '�������λ�ã�����ʾת����󴦲˵�
        APPENDMENU hPopupMenu, %MF_SEPARATOR OR %MF_OWNERDRAW, 0, ""
        APPENDMENU hPopupMenu, %MF_OWNERDRAW OR %MF_ENABLED, _
                %IDM_GOTOLASTPOSITION, ""
      END IF
      ' ����Ҫ����ʱ�����Ƴ������˵���
      IF ISFALSE SENDMESSAGE( GetEdit, %SCI_CANUNDO, 0, 0 ) THEN
        REMOVEMENU hPopupMenu, %IDM_UNDO, %MF_BYCOMMAND
      END IF
      ' ����Ҫ����ʱ�����Ƴ������˵���
      IF ISFALSE SENDMESSAGE( GetEdit, %SCI_CANREDO, 0, 0 ) THEN
        REMOVEMENU hPopupMenu, %IDM_REDO, %MF_BYCOMMAND
      END IF
      ' ����Ҫճ��ʱ�����Ƴ�ճ����ز˵���
      IF ISFALSE SENDMESSAGE( GetEdit, %SCI_CANPASTE, 0, 0 ) THEN
        REMOVEMENU hPopupMenu, %IDM_PASTE, %MF_BYCOMMAND
        REMOVEMENU hPopupMenu, %IDM_PASTEIE, %MF_BYCOMMAND
      END IF
      GETCURSORPOS Pt
      TRACKPOPUPMENU hPopupMenu, 0, Pt.x, Pt.y, 0, hWnd, BYVAL 0
      DESTROYMENU hPopupMenu
      UPDATEWINDOW GetEdit
    CASE %WM_COMMAND
      IF HIWRD( wParam ) = %SCEN_SETFOCUS THEN
        LOCAL dwEnable AS DWORD
        @wi.hFocus = lParam
        ' �����������ڱ༭����ʼ���Ŀؼ�(hBotR)
        ENABLEMENUITEM g_hMenuEdit, %IDM_INITASNEW, _
            IIF& ( lParam = @wi.hBotR, %MF_GRAYED, %MF_ENABLED )
        dwEnable = %MF_ENABLED
        IF @wi.bInit AND lParam <> @wi.hBotR THEN
          ' ����˿��Ʊ�����ʼ���������������Ϊ���͡��򿪡�
          ' IF BIT(@wi.bInit, GetDlgCtrlID(lParam) - %IDC_EDIT2) THEN
          ' -- Jos?Roca: ���Ӱ�ȫ�ж�
          bitnum = GETDLGCTRLID( lParam ) - %IDC_EDIT2
          IF bitnum = > 0 AND bitnum < 16 THEN
            IF BIT( @wi.bInit, bitnum ) THEN
              dwEnable = %MF_GRAYED
              ' ʹ�����桯ѡ����á�
              ' SendMessage @wi.hFocus, %SCI_EMPTYUNDOBUFFER, %NULL, %NULL
              ' ��ճ����Ļ������,ʹ����������
              ' ���ʹ�� %SCI_SETSAVEPOINT ������δ�޸ĵ�״̬
              SENDMESSAGE @wi.hFocus, %SCI_SETSAVEPOINT, 0, 0
            END IF
          END IF
        END IF
        IF ISFALSE SENDMESSAGE( @wi.hFocus, %SCI_GETMODIFY, %NULL, %NULL ) THEN
          IF ISTRUE SENDMESSAGE( ghToolbar, %TB_ISBUTTONENABLED, %IDM_SAVE, 0 ) THEN
            SENDMESSAGE ghToolbar, %TB_ENABLEBUTTON, %IDM_SAVE, %FALSE
          END IF
          ENABLEMENUITEM g_hMenuFile, %IDM_SAVE, %MF_GRAYED
        ELSE
          IF ISFALSE SENDMESSAGE( ghToolbar, %TB_ISBUTTONENABLED, %IDM_SAVE, 0 ) THEN
            SENDMESSAGE ghToolbar, %TB_ENABLEBUTTON, %IDM_SAVE, %TRUE
          END IF
          ENABLEMENUITEM g_hMenuFile, %IDM_SAVE, %MF_ENABLED
        END IF
        '            EnableMenuItem GetMenu(hWndMain), %IDM_OPEN, dwEnable
        ENABLEMENUITEM g_hMenuFile, %IDM_OPEN, dwEnable
        dwEnable = dwEnable XOR 1
        SENDMESSAGE ghToolBar, %TB_ENABLEBUTTON, %IDM_OPEN, dwEnable
        EXIT FUNCTION
      END IF
      ' *** NTM ***----------------------------------------------------------------------------------
      SELECT CASE LOWRD( wParam )
        CASE %IDM_INITASNEW :       SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_INITASNEW, 0
        CASE %IDM_UNDO :            SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_UNDO, 0
        CASE %IDM_REDO :            SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_REDO, 0
        CASE %IDM_CLEAR :           SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_CLEAR, 0
        CASE %IDM_CLEARALL :        SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_CLEARALL, 0
        CASE %IDM_CUT :             SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_CUT, 0
        CASE %IDM_COPY :            SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_COPY, 0
        CASE %IDM_PASTE :           SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_PASTE, 0
        CASE %IDM_PASTEIE :         SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_PASTEIE, 0
        CASE %IDM_SELALL :          SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_SELALL, 0
        CASE %IDM_LINEDELETE :      SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_LINEDELETE, 0
        CASE %IDM_INDENT :          SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_INDENT, 0
        CASE %IDM_OUTDENT :         SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_OUTDENT, 0
        CASE %IDM_COMMENT :         SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_COMMENT, 0
        CASE %IDM_UNCOMMENT :       SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_UNCOMMENT, 0
        CASE %IDM_FORMATREGION :    SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_FORMATREGION, 0
        CASE %IDM_TABULATEREGION :  SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_TABULATEREGION, 0
        CASE %IDM_SELTOUPPERCASE :  SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_SELTOUPPERCASE, 0
        CASE %IDM_SELTOLOWERCASE :  SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_SELTOLOWERCASE, 0
        CASE %IDM_SELTOMIXEDCASE :  SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_SELTOMIXEDCASE, 0
        CASE %IDM_INSERTINPROJECT
          GETWINDOWTEXT hWnd, szText, SIZEOF( szText )
          strTxt = szText
          ProjectFileInsert strTxt
          ProjectSaveTree %FALSE
        CASE %IDM_GOTOSELFILE
          ' -----------------------------------------------------------------
          ' �õ������ַ����µ��ļ���
          ' ��Ҫʹ��ALT+F7������
          ' -----------------------------------------------------------------
          x = 0 : y = 0 : buffer = ""
          '�õ���ǰλ��
          curPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )
          '�õ����ʵĿ�ʼλ��
          x = SENDMESSAGE( GetEdit, %SCI_WORDSTARTPOSITION, curPos, %TRUE )
          '�õ����ʵĽ���λ��
          y = SENDMESSAGE( GetEdit, %SCI_WORDENDPOSITION, curPos, %FALSE )
          IF y > x THEN
            buffer = SPACE$( y - x + 1 )
            '�ı���Χ
            txtrg.chrg.cpMin = x
            txtrg.chrg.cpMax = y
            txtrg.lpstrText = STRPTR( buffer )
            SENDMESSAGE GetEdit, %SCI_GETTEXTRANGE, 0, BYVAL VARPTR( txtrg )
            ' ɾ��$NUL
            p = INSTR( buffer, CHR$( 0 ))
            IF p THEN buffer = LEFT$( buffer, p - 1 )
          END IF
          buffer = REMOVE$( buffer, ANY CHR$( 13, 10 ))
          buffer = REMOVE$( buffer, ANY "()%," & $DQ )
          IF UCASE$( RIGHT$( buffer, 4 )) <> ".BAS" AND _
                  UCASE$( RIGHT$( buffer, 4 )) <> ".INC" THEN
            EXIT SELECT
          END IF
          ' -----------------------------------------------------------------
          GETWINDOWTEXT hWnd, szText, SIZEOF( szText )
          szText = GetFilePath( szText )
          szText = szText & buffer
          IF FINDFIRSTFILE( szText, WFD ) = %INVALID_HANDLE_VALUE THEN
            GetCompilerOptions CpOpt
            strTxt = CpOpt.PBWINIncPath
            RetVal = PARSECOUNT( strTxt, ";" )
            FOR i = 1 TO RetVal
              szText = PARSE$( strTxt, ";", i ) & "\" & buffer
              IF FINDFIRSTFILE( szText, WFD ) <> %INVALID_HANDLE_VALUE THEN
                OpenThisFile szText
                EXIT FUNCTION
              END IF
            NEXT
            SED_MsgBox( g_hWndMain, buffer & " δ�ҵ�", _
                    %MB_OK OR %MB_ICONINFORMATION, "�ļ�δ�ҵ�" )
            EXIT FUNCTION
          END IF
          OpenThisFile szText
        CASE %IDM_GOTOSELPROC
          ' -----------------------------------------------------------------
          ' �õ����ַ����µ��ļ���
          ' ��Ҫʹ��Ctrl+F7������
          ' -----------------------------------------------------------------
          x = 0 : y = 0 : buffer = ""
          ' �õ���ǰλ��
          curPos = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )
          ' �õ����ʵĿ�ʼλ��
          x = SENDMESSAGE( GetEdit, %SCI_WORDSTARTPOSITION, curPos, %TRUE )
          ' �õ����ʵĽ���λ��
          y = SENDMESSAGE( GetEdit, %SCI_WORDENDPOSITION, curPos, %FALSE )
          IF y > x THEN
            buffer = SPACE$( y - x + 1 )
            ' ���ʷ�Χ
            txtrg.chrg.cpMin = x
            txtrg.chrg.cpMax = y
            txtrg.lpstrText = STRPTR( buffer )
            SENDMESSAGE GetEdit, %SCI_GETTEXTRANGE, 0, BYVAL VARPTR( txtrg )
            '  ɾ��$NUL
            p = INSTR( buffer, CHR$( 0 ))
            IF p THEN buffer = LEFT$( buffer, p - 1 )
          END IF
          buffer = REMOVE$( buffer, ANY CHR$( 13, 10 ))
          buffer = REMOVE$( buffer, ANY "()%," )
          ' -----------------------------------------------------------------
          SED_CodeFinder
          x = SENDMESSAGE( ghComboBox, %CB_FINDSTRINGEXACT, - 1, STRPTR( buffer ))
          SLEEP 1
          IF x <> %CB_ERR THEN
            SED_LastPosition.Position = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )
            GETWINDOWTEXT hWnd, SED_LastPosition.FileName, SIZEOF( SED_LastPosition.FileName )
            y = SENDMESSAGE( GetEdit, %SCI_GETTEXTLENGTH, 0, 0 )
            ' ���õ��ĵ���ĩβλ��
            SENDMESSAGE GetEdit, %SCI_GOTOLINE, y, 0
            ' ���ú�������̵��༭���Ķ���
            p = SENDMESSAGE( ghComboBox, %CB_GETITEMDATA, x, 0 )
            SENDMESSAGE GetEdit, %SCI_GOTOLINE, p, 0
            SETFOCUS GetEdit
          ELSE
            RetVal = FINDWINDOW( $LYNXWINDOW, "" )
            IF ( RetVal = 0 ) OR ( ISWINDOW( RetVal ) = 0 ) THEN
              ' ���û�ҵ�������һ�����
              SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_IDEHELP, 0
            ELSE
              LOCAL uData AS COPYDATASTRUCT
              uData.dwData = %ID_LYNX_FIND
              uData.cbData = LEN( buffer )
              uData.lpData = STRPTR( buffer )
              IF SENDMESSAGE( RetVal, %WM_COPYDATA, 0, VARPTR( uData )) = 0 THEN
                ' ���û�ҵ�������һ�����
                SENDMESSAGE g_hWndMain, %WM_COMMAND, %IDM_HELP, 0
              ELSE
                SED_LastPosition.Position = SENDMESSAGE( GetEdit, %SCI_GETCURRENTPOS, 0, 0 )
                GETWINDOWTEXT hWnd, SED_LastPosition.FileName, SIZEOF( SED_LastPosition.FileName )
              END IF
            END IF
          END IF
        CASE %IDM_GOTOLASTPOSITION
          dwRes = GETWINDOW( g_hWndClient, %GW_CHILD )
          ' �ҵ�ƥ����Ӵ���
          DO WHILE dwRes <> 0
            GETWINDOWTEXT dwRes, szText, SIZEOF( szText )
            IF UCASE$( strTxt ) = UCASE$( SED_LastPosition.FileName ) THEN
              nTab = MdiGetActive( dwRes )
              MdiNext g_hWndClient, nTab, 0
            END IF
            dwRes = GETWINDOW( dwRes, %GW_HWNDNEXT )
          LOOP
          IF dwRes = 0 THEN OpenThisFile( SED_LastPosition.FileName )
          y = SENDMESSAGE( GetEdit, %SCI_GETTEXTLENGTH, 0, 0 )
          ' �����ĵ���ĩβλ��
          SENDMESSAGE GetEdit, %SCI_GOTOLINE, y, 0
          ' ���ú�������̵��༭�����ϲ���
          SENDMESSAGE GetEdit, %SCI_GOTOPOS, SED_LastPosition.Position, 0
          SED_LastPosition.Position = - 1
          SETFOCUS GetEdit
      END SELECT      'LOWRD(wParam)
    CASE %WM_CLOSE
      ' Ignore all edit controls except main in case some are initialized.
      ' ֻ�����ʼ��ʱָ���ı༭�ؼ�
      @wi.hFocus = @wi.hBotR
      ' ����ĵ��Ƿ��޸�
      IF ISTRUE SENDMESSAGE( GetEdit, %SCI_GETMODIFY, 0, 0 ) THEN
        ' �Ӵ��ڱ���õ��ļ�·��
        GETWINDOWTEXT hWnd, szText, SIZEOF( szText )
        ' ѯ���Ƿ񱣴�ı䣬��YES����ı䡣
        ' ������ȫ�ֱ���fClosed��ʶΪ�档
        BEEP
        RetVal = MESSAGEBOX( BYVAL hWnd, " ���浱ǰ�ĸı䣿 " & PATHNAME$(NAME,szText), _
                "�༭��", %MB_YESNOCANCEL OR %MB_ICONQUESTION OR %MB_APPLMODAL )
        IF RetVal = %IDCANCEL THEN
          fClosed = %FALSE
          EXIT FUNCTION
        ELSE
          fClosed = %TRUE
          IF RetVal = %IDYES THEN SAVEFILEs hWnd, %FALSE
        END IF
      ELSE
        fClosed = %TRUE
      END IF
      ' ɾ��FTIME����
      REMOVEPROP( hWnd, "FTIME" )
      ' ɾ��������tab
      GETWINDOWTEXT hWnd, szText, SIZEOF( szText )
      IF UBOUND( gTabFilePaths ) - LBOUND( gTabFilePaths ) + 1 = 1 THEN
        nTab = 0
        ERASE gTabFilePaths
        SENDMESSAGE g_hTabMdi, %TCM_DELETEITEM, nTab, 0
      ELSE
        ARRAY SCAN gTabFilePaths( ), = szText, TO nTab
        IF nTab THEN
          nTab = nTab - 1
          ARRAY DELETE gTabFilePaths( nTab )
          REDIM PRESERVE gTabFilePaths( UBOUND( gTabFilePaths ) - 1 )
          SENDMESSAGE g_hTabMdi, %TCM_DELETEITEM, nTab, 0
          ' �ػ�����tab
          UpdateSysTab
          ' ѡ��ɾ��һ�������һ��tab
          IF nTab < UBOUND( gTabFilePaths ) THEN
            SED_ActivateMdiWindow( gTabFilePaths( nTab ))
          END IF
        END IF
      END IF
      ' �鿴�Ƿ��ж���Ӵ��ڴ�
      dwRes = GETWINDOW( g_hWndClient, %GW_CHILD )
      IF GETWINDOW( dwRes, %GW_HWNDNEXT ) = 0 THEN        ' �Ƿ��������ĵ�
        ' ����״̬��
        ClearStatusbar
        ' ��С��Tab�ؼ�
        MOVEWINDOW g_hTabMdi, 0, 0, 0, 0, %TRUE
        SENDMESSAGE g_hWndMain, %WM_SIZE, 0, 0
        GetWindowText hWnd, zText, SIZEOF(zText)
        lastfolder=PATHNAME$(PATH,zText)
        HideButtons %TRUE
        IF LEN( sProjectPrimary ) THEN
          TreeView_FindItem( ghProjectTV, PATHNAME$(NAME, sProjectPrimary ), %FALSE, hPTROOT, %TRUE, %TRUE )
        END IF
        cUntitledFiles=0
      END IF
      ResetCodeFinder         ' ���ô������������
      IF UCASE$( PrimarySourceFile ) = UCASE$( szText ) THEN  '�����ļ�
        IF LEN( sProjectPrimary ) THEN      '��ǰ�л��Ŀ���������ļ�
          PrimarySourceFile = sProjectPrimary
        END IF
      END IF
    CASE %WM_DESTROY
      ' �ͷ��ڴ�
      HEAPFREE GETPROCESSHEAP, %NULL, BYVAL GETWINDOWLONG( hWnd, %GWL_USERDATA )
    CASE %WM_NOTIFY
      ' ����Scintilla�ؼ�֪ͨ��Ϣ
      SELECT CASE LOWRD( wParam )
        CASE %IDC_EDIT1, %IDC_EDIT2, %IDC_EDIT3, %IDC_EDIT4
          Sci_OnNotify hWnd, wParam, lParam
          EXIT FUNCTION
      END SELECT
  END SELECT
  ' ������Ϣ��Ĭ�ϵ�MDI���̡�
  FUNCTION = DEFMDICHILDPROC( hWnd, wMsg, wParam, lParam )
END FUNCTION
'==============================================================================
FUNCTION GetLineText (BYVAL hEdit AS DWORD, BYVAL ln AS LONG) AS STRING
  '------------------------------------------------------------------------------
  ' �ӱ༭�ؼ����ָ������
  '----------------------------------------------------------------------------
  LOCAL lnStart, lnLen AS LONG, Buf AS STRING

  lnStart = SendMessage(hEdit, %EM_LINEINDEX, ln, 0)       'һ�п�ʼ
  lnLen   = SendMessage(hEdit, %EM_LINELENGTH, lnStart, 0) 'һ�г���

  IF lnLen THEN
     Buf = SPACE$(lnLen)
     lnLen = SendMessage(hEdit, %EM_GETLINE, ln, STRPTR(Buf)) '��ȡ����
     IF lnLen THEN FUNCTION = Buf
  END IF

END FUNCTION
'==============================================================================
SUB PrintDocument (BYVAL hEdit AS DWORD)
'------------------------------------------------------------------------------
  ' ������ӡ���� - ��ӡ�༭�ؼ� (TextBox)�е�����
  ' with 0.8 inch margins (about 20 mm) and a header at top of each page.
  '----------------------------------------------------------------------------
  LOCAL c, x, y, x1, y1, x2, y2, w, h, w1, h1, pgNum, ppiX, ppiY AS LONG
  LOCAL sHeader  AS STRING
  LOCAL zDocName AS ASCIIZ * %MAX_PATH

  '----------------------------------------------------------------------------
  ' Grab document title (file name) for the header
  GetWindowText MdiGetActive(g_hWndClient), zDocName, SIZEOF(zDocName)
  zDocName = FileNam(zDocName)

  '----------------------------------------------------------------------------
  ERRCLEAR
  XPRINT ATTACH CHOOSE, "PBNote2 - " + TRIM$(zDocName)  ' Attach the default host printer

  IF ERR = 0 AND LEN(XPRINT$) THEN          ' On success
      XPRINT GET SIZE TO w, h               ' Get page size in pixels
      XPRINT GET MARGIN TO x1, y1, x2, y2   ' Get printer margins
      XPRINT GET PPI TO ppiX, ppiY          ' Get resolution in pixels/inch
      ' (For resolution in pixels/centimeter, divide ppiX and ppiY with 25.4)

      x1 = 0.8 * (ppiX - x1)            ' 0.8 inch user-defined left margin
      y1 = 0.8 * (ppiY - y1)            ' 0.8 inch user-defined top margin

      XPRINT FONT "Courier New", 10, 0  ' Set a 10 p fixed-width font
      GOSUB PrintPageHeader             ' Begin by printing a header

      y2 = h - (0.8 * ppiY) - y2 - h1   ' calculate the bottom margin pos

      '------------------------------------------------------------------------
      FOR c = 0 TO SendMessage(hEdit, %EM_GETLINECOUNT, 0, 0) - 1
          XPRINT GET POS TO x, y        ' check position

          IF y > y2 THEN                ' if we end up below bottom margin pos
              XPRINT FORMFEED           ' eject paper and start new page
              GOSUB PrintPageHeader     ' Print header part
          ELSE
              XPRINT SET POS (x1, y)    ' else simply set position
          END IF

          XPRINT GetLineText(hEdit, c)  ' and print line
      NEXT
      '------------------------------------------------------------------------

      XPRINT CLOSE                      ' End printjob and detach the printer
  END IF

EXIT SUB

'--------------------------------------------------------------------
PrintPageHeader:
  INCR pgNum      ' print a right-aligned header at top of page
  sHeader = TRIM$(zDocName) + "  -  Page" + STR$(pgNum)
  XPRINT TEXT SIZE sHeader TO w1, h1
  XPRINT SET POS (w - w1 - x1, 0.3 * y1)
  XPRINT sHeader
  y = y1
  XPRINT SET POS (x1, y) ' set position for actual text
RETURN

END SUB


'==============================================================================
SUB SAVEFILE (BYVAL Ask AS LONG)
'------------------------------------------------------------------------------
  ' �������
  '----------------------------------------------------------------------------
  LOCAL dwStyle AS DWORD
  LOCAL nFile   AS DWORD
  LOCAL PATH    AS STRING
  LOCAL f       AS STRING
  LOCAL Buffer  AS STRING
  LOCAL zText   AS ASCIIZ * %MAX_PATH

  '----------------------------------------------------------------------------
  GetWindowText MdiGetActive(g_hWndClient), zText, SIZEOF(zText)
  IF INSTR(zText, ANY ":\/") = 0 THEN '���û��·�������ʾ�����ĵ� (Untitled 1, ��)
       PATH = CURDIR$
       IF RIGHT$(PATH, 1) <> "\" THEN PATH = PATH + "\"
       f    = REMOVE$(zText, " ") & ".txt"  'suggest this name - remove space
       Ask  = %TRUE           'we need the dialog for new docs
  ELSE
       PATH = FilePath(zText)
       f    = FileNam(zText)
  END IF
  dwStyle = %OFN_FILEMUSTEXIST OR %OFN_HIDEREADONLY OR _
            %OFN_EXPLORER OR %OFN_OVERWRITEPROMPT

  '----------------------------------------------------------------------------
  IF Ask THEN
     IF SaveFileDialog(g_hWndMain, "", f, PATH, _
        "Text Files|*.txt|All Files|*.*", "txt", dwStyle) = 0 THEN EXIT SUB
  ELSE
     f = PATH & f
  END IF

  '----------------------------------------------------------------------------
  nFile = FREEFILE  'time to save the file
  OPEN f FOR BINARY AS nFile
    IF ERR THEN   ' if something went wrong
        MSGBOX "Following error occured while trying to save the file:" + $CRLF + _
               "Error:" + STR$(ERR) + ", " + ERROR$(ERR)+ $CRLF + $CRLF + _
               "The file could not be saved.", _
               %MB_ICONWARNING OR %MB_OK OR %MB_TASKMODAL, _
               "Save file error!"
        EXIT SUB
    END IF

     Buffer = SPACE$(GetWindowTextLength(GetEdit) + 1)
     GetWindowText GetEdit, BYVAL STRPTR(Buffer), LEN(Buffer)
     PUT$ nFile, LEFT$(Buffer, LEN(Buffer) - 1)
     SETEOF nFile
  CLOSE nFile

  '----------------------------------------------------------------------------
  IF Ask THEN 'if dialog, update caption in case name was changed
      SetWindowText MdiGetActive(g_hWndClient), BYVAL STRPTR(f)
  END IF
  SendMessage GetEdit, %EM_SETMODIFY, 0, 0
  WriteRecentFiles f   'finally, update reopen file list (MRU menu)

END SUB
' *********************************************************************************************
' Loads the global sAutoCompletionTypes string from the tsunami database
' *********************************************************************************************
SUB SED_GetTypes( )
  LOCAL hFile AS LONG
  LOCAL strPath AS STRING
  LOCAL Record AS STRING
  hFile = trm_Open( EXE.PATH$ & $SCI_TYPEDB, %TRUE )
  IF hFile < 1 THEN EXIT SUB
  sAutoCompletionTypes = ""
  Record = trm_GetFirst( hFile, 1 )
  WHILE LEN( Record )
    strPath = strPath & TRIM$( Record, ANY CHR$( 32, 0 ))
    Record = trm_GetNext( hFile )
    IF LEN( Record ) THEN strPath = strPath & $SPC
  WEND
  trm_Close( hFile )
  sAutoCompletionTypes = strPath
END SUB
' *********************************************************************************************
' *********************************************************************************************
' GenerateTypes - Generates types from the inc file passed in
' *********************************************************************************************
SUB GenerateTypes( sDbfileName AS STRING, sIncFileName AS STRING )
  LOCAL lFilNum AS LONG
  LOCAL sFileStr AS STRING
  LOCAL PosVar AS LONG
  LOCAL StartVar AS LONG
  LOCAL LenVar AS LONG
  LOCAL Mask AS STRING
  LOCAL TheWord AS STRING
  LOCAL i AS LONG, RetVal AS LONG
  LOCAL hFile AS LONG
  LOCAL Result AS LONG
  DIM Record AS STRING * 64
  ON ERROR GOTO errRtn
  IF FileExist( sIncFileName ) THEN
    lFilNum = FREEFILE
    OPEN sIncFileName FOR BINARY AS lFilNum
    GET$ lFilNum, LOF( lFilNum ), sFileStr
    CLOSE lFilNum
    IF LEN( sFileStr ) THEN
      hFile = trm_Open( sDbfileName, %TRUE )
      IF hFile < 1 THEN
        SED_MsgBox g_hWndMain, "Could not open the database"
        EXIT SUB
      END IF
      Mask = "^[\x20]*(TYPE)[\x20]+([a-z_]+)([a-z0-9_\*(\) ,]$)"
      StartVar = 1 : PosVar = 1
      WHILE PosVar
        REGEXPR Mask IN sFileStr AT StartVar TO PosVar, LenVar
        IF PosVar = 0 THEN EXIT
        TheWord = TRIM$( MID$( sFileStr, Posvar, Lenvar ))
        TheWord = MID$( TheWord, 6 )
        'add type word
        Record = TheWord
        Result = trm_Insert( hFile, BYCOPY Record )
        IF Result = 5 THEN    ' Duplicate Key Value so use trm_Update.
          Result = trm_Update( hFile, BYCOPY Record )
        END IF
        StartVar = PosVar + LenVar + 1
      WEND
      'Always add PB Types
      FOR i = 1 TO DATACOUNT
        Record = READ$( i )
        Result = trm_Insert( hFile, BYCOPY Record )
        IF Result = 5 THEN
          Result = trm_Update( hFile, BYCOPY Record )
        END IF
      NEXT
      RetVal = trm_Close( hFile )
      IF RetVal THEN
        SED_MsgBox g_hWndMain, "Error while closing the file"
      END IF
    END IF
  END IF
  SED_MsgBox g_hWndMain, "File Processed."
  '''''PB's Variables
  DATA "PTR", "POINTER", "INTEGER", "LONG", "QUAD"
  DATA "BYTE", "WORD", "DWORD", "SINGLE", "DOUBLE", "VARIANT", "GUID"
  DATA "EXT", "EXTENDED", "CUR", "CURRENCY", "CUX", "CURRENCYX", "STRING", "ASCIIZ", "ASCIZ"
errRtn :
END SUB
' *********************************************************************************************
' *********************************************************************************************
' Function    : CreateTypes ()
' Description : Creates the Tsunami Types Database.
'               The key value is the Type name.
' *********************************************************************************************
FUNCTION CreateTypesDatabase( BYVAL hWnd AS DWORD, BYVAL strDbPath AS STRING ) AS LONG
  LOCAL FileDef AS STRING, Keys AS STRING
  LOCAL strSegment AS STRING * 25
  ON ERROR RESUME NEXT
  strSegment = "Type Name"
  FileDef$ = MKBYT$( 1 ) + _    ' Page Size
          MKBYT$( 1 ) + _     ' Compress Record
          MKBYT$( 1 )     ' Number of Key Segments
  Keys = ""
  Keys = Keys + strSegment + _    ' Description
          MKBYT$( 1 ) + _     ' Segment number
          MKWRD$( 1 ) + _     ' Segment Offset
          MKBYT$( 64 ) + _    ' Segment Size
          MKBYT$( 2 )     ' Flags (No duplicates)
  FileDef = FileDef + Keys
  ' Create Tsunami database file
  IF FileExist( strDbPath ) THEN
    IF MSGBOX( strDbPath & $CRLF & "File already exist. Overwrite it?", %MB_OKCANCEL, "Warning!" ) = %IDOK THEN
      FUNCTION = trm_Create( strDbPath, FileDef, %TRUE )
    ELSE
      EXIT FUNCTION
    END IF
  ELSE
    FUNCTION = trm_Create( strDbPath, FileDef, %FALSE )
  END IF
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
'   ** CallBack **
' *********************************************************************************************
CALLBACK FUNCTION ShowGetTypesDlgProc( )
  STATIC hFocus AS DWORD
  LOCAL szPath AS ASCIIZ * %MAX_PATH
  LOCAL strPath AS STRING
  LOCAL strText AS STRING
  SELECT CASE AS LONG CBMSG
    CASE %WM_NCACTIVATE
      ' Save the handle of the control that has the focus
      IF ISFALSE CBWPARAM THEN hFocus = GETFOCUS
    CASE %WM_SETFOCUS
      ' Post a message to set the focus later, since some
      ' Window's actions can steal it if we set it here
      IF hFocus THEN
        POSTMESSAGE CBHNDL, %WM_USER + 999, hFocus, 0
        hFocus = 0
      END IF
    CASE %WM_USER + 999
      ' Set the focus and show the line an column in the status bar
      IF ISTRUE CBWPARAM THEN SETFOCUS CBWPARAM
      EXIT FUNCTION
    CASE %WM_COMMAND
      ' Process control notifications
      SELECT CASE AS LONG CBCTL
        CASE %IDCANCEL
          IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
            DIALOG END CBHNDL
          END IF
        CASE %IDB_CODETYPES_BROWSE
          IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
            szPath = SearchForIncFilesPath
            IF LEN( szPath ) THEN
              CONTROL SET TEXT CBHNDL, %IDB_CODETYPES_FIND, szPath
            END IF
            CONTROL SET FOCUS CBHNDL, %IDB_CODETYPES_FIND
            hFocus = GETDLGITEM( CBHNDL, %IDB_CODETYPES_FIND )
          END IF
        CASE %IDB_CODETYPES_APPEND
          IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
            IF ISFALSE( FileExist( EXE.PATH$ + $SCI_TYPEDB )) THEN
              CreateTypesDatabase( CBHNDL, EXE.PATH$ + $SCI_TYPEDB )
            END IF
            CONTROL GET TEXT CBHNDL, %IDB_CODETYPES_FIND TO strPath
            IF LEN( strPath ) THEN
              GenerateTypes( EXE.PATH$ + $SCI_TYPEDB, strPath )
              SED_GetTypes
            ELSE
              SED_MsgBox g_hWndMain, "Select an INC file first"
            END IF
            CONTROL SET FOCUS CBHNDL, %IDB_CODETYPES_FIND
            hFocus = GETDLGITEM( CBHNDL, %IDB_CODETYPES_FIND )
          END IF
        CASE %IDB_CODETYPES_HELP
          IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
            szPath = EXE.PATH$ & "SED.HLP"
            strText = "Codetypes Builder"
            IF FileExist( szPath ) THEN WINHELP CBHNDL, szPath, %HELP_KEY, BYVAL STRPTR( strText )
          END IF
      END SELECT
  END SELECT
END FUNCTION
' *********************************************************************************************
' *********************************************************************************************
'   ** Dialog **
' *********************************************************************************************
FUNCTION ShowGetTypesDlg( BYVAL hParent AS DWORD ) AS LONG
  LOCAL lRslt AS LONG
  LOCAL hDlg AS LONG
  DIALOG NEW hParent, "Codetypes Builder", 151, 121, 199, 78, %WS_POPUP OR _
          %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR %WS_CLIPSIBLINGS OR _
          %WS_VISIBLE OR %DS_MODALFRAME OR %DS_CENTER OR %DS_3DLOOK OR _
          %DS_NOFAILCREATE OR %DS_SETFONT, %WS_EX_WINDOWEDGE OR _
          %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
          %WS_EX_RIGHTSCROLLBAR, TO hDlg
  DIALOG SET TEXT hDlg, "�������͹���"
  CONTROL ADD LABEL, hDlg, %IDL_CODETYPES_LABEL, "ͷ�ļ�·��", 10, 9, 154, 10, %WS_CHILD OR %WS_VISIBLE
  CONTROL ADD TEXTBOX, hDlg, %IDB_CODETYPES_FIND, "", 10, 20, 154, 14
  CONTROL ADD BUTTON, hDlg, %IDB_CODETYPES_BROWSE, "...", 171, 20, 18, 14
  CONTROL ADD BUTTON, hDlg, %IDB_CODETYPES_APPEND, "׷��", 10, 50, 49, 14
  CONTROL ADD BUTTON, hDlg, %IDCANCEL, "�˳�", 75, 50, 50, 14
  CONTROL ADD BUTTON, hDlg, %IDB_CODETYPES_HELP, "����", 140, 50, 49, 14
  'DIALOG  SEND hDlg, %DM_SETDEFID, %IDCANCEL, 0
  DIALOG SHOW MODAL hDlg, CALL ShowGetTypesDlgProc TO lRslt
  FUNCTION = lRslt
END FUNCTION
' *********************************************************************************************
' Search for include filesѰ�Ұ����ļ�
' *********************************************************************************************
FUNCTION SearchForIncFilesPath( ) AS STRING
  LOCAL PATH AS STRING, fOptions AS STRING, STYLE AS DWORD, f AS STRING
  PATH = CURDIR$
  fOptions = "INC �ļ� (*.INC)|*.INC|"
  fOptions = fOptions & "BAS �ļ� (*.BAS)|*.BAS|"
  fOptions = fOptions & "�����ļ� (*.*)|*.*"
  f = "*.INC"
  STYLE = %OFN_EXPLORER OR %OFN_FILEMUSTEXIST OR %OFN_HIDEREADONLY
  IF ISTRUE OpenFileDialog( g_hWndMain, "", f, PATH, fOptions, "INC", STYLE ) THEN FUNCTION = f
END FUNCTION
SUB SED_LoadPreviousFileSet
  LOCAL fNumber           AS LONG         ' // File number �ļ���
  LOCAL buffer            AS STRING       ' // Buffer ����
  LOCAL strPath           AS STRING       ' // Path ·��
  LOCAL caretPos              AS LONG         ' // Caret position���ַ���λ��
  LOCAL endPos            AS LONG         ' // End of file �ļ�ĩβ
  LOCAL LinesOnScreen     AS LONG         ' // Lines on screen ��Ļ�ϵ�����
  LOCAL LineToGo              AS LONG         ' // Line to go ��ת������
  LOCAL bk AS             LONG        ' // Number of bookmarks ��ǩ��
  LOCAL nLine             AS LONG         ' // bookmark line ��ǩ��
  LOCAL fProjectLoaded    AS LONG         ' // Flag  ��ʶ
  ON ERROR GOTO ErrHandler
  ' Read the saved projects window placement��ȡ�������Ŀ����λ��
  'CALL SED_ReadProjectWindowPlacement
  fNumber = FREEFILE
  OPEN EXE.PATH$ & "SED.LFS" FOR INPUT AS fNumber
  WHILE ISFALSE EOF( fNumber )
    LINE INPUT #fNumber, buffer
    IF UCASE$( buffer ) = "NONAME.PBP" THEN
      sProjectName = buffer
      bMsgWindow = %TRUE
'      TV_SetItemText ghTreeView, hPTRoot, BYCOPY sProjectName
'      SETPARENT ghTreeView, ghPrjDockC
'      SHOWWINDOW ghPrjDockC, %SW_SHOW
'      SHOWWINDOW ghTreeView, %SW_HIDE
'      'SHOWWINDOW ghMsgWindow, %SW_SHOW
'      SENDMESSAGE ghPrjDockC, %WM_SIZE, 0, 0
'      SENDMESSAGE g_hWndMain, %WM_SIZE, 0, 0
      fProjectLoaded = %TRUE
    ELSEIF UCASE$( RIGHT$( buffer, 3 )) = "SPF" THEN
      IF FileExist( buffer ) THEN
        'ProjectLoadTree( buffer )
        fProjectLoaded = %TRUE
      END IF
    ELSE
      strPath = PARSE$( buffer, "|", 1 )
      caretPos = VAL( PARSE$( buffer, "|", 2 ))
      bk = VAL( PARSE$( buffer, "|", 3 ))
      IF FileExist( strPath ) THEN
        OpenThisFile strPath
        DO WHILE bk <> 0
          nLine = VAL( PARSE$( buffer, "|", bk + 3 ))
          SENDMESSAGE( GetEdit, %SCI_MARKERADD, nLine, 0 )
          DECR bk
        LOOP
        endPos = SENDMESSAGE( GetEdit, %SCI_GETTEXTLENGTH, 0, 0 )
        SENDMESSAGE GetEdit, %SCI_GOTOPOS, endPos, 0
        LinesOnScreen = SENDMESSAGE( GetEdit, %SCI_LINESONSCREEN, 0, 0 )
        LineToGo = SENDMESSAGE( GetEdit, %SCI_LINEFROMPOSITION, caretPos, 0 )
        LineToGo = LineToGo - ( LinesOnScreen \ 2 )
        SENDMESSAGE GetEdit, %SCI_GOTOLINE, LineToGo, 0
        SENDMESSAGE GetEdit, %SCI_GOTOPOS, caretPos, 0
      END IF
    END IF
  WEND
  ' If there is a primary source file, activate it if it is loaded�������Դ�ļ����ұ������򼤻���
  'SED_GetPrimarySourceFile( %FALSE )
ErrHandler :
  CLOSE fNumber
END SUB
FUNCTION CreateTabMdiCtl() AS DWORD
  LOCAL hIcon AS DWORD
  g_hTabMdi = CREATEWINDOWEX( %WS_EX_STATICEDGE, "SysTabControl32", BYVAL %NULL, _
                  %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %TCS_SINGLELINE OR %TCS_RIGHTJUSTIFY _
                  OR %TCS_TABS OR %TCS_FOCUSNEVER OR %TCS_TOOLTIPS, _
                  0, 0, 0, 0, _
                  g_hWndMain, %IDC_TABMDI, g_hInst, BYVAL %NULL )
'  CONTROL ADD TAB, g_hWndMain, %IDC_TABMDI, "", 1, 1, 298, 138
'  control handle g_hWndMain,%IDC_TABMDI TO g_hTabMdi
  'msgbox "g_hTabMdi=" & str$(g_hTabMdi)
  hTabImageList = IMAGELIST_CREATE( 16, 16, %ILC_MASK, 2, 1 )
  hIcon = LOADICON( GETMODULEHANDLE( "" ), "SEDICON_EDIT" )
  ImageList_AddIcon hTabImageList, hIcon
  DESTROYICON hIcon
  TabCtrl_SetImageList g_hTabMdi, hTabImageList
  'SendMessage(g_hTabMdi, %TCM_SETIMAGELIST, 0, hTabImageList)
  IF hFont THEN SENDMESSAGE g_hTabMdi, %WM_SETFONT, hFont, %TRUE
  'SHOWWINDOW g_hTabMdi, %SW_HIDE
  'msgbox str$(g_hTabMdi)
  FUNCTION = g_hTabMdi
END FUNCTION
FUNCTION CreateSciControl( BYVAL hWnd AS DWORD, wi AS WININFOSTRUC ) AS DWORD           ' *** NTM ***
  LOCAL szText AS ASCIIZ * 255
  LOCAL nFile AS LONG
  LOCAL buffer AS STRING
  LOCAL hSci AS DWORD
  LOCAL pSci AS DWORD
  LOCAL fTime AS DWORD
  LOCAL strExt AS STRING
  LOCAL p AS LONG
  LOCAL ValidExt AS LONG
  LOCAL ToolHeight AS LONG
  LOCAL rctb AS RECT
  LOCAL rc AS RECT
'  GETWINDOWRECT g_hWndClient, rctb
'  ScreenToClientRect g_hWndMain,rctb
  'ToolHeight = rctb.nBottom - rctb.nTop
  GETCLIENTRECT g_hTabMdi, rc
'  msgbox "rc.nBottom=" & str$(rc.nBottom)
  IF rc.nBottom = 0 THEN
    GETCLIENTRECT g_hWndClient, rc
    MOVEWINDOW g_hTabMdi, rc.nLeft, rc.nTop, rc.nRight, 26, %TRUE
    SENDMESSAGE g_hWndMain, %WM_SIZE, 0, 0
  END IF
  ' *** NTM ***----------------------------------------------------------------------------------
  LOCAL pDoc AS LONG
  wi.hBotR = CREATEWINDOWEX( %NULL, "Scintilla", BYVAL %NULL, %WS_CHILD OR %WS_VISIBLE OR _
          %ES_MULTILINE OR %WS_VSCROLL OR %WS_HSCROLL OR _
          %ES_AUTOHSCROLL OR %ES_AUTOVSCROLL OR %ES_NOHIDESEL, _
          0, 0, 0, 0, hWnd, %IDC_EDIT1, g_hInst, BYVAL %NULL )
  ' This is our left bottom VIEW window
  wi.hBotL = CREATEWINDOWEX( %NULL, "Scintilla", BYVAL %NULL, %WS_CHILD OR %WS_VISIBLE OR _
          %ES_MULTILINE OR %WS_VSCROLL OR %WS_HSCROLL OR _
          %ES_AUTOHSCROLL OR %ES_AUTOVSCROLL OR %ES_NOHIDESEL, _
          0, 0, 0, 0, hWnd, %IDC_EDIT2, g_hInst, BYVAL %NULL )
  ' This is right top VIEW window
  wi.hTopR = CREATEWINDOWEX( %NULL, "Scintilla", BYVAL %NULL, %WS_CHILD OR %WS_VISIBLE OR _
          %ES_MULTILINE OR %WS_VSCROLL OR %WS_HSCROLL OR _
          %ES_AUTOHSCROLL OR %ES_AUTOVSCROLL OR %ES_NOHIDESEL, _
          0, 0, 0, 0, hWnd, %IDC_EDIT3, g_hInst, BYVAL %NULL )
  ' This is left top VIEW window
  wi.hTopL = CREATEWINDOWEX( %NULL, "Scintilla", BYVAL %NULL, %WS_CHILD OR %WS_VISIBLE OR _
          %ES_MULTILINE OR %WS_VSCROLL OR %WS_HSCROLL OR _
          %ES_AUTOHSCROLL OR %ES_AUTOVSCROLL OR %ES_NOHIDESEL, _
          0, 0, 0, 0, hWnd, %IDC_EDIT4, g_hInst, BYVAL %NULL )
  wi.hFocus = wi.hBotR
  pDoc = SENDMESSAGE( wi.hBotR, %SCI_GETDOCPOINTER, %NULL, %NULL )
  SENDMESSAGE wi.hTopR, %SCI_SETDOCPOINTER, %NULL, pDoc           ' All editors now share the SAME document
  SENDMESSAGE wi.hBotL, %SCI_SETDOCPOINTER, %NULL, pDoc
  SENDMESSAGE wi.hTopL, %SCI_SETDOCPOINTER, %NULL, pDoc
  hSci = wi.hFocus
  ' *** NTM ***----------------------------------------------------------------------------------
  ' Retrieve the path of the file from the window's caption
  GETWINDOWTEXT hWnd, szText, SIZEOF( szText )
  p = INSTR( - 1, szText, "." )
  IF p THEN strExt = MID$( szText, p )
  IF strExt = ".BAS" OR strExt = ".INC" OR strExt = ".RC" THEN ValidExt = %TRUE
  IF strExt = "" THEN strExt = ".BAS"
  ' Get a direct pointer for faster access
  pSci = SENDMESSAGE( hSci, %SCI_GETDIRECTPOINTER, 0, 0 )
  FUNCTION = hSci
  IF LEN( szText ) THEN
    ERRCLEAR
    nFile = FREEFILE
    OPEN szText FOR BINARY AS nFile
    IF ERR THEN
      MESSAGEBOX( hWnd, "����" & STR$( ERR ) & " �����ļ���������,������.   ", _
              FUNCNAME$, %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL )
      INCR cUntitledFiles
      SETWINDOWTEXT hWnd, "Untitled" & FORMAT$( cUntitledFiles ) & strExt
    ELSE
      GET$ nFile, LOF( nFile ), Buffer
      CLOSE nFile
      IF ISTRUE INSTR( buffer, ANY CHR$( 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27, 28, 29, 30, 31, 127 )) _
            AND ISFALSE ValidExt THEN
        MESSAGEBOX( hWnd, szText & $CRLF & "�ļ����ݲ�Ϊ ASCII �� ANSI ��ʽ." & $CRLF & _
                "���ܼ��ص�������.", "���ı��ļ�", %MB_OK OR %MB_ICONWARNING OR %MB_APPLMODAL )
      ELSE
        ' Trim trailing spaces and tabs
        IF TrimTrailingBlanks THEN
          DO
            p = LEN( buffer )
            REPLACE " " & $CR WITH $CR IN buffer
            REPLACE $TAB & $CR WITH $CR IN buffer
          LOOP UNTIL p = LEN( buffer )
        END IF
        WriteRecentFiles szText                 ' Refresh reopen file list
        ' Put the text in the edit control
        SENDMESSAGE hSci, %SCI_INSERTTEXT, 0, BYVAL STRPTR( Buffer )
        ' Empty the undo buffer (it al sets the state of the document as unmodified)
        SENDMESSAGE hSci, %SCI_EMPTYUNDOBUFFER, 0, 0
        ' Tell to Scintilla that the current state of the document is unmodified.
        ' SendMessage hSci, %SCI_SETSAVEPOINT, 0, 0
        ' Retrieve the file time and store it in the properties list of the window
        fTime = SED_GetFileTime( szText )
        SETPROP hWnd, "FTIME", fTime
      END IF
    END IF
  ELSE
    INCR cUntitledFiles
    SETWINDOWTEXT hWnd, "Untitled" & FORMAT$( cUntitledFiles ) & strExt
  END IF
  ' Retrieve the name again and set the control options
  GETWINDOWTEXT hWnd, szText, SIZEOF( szText )
  IF pSci THEN Scintilla_SetOptions pSci, szText
  ' *** NTM ***----------------------------------------------------------------------------------
  pSci = SENDMESSAGE( wi.hTopR, %SCI_GETDIRECTPOINTER, %NULL, %NULL )
  IF pSci THEN Scintilla_SetOptions pSci, szText
  pSci = SENDMESSAGE( wi.hTopL, %SCI_GETDIRECTPOINTER, %NULL, %NULL )
  IF pSci THEN Scintilla_SetOptions pSci, szText
  pSci = SENDMESSAGE( wi.hBotL, %SCI_GETDIRECTPOINTER, %NULL, %NULL )
  IF pSci THEN Scintilla_SetOptions pSci, szText
  ghEditStr=ghEditStr & FORMAT$(hSci) & ","
  ' *** NTM ***----------------------------------------------------------------------------------
END FUNCTION
FUNCTION InsertTabMdiItem( BYVAL hTab AS DWORD, BYVAL ITEM AS LONG, szTabText AS ASCIIZ ) AS LONG
  LOCAL ttc_item AS TC_ITEM
  IF SENDMESSAGE( g_hTabMdi, %TCM_GETITEMCOUNT, 0, 0 ) = 0 THEN
    SHOWWINDOW g_hTabMdi, %SW_SHOW
  END IF
  ' Insert a tab...
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
FUNCTION SED_GetFileTime( BYVAL FileSpec AS STRING ) AS DWORD
  LOCAL fd AS WIN32_FIND_DATA
  LOCAL ft AS SYSTEMTIME
  LOCAL hFound AS DWORD
  LOCAL strTime AS STRING
  hFound = FINDFIRSTFILE( BYVAL STRPTR( FileSpec ), fd )
  IF hFound = %INVALID_HANDLE_VALUE THEN EXIT FUNCTION
  FINDCLOSE hFound
  ' -- Convert the file time into a compatible system time
  FILETIMETOSYSTEMTIME fd.ftLastWriteTime, ft
  strTime = FORMAT$( ft.wHour ) & FORMAT$( ft.wMinute ) & FORMAT$( ft.wSecond )
  FUNCTION = VAL( strTime )
END FUNCTION
FUNCTION GetTabName(BYVAL nTab AS LONG)AS STRING
  'local hMdi as DWORD
  'hMdi=TabCtrl_GetCurSel(g_hTabMdi)
  DIM  ttc_item AS TC_ITEM
  DIM  tmpStr   AS ASCIIZ * 257
  ttc_item.mask=%TCIF_TEXT 'or %TCIF_IMAGE OR %TCIF_PARAM OR %TCIF_RTLREADING
  ttc_item.cchTextMax=256
  ttc_item.pszText=VARPTR(tmpStr)
  IF ISTRUE(TabCtrl_GetItem(g_hTabMdi,nTab,ttc_item)) THEN
    FUNCTION=TRIM$(tmpStr)
  ELSE
    FUNCTION=""
  END IF
END FUNCTION
SUB SetTabName( BYVAL nTab AS LONG, BYVAL strName AS STRING )
    LOCAL ttc_item AS TC_ITEM
    ttc_item.mask = %TCIF_TEXT
    ttc_item.pszText = STRPTR( strName )
    ttc_item.cchTextMax = 256
    SENDMESSAGE g_hTabMdi, %TCM_SETITEM, nTab, BYVAL VARPTR( ttc_item )
END SUB
' *********************************************************************************************
' Reads the Color/Fonts options from the .INI file
' *********************************************************************************************
SUB GetPrinterSetupOptions( PrnOpt AS PrinterSetupOptionsType )
  LOCAL rs AS STRING
  LOCAL g_zIni AS STRING
  g_zIni=EXE.PATH$ & "config.ini"
  rs = IniRead( g_zIni, "Printer options", "PaperSize", "" )
  IF LEN( rs ) THEN PrnOpt.PaperSize = VAL( rs ) ELSE PrnOpt.PaperSize = 2
  rs = IniRead( g_zIni, "Printer options", "PaperBin", "" )
  IF LEN( rs ) THEN PrnOpt.PaperBin = VAL( rs ) ELSE PrnOpt.PaperBin = 7
  rs = IniRead( g_zIni, "Printer options", "Orientation", "" )
  IF LEN( rs ) THEN PrnOpt.Orientation = VAL( rs ) ELSE PrnOpt.Orientation = 0
  rs = IniRead( g_zIni, "Printer options", "FontName", "" )
  IF LEN( rs ) THEN PrnOpt.FontName = rs ELSE PrnOpt.FontName = "Courier New"
  rs = IniRead( g_zIni, "Printer options", "FontSize", "" )
  IF LEN( rs ) THEN PrnOpt.FontSize = VAL( rs ) ELSE PrnOpt.FontSize = 3
  rs = IniRead( g_zIni, "Printer options", "FontBold", "" )
  IF LEN( rs ) THEN PrnOpt.FontBold = VAL( rs ) ELSE PrnOpt.FontBold = 0
  rs = IniRead( g_zIni, "Printer options", "FontItalic", "" )
  IF LEN( rs ) THEN PrnOpt.FontItalic = VAL( rs ) ELSE PrnOpt.FontItalic = 0
  rs = IniRead( g_zIni, "Printer options", "FontUnderline", "" )
  IF LEN( rs ) THEN PrnOpt.FontUnderline = VAL( rs ) ELSE PrnOpt.FontUnderline = 0
  rs = IniRead( g_zIni, "Printer options", "MarginLeft", "" )
  PrnOpt.MarginLeft = VAL( rs )
  rs = IniRead( g_zIni, "Printer options", "MarginRight", "" )
  PrnOpt.MarginRight = VAL( rs )
  rs = IniRead( g_zIni, "Printer options", "MarginTop", "" )
  PrnOpt.MarginTop = VAL( rs )
  rs = IniRead( g_zIni, "Printer options", "MarginBottom", "" )
  PrnOpt.MarginBottom = VAL( rs )
  rs = IniRead( g_zIni, "Printer options", "LineHeight", "" )
  PrnOpt.LineHeight = VAL( rs )
  IF PrnOpt.LineHeight < = 0 THEN PrnOpt.LineHeight = 0.17!
END SUB
SUB AssignFocus( wi AS WININFOSTRUC )
  ' Set focus to an appropriate editor when editor w/focus is no longer shown ���༭����w/���㲻����ʾʱ�����ý��㵽�ʵ��ı༭��
  LOCAL dwFocus AS DWORD
  dwFocus = GETFOCUS
  IF wi.SplitY = 0 AND ( dwFocus = wi.hTopR OR dwFocus = wi.hTopL ) THEN
    SETFOCUS wi.hBotR
  ELSE
    IF wi.SplitXT = 0 AND dwFocus = wi.hTopL THEN
      SETFOCUS wi.hTopR
    ELSE
      IF wi.SplitXB = 0 AND dwFocus = wi.hBotL THEN SETFOCUS wi.hBotR
    END IF
  END IF
END SUB
SUB DrawWindow( BYVAL hWnd AS LONG, BYVAL bMove AS LONG )
  LOCAL w AS LONG, h AS LONG, hDwp AS DWORD, hDC AS DWORD
  LOCAL rc AS RECT, wi AS WININFOSTRUC PTR
  wi = GETWINDOWLONG( hWnd, %GWL_USERDATA )
  GETCLIENTRECT hWnd, rc
  ' Get width & height of child client area �õ��ӿͻ�����Ⱥ͸߶�
  w = rc.nRight
  h = rc.nBottom
  IF bMove THEN
    ' Allocates memory for a multiple-window position structure and returns the
    ' handle to the structure. Ϊ�ര��λ�ýṹ�ͷ��ظ�Щ�ṹ�ľ�������ڴ档
    ' Specify the initial number of windows for which to store position information ���ݱ����λ����Ϣָ����ʼ���Ĵ��ڸ���
    ' For some reason, %SWP_NOSENDCHANGING fails under XP - so it's not used...����δ֪ԭ��%SWP����NOSENDCHANGING��XP�»�������Բ�ʹ����
    hDwp = BEGINDEFERWINDOWPOS( 4 )         ' 4 edit windows 4 ���༭����
    IF @wi.SplitY THEN      ' Y coordinate is > 0  Y���������
      hDwp = DEFERWINDOWPOS( hDwp, @wi.hTopR, %NULL, @wi.SplitXT + %SPLITSIZE, 0, w - ( @wi.SplitXT + %SPLITSIZE ), @wi.SplitY, _
          %SWP_NOZORDER OR %SWP_SHOWWINDOW OR %SWP_NOACTIVATE OR %SWP_NOCOPYBITS )        '
      IF @wi.SplitXT THEN         ' Top X coordinate is > 0  ��X�������0
        hDwp = DEFERWINDOWPOS( hDwp, @wi.hTopL, %NULL, 0, 0, @wi.SplitXT, @wi.SplitY, _
            %SWP_NOZORDER OR %SWP_SHOWWINDOW OR %SWP_NOACTIVATE OR %SWP_NOCOPYBITS )        '
      ELSE
        SHOWWINDOW @wi.hTopL, %SW_HIDE
      END IF
    ELSE        ' If no viewable top, hide left AND right editors  ���û�ж���ͼ���������ұ༭��
      SHOWWINDOW @wi.hTopR, %SW_HIDE
      SHOWWINDOW @wi.hTopL, %SW_HIDE
    END IF
    IF @wi.SplitY < h - %SPLITSIZE THEN
      hDwp = DEFERWINDOWPOS( hDwp, @wi.hBotR, %NULL, @wi.SplitXB + %SPLITSIZE, @wi.SplitY + %SPLITSIZE, w - ( @wi.SplitXB + %SPLITSIZE ), h - @wi.SplitY - %SPLITSIZE, _
          %SWP_NOZORDER OR %SWP_SHOWWINDOW OR %SWP_NOACTIVATE OR %SWP_NOCOPYBITS )        '
      IF @wi.SplitXB THEN         ' Bottom X coordinate is > 0 ��X�������0
        hDwp = DEFERWINDOWPOS( hDwp, @wi.hBotL, %NULL, 0, @wi.SplitY + %SPLITSIZE, @wi.SplitXB, h - @wi.SplitY - %SPLITSIZE, _
          %SWP_NOZORDER OR %SWP_SHOWWINDOW OR %SWP_NOACTIVATE OR %SWP_NOCOPYBITS )        '
      ELSE
        SHOWWINDOW @wi.hBotL, %SW_HIDE
      END IF
    ELSE
        SHOWWINDOW @wi.hBotR, %SW_HIDE
    END IF
    ENDDEFERWINDOWPOS hDwp
  END IF
  hDC = GETDC( hWnd )
  ' Draw the horizontal split bar��ˮƽ�ָ���
  SETRECT rc, 0, @wi.SplitY, w, @wi.SplitY + %SPLITSIZE
  DRAWFRAMECONTROL hDC, rc, %DFC_BUTTON, %DFCS_BUTTONPUSH
  ' Draw the bottom vertical split bar ���ײ�����ָ���
  SETRECT rc, @wi.SplitXB, @wi.SplitY + %SPLITSIZE, @wi.SplitXB + %SPLITSIZE, h
  DRAWFRAMECONTROL hDC, rc, %DFC_BUTTON, %DFCS_BUTTONPUSH
  ' Draw the top vertical split bar ����������ָ���
  SETRECT rc, @wi.SplitXT, 0, @wi.SplitXT + %SPLITSIZE, @wi.SplitY
  DRAWFRAMECONTROL hDC, rc, %DFC_BUTTON, %DFCS_BUTTONPUSH
  RELEASEDC hWnd, hDC
END SUB
SUB UpdateSysTab
    LOCAL ttc_item AS TC_ITEM
    LOCAL i AS LONG
    LOCAL szFileName AS ASCIIZ * 256        ' %MAX_PATH
    '   SendMessage ghTabMdi, %WM_SETREDRAW, 0, 0
    SENDMESSAGE g_hTabMdi, %TCM_DELETEALLITEMS, 0, 0
    ttc_item.Mask = %TCIF_TEXT OR %TCIF_IMAGE
    ttc_item.cchTextMax = SIZEOF( szFileName )      ' %MAX_PATH
    ttc_item.iImage = 0
    FOR i = LBOUND( gTabFilePaths ) TO UBOUND( gTabFilePaths )
      szFileName = GetFileName( gTabFilePaths( i ))
      ttc_item.pszText = VARPTR( szFilename )
      SENDMESSAGE g_hTabMdi, %TCM_INSERTITEM, i, BYVAL VARPTR( ttc_item )
    NEXT
    '   SendMessage ghTabMdi, %WM_SETREDRAW, 1, 0
    '   InvalidateRect ghTabMdi, BYVAL 0, 1
    '   UpdateWindow ghTabMdi
END SUB
' *********************************************************************************************
SUB SED_ActivateMdiWindow( BYVAL strFilePath AS STRING )
    LOCAL hMdi AS DWORD
    LOCAL szPath AS ASCIIZ * %MAX_PATH
    LOCAL nTab AS LONG
    GETWINDOWTEXT MdiGetActive( g_hWndClient ), szPath, SIZEOF( szPath )
    IF szPath = strFilePath THEN EXIT SUB
    ' Get the first handle
    hMdi = GETWINDOW( g_hWndClient, %GW_CHILD )
    ' Cycle thru all the open windows
    WHILE hMdi
        ' Get the caption text
        GETWINDOWTEXT( hMdi, szPath, %MAX_PATH )
        IF szPath = strFilePath THEN
            ARRAY SCAN gTabFilePaths( ), = szPath, TO nTab
            IF nTab THEN
                SENDMESSAGE g_hTabMdi, %TCM_SETCURSEL, nTab - 1, 0       ' activate the tab associated with the window
                IF ISICONIC( hMdi ) THEN
                    SENDMESSAGE g_hWndClient, %WM_MDIRESTORE, hMdi, 0
                ELSE
                    SENDMESSAGE g_hWndClient, %WM_MDIACTIVATE, hMdi, 0
                END IF
                EXIT LOOP
            END IF
        END IF
        hMdi = GETWINDOW( hMdi, %GW_HWNDNEXT )
    WEND
END SUB
' *********************************************************************************************
' Save file procedure
' *********************************************************************************************
FUNCTION SAVEFILEs( BYVAL hWnd AS DWORD, BYVAL Ask AS LONG ) AS LONG
  LOCAL PATH AS STRING
  LOCAL f AS STRING
  LOCAL dwStyle AS DWORD
  LOCAL nFile AS DWORD
  LOCAL Buffer AS STRING
  LOCAL szText AS ASCIIZ * %MAX_PATH
  LOCAL fOptions AS STRING
  LOCAL fBak AS STRING
  LOCAL fExt AS STRING
  LOCAL fPos AS LONG
  LOCAL fTime AS DWORD
  LOCAL nLen AS LONG
  LOCAL nTab AS LONG
  LOCAL szExt AS ASCIIZ * 255
  LOCAL p AS LONG
  LOCAL pSci AS DWORD
  LOCAL fIsInitAsNew AS LONG              ' // Flag - It is an Initialized as new window��ʶ�Ƿ���һ���´���
  LOCAL bitnum AS LONG            ' // Bit numberλ��
  LOCAL wi AS WININFOSTRUC PTR
  LOCAL g_zIni AS STRING
  LOCAL tmpStr AS STRING
  g_zIni=EXE.PATH$ & "config.ini"
  IF ISFALSE GetEdit THEN EXIT FUNCTION
  ' Get the handles of the view windows�õ���ͼ���ڵľ��
  wi = GETWINDOWLONG( MdiGetActive( g_hWndClient ), %GWL_USERDATA )
  ' See if it is an Initialized as new window����Ƿ����´���
  IF wi THEN
    IF @wi.binit THEN
      bitnum = GETDLGCTRLID( @wi.hFocus ) - %IDC_EDIT2
      IF bitnum = > 0 AND bitnum < 16 THEN
        IF BIT( @wi.bInit, bitnum ) THEN
          fIsInitAsNew = %TRUE
        END IF
      END IF
    END IF
  END IF
  ' Get the title bar of the window�õ����ڱ�����
  IF ISTRUE fIsInitAsNew THEN
    fOptions = fOptions & "�����ļ� (*.*)|*.*"
  ELSE
    GETWINDOWTEXT MdiGetActive( g_hWndClient ), szText, SIZEOF( szText )
    IF INSTR( szText, ANY ":\/" ) = 0 THEN          ' if no path, it's a new doc���û��·������ȡ���ļ��ĵ�ǰ·��
      'msgbox "SED_FILE.INC SAVEFILE() szText=" & szText
      PATH = CURDIR$
      IF RIGHT$( PATH, 1 ) <> "\" THEN PATH = PATH + "\"
      IF LEFT$( UCASE$( szText ), 8 ) = "UNTITLED"  THEN  '�ж��Ƿ���δ�����ļ��������ļ�
        IF INSTR( szText, "." ) = 0 THEN  '����ļ����в�����.����ļ�������.bas��׺
          f = szText & ".BAS"
        ELSE
          f = szText
        END IF
        Ask = %TRUE             ' we need the dialog for new docs������ҪΪ���ĵ�׼������
      ELSE                '�ļ����в�����δ������ʶ������Ϊ�����ļ���
        f = szText
      END IF
      'msgbox "aaaa"
    ELSE
      PATH = GetFilePath( szText )
      f = GetFileName( szText )
    END IF
    p = INSTR( - 1, f, "." )  '���ļ����д���������.
    IF p THEN szExt = UCASE$( MID$( f, p + 1 ))'��ȡ�ҵ���.������ַ���������д����
    IF szExt <> "BAS" AND szExt <> "INC" AND szExt <> "RC" AND szExt <> "HTML" AND szExt <> "HTM" _
          AND szExt <> "TXT" AND szExt <> "INI" THEN szExt = "" '�������Ĭ�ϵ�bas,inc,rc,html,ini����txt�ļ����ͣ�����պ�׺
    IF szExt = "BAS" THEN
      fOptions = fOptions & "PB �����ļ� (*.BAS)|*.BAS|"
      fOptions = fOptions & "PB ͷ�ļ� (*.INC)|*.INC|"
    ELSE
      fOptions = fOptions & "PB ͷ�ļ� (*.INC)|*.INC|"
      fOptions = fOptions & "PB �����ļ� (*.BAS)|*.BAS|"
    END IF
    fOptions = fOptions & "PB ģ���ļ� (*.PBTPL)|*.PBTPL|"
    fOptions = fOptions & "��Դ�ļ� (*.RC)|*.RC|"
    fOptions = fOptions & "��ҳ�ļ� (*.HTML)|*.HTML|"
    fOptions = fOptions & "��ҳ�ļ� (*.HTM)|*.HTM|"
    fOptions = fOptions & "�ı��ļ� (*.TXT)|*.TXT|"
    fOptions = fOptions & "�����ļ� (*.*)|*.*"
  END IF
  IF ISTRUE( Ask ) THEN
    dwStyle = %OFN_EXPLORER OR %OFN_FILEMUSTEXIST OR %OFN_HIDEREADONLY OR %OFN_OVERWRITEPROMPT
    IF ISFALSE( SaveFileDialog( g_hWndMain, "", f, PATH, fOptions, "BAS", dwStyle )) THEN EXIT FUNCTION
  ELSE
    f = PATH & f
  END IF
  ' Don't allow to save an Initialized as new window with����������ͬ���ļ�������δ����������ͬ��·�����������ļ�
  ' the same name as the file being edited.
  IF ISTRUE fIsInitAsNew THEN
    GETWINDOWTEXT MdiGetActive( g_hWndClient ), szText, SIZEOF( szText )
    IF f = szText THEN
      MESSAGEBOX( hWnd, "�����ѡ��һ����ͬ���ļ�����·��.   ", _
              " �����ļ�", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL )
      EXIT FUNCTION
    END IF
  END IF
  ' Backup the file�����ļ�
  IF ISTRUE VAL( IniRead( g_zIni, "Editor options", "BackupEditorFiles", "" )) THEN
    tmpStr=MID$(f,1,INSTR(-1,f,ANY "\/")) & DateTimeForFileName & "bak"
    MKDIR tmpStr
    'fPos = INSTR( - 1, f, "." )
    'IF fPos THEN fExt = MID$( f, fPos )
    'IF UCASE$( fExt ) <> ".BAK" THEN
      fBak = tmpStr & MID$(f,INSTR(-1,f,ANY "\/"))
      'REPLACE fExt WITH ".bak" IN fBak
      IF FileExist( f ) THEN FILECOPY f, fBak
    'END IF
  END IF
  ' Save the file�����ļ�
  TRY
    nFile = FREEFILE
    OPEN f FOR BINARY AS nFile
  CATCH
    MESSAGEBOX( hWnd, "����:" & STR$( ERR ) & " �����ļ���������,������.   ", _
            " �����ļ�", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL )
    EXIT FUNCTION
  END TRY
  nLen = SENDMESSAGE( GetEdit, %SCI_GETTEXTLENGTH, 0, 0 )'ȡ�ñ����洰�����ı�����
  Buffer = SPACE$( nLen + 1 )                 '������
  SENDMESSAGE GetEdit, %SCI_GETTEXT, BYVAL LEN( Buffer ), BYVAL STRPTR( Buffer )'ȡ���ı���������
  ' Remove trailing spaces and tabs   ɾ��β���Ŀո��Ʊ��
  IF TrimTrailingBlanks THEN
    DO
      p = LEN( buffer )
      REPLACE " " & $CR WITH $CR IN buffer
      REPLACE $TAB & $CR WITH $CR IN buffer
    LOOP UNTIL p = LEN( buffer )
  END IF
  TRY
    PUT$ nFile, LEFT$( Buffer, LEN( Buffer ) - 1 )
    SETEOF nFile
  CATCH
    MESSAGEBOX( hWnd, "����" & STR$( ERR ) & " [" & ERROR$ & "] �����ļ���������,������.   ", _
            " �����ļ�", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL )
  FINALLY
    CLOSE nFile
  END TRY
  ' Tell to Scintilla that the current state of the document is unmodified.���õ�ǰ�ĵ���״̬Ϊδ�޸�
  SENDMESSAGE GetEdit, %SCI_SETSAVEPOINT, 0, 0
  ' If it is an Initialized as new window do not continue������´��ڣ����ٽ�������Ĳ�����ֱ���˳�
  IF ISTRUE fIsInitAsNew THEN FUNCTION = - 1 : EXIT FUNCTION
  IF ISTRUE Ask THEN
    ' Update the array of full paths����ȫ·������
    nTab = TabCtrl_GetCurSel( g_hTabMdi )
    IF UBOUND( gTabFilePaths ) > - 1 AND nTab < = UBOUND( gTabFilePaths ) THEN
      gTabFilePaths( nTab ) = f
    END IF
    ' Update the window caption and the tab name���´��ڱ��⼰tab��
    SETWINDOWTEXT MdiGetActive( g_hWndClient ), BYVAL STRPTR( f )
    CALL SetTabName( nTab, GetFileName( f ))
  END IF
  ' Update reopen file list (MRU menu)�����ش��ļ��б�MRU�˵���
  WriteRecentFiles f
  ' Retrieve the file time and store it in the properties list of the window�õ��ļ�ʱ�䣬���洢�����ڵ������б���
  fTime = SED_GetFileTime( f )
  SETPROP MdiGetActive( g_hWndClient ), "FTIME", fTime
  fPos = INSTR( - 1, f, "." )
  IF fPos THEN
    fExt = MID$( f, fPos )
    IF UCASE$( fExt ) <> UCASE$( szExt ) THEN
      pSci = SENDMESSAGE( GetEdit, %SCI_GETDIRECTPOINTER, 0, 0 )
      Scintilla_SetOptions pSci, f
    END IF
  END IF
  DIM hMdi AS DWORD
  hMdi=TabCtrl_GetCurSel(g_hTabMdi)
  DIM  ttc_item AS TC_ITEM
  DIM  tmpStr1   AS ASCIIZ * 100
  ttc_item.mask=%TCIF_TEXT 'or %TCIF_IMAGE OR %TCIF_PARAM OR %TCIF_RTLREADING
  ttc_item.cchTextMax=99
  ttc_item.pszText=VARPTR(tmpStr1)
  IF ISTRUE(TabCtrl_GetItem(g_hTabMdi,hMdi,ttc_item)) THEN
    'tmpPtr=ttc_item.pszText
    'tmpStr=@tmpPtr ' & " *"
    'ttc_item.cchTextMax=100
    'tmpStr=mid$(tmpStr,1,len(tmpStr)-1)
    IF RIGHT$(TRIM$(tmpStr1),2)=" *" THEN
      tmpStr=MID$(TRIM$(tmpStr1),1,LEN(TRIM$(tmpStr1))-2)
      ttc_item.pszText=VARPTR(tmpStr1)
      TabCtrl_SetItem g_hTabMdi,hMdi,ttc_item
    END IF

  END IF
  FUNCTION = - 1
END FUNCTION
' *********************************************************************************************
' ��ȡ�ļ���
' *********************************************************************************************
FUNCTION GetFileName( BYVAL Src AS STRING ) AS STRING
  LOCAL x AS LONG
  x = INSTR( - 1, Src, ANY ":/\" )
  IF x THEN
    FUNCTION = MID$( Src, x + 1 )
  ELSE
    FUNCTION = Src
  END IF
END FUNCTION
' *********************************************************************************************
' ��ȡ·��
' *********************************************************************************************
FUNCTION GetFilePath( BYVAL Src AS STRING ) AS STRING
  LOCAL x AS LONG
  x = INSTR( - 1, Src, ANY ":\/" )
  IF x THEN FUNCTION = LEFT$( Src, x )
END FUNCTION
' *********************************************************************************************
' ת��ѡ��ĵ��ʵ�ָ���Ĵ�д��ʽ
' fCase = 1 (��д), 2 (Сд), 3 (����ĸ��д(�շ�))
' *********************************************************************************************
SUB ChangeSelectedTextCase (BYVAL fCase AS LONG)
  LOCAL txtrg       AS TEXTRANGE   ' �ı���Χ�ṹ
  LOCAL startSelPos AS LONG        ' ��ʼλ��
  LOCAL endSelPos   AS LONG        ' ����λ��
  LOCAL buffer      AS STRING      ' �������
  ' ��� startSelPos �� endSelPos ��ȣ������û��ѡ������
  startSelPos = SendMessage(GetEdit, %SCI_GETSELECTIONSTART, 0, 0)
  endSelPos = SendMessage(GetEdit, %SCI_GETSELECTIONEND, 0, 0)
  IF startSelPos = endSelPos THEN EXIT SUB
  ' �������������ռ�
  buffer = SPACE$(endSelPos - startSelPos + 1)
  ' �ı���Χ
  txtrg.chrg.cpMin = startSelPos
  txtrg.chrg.cpMax = endSelPos
  txtrg.lpstrText = STRPTR(buffer)
  ' ����ı�
  SendMessage GetEdit, %SCI_GETTEXTRANGE, 0, BYVAL VARPTR(txtrg)
  ' תΪ��Ӧ�Ĵ�Сд��ʽ
  IF fCase = 1 THEN
    buffer = UCASE$(buffer)
  ELSEIF fCase = 2 THEN
    buffer = LCASE$(buffer)
  ELSEIF fCase = 3 THEN
    buffer = MCASE$(buffer)
  END IF
  ' �滻ѡ�е��ı�
  SendMessage GetEdit, %SCI_REPLACESEL, 0, BYVAL STRPTR(buffer)
END SUB
' *********************************************************************************************
' �������� SUB/FUNCTION �ڣ��򷵻���
' *********************************************************************************************
SUB IsWithinProc (BYREF tmpPROC AS PROC)
  LOCAL curPos           AS LONG                 ' // ��ǰλ��
  LOCAL LineNumber       AS LONG                 ' // ����
  LOCAL LineLen          AS LONG                 ' // �г���
  LOCAL i                AS LONG                 ' // ѭ������
  LOCAL buffer           AS STRING               ' // ����
  LOCAL Namebuffer       AS STRING               ' // ���溯�����Ļ���
  LOCAL LineCount        AS LONG                 ' // ����


  curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)                    ' ��ǰλ��
  LineNumber = SendMessage(GetEdit, %SCI_LINEFROMPOSITION, curPos, 0)        ' ��ǰ����
  LineCount = SendMessage(GetEdit, %SCI_GETLINECOUNT, 0, 0)                  ' ������

  tmpPROC.WhatIsUp     = %ID_ProcNull                                        ' Start with false
  tmpPROC.UpLnNo       = 0
  tmpPROC.WhatIsDown   = %ID_ProcNull                                        ' Start with false
  tmpPROC.DnLnNo       = LineCount
'   tmpPROC.ProcName = "(Code finder)"
  tmpPROC.ProcName = ""

  ' Check current line
  LineLen = SendMessage(GetEdit, %SCI_LINELENGTH, LineNumber, 0)             ' Length of the line
  buffer = SPACE$(LineLen)                                                   ' Size the buffer
  SendMessage(GetEdit, %SCI_GETLINE, LineNumber , STRPTR(buffer))            ' Get the text of the line
  DO
    IF ISFALSE INSTR(buffer, "  ") THEN EXIT DO
    REPLACE "  " WITH " " IN buffer
  LOOP
  Namebuffer = LTRIM$(buffer)
  buffer = LTRIM$(UCASE$(buffer))
  IF LEFT$(buffer, 4) = "SUB " OR LEFT$(buffer, 9) = "FUNCTION " OR _
          (LEFT$(buffer, 7) = "STATIC " AND MID$(buffer, 8, 3) = "SUB") OR _
          (LEFT$(buffer, 7) = "STATIC " AND MID$(buffer, 8, 8) = "FUNCTION") OR _
          (LEFT$(buffer, 9) = "CALLBACK " AND MID$(buffer, 10, 8) = "FUNCTION") THEN
    IF LEFT$(buffer, 9) = "FUNCTION " THEN
      Namebuffer = LTRIM$(MID$(Namebuffer, 10))
      buffer = LTRIM$(MID$(buffer, 10))
      IF LEFT$(buffer, 1) <> "=" THEN
        tmpPROC.WhatIsUp = %ID_ProcStart ' Proc. Begins
        tmpPROC.UpLnNo = LineNumber
      END IF
    ELSE
      tmpPROC.WhatIsUp = %ID_ProcStart ' Proc. Begins
      tmpPROC.UpLnNo = LineNumber
    END IF
    Namebuffer = TRIM$(MID$(Namebuffer, 1, INSTR(Namebuffer, ANY CHR$(40,13,10)) - 1 ))
    tmpPROC.ProcName = PARSE$(Namebuffer,CHR$(32),PARSECOUNT(Namebuffer,CHR$(32)))
  ELSEIF (LEFT$(buffer, 4) = "END " AND MID$(buffer, 5, 3) = "SUB") OR _
          (LEFT$(buffer, 4) = "END " AND MID$(buffer, 5, 8) = "FUNCTION") THEN
    tmpPROC.WhatIsDown = %ID_ProcEnd ' Proc. Ends
    tmpPROC.DnLnNo = LineNumber
  END IF

  ' Check for begining if not found
  IF tmpPROC.WhatIsUp = %ID_ProcNull THEN
    FOR i = LineNumber - 1 TO 0 STEP -1
      LineLen = SendMessage(GetEdit, %SCI_LINELENGTH, i, 0)                   ' Length of the line
      buffer = SPACE$(LineLen)                                                ' Size the buffer
      SendMessage(GetEdit, %SCI_GETLINE, i , STRPTR(buffer))                  ' Get the text of the line
      DO
        IF ISFALSE INSTR(buffer, "  ") THEN EXIT DO
        REPLACE "  " WITH " " IN buffer
      LOOP
      Namebuffer = LTRIM$(buffer)
      buffer = LTRIM$(UCASE$(buffer))
      IF LEFT$(buffer, 4) = "SUB " OR LEFT$(buffer, 9) = "FUNCTION " OR _
              (LEFT$(buffer, 7) = "STATIC " AND MID$(buffer, 8, 3) = "SUB") OR _
              (LEFT$(buffer, 7) = "STATIC " AND MID$(buffer, 8, 8) = "FUNCTION") OR _
              (LEFT$(buffer, 9) = "CALLBACK " AND MID$(buffer, 10, 8) = "FUNCTION") THEN
        IF LEFT$(buffer, 9) = "FUNCTION " THEN
          Namebuffer = LTRIM$(MID$(Namebuffer, 10))
          buffer = LTRIM$(MID$(buffer, 10))
          IF LEFT$(buffer, 1) <> "=" THEN
            tmpPROC.WhatIsUp = %ID_ProcStart ' Proc. Begins
            tmpPROC.UpLnNo = i
            Namebuffer = TRIM$(MID$(Namebuffer, 1, INSTR(Namebuffer, ANY CHR$(40,13,10)) - 1 ))
            tmpPROC.ProcName = PARSE$(Namebuffer,CHR$(32),PARSECOUNT(Namebuffer,CHR$(32)))
            EXIT FOR
          END IF
        ELSE
          tmpPROC.WhatIsUp = %ID_ProcStart ' Proc. Begins
          tmpPROC.UpLnNo = i
          Namebuffer = TRIM$(MID$(Namebuffer, 1, INSTR(Namebuffer, ANY CHR$(40,13,10)) - 1 ))
          tmpPROC.ProcName = PARSE$(Namebuffer,CHR$(32),PARSECOUNT(Namebuffer,CHR$(32)))
          EXIT FOR
        END IF
      ELSEIF (LEFT$(buffer, 4) = "END " AND MID$(buffer, 5, 3) = "SUB") OR _
            (LEFT$(buffer, 4) = "END " AND MID$(buffer, 5, 8) = "FUNCTION") THEN
        tmpPROC.WhatIsUp = %ID_ProcEnd ' Proc. Ends
        tmpPROC.UpLnNo = i
        EXIT FOR
      END IF
    NEXT
  END IF

  ' Check for End
  IF tmpPROC.WhatIsDown  = %ID_ProcNull THEN
    FOR i = LineNumber + 1 TO LineCount
      LineLen = SendMessage(GetEdit, %SCI_LINELENGTH, i, 0)                   ' Length of the line
      buffer = SPACE$(LineLen)                                                ' Size the buffer
      SendMessage(GetEdit, %SCI_GETLINE, i , STRPTR(buffer))                  ' Get the text of the line
      DO
        IF ISFALSE INSTR(buffer, "  ") THEN EXIT DO
        REPLACE "  " WITH " " IN buffer
      LOOP
      buffer = LTRIM$(UCASE$(buffer))
      IF (LEFT$(buffer, 4) = "END " AND MID$(buffer, 5, 3) = "SUB") OR _
              (LEFT$(buffer, 4) = "END " AND MID$(buffer, 5, 8) = "FUNCTION") THEN
        tmpPROC.WhatIsDown = %ID_ProcEnd
        tmpPROC.DnLnNo = i
        EXIT FOR
      ELSE
        IF LEFT$(buffer, 4) = "SUB " OR LEFT$(buffer, 9) = "FUNCTION " OR _
                (LEFT$(buffer, 7) = "STATIC " AND MID$(buffer, 8, 3) = "SUB") OR _
                (LEFT$(buffer, 7) = "STATIC " AND MID$(buffer, 8, 8) = "FUNCTION") OR _
                (LEFT$(buffer, 9) = "CALLBACK " AND MID$(buffer, 10, 8) = "FUNCTION") THEN
          IF LEFT$(buffer, 9) = "FUNCTION " THEN
            buffer = LTRIM$(MID$(buffer, 10))
            IF LEFT$(buffer, 1) <> "=" THEN
              tmpPROC.WhatIsDown = %ID_ProcStart
              tmpPROC.DnLnNo = i
              EXIT FOR
            END IF
          ELSE
            tmpPROC.WhatIsDown = %ID_ProcStart
            tmpPROC.DnLnNo = i
            EXIT FOR
          END IF
        END IF
      END IF
     NEXT
  END IF
END SUB
' *********************************************************************************************
' ת����ǰ SUB/FUNCTION ��ͷ
' *********************************************************************************************
FUNCTION GotoBeginThisProc () AS LONG
  LOCAL curPos           AS LONG                 ' // ��ǰλ��
  LOCAL tmpPROC          AS PROC
  SED_WithinProc(tmpPROC)
  IF tmpPROC.WhatIsUp =  %ID_ProcEnd AND tmpPROC.WhatIsDown = %ID_ProcStart OR _
          tmpPROC.WhatIsUp =  %ID_ProcNull AND tmpPROC.WhatIsDown = %ID_ProcStart THEN EXIT FUNCTION

  IF tmpPROC.WhatIsUp =  %ID_ProcStart THEN
    SendMessage GetEdit, %SCI_GOTOLINE, tmpPROC.UpLnNo, 0             ' Set the caret position
    curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)           ' Set the current position
  ELSE
    SendMessage GetEdit, %SCI_GOTOLINE, tmpPROC.UpLnNo + 1, 0         ' Set the caret position
    curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)           ' Set the current position
    SED_MSGBOX g_hWndMain, "�Ҳ��� FUNCTION/SUB �Ŀ�ͷ"
  END IF
  FUNCTION = curPos
END FUNCTION
' *********************************************************************************************
' ת����ǰ SUB/FUNCTION ĩβ
' *********************************************************************************************
FUNCTION GotoEndThisProc () AS LONG
  LOCAL curPos           AS LONG                 ' // Current position
  LOCAL tmpPROC          AS PROC
  SED_WithinProc(tmpPROC)
  IF tmpPROC.WhatIsUp =  %ID_ProcEnd AND tmpPROC.WhatIsDown = %ID_ProcStart OR _
          tmpPROC.WhatIsUp =  %ID_ProcNull AND tmpPROC.WhatIsDown = %ID_ProcStart THEN EXIT FUNCTION
  IF tmpPROC.WhatIsDown =  %ID_ProcEnd THEN
    SendMessage GetEdit, %SCI_GOTOLINE, tmpPROC.DnLnNo , 0            ' Set the caret position
    curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)           ' Set the current position
  ELSE
    SendMessage GetEdit, %SCI_GOTOLINE, tmpPROC.DnLnNo - 1, 0         ' Set the caret position
    curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)           ' Set the current position
    SED_MSGBOX g_hWndMain, "û�ҵ� FUNCTION/SUB ��ĩβ"
  END IF
  FUNCTION = curPos
END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' ת�� SUB/FUNCTION ��ͷ
' *********************************************************************************************
FUNCTION GotoBeginProc () AS LONG

  LOCAL curPos           AS LONG                 ' // Current position
  LOCAL LineNumber       AS LONG                 ' // Number of line
  LOCAL LineLen          AS LONG                 ' // Length of the line
  LOCAL fIsNotFunction   AS LONG                 ' // Flag
  LOCAL i                AS LONG                 ' // Loop counter
  LOCAL buffer           AS STRING               ' // buffer

  curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)                    ' Current position
  LineNumber = SendMessage(GetEdit, %SCI_LINEFROMPOSITION, curPos, 0) - 1    ' Line number
  FOR i = LineNumber TO 0 STEP -1
    LineLen = SendMessage(GetEdit, %SCI_LINELENGTH, i, 0)                   ' Length of the line
    buffer = SPACE$(LineLen)                                                ' Size the buffer
    SendMessage(GetEdit, %SCI_GETLINE, i, STRPTR(buffer))                   ' Get the text of the line
    fIsNotFunction = %FALSE
    buffer = LTRIM$(UCASE$(buffer))
    DO
      IF ISFALSE INSTR(buffer, "  ") THEN EXIT DO
      REPLACE "  " WITH " " IN buffer
    LOOP
    IF LEFT$(buffer, 4) = "SUB " OR LEFT$(buffer, 9) = "FUNCTION " OR _
            (LEFT$(buffer, 7) = "STATIC " AND MID$(buffer, 8, 3) = "SUB") OR _
            (LEFT$(buffer, 7) = "STATIC " AND MID$(buffer, 8, 8) = "FUNCTION") OR _
            (LEFT$(buffer, 9) = "CALLBACK " AND MID$(buffer, 10, 8) = "FUNCTION") THEN
      IF LEFT$(buffer, 9) = "FUNCTION " THEN
        buffer = LTRIM$(MID$(buffer, 10))
        IF LEFT$(buffer, 1) = "=" THEN fIsNotFunction = %TRUE
      END IF
      IF ISFALSE fIsNotFunction THEN
        SendMessage GetEdit, %SCI_GOTOLINE, i, 0                          ' Set the caret position
        curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)           ' Set the current position
        EXIT FOR
      END IF
    END IF
  NEXT
  FUNCTION = curPos
END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' ת�� SUB/FUNCTION ĩβ
' *********************************************************************************************
FUNCTION GotoEndProc () AS LONG
  LOCAL curPos           AS LONG                 ' // Current position
  LOCAL LineNumber       AS LONG                 ' // Number of line
  LOCAL LineCount        AS LONG                 ' // Number of lines
  LOCAL LineLen          AS LONG                 ' // Length of the line
  LOCAL i                AS LONG                 ' // Loop counter
  LOCAL buffer           AS STRING               ' // buffer
  LineCount = SendMessage(GetEdit, %SCI_GETLINECOUNT, 0, 0)                  ' Number of lines
  curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)                    ' Current position
  LineNumber = SendMessage(GetEdit, %SCI_LINEFROMPOSITION, curPos, 0) + 1    ' Line number
  FOR i = LineNumber TO LineCount
    LineLen = SendMessage(GetEdit, %SCI_LINELENGTH, i, 0)                   ' Length of the line
    buffer = SPACE$(LineLen)                                                ' Size the buffer
    SendMessage(GetEdit, %SCI_GETLINE, i , STRPTR(buffer))                  ' Get the text of the line
    buffer = LTRIM$(UCASE$(buffer))
    DO
      IF ISFALSE INSTR(buffer, "  ") THEN EXIT DO
      REPLACE "  " WITH " " IN buffer
    LOOP
    IF (LEFT$(buffer, 4) = "END " AND MID$(buffer, 5, 3) = "SUB") OR _
      (LEFT$(buffer, 4) = "END " AND MID$(buffer, 5, 8) = "FUNCTION") THEN
      SendMessage GetEdit, %SCI_GOTOLINE, i, 0                             ' Set the caret position
      curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)              ' Set the current position
      EXIT FOR
    END IF
  NEXT
  FUNCTION = curPos
END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' ���Ʊ���滻�ո�
' *********************************************************************************************
SUB ReplaceSpacesWithTabs
  LOCAL nLen        AS LONG
  LOCAL Buffer      AS STRING
  LOCAL TabSize     AS LONG
  LOCAL LineNumber  AS LONG
  LineNumber = GetCurrentLine
  TabSize = SendMessage(GetEdit, %SCI_GETTABWIDTH, 0, 0)
  IF TabSize < 1 THEN EXIT SUB
  nLen = SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
  Buffer = SPACE$(nLen + 1)
  SendMessage GetEdit, %SCI_GETTEXT, BYVAL LEN(Buffer), BYVAL STRPTR(Buffer)
  REPLACE SPACE$(TabSize) WITH $TAB IN Buffer
  SendMessage GetEdit, %SCI_SETTEXT, 0, STRPTR(buffer)
  SendMessage GetEdit, %SCI_GOTOLINE, LineNumber, 0    ' Set the caret position
END SUB
' *********************************************************************************************

' *********************************************************************************************
' �ÿո��滻�Ʊ��
' *********************************************************************************************
SUB ReplaceTabsWithSpaces
  LOCAL nLen        AS LONG
  LOCAL Buffer      AS STRING
  LOCAL TabSize     AS LONG
  LOCAL LineNumber  AS LONG
  LineNumber = GetCurrentLine
  TabSize = SendMessage(GetEdit, %SCI_GETTABWIDTH, 0, 0)
  IF TabSize < 1 THEN EXIT SUB
  nLen = SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
  Buffer = SPACE$(nLen + 1)
  SendMessage GetEdit, %SCI_GETTEXT, BYVAL LEN(Buffer), BYVAL STRPTR(Buffer)
  REPLACE $TAB WITH SPACE$(TabSize) IN Buffer
  SendMessage GetEdit, %SCI_SETTEXT, 0, STRPTR(buffer)
  SendMessage GetEdit, %SCI_GOTOLINE, LineNumber, 0    ' Set the caret position
END SUB
' *********************************************************************************************

' *********************************************************************************************
' ճ����ҳ����
' *********************************************************************************************
SUB PasteIE
  STATIC ClipFormat       AS LONG                 ' // Registered clipboard format
  LOCAL  cp               AS DWORD                ' // Retrieved clipboard format
  LOCAL  hClipData        AS ASCIIZ PTR           ' // Clipboard's handle
  LOCAL  strClipData      AS STRING               ' // Clipboard data
  LOCAL  hEdit            AS DWORD                ' // Handle of the edit control
  hEdit = GetEdit
  IF ISFALSE hEdit THEN EXIT SUB
  ' Open the clipboard
  OpenClipboard %NULL
  ' Register the HTML clipboard format name
  IF ISFALSE ClipFormat THEN ClipFormat = RegisterClipboardFormat("HTML Format")
  cp = 0
  DO
    ' Enumerate the data formats available in the clipboard
    cp = EnumClipboardFormats(cp)
    ' Exit on failure
    IF ISFALSE cp THEN EXIT DO
    ' If it contains an HTML document...
    IF cp = ClipFormat THEN
      ' Retrieves his handke...
      hClipData = GetClipboardData(cp)
      IF hClipData THEN
        ' ..and reads his contents
        strClipData = @hClipData
        EXIT DO
      END IF
    END IF
  LOOP
  ' Closes the clipboard
  CloseClipboard
  IF LEN(strClipData) THEN
    ' Remove HTML tags
    strClipData = RemoveHtml(strClipData)
    ' Insert the text at the current position
    SendMessage hEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(strClipData)
  ELSE
    ' Normal paste action
    SendMessage hEdit, %SCI_PASTE, 0, 0
  END IF
  ' Show the line an column
  ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' ���������İ���
' *********************************************************************************************
SUB SearchContextHelp (BYVAL hWnd AS DWORD)

   ' Context help
   LOCAL x                AS LONG                  ' // Starting position of a word
   LOCAL y                AS LONG                  ' // Ending position of a word
   LOCAL txtrg            AS TEXTRANGE             ' // Text range
   LOCAL DefCompiler      AS LONG                  ' // Default compiler
   LOCAL curPos           AS LONG                  ' // Current position
   LOCAL buffer           AS STRING                ' // General purpose buffer
   LOCAL p                AS LONG                  ' // General purpose variable
   LOCAL szPath           AS ASCIIZ * %MAX_PATH    ' // Path
   LOCAL hr               AS LONG                  ' // Result code

   ' Retrieve the word under the caret and activate the compiler's help file
   ' if it is a Power Basic keyword or the Windows 32 help file is it is not.
   x = 0 : y = 0
   ' Retrieve the current position
   curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)
   ' Retrieve the starting position of the word
   x = SendMessage(GetEdit, %SCI_WORDSTARTPOSITION, curPos, %TRUE)
   ' Retrieve the ending position of the word
   y = SendMessage(GetEdit, %SCI_WORDENDPOSITION, curPos, %TRUE)
   IF y > x THEN
    ' Prepare the buffer
    buffer = SPACE$(y - x + 1)
    ' Read also the characters before and after the word
    IF x > 0 THEN x = x - 1 : buffer = buffer + " "
    y = y + 1 : buffer = buffer + " "
    ' Text range
    txtrg.chrg.cpMin = x
    txtrg.chrg.cpMax = y
    txtrg.lpstrText = STRPTR(buffer)
    SendMessage GetEdit, %SCI_GETTEXTRANGE, 0, BYVAL VARPTR(txtrg)
    ' Remove the $NUL
    p = INSTR(buffer, CHR$(0))
    IF p THEN  buffer = LEFT$(buffer, p - 1)
    ' We need to preserve $ and # for PowerBasic keywords
    IF RIGHT$(buffer, 1) <> "$" THEN buffer = LEFT$(buffer, LEN(buffer) - 1)
    IF x > 0 THEN
     IF LEFT$(buffer, 1) <> "%" AND LEFT$(buffer, 1) <> "#" THEN buffer = MID$(buffer, 2)
    END IF
    ' If it begins with % and it is not a PB keyword, then remove it to call Win32Help
    IF LEFT$(buffer, 1) = "%" AND INSTR(strPBKeywords, " " & LCASE$(MID$(buffer, 2)) & " ") = 0 THEN buffer = MID$(buffer, 2)
    IF LEFT$(buffer, 1) = "#" AND INSTR(strPBKeywords, " " & LCASE$(MID$(buffer, 2)) & " ") = 0 THEN buffer = MID$(buffer, 2)
    buffer = TRIM$(buffer)
    IF LEN(buffer) THEN
     IF INSTR(strPBKeywords, LCASE$(buffer)) THEN
      ' If it is a Power Basic keyword, get the Default compiler and
      ' call the appropiate help file, else call the Win32 help file
      DefCompiler = VAL(IniRead(g_zIni, "Compiler options", "DefaultCompiler", ""))
      szPath = ""
      IF DefCompiler = 1 THEN
         szPath = IniRead(g_zIni, "Tools options", "PBWinHelp", "")
      ELSE
         szPath = IniRead(g_zIni, "Tools options", "PBCCHelp", "")
      END IF
      IF LEN(szPath) THEN WinHelp hWnd, szPath, %HELP_KEY, BYVAL STRPTR(buffer)
     ELSE
      ' Get the path of the file
      GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
      IF INSTR(UCASE$(szPath), ".RC") THEN     ' It is a resource file
         szPath = IniRead(g_zIni, "Tools options", "RCHelp", "")
         IF LEN(szPath) THEN WinHelp hWnd, szPath, %HELP_KEY, BYVAL STRPTR(buffer)
      ELSE
         szPath = IniRead(g_zIni, "Tools options", "Win32Help", "")
         IF LEN(szPath) THEN WinHelp hWnd, szPath, %HELP_KEY, BYVAL STRPTR(buffer)
      END IF
     END IF
    END IF
   ELSE
    buffer = ""
    DefCompiler = VAL(IniRead(g_zIni, "Compiler options", "DefaultCompiler", ""))
    szPath = ""
    IF DefCompiler = 1 THEN
     szPath = IniRead(g_zIni, "Tools options", "PBWinHelp", "")
    ELSE
     szPath = IniRead(g_zIni, "Tools options", "PBCCHelp", "")
    END IF
    IF LEN(szPath) THEN hr = ShellExecute(hWnd, "open", szPath, "", "", %SW_SHOWNORMAL)
   END IF

END SUB
' *********************************************************************************************
' *********************************************************************************************
' Get the primary source file (if any)
' *********************************************************************************************
FUNCTION GetPrimarySourceFile (BYVAL fLoad AS LONG) AS STRING

  LOCAL hWndActive AS DWORD                ' // Handle of the active child window
  LOCAL szPath     AS ASCIIZ * %MAX_PATH   ' // File path
  LOCAL p          AS LONG                 ' // Number of child windows
  LOCAL p1         AS LONG                 ' // Window handle
  LOCAL nTab       AS LONG                 ' // Tab number

  ' Count the number of child windows
  p = 0
  p1 = GetWindow(g_hWndClient, %GW_CHILD)
  DO WHILE p1 <> 0
    p1 = GetWindow(p1, %GW_HWNDNEXT)
    INCR p
  LOOP

  ' Save files as needed
  DO WHILE p <> 0
    hWndActive = MdiGetActive(g_hWndClient)
    IF ISTRUE SendMessage(GetEdit, %SCI_GETMODIFY, 0, 0) THEN ' Is it modified ?
      GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
      ' A previous saved file has a path in the child window
      IF (INSTR(szPath, ANY ":\/") = 0) THEN   ' Path available?
        ' No Path - it is a new file - save with prompt, except untitled files
        IF LEFT$(UCASE$(GetFileName(szPath)), 8) <> "UNTITLED" THEN
            SAVEFILE(%TRUE)
        END IF
      ELSE
        SAVEFILE(%FALSE)     ' Modified file found - Save it
      END IF
      UpdateWindow g_hToolbar
    END IF
    MdiNext g_hWndClient, hWndActive, 0      ' Set focus to next child window
    DECR p                                 ' Decrement window counter
  LOOP

  p = GetWindow(g_hWndClient, %GW_CHILD) ' first look at already opened docs
  WHILE p
    GetWindowText p, szPath, %MAX_PATH ' get name of child window
    IF UCASE$(szPath) = UCASE$(PrimarySourceFile) THEN ' is it the same as primary source file ?
      ARRAY SCAN gTabFilePaths(), = szPath, TO nTab
      IF nTab THEN
        SendMessage g_hTabMdi, %TCM_SETCURSEL, nTab - 1, 0  ' activate the tab associated with the window
        SendMessage g_hWndClient, %WM_MDIACTIVATE, p, 0 ' yes, activate it to top window before compile.
      END IF
      FUNCTION = szPath
      EXIT FUNCTION
    END IF
    p = GetWindow(p, %GW_HWNDNEXT)
   WEND

  ' If it is not loaded, try to load it
  IF ISTRUE fLoad THEN
    IF LEN(PrimarySourceFile) THEN
      IF FileExist(PrimarySourceFile) THEN
        IF OpenThisFile(PrimarySourceFile) THEN
          FUNCTION = PrimarySourceFile
        END IF
      END IF
    END IF
  END IF
END FUNCTION
' *********************************************************************************************
' �� Windows ��ʱ�ļ����б���δ�����ļ�
' *********************************************************************************************
FUNCTION SaveUntitledFile (BYVAL hWnd AS DWORD, szPath AS ASCIIZ) AS LONG

  LOCAL buffer AS STRING               ' // Buffer
  LOCAL szStr  AS ASCIIZ * %MAX_PATH   ' // Temporary file path
  LOCAL nFile  AS LONG                 ' // File number
  LOCAL nLen   AS LONG                 ' // Text lenght

  IF LEN(szPath) = 0 THEN EXIT FUNCTION
  buffer = SPACE$(%MAX_PATH)
  GetTempPath(SIZEOF(szStr), szStr)
  CHDIR szStr
  szPath = szStr + szPath
  nFile = FREEFILE
  OPEN szPath FOR BINARY AS nFile
  IF ERR THEN
    MessageBox(hWnd, "�����ļ�ʱ����" & STR$(ERR)," SaveFile", _
            %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL)
    EXIT FUNCTION
  END IF
  nLen = SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
  Buffer = SPACE$(nLen + 1)
  SendMessage GetEdit, %SCI_GETTEXT, BYVAL LEN(Buffer), BYVAL STRPTR(Buffer)
  PUT$ nFile, LEFT$(Buffer, LEN(Buffer) - 1)
  SETEOF nFile
  CLOSE nFile
  FUNCTION = %TRUE
END FUNCTION
' *********************************************************************************************
' Shows the standard file properties dialog. Returns an HRESULT code.
' *********************************************************************************************
FUNCTION ShowFileProperties(BYVAL hWnd AS DWORD, BYREF szFileName AS ASCIIZ) AS LONG

  LOCAL SEI AS SHELLEXECUTEINFO
  LOCAL szVerb AS ASCIIZ * %MAX_PATH
  LOCAL szNull AS ASCIIZ * 1
  LOCAL ecode AS DWORD

  szVerb = "Properties"
  SEI.cbSize = SIZEOF(SEI)
  SEI.fmask = _ '%SEE_MASK_NOCLOSEPROCESS OR _
          %SEE_MASK_INVOKEIDLIST OR _
          %SEE_MASK_FLAG_NO_UI
  SEI.hWnd = hWnd
  SEI.lpVerb = VARPTR(szVerb)
  SEI.lpFile = VARPTR(szFileName)
  SEI.lpParameters = VARPTR(szNull)
  SEI.lpDirectory = VARPTR(szNull)
  SEI.nShow = 0
  SEI.hInstApp = 0
  SEI.lpIDList = 0
  ShellExecuteEx SEI
  IF SEI.hInstApp < 32 THEN
    ecode = GetLastError
    Messagebox 0, SED_WinErrorMsg(ecode), "Operation Failed", %MB_ICONINFORMATION
    FUNCTION = SEI.hInstApp                       'error code
    EXIT FUNCTION
  END IF
END FUNCTION
