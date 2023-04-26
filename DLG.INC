' *********************************************************************************************
' SED_DLG.INC - Include file for the SED code editor (popup dialogs)
' *********************************************************************************************

' *********************************************************************************************
' Declares
' *********************************************************************************************
DECLARE FUNCTION ShowGotoLinePopupDialog (BYVAL hParent AS LONG) AS LONG
DECLARE FUNCTION GotoLinePopupDlgProc (BYVAL hWnd AS DWORD, BYVAL wMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG
DECLARE FUNCTION SED_InputBox (BYVAL hParent AS DWORD, BYVAL strCaption AS STRING, BYVAL strPrompt AS STRING, BYVAL strText AS STRING, OPTIONAL BYVAL nLen AS LONG) AS STRING
DECLARE CALLBACK FUNCTION SED_InputBoxDlgProc
DECLARE FUNCTION TemplateCodePopupDialog (BYVAL hParent AS LONG) AS LONG
DECLARE FUNCTION TemplateCodePopupDlgProc (BYVAL hWnd AS DWORD, BYVAL wMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG
DECLARE SUB InsertTemplateCode (BYVAL hWnd AS DWORD, BYVAL curSel AS LONG)
DECLARE SUB MainDialogTemplate (BYVAL hWnd AS DWORD, BYVAL fGradient AS LONG)
DECLARE SUB PopupDialogTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ButtonTemplate (BYVAL hWnd AS DWORD, BYVAL fFlat AS LONG)
DECLARE SUB CheckboxTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB CheckStateTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ComboboxDropdownTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ComboboxDropdownListTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ComboboxTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB FrameTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ImageTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ImageXTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ImageButtonTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ImageButtonXTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB LabelTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB LineTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB ListboxTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB RadioButtonTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB HorizScrollbarTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB VertScrollbarTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB TextboxTemplate (BYVAL hWnd AS DWORD)
DECLARE SUB MultilineTextboxTemplate (BYVAL hWnd AS DWORD)
DECLARE FUNCTION SED_CreateCodeKeeperDatabase (BYVAL hWnd AS DWORD, BYVAL strDbPath AS STRING) AS LONG
DECLARE CALLBACK FUNCTION ShowCodeKeeperDialogProc()
DECLARE FUNCTION ShowCodeKeeperDialog (BYVAL hParent AS DWORD) AS LONG
DECLARE SUB SED_ShowSysInfo (BYVAL hWnd AS DWORD)
' *********************************************************************************************

' *********************************************************************************************
' SED_InputBox - Input box dialog
' *********************************************************************************************
CALLBACK FUNCTION SED_InputBoxDlgProc

   STATIC lpText AS ASCIIZ PTR

   SELECT CASE CBMSG
   CASE %WM_USER + 1    : lpText = CBWPARAM
   CASE %WM_COMMAND
      SELECT CASE LOWRD (CBWPARAM)
         CASE %IDOK     : CONTROL GET TEXT CBHNDL, 101 TO @lpTEXT
                          DIALOG END CBHNDL, %TRUE
         CASE %IDCANCEL : DIALOG END CBHNDL, %FALSE
      END SELECT
   END SELECT

END FUNCTION

FUNCTION SED_InputBox (BYVAL hParent AS DWORD, BYVAL strCaption AS STRING, BYVAL strPrompt AS STRING, BYVAL strText AS STRING, OPTIONAL BYVAL nLen AS LONG) AS STRING

   #REGISTER NONE
   LOCAL hiDlg AS DWORD, lResult AS LONG, szUser AS ASCIIZ * 2048
   LOCAL dwStyle AS DWORD, dwExStyle AS DWORD

   dwStyle = %WS_CHILD OR %WS_VISIBLE OR %ES_LEFT OR %ES_AUTOHSCROLL
   dwExStyle = %WS_EX_CLIENTEDGE OR %WS_EX_LEFT OR %WS_EX_LTRREADING

   DIALOG NEW hParent, strCaption, ,, 180, 70, %WS_CAPTION OR %WS_POPUPWINDOW OR %WS_VISIBLE, %WS_EX_DLGMODALFRAME OR %WS_EX_CONTROLPARENT TO hiDlg
   CONTROL ADD LABEL, hiDlg, -1, strPrompt, 14, 6, 160, 12, %SS_LEFT
   CONTROL ADD TEXTBOX, hiDlg, 101, "", 14, 20, 153, 12, dwStyle, dwExStyle
   CONTROL ADD BUTTON,  hiDlg, %IDOK, "È·¶¨", 14, 44, 50, 14, 1
   CONTROL ADD BUTTON,  hiDlg, %IDCANCEL, "È¡Ïû", 117, 44, 50, 14
   CONTROL SET TEXT hiDlg, 101, strText
   IF nLen < 1 OR nLen > 2047 THEN nLen = 2047
   SendMessage GetDlgItem(hiDlg, 101), %EM_LIMITTEXT, nLen, 0
   IF LEN(strText) > %MAX_PATH - 1 THEN strText = LEFT$(strText, %MAX_PATH - 1)
   SetFocus GetDlgItem (hiDlg, 101)
   ' Set pointer for return string in callback function
   CallWindowProc CODEPTR (SED_InputBoxDlgProc), hiDlg, %WM_USER+1, VARPTR (szUser), 0
   DIALOG SHOW MODAL hiDlg CALL SED_InputBoxDlgProc TO lResult
   IF lResult THEN FUNCTION = szUser

END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' Go to line popup dialog
' *********************************************************************************************
FUNCTION ShowGotoLinePopupDialog (BYVAL hParent AS LONG) AS LONG

   LOCAL hDlg AS LONG
   LOCAL hCtl AS LONG
   LOCAL rc AS RECT
   LOCAL wc AS WndClassEx
   LOCAL szClassName AS ASCIIZ * 80
   LOCAL hFnt AS DWORD

   LOCAL  curPos     AS LONG
   LOCAL  nLine      AS LONG
   LOCAL  nLines     AS LONG

   curPos = SendMessage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)
   nLine = SendMessage(GetEdit, %SCI_LINEFROMPOSITION, curPos, 0) + 1
   nLines = SendMessage(GetEdit, %SCI_GETLINECOUNT, 0, 0)

   hFnt = GetStockObject(%ANSI_VAR_FONT)

   szClassName      = "GotoLine"
   wc.cbSize        = SIZEOF(wc)
   wc.style         = %CS_HREDRAW OR %CS_VREDRAW
   wc.lpfnWndProc   = CODEPTR(GotoLinePopupDlgProc)
   wc.cbClsExtra    = 0
   wc.cbWndExtra    = 0
   wc.hInstance     = GetModuleHandle("")
   wc.hCursor       = LoadCursor(%NULL, BYVAL %IDC_ARROW)
   wc.hbrBackground = %COLOR_3DFACE + 1
   wc.lpszMenuName  = %NULL
   wc.lpszClassName = VARPTR(szClassName)
   wc.hIcon         = 0
   wc.hIconSm       = 0
   CALL RegisterClassEx (wc)

   GetWindowRect hParent, rc          ' for centering child in parent
   rc.nRight = rc.nRight - rc.nLeft   ' parent's width
   rc.nBottom = rc.nBottom - rc.nTop  ' parent's height

   hDlg = CreateWindowEx(%WS_EX_DLGMODALFRAME OR %WS_EX_CONTROLPARENT, szClassName, "×ªµ½ÐÐ", _
          %WS_CAPTION OR %WS_POPUPWINDOW OR %WS_VISIBLE, rc.nLeft + (rc.nRight - 261) / 2, _
          rc.nTop + (rc.nBottom - 123) / 2, 261, 123, hParent, 0, GetModuleHandle(""), BYVAL %NULL)

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Static", BYVAL %NULL, %WS_CHILD OR %WS_VISIBLE OR %SS_ETCHEDFRAME, _
          15, 13, 135, 68, hDlg, -1, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Static", "×î½üÐÐ :", %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %SS_RIGHT, _
          18, 20, 68, 13, hDlg,  -1, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Static", "", %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %SS_RIGHT, _
          90, 20, 48, 13, hDlg, -1, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0
   SetWindowText hCtl, FORMAT$(nLines)

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Static", "µ±Ç°ÐÐ :", %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %SS_RIGHT, _
          18, 37, 68, 13, hDlg,  -1, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Static", "", %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %SS_RIGHT, _
          90, 37, 48, 13, hDlg, -1, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0
   SetWindowText hCtl, FORMAT$(nLine)

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Static", "×ªµ½ÐÐ :", %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %SS_RIGHT, _
          18, 57, 68, 13, hDlg,  -1, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   hCtl = CreateWindowEx(%WS_EX_CLIENTEDGE OR %WS_EX_NOPARENTNOTIFY, "Edit", BYVAL %NULL, %WS_CHILD OR %WS_VISIBLE OR %WS_GROUP OR %WS_TABSTOP OR _
          %ES_NUMBER OR %ES_AUTOHSCROLL, 90, 55, 51, 20, hDlg, 100, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0
   SendMessage hCtl, %EM_LIMITTEXT, LEN(FORMAT$(nLines)), 0
   SetFocus hCtl

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Button", "È·¶¨", %WS_CHILD OR %WS_VISIBLE OR %WS_DISABLED OR _
          %WS_TABSTOP OR %BS_DEFPUSHBUTTON, 165, 13, 75, 23, hDlg, %IDOK, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Button", "È¡Ïû", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
          165, 41, 75, 23, hDlg, %IDCANCEL, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   ShowWindow hDlg, %SW_SHOW
   UpdateWindow hDlg

   LOCAL Msg AS tagMsg
   WHILE GetMessage(Msg, %NULL, 0, 0)
      IF ISFALSE IsDialogMessage(hDlg, Msg) THEN
         TranslateMessage Msg
         DispatchMessage Msg
      END IF
   WEND

   FUNCTION = msg.wParam

END FUNCTION

' *********************************************************************************************
' GotoLine Popup dialog procedure
' *********************************************************************************************
FUNCTION GotoLinePopupDlgProc (BYVAL hWnd AS DWORD, BYVAL wMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG

  LOCAL szBuffer AS ASCIIZ * 255
  LOCAL nLine AS LONG

  SELECT CASE wMsg
     CASE %WM_CREATE   ' Entrance
        EnableWindow GetParent(hWnd), %FALSE   ' To make popup dialog modal

     CASE %WM_COMMAND
        SELECT CASE LOWRD(wParam)
           CASE %IDCANCEL
              IF HIWRD(wParam) = %BN_CLICKED THEN
                 SendMessage hWnd, %WM_CLOSE, wParam, lParam
                 FUNCTION = 0 : EXIT FUNCTION
              END IF
           CASE %IDOK
              IF HIWRD(wParam) = %BN_CLICKED THEN
                 GetWindowText GetDlgItem(hWnd, 100), BYVAL VARPTR(szBuffer), SIZEOF(szBuffer)
                 nLine = VAL(szBuffer) - 1
                 SendMessage GetEdit, %SCI_GOTOLINE, nLine, 0
                 SendMessage hWnd, %WM_CLOSE, wParam, lParam
                 FUNCTION = 0 : EXIT FUNCTION
              END IF
           CASE 100   ' Go to line edit control
              IF HIWRD(wParam) = %EN_CHANGE THEN
                 GetWindowText GetDlgItem(hWnd, 100), BYVAL VARPTR(szBuffer), SIZEOF(szBuffer)
                 IF LEN(szBuffer) THEN
                    IF VAL(szBuffer) > SendMessage(GetEdit, %SCI_GETLINECOUNT, 0, 0) THEN
                       EnableWindow GetDlgItem(hWnd, %IDOK), %FALSE
                    ELSE
                       EnableWindow GetDlgItem(hWnd, %IDOK), %TRUE
                    END IF
                 ELSE
                    EnableWindow GetDlgItem(hWnd, %IDOK), %FALSE
                 END IF
              END IF
        END SELECT

     CASE %WM_CLOSE
        EnableWindow GetParent(hWnd), %TRUE    ' Maintains parent's zorder

     CASE %WM_DESTROY  ' Exit
        PostQuitMessage 0
        FUNCTION = 0 : EXIT FUNCTION

  END SELECT

  FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam)

END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' MainDialogTemplate popup dialog
' *********************************************************************************************
SUB MainDialogTemplate (BYVAL hWnd AS DWORD, BYVAL fGradient AS LONG)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
       s = "#COMPILE EXE" & $CRLF & _
       "#DIM ALL" & $CRLF & _
       "#DEBUG ERROR ON" & $CRLF & _
       "#INCLUDE ""WIN32API.INC""" & $CRLF & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "' Main" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "FUNCTION WINMAIN (BYVAL hInstance AS DWORD, BYVAL hPrevInstance AS DWORD, BYVAL lpszCmdLine AS ASCIIZ PTR, BYVAL nCmdShow AS LONG) AS LONG"  & $CRLF & $CRLF & _
       "   LOCAL hWndMain AS dword" & $CRLF & _
       "   LOCAL hCtl AS DWORD" & $CRLF & _
       "   LOCAL hFont AS DWORD" & $CRLF & _
       "   LOCAL wcex AS WndClassEx" & $CRLF & _
       "   LOCAL szClassName AS ASCIIZ * 80" & $CRLF & _
       "   LOCAL rc AS RECT" & $CRLF & _
       "   LOCAL szCaption AS asciiz * 255" & $CRLF & _
       "   LOCAL nLeft AS LONG" & $CRLF & _
       "   LOCAL nTop AS LONG" & $CRLF & _
       "   LOCAL nWidth AS LONG" & $CRLF & _
       "   LOCAL nHeight AS LONG"  & $CRLF & $CRLF & _
       "   hFont = GetStockObject(%ANSI_VAR_FONT)" & $CRLF & $CRLF & _
       "   ' Register the window class" & $CRLF & _
       "   szClassName        = ""MyClassName""" & $CRLF & _
       "   wcex.cbSize        = SIZEOF(wcex)" & $CRLF & _
       "   wcex.style         = %CS_HREDRAW OR %CS_VREDRAW" & $CRLF & _
       "   wcex.lpfnWndProc   = CODEPTR(WndProc)" & $CRLF & _
       "   wcex.cbClsExtra    = 0" & $CRLF & _
       "   wcex.cbWndExtra    = 0" & $CRLF & _
       "   wcex.hInstance     = hInstance" & $CRLF & _
       "   wcex.hCursor       = LoadCursor (%NULL, BYVAL %IDC_ARROW)" & $CRLF & _
       "   wcex.hbrBackground = %COLOR_3DFACE + 1" & $CRLF & _
       "   wcex.lpszMenuName  = %NULL" & $CRLF & _
       "   wcex.lpszClassName = VARPTR(szClassName)" & $CRLF & _
       "   wcex.hIcon         = LoadIcon (%NULL, BYVAL %IDI_APPLICATION) ' Sample, if resource icon: LoadIcon(hInst, ""APPICON"")" & $CRLF & _
       "   wcex.hIconSm       = LoadIcon (%NULL, BYVAL %IDI_APPLICATION) ' Remember to set small icon too.." & $CRLF & _
       "   RegisterClassEx wcex" & $CRLF & $CRLF & _
       "   ' Window caption" & $CRLF & _
       "   szCaption = ""SDK Main Window""" & $CRLF & $CRLF & _
       "   ' Retrieve the size of the working area" & $CRLF & _
       "   SystemParametersInfo %SPI_GETWORKAREA, 0, BYVAL VARPTR(rc), 0" & $CRLF & $CRLF & _
       "   ' Calculate the position and size of the window" & $CRLF & _
       "   nWidth  = (((rc.nRight - rc.nLeft)) + 2) * 0.75   ' 75% of the client screen width" & $CRLF & _
       "   nHeight = (((rc.nBottom - rc.nTop)) + 2) * 0.70   ' 70% of the client screen height" & $CRLF & _
       "   nLeft   = ((rc.nRight - rc.nLeft) \ 2) - nWidth \ 2" & $CRLF & _
       "   nTop    = ((rc.nBottom - rc.nTop) \ 2) - (nHeight \ 2)" & $CRLF & $CRLF & _
       "   ' Create a window using the registered class" & $CRLF & _
       "   hWndMain = CreateWindowEx(%WS_EX_CONTROLPARENT, _           ' extended style" & $CRLF & _
       "                             szClassName, _                    ' window class name" & $CRLF & _
       "                             szCaption, _                      ' window caption" & $CRLF & _
       "                             %WS_OVERLAPPEDWINDOW OR _" & $CRLF & _
       "                             %WS_CLIPCHILDREN, _               ' window style" & $CRLF & _
       "                             nLeft, _                          ' initial x position" & $CRLF & _
       "                             nTop, _                           ' initial y position" & $CRLF & _
       "                             nWidth, _                         ' initial x size" & $CRLF & _
       "                             nHeight, _                        ' initial y size" & $CRLF & _
       "                             %NULL, _                          ' parent window handle" & $CRLF & _
       "                             0, _                              ' window menu handle" & $CRLF & _
       "                             hInstance, _                      ' program instance handle" & $CRLF & _
       "                             BYVAL %NULL)                      ' creation parameters" & $CRLF & $CRLF & _
       "   hCtl = CreateWindowEx(0, ""BUTTON"", ""&Ok"", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT, _" & $CRLF & _
       "          0, 0, 0, 0, hWndMain, %IDOK, hInstance, BYVAL %NULL)" & $CRLF & _
       "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF & $CRLF & _
       "   hCtl = CreateWindowEx(0, ""BUTTON"", ""&Close"", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT, _" & $CRLF & _
       "          0, 0, 0, 0, hWndMain, %IDCANCEL, hInstance, BYVAL %NULL)" & $CRLF & _
       "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF & $CRLF & _
       "   ' Show the window" & $CRLF & _
       "   ShowWindow hWndMain, nCmdShow" & $CRLF & _
       "   UpdateWindow hWndMain" & $CRLF & $CRLF & _
       "   ' Message handler loop" & $CRLF & _
       "   LOCAL Msg AS tagMsg" & $CRLF & _
       "   WHILE GetMessage(Msg, %NULL, 0, 0)" & $CRLF & _
       "      IF ISFALSE IsDialogMessage(hWndMain, Msg) THEN" & $CRLF & _
       "         TranslateMessage Msg" & $CRLF & _
       "         DispatchMessage Msg" & $CRLF & _
       "      END IF" & $CRLF & _
       "   WEND" & $CRLF & $CRLF & _
       "   FUNCTION = msg.wParam" & $CRLF & $CRLF & _
       "END FUNCTION" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & $CRLF

       IF fGradient THEN
          s = s & _
          "' *********************************************************************************************" & $CRLF & _
          "' Gradient fill function" & $CRLF & _
          "' *********************************************************************************************" & $CRLF & _
          "SUB DrawGradient (BYVAL hDC AS DWORD)" & $CRLF & $CRLF & _
          "   LOCAL rectFill AS RECT" & $CRLF & _
          "   LOCAL rectClient AS RECT" & $CRLF & _
          "   LOCAL fStep AS SINGLE" & $CRLF & _
          "   LOCAL hBrush AS DWORD" & $CRLF & _
          "   LOCAL lOnBand AS LONG" & $CRLF & $CRLF & _
          "   GetClientRect WindowFromDC(hDC), rectClient" & $CRLF & _
          "   fStep = rectClient.nbottom / 200" & $CRLF & $CRLF & _
          "   FOR lOnBand = 0 TO 199" & $CRLF & _
          "      SetRect rectFill, 0, lOnBand * fStep, rectClient.nright + 1, (lOnBand + 1) * fStep" & $CRLF & _
          "      hBrush = CreateSolidBrush(RGB(0, 0, (255 - lOnBand)))" & $CRLF & _
          "      Fillrect hDC, rectFill, hBrush" & $CRLF & _
          "      DeleteObject hBrush" & $CRLF & _
          "   NEXT" & $CRLF & $CRLF & _
          "END SUB" & $CRLF & _
          "' *********************************************************************************************" & $CRLF & $CRLF
       END IF

       s = s & _
       "' *********************************************************************************************" & $CRLF & _
       "' Main Window procedure" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "FUNCTION WndProc (BYVAL hWnd AS DWORD, BYVAL wMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG" & $CRLF & $CRLF

       IF fGradient THEN
          s = s & _
          "   LOCAL hDC AS DWORD" & $CRLF & _
          "   LOCAL pPaint AS PAINTSTRUCT" & $CRLF
       END IF

       S = s & _
       "   LOCAL rc AS RECT" & $CRLF & $CRLF & _
       "   SELECT CASE wMsg" & $CRLF & $CRLF & _
       "      CASE %WM_CREATE" & $CRLF & _
       "         ' -------------------------------------------------------" & $CRLF & _
       "         ' A good place to initiate things, declare variables," & $CRLF & _
       "         ' create controls and read/set settings from a file, etc." & $CRLF & _
       "         ' -------------------------------------------------------" & $CRLF & $CRLF

       IF fGradient THEN
          s = s & _
          "      CASE %WM_PAINT" & $CRLF & _
          "         hDC = BeginPaint(hWnd, pPaint)" & $CRLF & _
          "         GetClientRect hWnd, rc" & $CRLF & _
          "         SetBkMode hDC, %TRANSPARENT" & $CRLF & _
          "         SetTextColor hDC, %WHITE" & $CRLF & _
          "         DrawText hDC, ""Hello, Windows!"", -1, rc, %DT_SINGLELINE OR %DT_CENTER OR %DT_VCENTER" & $CRLF & _
          "         EndPaint hWnd, pPaint" & $CRLF & _
          "         FUNCTION = %TRUE" & $CRLF & _
          "         EXIT FUNCTION" & $CRLF & $CRLF & _
          "      CASE %WM_ERASEBKGND" & $CRLF & _
          "         hDC = wParam" & $CRLF & _
          "         DrawGradient hDC              ' Pass the DC of the region to repaint" & $CRLF & _
          "         FUNCTION = %TRUE" & $CRLF & _
          "         EXIT FUNCTION" & $CRLF & $CRLF
       END IF

       s = s & _
       "      CASE %WM_SIZE" & $CRLF & _
       "         ' ----------------------------------------------------------------------" & $CRLF & _
       "         ' If you have a Toolbar and/or a Statusbar, send the following messages:" & $CRLF & _
       "         ' SendMessage hStatusbar, wMsg, wParam, lParam" & $CRLF & _
       "         ' SendMessage hToolbar, wMsg, wParam, lParam" & $CRLF & _
       "         ' ----------------------------------------------------------------------" & $CRLF & $CRLF & _
       "         ' Resize the two sample buttons of the dialog" & $CRLF & _
       "         IF wParam <> %SIZE_MINIMIZED THEN" & $CRLF & _
       "            GetClientRect hWnd, rc" & $CRLF & _
       "            MoveWindow GetDlgItem(hWNd, %IDOK), (rc.nRight - rc.nLeft) - 185, (rc.nBottom - rc.nTop) - 35, 75, 23, %TRUE" & $CRLF & _
       "            MoveWindow GetDlgItem(hWNd, %IDCANCEL), (rc.nRight - rc.nLeft) - 95, (rc.nBottom - rc.nTop) - 35, 75, 23, %TRUE" & $CRLF & _
       "         END IF" & $CRLF & $CRLF & _
       "      CASE %WM_COMMAND" & $CRLF & _
       "         ' -------------------------------------------------------" & $CRLF & _
       "         ' Messages from controls and menu items are handled here." & $CRLF & _
       "         ' -------------------------------------------------------" & $CRLF & _
       "         SELECT CASE LOWRD(wParam)" & $CRLF & $CRLF & _
       "            CASE %IDOK" & $CRLF & _
       "               IF HIWRD(wParam) = %BN_CLICKED THEN" & $CRLF & _
       "               END IF" & $CRLF & $CRLF & _
       "            CASE %IDCANCEL" & $CRLF & _
       "               IF HIWRD(wParam) = %BN_CLICKED THEN" & $CRLF & _
       "                  SendMessage hWnd, %WM_DESTROY, wParam, lParam" & $CRLF & _
       "                  FUNCTION = 0" & $CRLF & _
       "                  EXIT FUNCTION" & $CRLF & _
       "               END IF" & $CRLF & $CRLF & _
       "         END SELECT" & $CRLF & $CRLF & _
       "      CASE %WM_CTLCOLORBTN, %WM_CTLCOLOREDIT, %WM_CTLCOLORLISTBOX, %WM_CTLCOLORSTATIC" & $CRLF & _
       "         ' --------------------------------------------------------" & $CRLF & _
       "         ' wParam is the handle of the control's display context (hDC)" & $CRLF & _
       "         ' lParam is the handle of the control" & $CRLF & _
       "         ' Example of how to set the colors of an specific control:" & $CRLF & _
       "         ' --------------------------------------------------------" & $CRLF & _
       "         ' IF lParam = GetDlgItem(hWnd, CtlId) THEN" & $CRLF & _
       "         '    SetBkColor wParam, GetSysColor(%COLOR_INFOBK)" & $CRLF & _
       "         '    SetTextColor wParam, GetSysColor(%COLOR_INFOTEXT)" & $CRLF & _
       "         '    FUNCTION = GetSysColorBrush(%COLOR_INFOBK)" & $CRLF & _
       "         '    EXIT FUNCTION" & $CRLF & _
       "         ' END IF" & $CRLF & _
       "         ' --------------------------------------------------------" & $CRLF & $CRLF & _
       "      CASE %WM_SYSCOMMAND" & $CRLF & _
       "         ' Capture this message and send a WM_CLOSE message" & $CRLF & _
       "         IF (wParam AND &HFFF0) = %SC_CLOSE THEN" & $CRLF & _
       "            SendMessage hWnd, %WM_CLOSE, wParam, lParam" & $CRLF & _
       "            EXIT FUNCTION" & $CRLF & _
       "         END IF" & $CRLF & $CRLF & _
       "      CASE %WM_CLOSE" & $CRLF &  _
       "         ' --------------------------------------------------------" & $CRLF & _
       "         ' The WM_CLOSE message is processed BEFORE the WM_DESTROY" & $CRLF & _
       "         ' message and it can be used to confirm program exit or" & $CRLF & _
       "         ' tasks like deallocating memory or similar tasks." & $CRLF & _
       "         ' --------------------------------------------------------" & $CRLF & $CRLF & _
       "      CASE %WM_DESTROY" & $CRLF & _
       "         ' ---------------------------------------------------------------------------" & $CRLF & _
       "         ' Is sent when program ends - a good place to delete any created objects and" & $CRLF & _
       "         ' store settings in file for next run, etc. Must send PostQuitMessage to end" & $CRLF & _
       "         ' properly in SDK-style dialogs. The PostQuitMessage function sends a WM_QUIT" & $CRLF & _
       "         ' message to the program's (thread's) message queue, and then WM_QUIT causes" & $CRLF & _
       "         ' the GetMessage function to return zero in WINMAIN's message loop." & $CRLF & _
       "         ' ---------------------------------------------------------------------------" & $CRLF & _
       "         PostQuitMessage 0    ' This function closes the main window" & $CRLF & _
       "         FUNCTION = 0         ' by sending zero to the main message loop" & $CRLF & _
       "         EXIT FUNCTION" & $CRLF & $CRLF & _
       "   END SELECT" & $CRLF & $CRLF & _
       "   FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam)" & $CRLF & $CRLF & _
       "END FUNCTION" & $CRLF & _
       "' *********************************************************************************************" & $CRLF
   ELSE
       s = "#COMPILE EXE" & $CRLF & _
       "#DIM ALL" & $CRLF & _
       "#DEBUG ERROR ON" & $CRLF & _
       "#INCLUDE ""WIN32API.INC""" & $CRLF & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "' Main" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "FUNCTION WINMAIN (BYVAL hInstance AS DWORD, BYVAL hPrevInstance AS DWORD, BYVAL lpszCmdLine AS ASCIIZ PTR, BYVAL nCmdShow AS LONG) AS LONG" & $CRLF & $CRLF & _
       "   LOCAL hDlg AS LONG" & $CRLF & $CRLF & _
       "   DIALOG NEW 0, ""DDT MAIN DIALOG"", , , 400, 240, %WS_OVERLAPPED OR %WS_THICKFRAME OR %WS_SYSMENU OR _" & $CRLF & _
       "   %WS_MINIMIZEBOX OR %WS_MAXIMIZEBOX OR %WS_VISIBLE OR %DS_CENTER TO hDlg" & $CRLF & _
       "   ' For icon from resource, instead use something like, LoadIcon(hInst, ""APPICON"")" & $CRLF & _
       "   DIALOG SEND hDlg, %WM_SETICON, %ICON_SMALL, LoadIcon(%NULL, BYVAL %IDI_APPLICATION)" & $CRLF & _
       "   DIALOG SEND hDlg, %WM_SETICON, %ICON_BIG, LoadIcon(%NULL, BYVAL %IDI_APPLICATION)" & $CRLF & $CRLF & _
       "   CONTROL ADD BUTTON, hDlg, %IDOK, ""&Ok"", 0, 0, 50, 14, %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT" & $CRLF & _
       "   CONTROL ADD BUTTON, hDlg, %IDCANCEL, ""&Close"", 0, 0, 50, 14, %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT" & $CRLF & $CRLF & _
       "   DIALOG SHOW MODAL hDlg, CALL DlgProc" & $CRLF & $CRLF & _
       "END FUNCTION" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & $CRLF

       IF fGradient THEN
          s = s & _
          "' *********************************************************************************************" & $CRLF & _
          "' Gradient fill function" & $CRLF & _
          "' *********************************************************************************************" & $CRLF & _
          "SUB DrawGradient (BYVAL hDC AS DWORD)" & $CRLF & $CRLF & _
          "   LOCAL rectFill AS RECT" & $CRLF & _
          "   LOCAL rectClient AS RECT" & $CRLF & _
          "   LOCAL fStep AS SINGLE" & $CRLF & _
          "   LOCAL hBrush AS DWORD" & $CRLF & _
          "   LOCAL lOnBand AS LONG" & $CRLF & $CRLF & _
          "   GetClientRect WindowFromDC(hDC), rectClient" & $CRLF & _
          "   fStep = rectClient.nbottom / 200" & $CRLF & $CRLF & _
          "   FOR lOnBand = 0 TO 199" & $CRLF & _
          "      SetRect rectFill, 0, lOnBand * fStep, rectClient.nright + 1, (lOnBand + 1) * fStep" & $CRLF & _
          "      hBrush = CreateSolidBrush(RGB(0, 0, (255 - lOnBand)))" & $CRLF & _
          "      Fillrect hDC, rectFill, hBrush" & $CRLF & _
          "      DeleteObject hBrush" & $CRLF & _
          "   NEXT" & $CRLF & $CRLF & _
          "END SUB" & $CRLF & _
          "' *********************************************************************************************" & $CRLF & $CRLF
       END IF

       s = s & _
       "' *********************************************************************************************" & $CRLF & _
       "' Main Dialog procedure" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "CALLBACK FUNCTION DlgProc() AS LONG" & $CRLF & $CRLF

       IF fGradient THEN
          s = s & _
          "   LOCAL hDC    AS DWORD" & $CRLF & _
          "   LOCAL pPaint AS PAINTSTRUCT" & $CRLF
       END IF

       s = s & _
       "   LOCAL rc     AS RECT" & $CRLF & $CRLF & _
       "   SELECT CASE CBMSG" & $CRLF & $CRLF & _
       "      CASE %WM_INITDIALOG" & $CRLF & _
       "         ' --------------------------------------------------------------------" & $CRLF & _
       "         ' First message - a good place to initiate things, declare variables," & $CRLF & _
       "         ' create controls and read/set settings from a file, etc." & $CRLF & _
       "         ' --------------------------------------------------------------------" & $CRLF & $CRLF & _
       "      CASE %WM_SIZE" & $CRLF & _
       "         ' ----------------------------------------------------------------------" & $CRLF & _
       "         ' If you have a Toolbar and/or a Statusbar, send the following messages:" & $CRLF & _
       "         ' SendMessage hStatusbar, wMsg, wParam, lParam" & $CRLF & _
       "         ' SendMessage hToolbar, wMsg, wParam, lParam" & $CRLF & _
       "         ' ----------------------------------------------------------------------" & $CRLF & $CRLF & _
       "         ' Resize the two sample buttons of the dialog" & $CRLF & _
       "         IF CBWPARAM <> %SIZE_MINIMIZED THEN" & $CRLF & _
       "            GetClientRect CBHNDL, rc" & $CRLF & _
       "            MoveWindow GetDlgItem(CBHNDL, %IDOK), (rc.nRight - rc.nLeft) - 185, (rc.nBottom - rc.nTop) - 35, 75, 23, %TRUE" & $CRLF & _
       "            MoveWindow GetDlgItem(CBHNDL, %IDCANCEL), (rc.nRight - rc.nLeft) - 95, (rc.nBottom - rc.nTop) - 35, 75, 23, %TRUE" & $CRLF & _
       "            DIALOG REDRAW CBHNDL" & $CRLF & _
       "         END IF" & $CRLF & $CRLF

       IF fGradient THEN
          s = s & _
          "      CASE %WM_PAINT" & $CRLF & _
          "         hDC = BeginPaint(CBHNDL, pPaint)" & $CRLF & _
          "         GetClientRect CBHNDL, rc" & $CRLF & _
          "         SetBkMode hDC, %TRANSPARENT" & $CRLF & _
          "         SetTextColor hDC, %WHITE" & $CRLF & _
          "         DrawText hDC, ""Hello, Windows!"", -1, rc, %DT_SINGLELINE OR %DT_CENTER OR %DT_VCENTER" & $CRLF & _
          "         EndPaint CBHNDL, pPaint" & $CRLF & _
          "         FUNCTION = %TRUE" & $CRLF & _
          "         EXIT FUNCTION" & $CRLF & $CRLF & _
          "      CASE %WM_ERASEBKGND" & $CRLF & _
          "         hDC = CBWPARAM" & $CRLF & _
          "         DrawGradient hDC              ' Pass the DC of the region to repaint" & $CRLF & _
          "         FUNCTION = %TRUE" & $CRLF & _
          "         EXIT FUNCTION" & $CRLF & $CRLF
       END IF

       s = s & _
       "      CASE %WM_COMMAND" & $CRLF & _
       "         ' -------------------------------------------------------------------" & $CRLF & _
       "         ' Messages from controls and menu items are handled here." & $CRLF & _
       "         ' -------------------------------------------------------------------" & $CRLF & _
       "         SELECT CASE CBCTL" & $CRLF & _
       "            CASE %IDOK" & $CRLF & _
       "               IF CBCTLMSG = %BN_CLICKED THEN" & $CRLF & _
       "               END IF" & $CRLF & _
       "            CASE %IDCANCEL"  & $CRLF & _
       "               IF CBCTLMSG = %BN_CLICKED THEN DIALOG END CBHNDL, 0" & $CRLF & _
       "         END SELECT" & $CRLF & $CRLF & _
       "      CASE %WM_CTLCOLORBTN, %WM_CTLCOLOREDIT, %WM_CTLCOLORLISTBOX, %WM_CTLCOLORSTATIC"  & $CRLF & _
       "         ' -------------------------------------------------------------------" & $CRLF & _
       "         ' CBWPARAM is the handle of the control's display context (hDC)" & $CRLF & _
       "         ' CBLPARAM is the handle of the control" & $CRLF & _
       "         ' Example of how to set the colors of an specific control:" & $CRLF & _
       "         ' -------------------------------------------------------------------" & $CRLF & _
       "         ' IF CBLPARAM = GetDlgItem(CBHNDL, CtlId) THEN" & $CRLF & _
       "         '    SetBkColor CBWPARAM, GetSysColor(%COLOR_INFOBK)" & $CRLF & _
       "         '    SetTextColor CBWPARAM, GetSysColor(%COLOR_INFOTEXT)" & $CRLF & _
       "         '    FUNCTION = GetSysColorBrush(%COLOR_INFOBK)" & $CRLF & _
       "         ' END IF" & $CRLF & _
       "         ' -------------------------------------------------------------------" & $CRLF & $CRLF & _
       "      CASE %WM_DESTROY" & $CRLF & _
       "         ' -------------------------------------------------------------------" & $CRLF & _
       "         ' Is sent when program ends - a good place to delete created" & $CRLF & _
       "         ' objects and store settings in file for next run, etc." & $CRLF & _
       "         ' -------------------------------------------------------------------" & $CRLF & $CRLF & _
       "   END SELECT" & $CRLF & $CRLF & _
       "END FUNCTION" & $CRLF & _
       "' *********************************************************************************************" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Popup dialog template
' *********************************************************************************************
SUB PopupDialogTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB

   IF IsDlgButtonChecked(hWnd, 101) THEN
       s = "' *********************************************************************************************" & $CRLF & _
       "' Popup dialog - Calling example: ShowPopupDialog hWnd" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "FUNCTION ShowPopupDialog (BYVAL hParent AS LONG) AS LONG" & $CRLF & $CRLF & _
       "   LOCAL hWndPopup AS LONG" & $CRLF & _
       "   LOCAL hCtl AS LONG" & $CRLF & _
       "   LOCAL hFont AS LONG" & $CRLF & _
       "   LOCAL rc AS RECT" & $CRLF & _
       "   LOCAL wcex AS WndClassEx" & $CRLF & _
       "   LOCAL szClassName AS ASCIIZ * 80" & $CRLF & _
       "   LOCAL szCaption AS ASCIIZ * 255" & $CRLF & $CRLF & _
       "   hFont = GetStockObject(%ANSI_VAR_FONT)" & $CRLF & $CRLF & _
       "   szClassName        = ""MyPopupClassName""" & $CRLF & _
       "   wcex.cbSize        = SIZEOF(wcex)" & $CRLF & _
       "   wcex.style         = %CS_HREDRAW OR %CS_VREDRAW" & $CRLF & _
       "   wcex.lpfnWndProc   = CODEPTR(PopupDlgProc)" & $CRLF & _
       "   wcex.cbClsExtra    = 0" & $CRLF & _
       "   wcex.cbWndExtra    = 0" & $CRLF & _
       "   wcex.hInstance     = GetModuleHandle("""")" & $CRLF & _
       "   wcex.hCursor       = LoadCursor(%NULL, BYVAL %IDC_ARROW)" & $CRLF & _
       "   wcex.hbrBackground = %COLOR_3DFACE + 1" & $CRLF & _
       "   wcex.lpszMenuName  = %NULL" & $CRLF & _
       "   wcex.lpszClassName = VARPTR(szClassName)" & $CRLF & _
       "   wcex.hIcon         = 0" & $CRLF & _
       "   wcex.hIconSm       = 0" & $CRLF & _
       "   RegisterClassEx wcex" & $CRLF & $CRLF & _
       "   GetWindowRect hParent, rc          ' For centering child in parent" & $CRLF & _
       "   rc.nRight = rc.nRight - rc.nLeft   ' Parent's width" & $CRLF & _
       "   rc.nBottom = rc.nBottom - rc.nTop  ' Parent's height" & $CRLF & $CRLF & _
       "   szCaption = ""Popup dialog""" & $CRLF & _
       "   hWndPopup = CreateWindowEx(%WS_EX_DLGMODALFRAME OR %WS_EX_CONTROLPARENT, _" & $CRLF & _
       "               szClassName, szCaption, %WS_CAPTION OR %WS_POPUPWINDOW OR %WS_VISIBLE, _" & $CRLF & _
       "               rc.nLeft + (rc.nRight - 290) / 2, _" & $CRLF & _
       "               rc.nTop + (rc.nBottom - 180) / 2, _" & $CRLF & _
       "              290, 180, hParent, 0, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & $CRLF & _
       "   hCtl = CreateWindowEx(0, ""BUTTON"", ""&Close"", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT, _" & $CRLF & _
       "               200, 124, 75, 23, hWndPopup, %IDCANCEL, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
       "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF & $CRLF & _
       "   ShowWindow hWndPopup, %SW_SHOW" & $CRLF & _
       "   UpdateWindow hWndPopup" & $CRLF & _
       "   ' Message handler loop" & $CRLF & _
       "   LOCAL Msg AS tagMsg" & $CRLF & _
       "   WHILE GetMessage(Msg, %NULL, 0, 0)" & $CRLF & _
       "      IF ISFALSE IsDialogMessage(hWndPopup, Msg) THEN" & $CRLF & _
       "         TranslateMessage Msg" & $CRLF & _
       "         DispatchMessage Msg" & $CRLF & _
       "      END IF" & $CRLF & _
       "   WEND" & $CRLF & $CRLF & _
       "   FUNCTION = msg.wParam" & $CRLF & $CRLF & _
       "END FUNCTION" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "' Popup dialog procedure" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "FUNCTION PopupDlgProc (BYVAL hWnd AS DWORD, BYVAL wMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG" & $CRLF & $CRLF & _
       "   SELECT CASE wMsg" & $CRLF & _
       "      CASE %WM_CREATE" & $CRLF & _
       "         EnableWindow GetParent(hWnd), 0   ' To make the popup dialog modal" & $CRLF & $CRLF & _
       "      CASE %WM_COMMAND" & $CRLF & _
       "         SELECT CASE LOWRD(wParam)" & $CRLF & _
       "            CASE %IDCANCEL" & $CRLF & _
       "               IF HIWRD(wParam) = %BN_CLICKED THEN" & $CRLF & _
       "                  SendMessage hWnd, %WM_CLOSE, wParam, lParam" & $CRLF & _
       "                  FUNCTION = 0" & $CRLF & _
       "                  EXIT FUNCTION" & $CRLF & _
       "               END IF" & $CRLF & _
       "         END SELECT" & $CRLF & $CRLF & _
       "      CASE %WM_CLOSE" & $CRLF & _
       "         EnableWindow GetParent(hWnd), 1  ' Maintains parent's zorder" & $CRLF & $CRLF & _
       "      CASE %WM_DESTROY" & $CRLF & _
       "         PostQuitMessage 0                ' This function closes the main window" & $CRLF & _
       "         FUNCTION = 0                     ' by sending zero to the main message loop" & $CRLF & _
       "         EXIT FUNCTION" & $CRLF & $CRLF & _
       "   END SELECT" & $CRLF & $CRLF & _
       "   FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam)" & $CRLF & $CRLF & _
       "END FUNCTION" & $CRLF & _
       "' *********************************************************************************************" & $CRLF
   ELSE
       s = "' *********************************************************************************************" & $CRLF & _
       "' Popup (child) dialog template." & $CRLF & _
       "' Call from main program: ShowPopupDialog CBHNDL" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "FUNCTION ShowPopupDialog (BYVAL hParent AS DWORD) AS LONG" & $CRLF & $CRLF & _
       "   LOCAL hDlg2 AS LONG" & $CRLF & _
       "   DIALOG NEW hParent, ""Popup dialog"", , , 180, 90, %WS_POPUP OR %WS_CAPTION OR %WS_SYSMENU OR %DS_MODALFRAME, 0 TO hDlg2" & $CRLF & $CRLF & _
       "   CONTROL ADD BUTTON, hDlg2, %IDCANCEL, ""&Close"", 126, 72, 50, 14, %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT" & $CRLF & $CRLF & _
       "   DIALOG SHOW MODAL hDlg2, CALL PopupDlgProc" & $CRLF & $CRLF & _
       "END FUNCTION" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "' Popup Dialog procedure" & $CRLF & _
       "' *********************************************************************************************" & $CRLF & _
       "CALLBACK FUNCTION PopupDlgProc() AS LONG" & $CRLF & $CRLF & _
       "   SELECT CASE CBMSG" & $CRLF & $CRLF & _
       "      CASE %WM_INITDIALOG" & $CRLF & $CRLF & _
       "      CASE %WM_COMMAND" & $CRLF & $CRLF & _
       "         SELECT CASE CBCTL" & $CRLF & _
       "            CASE %IDCANCEL" & $CRLF & _
       "               IF CBCTLMSG = %BN_CLICKED THEN DIALOG END CBHNDL, 0" & $CRLF & _
       "         END SELECT" & $CRLF & $CRLF & _
       "      CASE %WM_DESTROY" & $CRLF & $CRLF & _
       "   END SELECT" & $CRLF & $CRLF & _
       "END FUNCTION" & $CRLF & _
       "' *********************************************************************************************" & $CRLF
   END IF
   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Button template
' *********************************************************************************************
SUB ButtonTemplate (BYVAL hWnd AS DWORD, BYVAL fFlat AS LONG)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      IF ISTRUE fFlat THEN
         s = "   hCtl = CreateWindowEx(0, ""BUTTON"", ""&Caption"", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT, _" & $CRLF & _
             "          5, 5, 75, 23, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
             "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
         ELSE
         s = "   hCtl = CreateWindowEx(0, ""BUTTON"", ""&Caption"", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _" & $CRLF & _
             "          5, 5, 75, 23, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
             "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
      END IF
   ELSE
      IF ISTRUE fFlat THEN
         s = "   CONTROL ADD BUTTON, hDlg, cID, ""&Caption"", 5, 5, 50, 14, %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT" & $CRLF
      ELSE
         s = "   CONTROL ADD BUTTON, hDlg, cID, ""&Caption"", 5, 5, 50, 14" & $CRLF
      END IF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Checkbox template
' *********************************************************************************************
SUB CheckboxTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""BUTTON"", ""&Caption "", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_AUTOCHECKBOX, _" & $CRLF & _
          "          5, 5, 90, 20, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD CHECKBOX, hDlg, cID, ""&Caption"", 5, 5, 60, 12" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' CheckState template
' *********************************************************************************************
SUB CheckStateTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""BUTTON"", ""&Caption "", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_AUTO3STATE, _" & $CRLF & _
          "          5, 5, 90, 20, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD CHECK3STATE, hDlg, cID, ""&Caption"", 5, 5, 60, 12" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Combobox Dropdown template
' *********************************************************************************************
SUB ComboboxDropdownTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""COMBOBOX"", """", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %CBS_DROPDOWN, _" & $CRLF & _
          "   5, 5, 180, 146, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD COMBOBOX, hDlg, cID, , 5, 5, 120, 90" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Combobox Dropdown template
' *********************************************************************************************
SUB ComboboxDropdownListTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""COMBOBOX"", """", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %CBS_DROPDOWNLIST, _" & $CRLF & _
          "          5, 5, 180, 146, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD COMBOBOX, hDlg, cID, , 5, 5, 120, 90, %WS_TABSTOP OR %CBS_DROPDOWNLIST" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Combobox template
' *********************************************************************************************
SUB ComboboxTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""COMBOBOX"", """", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %CBS_SIMPLE, _" & $CRLF & _
          "          5, 5, 180, 146, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD COMBOBOX, hDlg, cID, , 5, 5, 120, 90, %WS_TABSTOP OR %CBS_SIMPLE" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Frame template
' *********************************************************************************************
SUB FrameTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(%WS_EX_TRANSPARENT, ""BUTTON"", "" Caption "", %WS_CHILD OR %WS_VISIBLE OR %BS_GROUPBOX, _" & $CRLF & _
          "          10, 5, 180, 146, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD FRAME, hDlg, cID, "" Caption "", 9, 5, 120, 90" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Image template
' *********************************************************************************************
SUB ImageTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""STATIC"", """", %WS_CHILD OR %WS_VISIBLE OR %SS_BITMAP, _" & $CRLF & _
          "          5, 5, 75, 23, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   ' Note: replace %SS_BITMAP with %SS_ICON for icons" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD IMAGE, hDlg, cID, """", 5, 5, 50, 14" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************


' *********************************************************************************************
' ImageX template
' *********************************************************************************************
SUB ImageXTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""STATIC"", """", %WS_CHILD OR %WS_VISIBLE OR %SS_BITMAP, _" & $CRLF & _
          "          5, 5, 75, 23, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   ' Note: replace %SS_BITMAP with %SS_ICON for icons" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD IMAGEX, hDlg, cID, """", 5, 5, 50, 14" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' ImageButton template
' *********************************************************************************************
SUB ImageButtonTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""BUTTON"", """", %WS_CHILD OR %WS_VISIBLE OR %BS_BITMAP, _" & $CRLF & _
          "          5, 5, 75, 23, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   ' Note: replace %SS_BITMAP with %SS_ICON for icons" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD IMGBUTTON, hDlg, cID, """", 5, 5, 50, 14" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' ImageButtonX template
' *********************************************************************************************
SUB ImageButtonXTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""BUTTON"", """", %WS_CHILD OR %WS_VISIBLE OR %BS_BITMAP, _" & $CRLF & _
          "          5, 5, 75, 23, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   ' Note: replace %SS_BITMAP with %SS_ICON for icons" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD IMGBUTTONX, hDlg, cID, """", 5, 5, 50, 14" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Label template
' *********************************************************************************************
SUB LabelTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""STATIC"", ""Caption"", %WS_CHILD OR %WS_VISIBLE, _" & $CRLF & _
          "          5, 5, 90, 20, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD LABEL, hDlg, cID, ""Caption"", 5, 5, 60, 12" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Line template
' *********************************************************************************************
SUB LineTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""STATIC"", """", %WS_CHILD OR %WS_VISIBLE OR %SS_ETCHEDHORZ, _" & $CRLF & _
          "          5, 5, 180, 2, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD LINE, hDlg, cID, """", 5, 5, 120, 1" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Listbox template
' *********************************************************************************************
SUB ListboxTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(%WS_EX_CLIENTEDGE, ""LISTBOX"", """", %WS_CHILD OR %WS_VISIBLE, _" & $CRLF & _
          "          5, 5, 180, 146, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD LISTBOX, hDlg, cID, , 5, 5, 120, 90" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Radiobutton template
' *********************************************************************************************
SUB RadiobuttonTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""BUTTON"", ""&Caption"", %WS_CHILD OR %WS_VISIBLE OR %BS_AUTORADIOBUTTON, _" & $CRLF & _
          "          5, 5, 90, 20, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   ' Note: if several groups of Option controls are needed in a dialog," & $CRLF & _
          "   ' add %WS_TABSTOP OR %WS_GROUP style to first control in each group." & $CRLF & _
          "   ' and make sure at least one control in each group is checked via:" & $CRLF & _
          "   ' SendMessage hCtl, %BM_SETCHECK, %BST_CHECKED, 0" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD OPTION, hDlg, cID&, ""&Caption"", 5, 5, 60, 12" & $CRLF & _
          "   ' Note: if several groups of Option controls are needed in a dialog," & $CRLF & _
          "   ' add %WS_TABSTOP OR %WS_GROUP style to first control in each group." & $CRLF & _
          "   ' and make sure at least one control in each group is checked via:" & $CRLF & _
          "   ' ONTROL SET CHECK hDlg, cID, 1" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Horizontal scrollbar template
' *********************************************************************************************
SUB HorizScrollbarTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""SCROLLBAR"", """", %WS_CHILD OR %WS_VISIBLE OR %SBS_HORZ, _" & $CRLF & _
          "          5, 5, 360, 18, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD SCROLLBAR, hDlg, cID, """", 5, 5, 230, 11" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Vertical scrollbar template
' *********************************************************************************************
SUB VertScrollbarTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(0, ""SCROLLBAR"", """", %WS_CHILD OR %WS_VISIBLE OR %SBS_VERT, _" & $CRLF & _
          "          5, 5, 360, 18, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD SCROLLBAR, hDlg, cID, """", 5, 5, 12, 140, %SBS_VERT" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Textbox template
' *********************************************************************************************
SUB TextboxTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(%WS_EX_CLIENTEDGE, ""EDIT"", """", %WS_CHILD OR %WS_VISIBLE, _" & $CRLF & _
          "          5, 5, 90, 20, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD TEXTBOX, hDlg, cID, """", 5, 5, 60, 12" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Textbox template
' *********************************************************************************************
SUB MultilineTextboxTemplate (BYVAL hWnd AS DWORD)

   LOCAL s AS STRING

   IF ISFALSE GetEdit THEN EXIT SUB
   IF IsDlgButtonChecked(hWnd, 101) THEN
      s = "   hCtl = CreateWindowEx(%WS_EX_CLIENTEDGE, ""EDIT"", """", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR %ES_MULTILINE, _" & $CRLF & _
          "          5, 5, 180, 146, hWnd, cID, GetModuleHandle(""""), BYVAL %NULL)" & $CRLF & _
          "   IF hFont THEN SendMessage hCtl, %WM_SETFONT, hFont, 0" & $CRLF
   ELSE
      s = "   CONTROL ADD TEXTBOX, hDlg, cID, """", 5, 5, 120, 90, _" & $CRLF & _
          "      %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR %ES_MULTILINE, %WS_EX_CLIENTEDGE" & $CRLF
   END IF

   ' Insert the text at the current position
   SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(s)
   ' Change statusbar information
   ShowLinCol

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Templates
' *********************************************************************************************
SUB InsertTemplateCode (BYVAL hWnd AS DWORD, BYVAL curSel AS LONG)

   LOCAL curPos AS LONG

   IF ISFALSE GetEdit THEN EXIT SUB
   curPos = SendMEssage(GetEdit, %SCI_GETCURRENTPOS, 0, 0)

   SELECT CASE AS LONG curSel

      CASE  0  : MainDialogTemplate hWnd, %FALSE
      CASE  1  : MainDialogTemplate hWnd, %TRUE
      CASE  2  : PopupDialogTemplate hWnd
      CASE  3  : ButtonTemplate hWnd, %FALSE
      CASE  4  : ButtonTemplate hWnd, %TRUE
      CASE  5  : CheckboxTemplate hWNd
      CASE  6  : CheckStateTemplate hWNd
      CASE  7  : ComboboxDropdownTemplate hWnd
      CASE  8  : ComboboxDropdownListTemplate hWnd
      CASE  9  : ComboboxTemplate hWnd
      CASE 10  : FrameTemplate hWnd
      CASE 11  : ImageTemplate hWnd
      CASE 12  : ImageXTemplate hWnd
      CASE 13  : ImageButtonTemplate hWnd
      CASE 14  : ImageButtonXTemplate hWnd
      CASE 15  : LabelTemplate hWnd
      CASE 16  : LineTemplate hWNd
      CASE 17  : ListboxTemplate hWNd
      CASE 18  : RadioButtonTemplate hWnd
      CASE 19  : HorizScrollbarTemplate hWnd
      CASE 20  : VertScrollbarTemplate hWnd
      CASE 21  : TextboxTemplate hWnd
      CASE 22  : MultilineTextboxTemplate hWnd

   END SELECT

END SUB
' *********************************************************************************************

' *********************************************************************************************
' Template code popup dialog
' *********************************************************************************************
FUNCTION TemplateCodePopupDialog (BYVAL hParent AS LONG) AS LONG

   LOCAL hDlg AS LONG
   LOCAL hCtl AS LONG
   LOCAL rc AS RECT
   LOCAL wc AS WndClassEx
   LOCAL szClassName AS ASCIIZ * 80
   LOCAL hFnt AS DWORD

   hFnt = GetStockObject(%ANSI_VAR_FONT)

   szClassName      = "TemplateCode"
   wc.cbSize        = SIZEOF(wc)
   wc.style         = %CS_HREDRAW OR %CS_VREDRAW
   wc.lpfnWndProc   = CODEPTR(TemplateCodePopupDlgProc)
   wc.cbClsExtra    = 0
   wc.cbWndExtra    = 0
   wc.hInstance     = GetModuleHandle("")
   wc.hCursor       = LoadCursor(%NULL, BYVAL %IDC_ARROW)
   wc.hbrBackground = %COLOR_3DFACE + 1
   wc.lpszMenuName  = %NULL
   wc.lpszClassName = VARPTR(szClassName)
   wc.hIcon         = 0
   wc.hIconSm       = 0
   CALL RegisterClassEx (wc)

   GetWindowRect hParent, rc          ' for centering child in parent
   rc.nRight = rc.nRight - rc.nLeft   ' parent's width
   rc.nBottom = rc.nBottom - rc.nTop  ' parent's height

   hDlg = CreateWindowEx(%WS_EX_DLGMODALFRAME OR %WS_EX_CONTROLPARENT, szClassName, "Template code", _
          %WS_CAPTION OR %WS_POPUPWINDOW OR %WS_VISIBLE, rc.nLeft + (rc.nRight - 261) / 2, _
          rc.nTop + (rc.nBottom - 298) / 2, 261, 298, hParent, 0, GetModuleHandle(""), BYVAL %NULL)

   hCtl = CreateWindowEx(%WS_EX_CLIENTEDGE, "ListBox", BYVAL %NULL, %WS_CHILD OR %WS_VISIBLE OR %WS_BORDER OR %WS_VSCROLL OR %WS_TABSTOP OR %LBS_NOTIFY, _
          8, 8, 240, 214, hDlg, 100, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   ListBox_AddString hCtl, "MAIN DIALOG"
   Listbox_AddString hCtl, "MAIN DIALOG - with gradient"
   Listbox_AddString hCtl, "POPUP DIALOG"
   Listbox_AddString hCtl, "BUTTON"
   Listbox_AddString hCtl, "BUTTON - Flat style"
   Listbox_AddString hCtl, "CHECKBOX"
   Listbox_AddString hCtl, "CHECKSTATE"
   Listbox_AddString hCtl, "COMBOBOX - Dropdown"
   Listbox_AddString hCtl, "COMBOBOX - Dropdownlist"
   Listbox_AddString hCtl, "COMBOBOX - Simple"
   Listbox_AddString hCtl, "FRAME"
   Listbox_AddString hCtl, "IMAGE"
   Listbox_AddString hCtl, "IMAGEX"
   Listbox_AddString hCtl, "IMAGEBUTTON"
   Listbox_AddString hCtl, "IMAGEBUTTONX"
   Listbox_AddString hCtl, "LABEL"
   Listbox_AddString hCtl, "LINE"
   Listbox_AddString hCtl, "LISTBOX"
   Listbox_AddString hCtl, "OPTION (Radio button)"
   Listbox_AddString hCtl, "SCROLLBAR - Horizontal"
   Listbox_AddString hCtl, "SCROLLBAR - Vertical"
   Listbox_AddString hCtl, "TEXTBOX - Single line"
   Listbox_AddString hCtl, "TEXTBOX - Multiline"

   hCtl = CreateWindowEx(0, "Button", "SDK", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_AUTOCHECKBOX OR %BS_LEFT OR %BS_VCENTER, _
          9, 238, 60, 16, hDlg, 101, GetModuleHandle(""), BYVAL %NULL)
   IF ISTRUE VAL(IniRead(g_zIni, "Templates option", "SDKTemplates", "0"))  THEN CheckDlgButton hDlg, 101, %BST_CHECKED

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Button", "È·¶¨", %WS_CHILD OR %WS_VISIBLE OR _
          %WS_TABSTOP OR %BS_FLAT OR %BS_DEFPUSHBUTTON, 83, 234, 75, 23, hDlg, %IDOK, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   hCtl = CreateWindowEx(%WS_EX_NOPARENTNOTIFY, "Button", "È¡Ïû", %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_FLAT, _
          171, 234, 75, 23, hDlg, %IDCANCEL, GetModuleHandle(""), BYVAL %NULL)
   IF hFnt THEN SendMessage hCtl, %WM_SETFONT, hFnt, 0

   ShowWindow hDlg, %SW_SHOW
   UpdateWindow hDlg

   LOCAL Msg AS tagMsg
   WHILE GetMessage(Msg, %NULL, 0, 0)
      IF ISFALSE IsDialogMessage(hDlg, Msg) THEN
         TranslateMessage Msg
         DispatchMessage Msg
      END IF
   WEND

   FUNCTION = msg.wParam

END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' Popup dialog procedure
' *********************************************************************************************
FUNCTION TemplateCodePopupDlgProc (BYVAL hWnd AS DWORD, BYVAL wMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG

   LOCAL curSel AS LONG
   LOCAL SDKTemplates AS LONG

   SELECT CASE wMsg
      CASE %WM_CREATE   ' Entrance
         EnableWindow GetParent(hWnd), %FALSE   ' To make popup dialog modal

      CASE %WM_COMMAND
         SELECT CASE LOWRD(wParam)
            CASE %IDCANCEL
               IF HIWRD(wParam) = %BN_CLICKED THEN
                  SendMessage hWnd, %WM_CLOSE, wParam, lParam
                  FUNCTION = 0 : EXIT FUNCTION
               END IF
            CASE %IDOK
               IF HIWRD(wParam) = %BN_CLICKED THEN
                  ' Get the templates option
                  SDKTemplates = IsDlgButtonChecked(hWnd, 101)
                  'WriteTemplatesOption SDKTemplates
                  IniWrite g_zIni,"Templates option", "SDKTemplates", FORMAT$(SDKTemplates)
                  ' Get current selection
                  curSel = Listbox_GetCurSel(GetDlgItem(hWnd, 100))
                  InsertTemplateCode hWnd, curSel
                  SendMessage hWnd, %WM_CLOSE, wParam, lParam
                  FUNCTION = 0 : EXIT FUNCTION
               END IF
            CASE 100   ' Listbox control
               IF HIWRD(wParam) = %LBN_DBLCLK THEN
                  ' Get the templates option
                  SDKTemplates = IsDlgButtonChecked(hWnd, 101)
                  'WriteTemplatesOption SDKTemplates
                  IniWrite g_zIni,"Templates option", "SDKTemplates", FORMAT$(SDKTemplates)
                  ' Get current selection
                  curSel = Listbox_GetCurSel(GetDlgItem(hWnd, 100))
                  InsertTemplateCode hWnd, curSel
                  SendMessage hWnd, %WM_CLOSE, wParam, lParam
                  FUNCTION = 0 : EXIT FUNCTION
               END IF
         END SELECT

      CASE %WM_CLOSE
         EnableWindow GetParent(hWnd), %TRUE    ' Maintains parent's zorder

      CASE %WM_DESTROY  ' Exit
         PostQuitMessage 0
         FUNCTION = 0 : EXIT FUNCTION

   END SELECT

   FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam)

END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' Function    : SED_CreateCodeKeeperDatabase ()
' Description : Creates the Tsunami CodeKeepiations Database.
'               The key value is the CodeKeepiation name.
' *********************************************************************************************
FUNCTION SED_CreateCodeKeeperDatabase (BYVAL hWnd AS DWORD, BYVAL strDbPath AS STRING) AS LONG

   LOCAL FileDef AS STRING, Keys AS STRING
   LOCAL strSegment AS STRING * 25

   ON ERROR RESUME NEXT

   strSegment = "CodeKeeper Name"

   FileDef$ = MKBYT$(1) + _ ' Page Size
              MKBYT$(1) + _ ' Compress Record
              MKBYT$(1)     ' Number of Key Segments

   Keys = ""

   Keys = Keys + strSegment  + _   ' Description
                 MKBYT$(1)   + _   ' Segment number
                 MKWRD$(1)   + _   ' Segment Offset
                 MKBYT$(40)  + _   ' Segment Size
                 MKBYT$(2)         ' Flags (No duplicates)

   FileDef = FileDef + Keys

   ' Create Tsunami database file
   IF FileExist(strDbPath) THEN
      IF SED_MsgBox(hWnd, strDbPath & $CRLF & "File already exist. Overwrite it?", %MB_OK, "Warning!") = %IDYES THEN
         FUNCTION = trm_Create(strDbPath, FileDef, %TRUE)
      ELSE
         EXIT FUNCTION
      END IF
   ELSE
      FUNCTION = trm_Create(strDbPath, FileDef, %FALSE)
   END IF

END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
'   ** Constants **
' *********************************************************************************************
%IDC_CODEKEEPNAME   = 6001
%IDC_CODEKEEPTEXT   = 6002
%IDC_CODEKEEPSAVE   = 6003
%IDC_CODEKEEPDELETE = 6004
%IDC_CODEKEEPINSERT = 6005
%IDC_CODEKEEPEXIT   = 6006
%IDC_CODEKEEPHELP   = 6007
' *********************************************************************************************

' *********************************************************************************************
' CallBacks of the ShowCodeKeeper popup dialog
' *********************************************************************************************
CALLBACK FUNCTION ShowCodeKeeperDialogProc()

   STATIC hFocus AS DWORD
   LOCAL  strCodeKeepName AS STRING * 40
   LOCAL  buffer AS STRING
   LOCAL  Result AS LONG
   LOCAL  RetVal AS STRING
   LOCAL  szPath AS ASCIIZ * %MAX_PATH
   LOCAL  strText AS STRING

   SELECT CASE CBMSG

      CASE %WM_NCACTIVATE
         ' Save the handle of the control that has the focus
         IF ISFALSE CBWPARAM THEN hFocus = GetFocus

      CASE %WM_SETFOCUS
         ' Post a message to set the focus later, since some
         ' Window's actions can steal it if we set it here
         IF hFocus THEN
            PostMessage CBHNDL, %WM_USER + 999, hFocus, 0
            hFocus = 0
         END IF

      CASE %WM_USER + 999
         ' Set the focus and show the line an column in the status bar
         IF ISTRUE CBWPARAM THEN SetFocus CBWPARAM
         EXIT FUNCTION

      CASE %WM_USER + 1000
         ' Fill the combobox with the stored abbreviations
         IF ISTRUE trm_FileIsOpen(hCodeKeepFile) THEN
            MOUSEPTR 11
            RetVal = trm_GetFirst(hCodeKeepFile + %KEY_ONLY, 1)
            DO
               IF LEN(RetVal) = 0 THEN EXIT DO
               COMBOBOX ADD CBHNDL, %IDC_CODEKEEPNAME, RTRIM$(MID$(RetVal, 13))
               RetVal = trm_GetNext(hCodeKeepFile + %KEY_ONLY)
            LOOP
            MOUSEPTR 1
         END IF

      CASE %WM_COMMAND
         SELECT CASE CBCTL
            CASE %IDOK
               IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
                  CONTROL GET TEXT CBHNDL, %IDC_CODEKEEPNAME TO strCodeKeepName
                  IF TRIM$(strCodeKeepName) = "" THEN
                     SED_MsgBox(CBHNDL, "You must enter a name", %MB_OK, "Error!")
                     CONTROL SET FOCUS CBHNDL, %IDC_CODEKEEPNAME
                  ELSE
                     buffer = trm_GetEqualOrGreater(hCodeKeepFile, 1, BYCOPY strCodeKeepName)
                     IF UCASE$(strCodeKeepName) = UCASE$(LEFT$(buffer, 40)) THEN
                        CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPTEXT, MID$(buffer, 41)
                        buffer = ""
                     ELSE
                        CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPTEXT, ""
                     END IF
                     CONTROL SET FOCUS CBHNDL, %IDC_CODEKEEPTEXT
                  END IF
               END IF
            CASE %IDC_CODEKEEPSAVE
               IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
                  CONTROL GET TEXT CBHNDL, %IDC_CODEKEEPNAME TO strCodeKeepName
                  IF TRIM$(strCodeKeepName) = "" THEN
                     SED_MsgBox(CBHNDL, "You must enter a name", %MB_OK, "Error!")
                     CONTROL SET FOCUS CBHNDL, %IDC_CODEKEEPNAME
                  ELSE
                     CONTROL GET TEXT CBHNDL, %IDC_CODEKEEPTEXT TO buffer
                     Result = trm_Insert(hCodeKeepFile, strCodeKeepName & buffer)
                     IF Result = 5 THEN        ' Duplicate Key Value so use trm_Update.
                        Result = trm_Update(hCodeKeepFile, strCodeKeepName & buffer)
                        IF Result THEN SED_MsgBox(CBHNDL, "Error " & FORMAT$(Result) & " updating the file", %MB_OK, "Error!")
                     ELSE
                        COMBOBOX ADD CBHNDL, %IDC_CODEKEEPNAME, RTRIM$(strCodeKeepName)
                     END IF
                     CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPNAME, ""
                     CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPTEXT, ""
                     CONTROL SET FOCUS CBHNDL, %IDC_CODEKEEPNAME
                  END IF
               END IF
            CASE %IDC_CODEKEEPDELETE
               IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
                  CONTROL GET TEXT CBHNDL, %IDC_CODEKEEPNAME TO strCodeKeepName
                  buffer = trm_GetEqualOrGreater(hCodeKeepFile, 1, BYCOPY strCodeKeepName)
                  IF UCASE$(strCodeKeepName) = UCASE$(LEFT$(buffer, 40)) THEN
                     CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPTEXT, MID$(buffer, 41)
                     buffer = ""
                     IF SED_MsgBox(CBHNDL, "Are you sure you want to delete it?", %MB_YESNO, "Delete record") = %IDYES THEN
                        Result = trm_Delete(hCodeKeepFile)
                        IF Result THEN
                           SED_MsgBox(CBHNDL, "Error " & FORMAT$(Result) & " updating the file", %MB_OK, "Error!")
                        ELSE
                           'Result = Combo_FindStringExact(GetDlgItem(CBHNDL, %IDC_CODEKEEPNAME), 0, RTRIM$(strCodeKeepName))
                           COMBOBOX FIND EXACT CB.HNDL,%IDC_CODEKEEPNAME,1,RTRIM$(strCodeKeepName) TO Result
                           IF Result>0 THEN
                             'Result = Combo_DeleteString(GetDlgItem(CBHNDL, %IDC_CODEKEEPNAME), Result)
                             COMBOBOX DELETE CB.HNDL,%IDC_CODEKEEPNAME,Result
                           END IF
                           'IF Result = 0 THEN Combo_ResetContent(GetDlgItem(CBHNDL, %IDC_CODEKEEPNAME))
                           COMBOBOX RESET CB.HNDL,5=%IDC_CODEKEEPNAME
                        END IF
                        CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPNAME, ""
                        CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPTEXT, ""
                        PostMessage CBHNDL, %WM_USER + 999, GetDlgItem(CBHNDL, %IDC_CODEKEEPNAME), 0
                     END IF
                  ELSE
                     CONTROL SET FOCUS CBHNDL, %IDC_CODEKEEPNAME
                  END IF
               END IF
            CASE %IDC_CODEKEEPINSERT
               ' Insert the text at the current position
               IF GetEdit THEN
                  CONTROL GET TEXT CBHNDL, %IDC_CODEKEEPNAME TO strCodeKeepName
                  buffer = trm_GetEqualOrGreater(hCodeKeepFile, 1, BYCOPY strCodeKeepName)
                  IF UCASE$(strCodeKeepName) = UCASE$(LEFT$(buffer, 40)) THEN
                     buffer = MID$(buffer, 41)
                     SendMessage GetEdit, %SCI_INSERTTEXT, -1, BYVAL STRPTR(Buffer)
                     DIALOG END CBHNDL
                  END IF
               END IF
            CASE %IDCANCEL
               CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPNAME, ""
               CONTROL SET TEXT CBHNDL, %IDC_CODEKEEPTEXT, ""
               CONTROL SET FOCUS CBHNDL, %IDC_CODEKEEPNAME
            CASE %IDC_CODEKEEPEXIT
               IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
                  DIALOG END CBHNDL
               END IF
            CASE %IDC_CODEKEEPHELP
               IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
                  szPath = EXE.PATH$ & "SED.HLP"
                  strText = "Code Keeper"
                  IF FileExist(szPath) THEN WinHelp CBHNDL, szPath, %HELP_KEY, BYVAL STRPTR(strText)
               END IF

         END SELECT
   END SELECT

END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' Code Keeper popup dialog
' *********************************************************************************************
FUNCTION ShowCodeKeeperDialog (BYVAL hParent AS DWORD) AS LONG

   LOCAL lRslt AS LONG
   LOCAL hDlg AS DWORD
   LOCAL hModRE AS DWORD
   LOCAL strDbPath AS STRING
   LOCAL Result AS LONG
   LOCAL hFnt AS DWORD

   hFnt = GetStockObject(%ANSI_FIXED_FONT)

   ' Load the RichEdit library
   hModRE = LoadLibrary("RICHED32.DLL")
   IF ISFALSE hModRE THEN
      SED_MsgBox(hParent, "Unable to load RICHED32.DLL", %MB_OK, "Error!")
      EXIT FUNCTION
   END IF

   ' Create the databse file if it doesn't exist
   strDbPath = EXE.PATH$ & $SCI_CODEKEEPDB
   IF ISFALSE FileExist(strDbPath) THEN
      Result = SED_CreateCodeKeeperDatabase (hParent, strDbPath)
      IF Result THEN
         SED_MsgBox(hParent, strDbPath & $CRLF & "Error " & FORMAT$(Result) & " creating the file", %MB_OK, "Tsunami Error!")
         EXIT FUNCTION
      END IF
   END IF

   ' Open the file if it is closed
   IF ISFALSE trm_FileIsOpen(hCodeKeepFile) THEN
      hCodeKeepFile = trm_Open(strDbPath, 1)
      IF hCodeKeepFile < 1 THEN
         SED_MsgBox(hParent, strDbPath & $CRLF & "Error " & FORMAT$(hCodeKeepFile) & " opening the file", %MB_OK, "Tsunami Error!")
         EXIT FUNCTION
      END IF
   END IF

   DIALOG NEW hParent, "Code Keeper", , , 334, 181, %WS_POPUP OR _
      %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR %WS_CLIPSIBLINGS OR _
      %WS_VISIBLE OR %DS_MODALFRAME OR %DS_CENTER OR %DS_3DLOOK OR _
      %DS_NOFAILCREATE OR %DS_SETFONT, %WS_EX_WINDOWEDGE OR _
      %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
      %WS_EX_RIGHTSCROLLBAR, TO hDlg
   CONTROL ADD LABEL, hDlg, -1, " Name", 5, 5, 66, 10
   CONTROL ADD COMBOBOX, hDlg, %IDC_CODEKEEPNAME, , 5, 16, 323, 144, %WS_CHILD OR _
      %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR %CBS_DROPDOWN OR %CBS_SORT, _
      %WS_EX_LEFT OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR
   CONTROL ADD LABEL, hDlg, -1, " Text", 5, 31, 64, 10
   CONTROL ADD "RichEdit", hDlg, %IDC_CODEKEEPTEXT, "", 5, 41, 323, 113, %WS_CHILD _
      OR %WS_VISIBLE OR %WS_BORDER OR %WS_TABSTOP OR %WS_HSCROLL OR %WS_VSCROLL _
      OR %ES_LEFT OR %ES_MULTILINE OR %ES_AUTOHSCROLL OR %ES_AUTOVSCROLL OR _
      %ES_WANTRETURN, %WS_EX_LEFT OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR
   CONTROL ADD BUTTON, hDlg, %IDOK, "&Ok", 6, 160, 50, 15, %WS_CHILD OR _
      %WS_VISIBLE OR %WS_TABSTOP OR %BS_TEXT OR %BS_DEFPUSHBUTTON OR %BS_FLAT OR _
      %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT OR %WS_EX_LTRREADING
   CONTROL ADD BUTTON, hDlg, %IDC_CODEKEEPSAVE, "&Save", 60, 160, 50, 15, %WS_CHILD OR _
      %WS_VISIBLE OR %WS_TABSTOP OR %BS_TEXT OR %BS_PUSHBUTTON OR %BS_FLAT OR _
      %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT OR %WS_EX_LTRREADING
   CONTROL ADD BUTTON, hDlg, %IDC_CODEKEEPDELETE, "&Delete", 114, 160, 50, 15, %WS_CHILD _
      OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_TEXT OR %BS_PUSHBUTTON OR %BS_FLAT _
      OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT OR %WS_EX_LTRREADING
   CONTROL ADD BUTTON, hDlg, %IDC_CODEKEEPINSERT, "&Insert", 170, 160, 50, 15, _
      %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_TEXT OR %BS_PUSHBUTTON OR _
      %BS_FLAT OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT OR %WS_EX_LTRREADING
   CONTROL ADD BUTTON, hDlg, %IDC_CODEKEEPEXIT, "&Exit", 224, 160, 50, 15, _
      %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_TEXT OR %BS_PUSHBUTTON OR _
      %BS_FLAT OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT OR %WS_EX_LTRREADING
   CONTROL ADD BUTTON, hDlg, %IDC_CODEKEEPHELP, "&Help", 278, 160, 50, 15, _
      %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_TEXT OR %BS_PUSHBUTTON OR _
      %BS_FLAT OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT OR %WS_EX_LTRREADING

   IF hFnt THEN SendMessage GetDlgItem(hDlg, %IDC_CODEKEEPTEXT), %WM_SETFONT, hFnt, 0

   PostMessage hDlg, %WM_USER + 1000, 0, 0

   DIALOG SHOW MODAL hDlg, CALL ShowCodeKeeperDialogProc TO lRslt

   FreeLibrary hModRE

   FUNCTION = lRslt

END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
' Show MSInfo32 - first look in Registry for path, then SHELL..
' *********************************************************************************************
SUB SED_ShowSysInfo (BYVAL hWnd AS DWORD)

   LOCAL lRes AS LONG
   LOCAL hKey AS DWORD
   LOCAL szBuff AS ASCIIZ * 1024
   LOCAL strPath AS STRING

   IF RegOpenKeyEx(%HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Shared Tools\MSInfo", 0, _
                   %KEY_QUERY_VALUE, hKey) = %ERROR_SUCCESS THEN
      IF hKey THEN
         RegQueryValueEx hKey, "Path", 0, 0, szBuff, SIZEOF(szBuff)
         RegCloseKey hKey
         strPath = EXTRACT$(szBuff, CHR$(0))
      END IF
   END IF

   IF LEN(strPath) THEN
      lRes = SHELL(strPath)
   ELSE
      MSGBOX "Can't find MSInfo32.exe"
      SED_MsgBox hWnd, "Can't find MSInfo32.exe", %MB_OK, "Error"
   END IF

END SUB
' *********************************************************************************************
'   ** CallBack **
' *********************************************************************************************
CALLBACK FUNCTION ShowPrinterSetupDialogProc ()

   LOCAL PrnOpt AS PrinterSetupOptionsType
   LOCAL lFont AS LONG
   LOCAL lOrientation AS LONG
   LOCAL rs AS STRING

   SELECT CASE CBMSG
      CASE %WM_COMMAND
         SELECT CASE CBCTL
            CASE %IDOK
               IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
                  PrnOpt.PaperSize = Combo_GetCurSel(GetDlgItem(CBHNDL, %IDCPS_CBPAPERSIZE)) + 1
                  PrnOpt.PaperBin = Combo_GetCurSel(GetDlgItem(CBHNDL, %IDCPS_CBPAPERBIN)) + 1
                  lFont = Combo_GetCursel(GetDlgItem(CBHNDL, %IDCPS_CBFONTNAME)) + 1
                  SELECT CASE lFont
                     CASE  1 : PrnOpt.FontName = "Courier"
                     CASE  2 : PrnOpt.FontName = "Courier New"
                     CASE  3 : PrnOpt.FontName = "Courier CE"
                     CASE  4 : PrnOpt.FontName = "Courier CYR"
                     CASE  5 : PrnOpt.FontName = "Courier Greek"
                     CASE  6 : PrnOpt.FontName = "Courier TUR"
                     CASE  7 : PrnOpt.FontName = "Fixedsys"
                     CASE  8 : PrnOpt.FontName = "Lucida Console"
                     CASE  9 : PrnOpt.FontName = "Terminal"
                     CASE 10 : PrnOpt.FontName = "Andale Mono"
                     CASE 11 : PrnOpt.FontName = "FixedsysTTF"
                  END SELECT
                  PrnOpt.FontSize = Combo_GetCursel(GetDlgItem(CBHNDL, %IDCPS_CBFONTSIZE)) + 1
                  CONTROL GET CHECK CBHNDL, %IDCPS_LANDSCAPE TO lOrientation
                  IF lOrientation = 1 THEN PrnOpt.Orientation = 1 ELSE PrnOpt.Orientation = 0
                  CONTROL GET CHECK CBHNDL, %IDCPS_CKFONTBOLD TO PrnOpt.FontBold
                  CONTROL GET CHECK CBHNDL, %IDCPS_CKFONTITALIC TO PrnOpt.FontItalic
                  CONTROL GET CHECK CBHNDL, %IDCPS_CKFONTUNDERLINE TO PrnOpt.FontUnderline
                  CONTROL GET TEXT CBHNDL, %IDCPS_MARGINLEFT TO rs
                  PrnOpt.MarginLeft = VAL(rs)
                  IF PrnOpt.MarginLeft < 0 THEN PrnOpt.MarginLeft = 0
                  CONTROL GET TEXT CBHNDL, %IDCPS_MARGINRIGHT TO rs
                  PrnOpt.MarginRight = VAL(rs)
                  IF PrnOpt.MarginRight < 0 THEN PrnOpt.MarginRight = 0
                  CONTROL GET TEXT CBHNDL, %IDCPS_MARGINTOP TO rs
                  PrnOpt.MarginTop = VAL(rs)
                  IF PrnOpt.MarginTop < 0 THEN PrnOpt.MarginTop = 0
                  CONTROL GET TEXT CBHNDL, %IDCPS_MARGINBOTTOM TO rs
                  PrnOpt.MarginBottom = VAL(rs)
                  IF PrnOpt.MarginBottom < 0 THEN PrnOpt.MarginBottom = 0
                  CONTROL GET TEXT CBHNDL, %IDCPS_LINEHEIGHT TO rs
                  PrnOpt.LineHeight = VAL(rs)
                  IF PrnOpt.LineHeight <= 0 THEN PrnOpt.LineHeight = 0.17!
                  WritePrinterSetupOptions PrnOpt
                  DIALOG END CBHNDL
               END IF
            CASE %IDCANCEL
               IF CBCTLMSG = %BN_CLICKED OR CBCTLMSG = 1 THEN
                  DIALOG END CBHNDL
               END IF

         END SELECT
   END SELECT

END FUNCTION
' *********************************************************************************************

' *********************************************************************************************
'   ** Dialog **
' *********************************************************************************************
FUNCTION ShowPrinterSetupDialog (BYVAL hParent AS DWORD) AS LONG

   LOCAL lRslt AS LONG
   LOCAL hDlg AS DWORD
   LOCAL PrnOpt AS PrinterSetupOptionsType
   LOCAL lFont AS LONG

   DIALOG NEW hParent, "´òÓ¡ÉèÖÃ", , , 270, 208, %WS_POPUP OR %WS_BORDER _
      OR %WS_DLGFRAME OR %WS_SYSMENU OR %WS_CLIPSIBLINGS OR %WS_VISIBLE OR _
      %DS_MODALFRAME OR %DS_CENTER OR %DS_3DLOOK OR %DS_NOFAILCREATE OR _
      %DS_SETFONT, %WS_EX_WINDOWEDGE OR %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR _
      %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR, TO hDlg
   CONTROL ADD FRAME, hDlg, -1, " Ö½ÕÅ ", 5, 4, 156, 71
   CONTROL ADD LABEL, hDlg, -1, " ´óÐ¡", 10, 15, 100, 10
   CONTROL ADD COMBOBOX, hDlg, %IDCPS_CBPAPERSIZE, , 12, 25, 143, 100, %WS_CHILD OR _
      %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR %CBS_DROPDOWNLIST, %WS_EX_LEFT _
      OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR
   CONTROL ADD LABEL, hDlg, -2, " Bin", 10, 40, 100, 10
   CONTROL ADD COMBOBOX, hDlg, %IDCPS_CBPAPERBIN, , 12, 50, 143, 100, %WS_CHILD OR _
      %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR %CBS_DROPDOWNLIST, %WS_EX_LEFT _
      OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR
   CONTROL ADD OPTION, hDlg, %IDCPS_PORTRAIT, "&Portrait", 170, 22, 90, 10
   CONTROL ADD FRAME, hDlg, -1, " ·½Ïò ", 165, 4, 100, 71
   CONTROL ADD OPTION, hDlg, %IDCPS_LANDSCAPE, "&Landscape", 170, 41, 90, 10
   CONTROL ADD FRAME, hDlg, -1, " ×ÖÌå ", 5, 76, 260, 48
   CONTROL ADD COMBOBOX, hDlg, %IDCPS_CBFONTNAME, , 12, 89, 177, 100, %WS_CHILD OR _
      %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR %CBS_DROPDOWNLIST, %WS_EX_LEFT _
      OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR
   CONTROL ADD LABEL, hDlg, -1, "´óÐ¡", 193, 89, 25, 10, %WS_CHILD OR _
      %WS_VISIBLE OR %SS_LEFT, %WS_EX_RIGHT OR %WS_EX_LTRREADING
   CONTROL ADD COMBOBOX, hDlg, %IDCPS_CBFONTSIZE, , 224, 89, 35, 100, %WS_CHILD OR _
      %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR %CBS_DROPDOWNLIST, %WS_EX_LEFT _
      OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR
   CONTROL ADD CHECKBOX, hDlg, %IDCPS_CKFONTBOLD, "¼Ó´Ö", 12, 110, 50, 10
   CONTROL ADD CHECKBOX, hDlg, %IDCPS_CKFONTITALIC, "Ð±Ìå", 64, 110, 50, 10
   CONTROL ADD CHECKBOX, hDlg, %IDCPS_CKFONTUNDERLINE, "ÏÂ»®Ïß", 116, 110, 50, 10

   CONTROL ADD FRAME, hDlg, -1, " Margins [in inches] ", 5, 126, 160, 54
   CONTROL ADD LABEL, hDlg, -1, "×ó", 12, 142, 20, 10
   CONTROL ADD LABEL, hDlg, -1, "ÓÒ", 12, 162, 20, 10
   CONTROL ADD LABEL, hDlg, -1, "ÉÏ", 90, 142, 20, 10
   CONTROL ADD LABEL, hDlg, -1, "ÏÂ", 90, 162, 20, 10
   CONTROL ADD TEXTBOX, hDlg, %IDCPS_MARGINLEFT, "", 40, 140, 30, 13
   CONTROL ADD TEXTBOX, hDlg, %IDCPS_MARGINRIGHT, "", 40, 160, 30, 13
   CONTROL ADD TEXTBOX, hDlg, %IDCPS_MARGINTOP, "", 128, 140, 30, 13
   CONTROL ADD TEXTBOX, hDlg, %IDCPS_MARGINBOTTOM, "", 128, 160, 30, 13

   CONTROL ADD FRAME, hDlg, -1, " Line spacing ", 170, 126, 95, 54, %BS_TOP OR %BS_CENTER
   CONTROL ADD TEXTBOX, hDlg, %IDCPS_LINEHEIGHT, "", 202, 150, 30, 13

   CONTROL ADD BUTTON, hDlg, %IDOK, "È·¶¨", 153, 186, 50, 15, %WS_CHILD OR _
      %WS_VISIBLE OR %WS_TABSTOP OR %BS_TEXT OR %BS_DEFPUSHBUTTON OR %BS_FLAT OR _
      %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT OR %WS_EX_LTRREADING
   CONTROL ADD BUTTON, hDlg, %IDCANCEL, "ÍË³ö", 214, 186, 50, 15, %WS_CHILD _
      OR %WS_VISIBLE OR %WS_TABSTOP OR %BS_TEXT OR %BS_FLAT _
      OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT OR %WS_EX_LTRREADING

   COMBOBOX ADD hDlg, %IDCPS_CBPAPERSIZE, "Letter 8 ½ x 11"""
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERSIZE, "A4 210 x 297 mm (8.27x11.69)"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERSIZE, "Legal 8 ½ x 14"""
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERSIZE, "Folio 8 ½ x 13"""

   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Upper"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Lower"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Middle"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Manual"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Envelope"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Enum manual"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Auto"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Tractor"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Small font"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Large font"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Large capacity"
   COMBOBOX ADD hDlg, %IDCPS_CBPAPERBIN, "Cassette"

   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Courier"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Courier New"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Courier CE"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Courier CYR"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Courier Greek"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Courier TUR"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Fixedsys"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Lucida Console"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Terminal"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "Andale Mono"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTNAME, "FixedsysTTF"

   COMBOBOX ADD hDlg, %IDCPS_CBFONTSIZE, "8"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTSIZE, "9"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTSIZE, "10"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTSIZE, "11"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTSIZE, "12"
   COMBOBOX ADD hDlg, %IDCPS_CBFONTSIZE, "14"

   GetPrinterSetupOptions PrnOpt
   COMBOBOX SELECT hDlg, %IDCPS_CBPAPERSIZE, PrnOpt.PaperSize
   COMBOBOX SELECT hDlg, %IDCPS_CBPAPERBIN, PrnOpt.PaperBin
   SELECT CASE PrnOpt.FontName
      CASE "Courier"         : lFont = 1
      CASE "Courier New"     : lFont = 2
      CASE "Courier CE"      : lFont = 3
      CASE "Courier CYR"     : lFont = 4
      CASE "Courier Greek"   : lFont = 5
      CASE "Courier TUR"     : lFont = 6
      CASE "Fixedsys"        : lFont = 7
      CASE "Lucida Console"  : lFont = 8
      CASE "Terminal"        : lFont = 9
      CASE "Andale Mono"     : lFont = 10
      CASE "FixedsysTTF"     : lFont = 11
   END SELECT
   COMBOBOX SELECT hDlg, %IDCPS_CBFONTNAME, lFont
   COMBOBOX SELECT hDlg, %IDCPS_CBFONTSIZE, PrnOpt.FontSize
   IF PrnOpt.Orientation = 0 THEN
      CONTROL SET OPTION hDlg, %IDCPS_PORTRAIT, %IDCPS_PORTRAIT, %IDCPS_LANDSCAPE
   ELSE
      CONTROL SET OPTION hDlg, %IDCPS_LANDSCAPE, %IDCPS_PORTRAIT, %IDCPS_LANDSCAPE
   END IF

   CONTROL SET CHECK hDlg, %IDCPS_CKFONTBOLD, PrnOpt.FontBold
   CONTROL SET CHECK hDlg, %IDCPS_CKFONTITALIC, PrnOpt.FontItalic
   CONTROL SET CHECK hDlg, %IDCPS_CKFONTUNDERLINE, PrnOpt.FontUnderline

   CONTROL SET TEXT hDlg, %IDCPS_MARGINLEFT, USING$("#.###", PrnOpt.MarginLeft)
   CONTROL SET TEXT hDlg, %IDCPS_MARGINRIGHT, USING$("#.###", PrnOpt.MarginRight)
   CONTROL SET TEXT hDlg, %IDCPS_MARGINTOP, USING$("#.###", PrnOpt.MarginTop)
   CONTROL SET TEXT hDlg, %IDCPS_MARGINBOTTOM, USING$("#.###", PrnOpt.MarginBottom)
   CONTROL SET TEXT hDlg, %IDCPS_LINEHEIGHT, USING$("#.###", PrnOpt.LineHeight)

   DIALOG SHOW MODAL hDlg, CALL ShowPrinterSetupDialogProc TO lRslt

   FUNCTION = lRslt

END FUNCTION
' *********************************************************************************************
