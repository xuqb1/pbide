TYPE PREVIEWINFO
  hDCTo      AS DWORD      ' DC to preview into (bitmap)
  hBmp       AS DWORD      ' Handle to bitmap
  hChild     AS DWORD      ' Handle to preview child window
  hThread    AS DWORD      ' Handle to preview thread
  hCtl       AS DWORD      ' Handle to editor we're prviewing
  lFlag      AS LONG       ' Keeps a preview thread alive
  nPage      AS LONG       ' Current page #
  MaxPage    AS LONG       ' Last page #
  ScaleX     AS LONG       ' Client scaling width
  ScaleY     AS LONG       ' Client scaling height
  lWidth     AS LONG       ' Client width
  lHeight    AS LONG       ' Client height
  sci        AS SCROLLINFO
END TYPE

TYPE PAGENUMINFO
  fr                  AS RANGETOFORMAT
  hEdit               AS LONG
  HeaderLineHeight    AS LONG
  hFontHeader         AS LONG
  PrintFlag           AS LONG
  SelPage             AS LONG
  ms                  AS DWORD      ' Page # calculation run time
  pArray              AS DWORD
END TYPE
%USE_THREAD = 1
GLOBAL g_hModeless     AS LONG
GLOBAL g_AbortPrint    AS LONG
GLOBAL g_psDlg         AS PAGESETUPDLGA
GLOBAL g_hHook         AS LONG
GLOBAL PrnDC           AS DWORD      ' Current printer DC

'SUB CenterWindow (BYVAL hWnd AS DWORD)
'   LOCAL WndRect AS RECT
'   LOCAL x AS LONG, y AS LONG
'
'   GetWindowRect hWnd, WndRect
'
'   x = (GetSystemMetrics(%SM_CXSCREEN) - (WndRect.nRight-WndRect.nLeft)) \ 2
'   y = (GetSystemMetrics(%SM_CYSCREEN) - (WndRect.nBottom - WndRect.nTop + GetSystemMetrics(%SM_CYCAPTION))) \ 2
'
'   SetWindowPos hWnd, %NULL, x, y, 0, 0, %SWP_NOSIZE OR %SWP_NOZORDER
'END SUB
'------------------------------------------------------------------------------

FUNCTION BuildHeader(hWnd AS DWORD) AS STRING
   LOCAL sFormat   AS ASCIIZ * 20
   LOCAL sDate     AS ASCIIZ * 40
   LOCAL sTime     AS ASCIIZ * 15
   LOCAL sText     AS STRING
   LOCAL st        AS SYSTEMTIME, lLen AS LONG

   sText = STRING$(%MAX_PATH, CHR$(0))
   ' Retrieve the current local date and time
   GetLocalTime st
   sFormat = "HH':'mm':'ss"
   GetTimeFormat %LOCALE_SYSTEM_DEFAULT, %NULL, st, sFormat, sTime, SIZEOF(sTime)
   sFormat = "yyyy'/'MM'/'dd "
   GetDateFormat %LOCALE_SYSTEM_DEFAULT, %NULL, st, sFormat, sDate, SIZEOF(sDate)
   lLen = GetWindowText(hWnd, BYVAL STRPTR(sText), %MAX_PATH)
   sText = LEFT$(sText, lLen)
   FUNCTION = "文件: " & sText & "|" & "日期: " & sDate & " - 时间: " & sTime & " - 页码:"
END FUNCTION
'------------------------------------------------------------------------------

FUNCTION MBProc(BYVAL nCode AS LONG, BYVAL wParam AS LONG, BYVAL lParam AS LONG) AS LONG
   ' Centers MessageBox dialog within CLIENT area of our dialog, rather than the desktop.
   ' Monitors HCBT_CREATEWND rather than HCBT_ACTIVATE to prevent interference with the
   ' 'SnapTo default button' mouse pointer option.

   IF nCode = %HCBT_CREATEWND THEN
      LOCAL rc AS RECT, p AS POINTAPI
      LOCAL cw AS CBT_CREATEWND PTR, cst AS CREATESTRUCT PTR

      cw = lParam      ' Get pointer to CBT_CREATEWND struct so we can...
      cst = @cw.lpcs   ' get a pointer to the CREATESTRUCT struct
      GetClientRect @cst.hwndParent, rc
      ClientToScreen @cst.hwndParent, p
      @cst.x = (rc.nRight - @cst.cx) \ 2 + p.x
      @cst.y = (rc.nBottom - @cst.cy) \ 2 + p.y
      UnhookWindowsHookEx g_hHook
   END IF
   FUNCTION = %FALSE
END FUNCTION
'--------------------------------------------------------------------------------

SUB InitPageSetUp
   ' Initialize PAGESETUPDLG structure
   g_psDlg.lStructSize = SIZEOF(g_psDlg)
   g_psDlg.hwndOwner   = g_hWndMain
   'g_psDlg.hDevMode   = Obtained when GetPrinterDC() is called
   g_psDlg.hDevNames   = %NULL
   g_psDlg.Flags       = %PSD_DEFAULTMINMARGINS OR %PSD_INTHOUSANDTHSOFINCHES OR %PSD_MARGINS
   g_psDlg.hInstance   = g_hInst

   SetRect g_psDlg.rtMargin, _
            VAL(IniRead(g_zIni, "Printer options", "MarginLeft", ".5"))   * 1000, _
            VAL(IniRead(g_zIni, "Printer options", "MarginTop", ".5"))    * 1000, _
            VAL(IniRead(g_zIni, "Printer options", "MarginRight", ".5"))  * 1000, _
            VAL(IniRead(g_zIni, "Printer options", "MarginBottom", ".5")) * 1000

   g_psDlg.ptPaperSize.y   = 11000    ' 8.5 x 11 inches
   g_psDlg.ptPaperSize.x   = 8500
END SUB
'------------------------------------------------------------------------------

FUNCTION GetPrinterDC() AS LONG
   LOCAL pdlg AS PRINTDLGAPI, DvModePtr AS DEVMODE PTR

   IF g_psDlg.lStructSize = 0 THEN InitPageSetup
   IF g_psDlg.hDevMode THEN
      DvModePtr = GlobalLock(g_psDlg.hDevMode)
      ' Kill current printer DC, and replace with new based on any changes made in
      ' pagesetup dialog whether anything was changed or not.
      IF PrnDC THEN DeleteDC PrnDC
      FUNCTION = CreateDC("WINSPOOL", BYCOPY @DvModePtr.dmDeviceName & $NUL, BYVAL %NULL, BYVAL DvModePtr)
      GlobalUnLock g_psDlg.hDevMode
   ELSE
      ' Initialize the PRINTDLG structure.
      ZeroMemory VARPTR(pdlg), SIZEOF(pdlg)
      pdlg.lStructSize = SIZEOF(pdlg)

      ' Set the flag to return printer DC (no dialog)
      pdlg.Flags = %PD_RETURNDEFAULT OR %PD_RETURNDC

      ' Invoke the printer DIALOG box (fake).
      PrintDlg pdlg

      ' Get the handle of the Devicemode data
      g_psDlg.hDevMode = pdlg.hDevMode

      ' hDC member of the PRINTDLG structure contains the printer DC.
      FUNCTION = pdlg.hDC
   END IF
END FUNCTION
'------------------------------------------------------------------------------

SUB GeneralErrorMsg(msg AS STRING)
   ' A well-behaved message box
   LOCAL cap AS STRING

   msg = msg & $NUL
   cap   = "PBscite Error" & $NUL
   g_hHook = SetWindowsHookEx(%WH_CBT, CODEPTR(MBProc), GetWindowLong(g_hWndMain, %GWL_HINSTANCE), GetCurrentThreadId)
   MessageBoxEx   g_hWndMain, _
                  BYVAL STRPTR(msg), _
                  BYVAL STRPTR(cap), _
                  BYVAL %MB_OK OR %MB_TASKMODAL OR %MB_ICONWARNING OR %MB_TOPMOST, _
                  BYVAL %LANG_ENGLISH
END SUB
'------------------------------------------------------------------------------

' This sub allows two threads to share the same array. The array descriptor
' address is passed BYVAL to this sub. The array name must be the same as
' the one passed to the thread.
SUB TrickTheCompiler(PageIndex() AS LONG, NewDimension AS LONG, NewData AS LONG)
   REDIM PRESERVE PageIndex(NewDimension)
   PageIndex(NewDimension) = NewData
END SUB
'------------------------------------------------------------------------------

FUNCTION TrickTheCompiler2(PageIndex() AS LONG, nIndex AS LONG) AS LONG
   FUNCTION = PageIndex(nIndex)
END FUNCTION
'------------------------------------------------------------------------------

FUNCTION PreviewProc (BYVAL hWnd AS LONG, BYVAL wMsg AS LONG, _
                     BYVAL wParam AS LONG, BYVAL lParam AS LONG) AS LONG
  LOCAL pi AS PREVIEWINFO PTR
  LOCAL  rc AS RECT
  LOCAL  hPen AS DWORD, hDC AS DWORD, rslt AS DWORD, BorderW AS LONG
  STATIC hCursor AS LONG

  ' 获取子窗口数据
  pi = GetWindowLong(hWnd, %GWL_USERDATA)

  SELECT CASE wMsg
    CASE %WM_CREATE
      pi = HeapAlloc(GetProcessHeap(), %HEAP_ZERO_MEMORY, SIZEOF(@pi))
      IF pi THEN
        SetWindowLong hWnd, %GWL_USERDATA, pi   ' 保存最后使用的指针
      ELSE
        FUNCTION = -1 : EXIT FUNCTION           ' 分配内存失败
      END IF

      @pi.hChild = hWnd
      @pi.hCtl   = GetEdit
      PrnDC      = GetPrinterDC

      ' 缩放客户区尺寸到纸张尺寸 (使用设备单位).
      @pi.ScaleX = GetDeviceCaps(PrnDC, %PHYSICALWIDTH) - 1
      @pi.ScaleY = GetDeviceCaps(PrnDC, %PHYSICALHEIGHT) - 1

      BorderW = GetSystemMetrics(%SM_CXFIXEDFRAME) * 2
      GetClientRect GetDeskTopWindow(),rc 'g_hWndClient, rc
      rc.nBottom-=35
      @pi.lHeight = rc.nBottom - (GetSystemMetrics(%SM_CYCAPTION) + BorderW)
      @pi.lWidth = MulDiv(@pi.lHeight, @pi.ScaleX + 1, @pi.ScaleY + 1)

      ' Scintilla 渲染到画布上。不直接画到客户区上，因为这会导致一个问题
      ' renders to this canvas. We don't draw directly on
      ' the client area since painting would be an issue.
      hDC = GetDC(hWnd)
      @pi.hBmp = CreateCompatibleBitmap(hDC, @pi.lWidth, @pi.lHeight)
      @pi.hDCTo = CreateCompatibleDC(hDC)
      ReleaseDC hWnd, hDC
      hPen = CreatePen(%PS_SOLID, 2, %LTGRAY)   ' 用于显示矩形的空白 For the margin display rectangle
      SelectObject @pi.hDCTo, @pi.hBmp
      SelectObject @pi.hDCTo, hPen

      ' 这是关键部分. 缩放渲染DC以匹配被绘制页面的尺寸。
      '
      ' We scale our rendering DC to match the size of
      ' the printed page. We scale Scintilla's output to the ratio of the
      ' printer page size in Device Units vs the bitmap size in Device Units.
      SetMapMode @pi.hDCTo, %MM_ANISOTROPIC
      SetWindowExtEx @pi.hDCTo, @pi.ScaleX, @pi.ScaleY, BYVAL %NULL
      SetViewportExtEx @pi.hDCTo, @pi.lWidth, @pi.lHeight, BYVAL %NULL

      ' Initialize scroll bar
      @pi.sci.cbSize = SIZEOF(@pi.sci)
      @pi.sci.fMask  = %SIF_ALL OR %SIF_DISABLENOSCROLL
      @pi.sci.nMin   = 1
      @pi.sci.nPos   = 1
      @pi.sci.nPage  = 1
      SetScrollInfo hWnd, %SB_VERT, @pi.sci, %FALSE

      ClientToScreen g_hWndClient, BYVAL VARPTR(rc)
      ' Size & position child
      MoveWindow hWnd, rc.nLeft, 0, @pi.lWidth + GetSystemMetrics(%SM_CXVSCROLL) + BorderW, _
                  @pi.lHeight + GetSystemMetrics(%SM_CYCAPTION) + BorderW, %FALSE
      hCursor = LoadCursor(%NULL, BYVAL %IDC_WAIT)
      SetCursor hCursor
      THREAD CREATE PreviewText(pi) TO @pi.hThread
      EXIT FUNCTION
   CASE %WM_SIZE

     DIALOG REDRAW hWnd
   CASE %WM_PAINT
      LOCAL ps AS PAINTSTRUCT

      BeginPaint hWnd, ps
      IF @pi.lFlag THEN
         SaveDC @pi.hDCTo
         SetMapMode @pi.hDCTo, %MM_TEXT   ' Temporarily set default mapping mode
         BitBlt ps.hdc, ps.rcPaint.nLeft, ps.rcPaint.nTop, ps.rcPaint.nRight, ps.rcPaint.nBottom, _
               @pi.hDCTo, ps.rcPaint.nLeft, ps.rcPaint.nTop, %SRCCOPY
         RestoreDC @pi.hDCTo, (-1)        ' Restore our desired mapping mode
         SetWindowText hWnd, "打印预览 - 第" & STR$(@pi.nPage) & " 页,共" & STR$(@pi.MaxPage) & " 页"
      ELSE
         GetClientRect hWnd, rc
         FillRect ps.hdc, rc, GetSysColorBrush(%COLOR_WINDOW)
      END IF
      EndPaint hWnd, ps
      EXIT FUNCTION

   CASE %WM_SETTEXT
      IF @pi.nPage = 1 THEN
         hCursor = LoadCursor(%NULL, BYVAL %IDC_ARROW)
         SetCursor hCursor
      END IF

   CASE %WM_MOUSEMOVE
      SetCursor hCursor
      EXIT FUNCTION

   CASE %WM_MOUSEWHEEL
      PostMessage hWnd, %WM_VSCROLL, IIF&(HIINT(wParam) > 0, %SB_PAGEUP, %SB_PAGEDOWN), %NULL
      FUNCTION = 1 : EXIT FUNCTION

   CASE %WM_VSCROLL
      SELECT CASE LOWRD(wParam)
      CASE %SB_THUMBTRACK
         @pi.nPage = HIWRD(wParam)

      CASE %SB_LINEDOWN, %SB_PAGEDOWN
         IF @pi.nPage < @pi.MaxPage THEN INCR @pi.nPage

      CASE %SB_LINEUP, %SB_PAGEUP
         IF @pi.nPage > 1 THEN DECR @pi.nPage

      CASE %SB_TOP
         @pi.nPage = 1

      CASE %SB_BOTTOM
         @pi.nPage = @pi.MaxPage

      CASE ELSE
         EXIT FUNCTION
      END SELECT   'LOWRD(wParam)

      @pi.sci.nPos = @pi.nPage
      SetScrollInfo hwnd, %SB_VERT, @pi.sci, %TRUE
      EXIT FUNCTION

   CASE %WM_KEYDOWN
      LOCAL Action AS LONG

      SELECT CASE LOWRD(wParam)
      CASE %VK_HOME
         Action = %SB_TOP

      CASE %VK_END
         Action = %SB_BOTTOM

      CASE %VK_DOWN
         Action = %SB_LINEDOWN

      CASE %VK_UP
         Action = %SB_LINEUP

      CASE %VK_PGDN
         Action = %SB_PAGEDOWN

      CASE %VK_PGUP
         Action = %SB_PAGEUP

      CASE ELSE
         FUNCTION = 1 : EXIT FUNCTION
      END SELECT 'LOWRD(wParam)
      SendMessage hWnd, %WM_VSCROLL, Action, %NULL
      EXIT FUNCTION

   CASE %WM_CLOSE
      @pi.lFlag = %FALSE                      ' Force preview thread to exit
      WaitforsingleObject @pi.hThread, 2000   ' Then wait till it does so
      THREAD CLOSE @pi.hThread TO rslt

   CASE %WM_DESTROY
      IF PrnDC THEN DeleteDC PrnDC : PrnDC = 0
      DeleteDC @pi.hDCTo
      DeleteObject hPen
      DeleteObject @pi.hBmp
      HeapFree(GetProcessHeap, %NULL, BYVAL GetWindowLong(hWnd, %GWL_USERDATA))
   END SELECT   ' wMsg

   FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam)
END FUNCTION
'------------------------------------------------------------------------------

'FUNCTION AbortProc(hDC AS LONG, iCode AS LONG) AS LONG
'   ' Called when EndPage is called.
'   ' Can be useless for keeping a modeless message pump active.
'   ' One of the reasons why we don't use it.
'   LOCAL Msg AS tagMSG

'   WHILE PeekMessage(msg, %NULL, 0, 0, %PM_REMOVE)
'      IF g_hModeless <> 0 AND NOT(IsDialogMessage(g_hModeless, msg)) THEN
'         CALL TranslateMessage (msg)
'         CALL DispatchMessage (msg)
'      END IF
'   WEND

'   FUNCTION = ISFALSE(g_AbortPrint)
'END FUNCTION
'------------------------------------------------------------------------------

CALLBACK FUNCTION AbortDlgProc() AS LONG
   SELECT CASE CBMSG
   CASE %WM_COMMAND
      IF CBCTL = 303 AND CBCTLMSG = %BN_CLICKED THEN
         g_AbortPrint = %TRUE
         DIALOG END CBHNDL
      END IF
   END SELECT 'CBMSG
END FUNCTION
'------------------------------------------------------------------------------

FUNCTION ShowAbortDialog(BYVAL Junk AS LONG) AS DWORD   ' Structured for threading
   LOCAL COUNT AS LONG

   DIALOG NEW g_hWndMain, "Cancel Print Job",,,150, 90,, TO g_hModeless
   CONTROL ADD LABEL, g_hModeless, 301, "", 5,  5,  140, 30
   CONTROL ADD LABEL, g_hModeless, 300, "", 38, 40, 80,  15
   CONTROL ADD BUTTON, g_hModeless, 303, "&Cancel", 50, 65, 50, 15, , , CALL AbortDlgProc
   CenterWindow BYVAL g_hModeless
   DIALOG SHOW MODELESS g_hModeless

   DO      ' Message pump
      DIALOG DOEVENTS TO COUNT
   LOOP WHILE COUNT

   g_hModeless = 0
END FUNCTION
'------------------------------------------------------------------------------

FUNCTION GetPageNum(BYVAL pInfo AS DWORD) AS LONG
   ' Does all the margin calculations and determines the true # of pages in the document
   LOCAL Dpi AS POINTAPI, PAGESIZE AS POINTAPI, Printable AS POINTAPI
   LOCAL Unprintable AS RECT, MARGIN AS RECT
   LOCAL Temp AS LONG, pg AS LONG, szText AS ASCIIZ * 32
   LOCAL pni AS PAGENUMINFO PTR
   LOCAL tm AS TEXTMETRIC

   pni = pInfo

   ' Get dots per inch x/y (printer resolution)
   Dpi.x = GetDeviceCaps(PrnDC, %LOGPIXELSX)   ' dpi in X direction
   Dpi.y = GetDeviceCaps(PrnDC, %LOGPIXELSY)   ' dpi in Y direction

   ' Get the length/width of the page (in device units)
   PageSize.x = GetDeviceCaps(PrnDC, %PHYSICALWIDTH)
   PageSize.y = GetDeviceCaps(PrnDC, %PHYSICALHEIGHT)

   ' Get the PRINTABLE length/width of the page (in device units)
   Printable.x = GetDeviceCaps (PrnDC, %HORZRES)
   Printable.y = GetDeviceCaps (PrnDC, %VERTRES)

   ' Get page Unprintable width & height (in device units)
   SetRect Unprintable, _
            GetDeviceCaps(PrnDC, %PHYSICALOFFSETX), _
            GetDeviceCaps(PrnDC, %PHYSICALOFFSETY), _
            PageSize.x - Printable.x - GetDeviceCaps(PrnDC, %PHYSICALOFFSETX), _
            PageSize.y - Printable.y - GetDeviceCaps(PrnDC, %PHYSICALOFFSETY)
   ' At this point, Unprintable contains the widths of the Unprintable
   ' regions on all four sides of the page in device units

   ' Convert the hundredths of millimeters (HiMetric) or thousandths of inches (HiEnglish) margin
   ' values from the Page Setup dialog to device units (There are 2540 hundredths of a mm in an inch).
   ' Don't reduce margins below the unprintable area.
   GetLocaleInfo %LOCALE_USER_DEFAULT, %LOCALE_IMEASURE, szText, 3
   Temp = IIF&(szText = "1", 1000, 2540) ' 1 = US System. 0 = Metric system
   SetRect MARGIN, _
            MAX(Unprintable.nLeft,  MulDiv(g_psDlg.rtMargin.nLeft,   Dpi.x, Temp)), _
            MAX(Unprintable.nTop,   MulDiv(g_psDlg.rtMargin.nTop,    Dpi.y, Temp)), _
            MAX(Unprintable.nRight, MulDiv(g_psDlg.rtMargin.nRight,  Dpi.x, Temp)), _
            MAX(Unprintable.nBottom,MulDiv(g_psDlg.rtMargin.nBottom, Dpi.y, Temp))
   ' Margin now contains the values used to shrink the printable area of the page

   ' Convert device coordinates to logical coordinates (in case mapping mode changes)
   DPtoLP(PrnDC, BYVAL VARPTR(MARGIN), 2)
   DPtoLP(PrnDC, BYVAL VARPTR(Unprintable), 2)

   ' Convert page size to logical coordinates
   DPtoLP(PrnDC, PAGESIZE, 1)

   ' Set up the Physically printable page size. We use the entire area of measuring
   ' device, since margins will never be less than the unprintable margins.
   SetRect @pni.fr.rcPage, 0, 0, PageSize.x - 1, PageSize.y - 1

   ' Set up the Rectangle in which to print.
   ' We must subtract the physical margins from the printable page size.
   ' Scintilla appears to have problems with printing to the right margin.
   ' (Richedit has problems with the bottom margin). I believe this is due
   ' to the way Scintilla calculates word breaks.
   ' This issue means that the right margin will generally not be accurate.
   GetClassName GetForegroundWindow, szText, SIZEOF(szText)
   IF szText = "NPREVIEW32" THEN   ' Rendering to Print Preview window
      SetRect @pni.fr.rc, _
               Margin.nLeft, _
               Margin.nTop, _
               Pagesize.x - Margin.nRight, _
               Pagesize.y - Margin.nBottom
   ELSE                  ' Rendering to Printer
      ' Scintilla adds unprintable top & left margins to desired page coordinates
      SetRect @pni.fr.rc, _
               Margin.nLeft - Unprintable.nLeft, _
               Margin.nTop - Unprintable.nTop, _
               PageSize.x - Margin.nRight - Unprintable.nLeft, _
               PageSize.y - Margin.nBottom - Unprintable.nTop
   END IF

   ' Get the header area set up
   szText = IIF$(SciColorsAndFonts.DefaultFontName = "", "Courier New", SciColorsAndFonts.DefaultFontName)
   ' Always use a 9 point font size
   @pni.hFontHeader = CreateFont(-muldiv(9, Dpi.y, 72), _
                                 0, 0, 0, _
                                 %FW_BOLD, _
                                 %FALSE, _
                                 %FALSE, _
                                 0, 0, 0, _
                                 0, 0, 0, _
                                 BYVAL VARPTR(szText))
   SelectObject PrnDC, @pni.hFontHeader
   GetTextMetrics PrnDC, tm
   @pni.HeaderLineHeight = tm.tmHeight + tm.tmExternalLeading

   LOCAL ms AS DWORD, fr AS RANGETOFORMAT
   fr = @pni.fr   ' Use local structure for page # calc's
   fr.rc.nTop = fr.rc.nTop + @pni.HeaderLineHeight * 3   ' Adjust Scintilla's rendering area
   ' If printing, we always need entire document length
   IF @pni.PrintFlag THEN _
      fr.chrg.cpMin = 0 : fr.chrg.cpMax = SendMessage(@pni.hEdit, %SCI_GETLENGTH, %NULL, %NULL)
   @pni.ms = 0

   ' Now that we know the printable size, determine # of pages
   WHILE fr.chrg.cpMin < fr.chrg.cpMax
      ' Recover printer DC if printing or another preview window terminates
      IF PrnDC = 0 THEN PrnDC = GetPrinterDC : fr.hdcTarget = PrnDC
      ' Needed when printing selected text
      IF fr.chrg.cpMin <= @pni.fr.chrg.cpMin THEN @pni.SelPage = pg + 1
      ms = TimeGetTime
      fr.chrg.cpMin = SendMessage(@pni.hEdit, %SCI_FORMATRANGE, %FALSE, VARPTR(fr))
      ' Used for displaying how long it takes Scintilla to paginate
      @pni.ms = @pni.ms + (TimeGetTime - ms)
      INCR pg

      ' Keep track of each page's start char index
      TrickTheCompiler BYVAL @pni.pArray, pg, fr.chrg.cpMin
      SLEEP 0
   WEND

   ' The last two array elements must be identical
   TrickTheCompiler BYVAL @pni.pArray, pg, TrickTheCompiler2(BYVAL @pni.pArray, pg - 1)

   ' Tell Scintilla to release cached information
   PostMessage @pni.hEdit, %SCI_FORMATRANGE, %NULL, %NULL
   FUNCTION = pg
END FUNCTION
'------------------------------------------------------------------------------

FUNCTION GetPageNumThread(BYVAL Param AS DWORD) AS LONG
   ' Wrapper function for GetPageNum threading
   FUNCTION = GetPageNum(Param)
END FUNCTION
'------------------------------------------------------------------------------

FUNCTION PreviewText(BYVAL piPtr AS DWORD) AS DWORD
   LOCAL CurPage AS LONG, TotPage AS LONG, hEdit AS LONG, pSciData AS LONG, rslt AS LONG
   LOCAL rc AS RECT         ' General purpose rectangle
   LOCAL fr AS RANGETOFORMAT
   LOCAL pi   AS PREVIEWINFO PTR
   STATIC pni AS PAGENUMINFO
   LOCAL hPen AS DWORD, ta AS DWORD
   LOCAL HEADER AS STRING, sPath AS STRING, sDate AS STRING

   DIM ma(1) AS LOCAL RECT

   ' Get preview child window data
   pi = piPtr

   hEdit = @pi.hCtl   ' The Edit Control we're previewing
   pSciData = SendMessage(hEdit, %SCI_GETDIRECTPOINTER, %NULL, %NULL)
   IF ISFALSE pSciData THEN EXIT FUNCTION
   CurPage = 1

   fr.chrg.cpMin = SciMsg(pSciData, %SCI_GETSELECTIONSTART, %NULL, %NULL)
   fr.chrg.cpMax = SciMsg(pSciData, %SCI_GETSELECTIONEND, %NULL, %NULL)

   ' If no selection was made, set start & end = to the whole file
   IF fr.chrg.cpMax = fr.chrg.cpMin THEN _
      fr.chrg.cpMin = 0 : fr.chrg.cpMax = SciMsg(pSciData, %SCI_GETLENGTH, %NULL, %NULL)

   fr.hdc         = @pi.hDCTo  ' Device context to render to (where output goes)
   fr.hdcTarget   = PrnDC      ' Device context to measure against

   DIM PageIndex(0) AS LOCAL LONG   ' Shared with GetPageNum()
   PageIndex(0) = fr.chrg.cpMin     ' Set start position in document

   ' Init PAGENUMINFO
   pni.fr     = fr
   pni.hEdit  = hEdit
   pni.pArray = VARPTR(PageIndex()) ' Pass address of array descriptor as param
#IF %USE_THREAD
   THREAD CREATE GetPageNumThread(VARPTR(pni)) TO rslt
   THREAD CLOSE rslt TO rslt

   ' Wait for margin calculations to complete. We want to show the 1st page
   ' as soon as we can rather than wait for the page # calc's to complete.
   DO : TotPage = UBOUND(PageIndex) : SLEEP 1 : LOOP UNTIL TotPage > 0
#ELSE
   TotPage = GetPageNum(VARPTR(pni))
#ENDIF

   @pi.sci.nMax = TotPage
   SetScrollInfo @pi.hChild, %SB_VERT, @pi.sci, %TRUE   ' Update scroll bar

   ' Get our margin results
   fr = pni.fr

   ' Set up margin rectangle
   rc = fr.rc
   InflateRect rc, 1, 1
   SetRect ma(0), rc.nRight, rc.nTop, rc.nRight, rc.nBottom
   SetRect ma(1), rc.nLeft, rc.nBottom , rc.nLeft, rc.nTop
   MoveToEx @pi.hDCTo, rc.nLeft, rc.nTop, BYVAL %NULL

   ' Init PREVIEWINFO
   @pi.nPage = 1 : @pi.MaxPage = TotPage : @pi.lFlag = %TRUE

   ' Setup header stuff...
   HEADER = BuildHeader(GetParent(@pi.hCtl))
   rslt = INSTR(HEADER, "|")
   sPath = LEFT$(HEADER, rslt - 1)
   HEADER = RIGHT$(HEADER, LEN(HEADER) - rslt)
   fr.rc.nTop = fr.rc.nTop + pni.HeaderLineHeight * 3   ' Adjust Scintilla's rendering area
   SelectObject @pi.hDCTo, pni.hFontHeader

   GetClientRect @pi.hChild, rc

   ' This does the previewing. We send messages to Scintilla telling it to format its
   ' text to fit the printer's DC & send result to the preview child for painting.
   WHILE fr.chrg.cpMin < fr.chrg.cpMax AND @pi.lFlag
      ' Recover printer DC if printing or another preview window terminates
      IF PrnDC = 0 THEN PrnDC = GetPrinterDC

      ' Refresh device context to measure against in case PrintText()
      ' changes it (since we can preview & print at the same time).
      fr.hdcTarget = PrnDC

      ' Clear rendering DC canvas
      PatBlt @pi.hDCTo, 0, 0, @pi.ScaleX, @pi.ScaleY, %WHITENESS

      ' Print header to our rendering DC
      ta = SetTextAlign(@pi.hDCTo, %TA_BOTTOM)
      SetTextColor @pi.hDCTo, %BLACK : SetBkColor @pi.hDCTo, %WHITE
      ExtTextOut @pi.hDCTo, fr.rc.nLeft + 5, fr.rc.nTop - (pni.HeaderLineHeight * 2), _
                  %ETO_OPAQUE, BYVAL %NULL, BYVAL STRPTR(sPath), LEN(sPath), BYVAL %NULL
      sDate = HEADER & STR$(CurPage) & " /" & STR$(TotPage)
      ExtTextOut @pi.hDCTo, fr.rc.nLeft + 5, fr.rc.nTop - pni.HeaderLineHeight + 5, _
                  %ETO_OPAQUE, BYVAL %NULL, BYVAL STRPTR(sDate), LEN(sDate), BYVAL %NULL
      SetTextAlign @pi.hDCTo, ta
      SaveDC @pi.hDCTo
         hPen = CreatePen(%PS_SOLID, 1, %BLACK)
         SelectObject @pi.hDCTo, hPen
         MoveToEx @pi.hDCTo, fr.rc.nLeft, fr.rc.nTop - pni.HeaderLineHeight \ 4, BYVAL %NULL
         LineTo @pi.hDCTo, fr.rc.nRight, fr.rc.nTop - pni.HeaderLineHeight \ 4
         DeleteObject hPen
      RestoreDC @pi.hDCTo, (-1)

      ' Print formatted text to our rendering DC
      fr.chrg.cpMin = SciMsg(pSciData, %SCI_FORMATRANGE, %TRUE, VARPTR(fr))

      ' Draw the margin rectangle
      PolylineTo @pi.hDCTo, BYVAL VARPTR(ma(0)), 4

      ' Triggers painting
      InvalidateRect @pi.hChild, rc, %FALSE : UpdateWindow @pi.hChild

      ' Wait for new page command from preview window. We can take advantage of this idle
      ' time by capturing the remaining page indexes from the GetPageNum() thread...
      WHILE CurPage = @pi.nPage AND @pi.lFlag
         ' If user closes editor we're previewing, close preview child also- forcing this thread to terminate
         IF IsWindow(hEdit) = 0 THEN PostMessage @pi.hChild, %WM_SYSCOMMAND, %SC_CLOSE, %NULL : EXIT, EXIT
#IF %USE_THREAD
         ' Monitor the PageIndex array's size. It's being redimmed in the GetPageNum() thread.
         TotPage = UBOUND(PageIndex)
         IF TotPage > @pi.MaxPage THEN
            @pi.MaxPage = TotPage : @pi.sci.nMax = TotPage
            SetScrollInfo @pi.hChild, %SB_VERT, @pi.sci, %TRUE   ' Update preview child's scroll bar
            ' Update preview child's displayed total page count
            SetWindowText @pi.hChild, "Print Preview - Page" & STR$(@pi.nPage) & " of" & STR$(@pi.MaxPage) & _
                           " in" & STR$(ROUND(pni.ms / 1000, 3)) & " sec."
         END IF
#ENDIF
         SLEEP 1
      WEND

      ' If viewing a previous page...
      IF @pi.nPage < CurPage THEN _
         CurPage = @pi.nPage : fr.chrg.cpMin = PageIndex(CurPage - 1) : ITERATE LOOP

      ' If 'END' key pressed or using scroll handle to scroll up & down...
      IF @pi.nPage = TotPage OR @pi.nPage > CurPage + 1 THEN _
         CurPage = @pi.nPage : fr.chrg.cpMin = PageIndex(CurPage) : ITERATE LOOP

      INCR CurPage
   WEND

   DeleteObject pni.hFontHeader
   ' Tell Scintilla to release cached information
   PostMessage hEdit, %SCI_FORMATRANGE, %NULL, %NULL
END FUNCTION
'------------------------------------------------------------------------------


FUNCTION PrintText(BYVAL tbFlag AS LONG) AS DWORD
   LOCAL fr AS RANGETOFORMAT, rDocInfo AS DOCINFO, pd AS PRINTDLGAPI
   LOCAL ms AS DWORD
   LOCAL CurPage AS LONG, TotPage AS LONG, ORIENTATION AS LONG, RENDER AS LONG, hEdit AS LONG
   LOCAL Flags AS DWORD, vDevMode AS DWORD
   LOCAL zString AS ASCIIZ * 130, pName AS ASCIIZ * 32
   LOCAL pNamePTR      AS ASCIIZ PTR      ' Pointer to printer name
   LOCAL DvModePtr   AS DEVMODE PTR
   LOCAL DevNamePTR   AS DEVNAMES PTR   ' Pointer to printer names structure
   LOCAL hPen AS DWORD, ta AS DWORD, rslt AS DWORD
   LOCAL HEADER AS STRING, sPath AS STRING, sDate AS STRING
   STATIC pni AS PAGENUMINFO

   hEdit = GetEdit
   PrnDC = GetPrinterDC

   IF tbFlag THEN ' Toolbar selection - don't show print dialog
      Flags      = %PD_RETURNDEFAULT
      vDevMode   = %NULL
   ELSE
      Flags         = %PD_USEDEVMODECOPIESANDCOLLATE OR %PD_SHOWHELP OR %PD_RETURNDC
      vDevMode      = g_psDlg.hDevMode
      DvModePtr     = GlobalLock(g_psDlg.hDevMode)
      ORIENTATION   = @DvModePtr.dmOrientation   ' Log current orientation
      GlobalUnLock g_psDlg.hDevMode
   END IF

   fr.chrg.cpMin = SendMessage(hEdit, %SCI_GETSELECTIONSTART, %NULL, %NULL)
   fr.chrg.cpMax = SendMessage(hEdit, %SCI_GETSELECTIONEND, %NULL, %NULL)

   ' If nothing selected, disable the print dialog SELECTION radio button
   IF fr.chrg.cpMax = fr.chrg.cpMin THEN
      Flags = Flags OR %PD_NOSELECTION
      ' Since no selection was made, set start & end = to the whole file
      fr.chrg.cpMin = 0 : fr.chrg.cpMax = SendMessage(hEdit, %SCI_GETLENGTH, %NULL, %NULL)
   ELSE  ' Make the SELECTION button the default
      Flags = Flags OR %PD_SELECTION
   END IF

   DIM PageIndex(0) AS LOCAL LONG   ' Shared with GetPageNum(). Only needed for page #'s
   IF tbFlag THEN pni.PrintFlag = %TRUE

   fr.hdc       = PrnDC
   fr.hdcTarget = PrnDC

   pni.hEdit    = hEdit
   pni.fr       = fr
   pni.pArray   = VARPTR(PageIndex()) ' Pass address of array descriptor as param

   ' Setup page margins & determine how many pages the document will use
   TotPage = GetPageNum(VARPTR(pni))

   ' Setup the print common dialog
   FillMemory VARPTR(pd), SIZEOF(pd), 0
   pd.lStructSize   = SIZEOF(pd)
   pd.hwndOwner     = g_hWndMain
   pd.hDevMode      = vDevMode  ' Either NULL or set by page setup dlg
   pd.hDevNames     = %NULL
   pd.Flags         = Flags
   pd.nFromPage     = 1
   pd.nToPage       = TotPage
   pd.nMinPage      = 1
   pd.nMaxPage      = TotPage   ' Activates pages
   pd.nCopies       = 1

   IF PrintDlg(pd) THEN
      ' In case user may have changed settings from print dialog
      IF vDevMode THEN   ' The dialog was shown
         DvModePtr = GlobalLock(pd.hDevMode)
         ' Kill current DC, and replace with new.
         DeleteDC PrnDC
         PrnDC               = pd.hDC
         pni.fr.hdc         = PrnDC
         pni.fr.hdcTarget   = PrnDC
         ' See if we need to recalc margins & total pages
         IF ORIENTATION <> @DvModePtr.dmOrientation THEN _
            TotPage = GetPageNum(VARPTR(pni))
         GlobalUnLock pd.hDevMode
      END IF

      ' We maintain our own copy of the margin results for running multiple threads
      fr      = pni.fr
      CurPage = IIF&(tbFlag, pni.SelPage, 1)

      ' Fill rDocInfo structure - standard printing stuff
      FillMemory VARPTR(rDocInfo), SIZEOF(rDocInfo), 0
      rDocInfo.cbSize = SIZEOF(rDocInfo)

      ' Get the name of the current printer
      DevNamePTR    = GlobalLock(pd.hDevNames)
      pNamePTR      = DevNamePTR + @DevNamePTR.wDeviceOffset
      pName         = @pNamePTR
      GlobalUnLock pd.hDevNames

      ' Setup header stuff...
      HEADER     = BuildHeader(GetParent(hEdit))
      rslt       = INSTR(HEADER, "|")
      sPath      = LEFT$(HEADER, rslt - 1)
      HEADER     = RIGHT$(HEADER, LEN(HEADER) - rslt)
      fr.rc.nTop = fr.rc.nTop + pni.HeaderLineHeight * 3   ' Adjust Scintilla's rendering area

      ' Get the filename of what we're printing for use in print-job information
      zString = GetFileName(sPath)
      rDocInfo.lpszDocName = VARPTR(zString)

      ' Allow Abort dialog message processing...
      'SetAbortProc fr.hdc, CODEPTR(AbortProc)

      ' We don't use SetAbortProc for efficiency reasons.
      ' Generate Abort dialog and run in a separate thread. Easier on the forests since
      ' the worst case is usually 1 page printed when cancelled. You even get to see
      ' the button get pushed :) 'Flags' is reused here as a dummy variable
      THREAD CREATE ShowAbortDialog(Flags) TO Flags
      THREAD CLOSE Flags TO Flags
      DO UNTIL IsWindowEnabled(g_hModeless) : SLEEP 1 : LOOP   ' Wait for dialog to show
      SetWindowText GetDlgItem(g_hModeless, 301), "Printing " & zString & $CRLF & "on " & pName
      ShowWindowAsync g_hModeless, %SW_SHOW

      ' Emulate MODAL by disabling keyboard/mouse input to main dialog
      EnableWindow g_hWndMain, %FALSE

      ' As a precaution, tell Scintilla to release cached information (if any)
      SendMessage hEdit, %SCI_FORMATRANGE, %NULL, %NULL

      ' Here we go...
      IF StartDoc(fr.hdc, rDocInfo) < 0 THEN
         GeneralErrorMsg "Can not start print job."
      ELSE
         ' This does the printing. We send messages to Scintilla telling it to
         ' format its text to fit the printer's DC & send result to the printer.
         WHILE (fr.chrg.cpMin < fr.chrg.cpMax)
            IF g_AbortPrint THEN EXIT LOOP
            ms = TimeGetTime
            ' Wait for desired 1st page before printing
            IF CurPage => pd.nFromPage THEN
               RENDER = %TRUE      ' Allow Scintilla to render text to device
               StartPage fr.hdc   ' Mark beginning of current page
            END IF

            ' Recover printer DC if preview closed while printing
            IF PrnDC = 0 THEN PrnDC = GetPrinterDC

            ' Print header
            ta = SetTextAlign(PrnDC, %TA_BOTTOM)
            SetTextColor PrnDC, %BLACK : SetBkColor PrnDC, %WHITE
            SelectObject PrnDC, pni.hFontHeader
            ExtTextOut PrnDC, fr.rc.nLeft + 5, fr.rc.nTop - (pni.HeaderLineHeight * 2), _
                        %ETO_OPAQUE, BYVAL %NULL, BYVAL STRPTR(sPath), LEN(sPath), BYVAL %NULL
            sDate = HEADER & STR$(CurPage) & " of" & STR$(TotPage)
            ExtTextOut PrnDC, fr.rc.nLeft + 5, fr.rc.nTop - pni.HeaderLineHeight + 5, _
                        %ETO_OPAQUE, BYVAL %NULL, BYVAL STRPTR(sDate), LEN(sDate), BYVAL %NULL
            SetTextAlign PrnDC, ta
            SaveDC PrnDC
               hPen = CreatePen(%PS_SOLID, 2, %BLACK)
               SelectObject PrnDC, hPen
               MoveToEx PrnDC, fr.rc.nLeft, fr.rc.nTop - pni.HeaderLineHeight \ 4, BYVAL %NULL
               LineTo PrnDC, fr.rc.nRight, fr.rc.nTop - pni.HeaderLineHeight \ 4
               DeleteObject hPen
            RestoreDC PrnDC, (-1)

            ' Update Abort Dialog
            SetWindowText GetDlgItem(g_hModeless, 300), "Printing Page" & STR$(CurPage) & " of" & STR$(TotPage)

            ' Print as much text as can fit on a page. The return value is the index of the
            ' 1st char on the next page. Using %TRUE for wParam causes the text to be printed.
            fr.chrg.cpMin = SendMessage(hEdit, %SCI_FORMATRANGE, RENDER, VARPTR(fr))

            ' If there is more text to print, adjust the range of characters to start printing at the
            ' 1st character of the next page.
            IF (fr.chrg.cpMin < fr.chrg.cpMax) THEN
               ' Wait for desired 1st page- then send formfeed to printer to trigger printing of page
               IF (CurPage => pd.nFromPage) THEN EndPage fr.hdc
               IF CurPage = pd.nToPage THEN EXIT LOOP
               INCR CurPage
            END IF
            ms = TimeGetTime - ms
            ' Maintain a constant 150 ms delay between pages so user has some time to abort...
            ' BTW - it takes Scintilla a long time to format a page for printing
            IF ms < 150 THEN SLEEP 150 - ms
         WEND

         ' Tell Scintilla to release cached information
         SendMessage hEdit, %SCI_FORMATRANGE, %NULL, %NULL
      END IF

      IF g_AbortPrint THEN
         ' We need to use AbortDoc() since we don't use SetAbortProc()
         AbortDoc fr.hdc            ' Erases the entire spool job
         g_AbortPrint = %FALSE
      ELSE
         ' Clean up & finish print job
         EndPage fr.hdc
         EndDoc fr.hdc
         IF g_hModeless THEN _      ' Kill abort dialog
            SendMessage g_hModeless, %WM_SYSCOMMAND, %SC_CLOSE, %NULL
      END IF

      EnableWindow g_hWndMain, %TRUE      ' End MODAL emulation

      ' If we used print dialog, then dialog used psDlg.hDevMode.
      ' Print dialog always changes the hDevMode movable memory block location
      ' which would screw up any further calls to the page setup dialog. So...
      IF vDevMode THEN g_psDlg.hDevMode = pd.hDevMode
   ELSE
      ' User pressed Cancel from print dialog
   END IF

   IF PrnDC THEN DeleteDC PrnDC : PrnDC = 0
   IF pni.hFontHeader THEN DeleteObject pni.hFontHeader
   IF pd.hDevNames THEN GlobalFree pd.hDevNames
END FUNCTION
'------------------------------------------------------------------------------
