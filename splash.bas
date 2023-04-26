'splash.bas
'��������

GLOBAL hWndSplash    AS DWORD
GLOBAL timeSec       AS DWORD '����ͣ��������
FUNCTION CreateSplash(BYVAL hInst AS DWORD)AS LONG
  LOCAL wce          AS WndClassEx
  LOCAL szClassName  AS ASCIIZ * 80
  LOCAL STYLE        AS LONG
  LOCAL SplashWidth  AS LONG       ' ��������λͼ���.
  LOCAL SplashHeight AS LONG       ' ��������λͼ�߶�.
  ' ** ע��˽�д�����
  szClassName       = "SDKAPPSPLASHSCR"
  wce.cbSize        = SIZEOF(wce)
  wce.style         = %CS_HREDRAW OR %CS_VREDRAW
  wce.lpfnWndProc   = CODEPTR( SplashProc )
  wce.cbClsExtra    = 0
  wce.cbWndExtra    = 0
  wce.hInstance     = hInst
  wce.hIcon         = %NULL
  wce.hCursor       = LoadCursor( %NULL, BYVAL %IDC_ARROW )
  wce.hbrBackground = %COLOR_WINDOW + 1
  wce.lpszMenuName  = %NULL
  wce.lpszClassName = VARPTR( szClassName )
  wce.hIconSm       = %NULL

  IF ISFALSE(RegisterClassEx(wce)) THEN
   FUNCTION = %TRUE
   EXIT FUNCTION
  END IF

  '----------------------------------------------------------------------------
  ' ����������������.
  '----------------------------------------------------------------------------
  STYLE = %WS_POPUP OR %WS_CHILD

  '----------------------------------------------------------------------------
  ' ����λͼΪ350Wx150H�������ı�˳ߴ�����Ҫ����������������.
  '----------------------------------------------------------------------------
  SplashWidth  = 350
  SplashHeight = 150

  hWndSplash = CreateWindowEx(%WS_EX_TOPMOST, "SDKAPPSPLASHSCR", "", STYLE, _
              0, 0, SplashWidth, SplashHeight, _
              %HWND_DESKTOP, %NULL, hInst, BYVAL %NULL)   '

  ' ��ʾ��������.
  ShowWindow hWndSplash, %SW_SHOW
  UpdateWindow hWndSplash
  FUNCTION=%FALSE
END FUNCTION
'------------------------------------------------------------------------------
' ���� : SplashProc ()
'
' ���� : �����ص�����.
'
'------------------------------------------------------------------------------
FUNCTION SplashProc (BYVAL hWnd   AS DWORD, _
           BYVAL wMsg   AS DWORD, _
           BYVAL wParam AS DWORD, _
           BYVAL lParam AS LONG) AS LONG

  STATIC hBmp   AS DWORD
  STATIC hPal   AS DWORD
  STATIC hTimer AS LONG

  LOCAL  hDC    AS DWORD
  LOCAL  ps     AS PAINTSTRUCT
  LOCAL  r      AS RECT

  SELECT CASE wMsg
  CASE %WM_CREATE
    hBmp = LoadBitmapRes(0, "SPLASH", hPal)   'hInst
    IF timeSec=0 THEN
      timeSec=5
    END IF
    hTimer = SetTimer(hWnd, &HDABEDABE, timeSec*1000, %NULL)  '5 ��
    CenterWindow hWnd

  CASE %WM_PAINT
    hDC = BeginPaint(hWnd, ps)
    SelectObject hDC, hPal
    DrawBitmap hDC, hBmp, 0, 0
    RealizePalette hDC
    EndPaint hWnd, ps

    FUNCTION = 0
    EXIT FUNCTION

  CASE %WM_TIMER
    SendMessage hWnd, %WM_CLOSE, %NULL, %NULL
  CASE %WM_LBUTTONDOWN
    SendMessage hWnd, %WM_CLOSE, %NULL, %NULL
  CASE %WM_CLOSE
    KillTimer hWnd, hTimer
  CASE %WM_DESTROY
    DestroyWindow hWnd
    DeleteObject hBmp
    FUNCTION = 0
    EXIT FUNCTION

  END SELECT

  FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam)

END FUNCTION
