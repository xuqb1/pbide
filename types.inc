TYPE movewindow_type
  hWnd   AS LONG
  rStart AS rect
  rEnd   AS rect
END TYPE
'' TYPE DestroyWindow_type    ''' create, change, destroy
TYPE CreateWindowEx_type
  hWnd         AS LONG         '' retval
  dwExStyle    AS LONG         '' extended window style
  szClassName  AS ASCIIZ * 80  '' pointer TO registered class NAME
  szWindowName AS ASCIIZ * 80  '' pointer TO window NAME
  dwStyle      AS LONG         '' window style
  x            AS LONG         '' horizontal position of window
  y            AS LONG         '' vertical position of window
  nWidth       AS LONG         '' window width
  nHeight      AS LONG         '' window height
  hWndParent   AS LONG         '' HANDLE TO parent OR owner window
  hMenu        AS LONG         '' HANDLE TO MENU, OR child-window identifier
  hInstance    AS LONG         '' HANDLE TO application instance
  lpParam      AS LONG         '' always 0
END TYPE

TYPE SetWindowText_type
  hWnd  AS LONG
  szOld AS ASCIIZ * 256
  szNew AS ASCIIZ * 256
END TYPE

TYPE SetWindowLong_type
  hWnd   AS LONG
  lIndex AS LONG      '' %GWL_EXSTYLE ... %DWL_USER
  lOld   AS LONG
  lNew   AS LONG
END TYPE


UNION undo_UNION
  moveW   AS movewindow_type
  CreateW AS CreateWindowEx_type
  SetText AS SetWindowText_type
  SetLong AS SetWindowLong_type
END UNION


TYPE undoDMS_type
  action AS LONG
  undo   AS undo_union
END TYPE
TYPE settings
  showGrid AS LONG
  snapGrid AS LONG
  gridSize AS INTEGER             ' 4 to 16 ( step 2 )   default 8
  stickFrames AS LONG
  autoSizeImg AS LONG
  HANDLES AS LONG                 ' %use_HANDLESMALL ..  %use_HANDLELARGE
  backgrndType AS LONG            ' 0 = none, %use_backColor, %use_backBrush
  backColor AS LONG
  backBrush AS DWORD
  Include AS ASCIIZ * 255
END TYPE
TYPE flag_type          'used in WndClientProc as static
  Dragging AS LONG                'flags for WM_LBUTTONDWON, WM_MOUSEMOVE & WM_LBUTTONUP q.
  Sizing AS LONG          'flag, sizing a control
  Drawing AS LONG                 'flag, drawing a new control
  DrawNewControl AS LONG          'flag, Ok to drop a New Control
  ctype AS LONG           'control type
  LostCapture AS LONG             'flag, someone else took our mouse Capture
  fOneShot AS LONG                'flag, do something one time only
  fWM_RBUTTONUP AS LONG           'flag, Indicate Right ButtonUp on a Control
  IDhWndCtl AS LONG               'Right ButtonUp, ID's the control, for Context menu
  changed AS LONG                 'udpate flag - set by and new,move,del op.
END TYPE
TYPE GRIPPERS
       rcArea AS RECT          'Grippers coordinates
       dwNCHITVALUE AS DWORD           'NCHITTEST return value
END TYPE
GLOBAL DialogHandles() AS GRIPPERS
TYPE setting_type
  showGrid AS LONG                ' old : gGridOnOff
  snapGrid AS LONG                ' old : gSnapOnOff
  gridSize AS INTEGER             ' old : gGridSize  ' 4 to 16 ( step 2 )   default 8
  stickFrames AS LONG             ' old : gSticky
  autoSizeImg AS LONG             ' old : gImageAutoSize
  HANDLES AS LONG                 ' old : sz ' %use_HANDLESMALL ..  %use_HANDLELARGE
  backgrndType AS LONG            ' 0 = none, %use_backColor, %use_backBrush
  backColor AS LONG
  backBrush AS DWORD
  Include AS ASCIIZ * 255
END TYPE
TYPE st_types           ' used in WndClientProc as static
  rcDraw AS RECT
  rcDrag AS RECT
  nRect AS RECT
  Start_dnc AS POINTAPI           'draging
  Start_dwc AS POINTAPI
  ptLastPos AS POINTAPI           'For Sizing operation
  ptDragStart AS POINTAPI
  ptDragEnd AS POINTAPI
  ptDragLast AS POINTAPI
  ptDrawStart AS POINTAPI                 'For Drawing new control operation
  ptDrawEnd AS POINTAPI
  ptLastPos_dnc AS POINTAPI
END TYPE
'-------------------------------------------------------------------------------
' Listview control functions
'-------------------------------------------------------------------------------
TYPE dField
    fName     AS ASCIIZ * 80
    SIZE      AS INTEGER
    Alignment AS INTEGER
END TYPE

TYPE ListData
    Columns AS INTEGER
    Fields(0 TO 80) AS dField
END TYPE
GLOBAL hWndControl AS DWORD
