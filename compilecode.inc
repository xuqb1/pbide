' *****************************************************************************
' 枚举所有运行的应用程序，以查找PBWIN运行实例
' *****************************************************************************
FUNCTION SearchPbwinIde (BYVAL hWnd AS LONG, BYVAL lParam AS LONG) AS LONG
  LOCAL szTxt AS ASCIIZ * %MAX_PATH
  IF LEN(strCompilerTitle) = 0 THEN EXIT FUNCTION
  GetWindowText hWnd, szTxt, %MAX_PATH
  IF LEN(szTxt) THEN
    IF INSTR(UCASE$(szTxt), UCASE$(strCompilerTitle)) THEN
      hPbIde = hWnd
      EXIT FUNCTION
    END IF
  END IF
  FUNCTION = %TRUE
END FUNCTION
' *****************************************************************************
' 枚举所有运行的应用程序以确认PBCC是否有实例在运行
' *****************************************************************************
FUNCTION SearchPbccIde (BYVAL hWnd AS LONG, BYVAL lParam AS LONG) AS LONG
  LOCAL szTxt AS ASCIIZ * %MAX_PATH
  IF LEN(strCompilerTitle) = 0 THEN EXIT FUNCTION
  GetWindowText hWnd, szTxt, %MAX_PATH
  IF LEN(szTxt) AND INSTR(szTxt, strCompilerTitle) THEN
    hPbIde = hWnd
    EXIT FUNCTION
  END IF
  FUNCTION = %TRUE
END FUNCTION
' *****************************************************************************
' 使用PB IDE编译并调试
' 由于调试器集成在该IDE中，所以必须强制它加载要调试的文件，并激活编译并调试选项
' *****************************************************************************
SUB CompileAndDebug (BYVAL hWnd AS DWORD)
  LOCAL CpOpt   AS CompilerOptionsType  ' // 编译器选项
  LOCAL nLen    AS LONG                 ' // 缓存长度
  LOCAL Buffer  AS STRING               ' // 缓存
  LOCAL szPath  AS ASCIIZ * %MAX_PATH   ' // 文件路径
  LOCAL szDir   AS ASCIIZ * %MAX_PATH   ' // 文件夹
  LOCAL szText  AS ASCIIZ * 255         ' // 临时变量
  LOCAL hPbMenu AS DWORD                ' // PB IDE 菜单句柄
  LOCAL r       AS LONG                 ' // 结果代码
  LOCAL tmpStr  AS STRING

  LOCAL idPbRecentFiles AS LONG         ' // 最近文件指示
  LOCAL idPbCompileAndDebug AS LONG     ' // 编译并运行指示

  ' 获取当前文件路径
  szPath = ""
  GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
  IF LEN(szPath) = 0 THEN EXIT SUB

  ' 如果不是.BAS 文件，则退出
  tmpStr=szPath
  tmpStr=TRIM$(tmpStr)
  IF RIGHT$(tmpStr,1)="*" THEN
    tmpStr=MID$(tmpStr,1,LEN(tmpStr)-2)
  END IF
  tmpStr=TRIM$(tmpStr)
  IF RIGHT$(UCASE$(tmpStr),4)<>".BAS" THEN 'INSTR(UCASE$(szPath), ".BAS") = 0 THEN
    MessageBox(hWnd, "当前不是 .BAS 文件   ", _
            " 编译", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL)
    EXIT SUB
  END IF

  ' 如果需要，则保存
  IF ISTRUE SendMessage(GetEdit, %SCI_GETMODIFY, 0, 0) THEN
    IF ISFALSE SAVEFILEs(hWnd, %FALSE) THEN
      SetFocus GetEdit
      EXIT SUB
    END IF
    IF ISTRUE SendMessage(g_hToolbar, %TB_ISBUTTONENABLED, %IDM_SAVE, 0) THEN _
      SendMessage g_hToolbar, %TB_ENABLEBUTTON, %IDM_SAVE, %FALSE
    UpdateWindow g_hToolbar
    ' 重读标题 (路径可能会被修改)
    GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
  END IF
  ' 改变目录
  szDir = GetFilePath(szPath)
  IF szDir <> CURDIR$ THEN CHDIR szDir
  ' 获取编译器选项
  GetCompilerOptions CpOpt
  ' 检查是否已经在代码中定义了编译器
  nLen = SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
  Buffer = SPACE$(nLen + 1)
  SendMessage GetEdit, %SCI_GETTEXT, BYVAL LEN(Buffer), BYVAL STRPTR(Buffer)
  IF INSTR(Buffer, "SED_PBWIN") THEN
    IF CpOpt.DefaultCompiler <> 1 THEN
      CpOpt.DefaultCompiler = 1
      IniWrite g_zIni, "Compiler options", "DefaultCompiler", FORMAT$(CpOpt.DefaultCompiler)
      szText = "PBWIN"
      SendMessage g_hStatus, %SB_SETTEXT, 3, VARPTR(szText)
    END IF
  ELSEIF INSTR(Buffer, "SED_PBCC") THEN
    IF CpOpt.DefaultCompiler <> 2 THEN
      CpOpt.DefaultCompiler = 2
      IniWrite g_zIni, "Compiler options", "DefaultCompiler", FORMAT$(CpOpt.DefaultCompiler)
      szText = "PBCC"
      SendMessage g_hStatus, %SB_SETTEXT, 3, VARPTR(szText)
    END IF
  END IF

  ' 要激活的编译器标识
  IF CpOpt.DefaultCompiler = 1 THEN
    IF CpOpt.PBWINVersion = 0 THEN
      strCompilerTitle = $PBWINTITLE
      idPbRecentFiles = 3081
      idPbCompileAndDebug = 3115
    ELSEIF CpOpt.PBWINVersion = 1 THEN
      strCompilerTitle = $PBWIN80TITLE
      idPbRecentFiles = 33802
      idPbCompileAndDebug = 33827
    END IF
  ELSEIF CpOpt.DefaultCompiler = 2 THEN
    IF CpOpt.PBCCVersion = 0 THEN
      strCompilerTitle = $PBCCTITLE
      idPbRecentFiles = 3081
      idPbCompileAndDebug = 3115
    ELSEIF CpOpt.PBCCVersion = 1 THEN
      strCompilerTitle = $PBCC80TITLE
      idPbRecentFiles = 33802
      idPbCompileAndDebug = 33827
    END IF
  END IF
  ' 查找IDE运行实例
  hPbIde = 0
  IF CpOpt.DefaultCompiler = 1 THEN
    EnumWindows CODEPTR(SearchPbWinIde), 0
  ELSEIF CpOpt.DefaultCompiler = 2 THEN
    EnumWindows CODEPTR(SearchPbccIde), 0
  END IF
  ' 如果没找到，则启动该 IDE
  IF hPbIde = 0 THEN
    IF CpOpt.DefaultCompiler = 1 AND LEN(CpOpt.PBWINPath) > 0 THEN
      r = SHELL(GetFilePath(CpOpt.PBWINPath) & "PBEDIT.EXE", 1)
      SLEEP 2000   ' Take a breath
    ELSEIF CpOpt.DefaultCompiler = 2 AND LEN(CpOpt.PBCCPath) > 0 THEN
      r = SHELL(GetFilePath(CpOpt.PBCCPath) & "CCEDIT.EXE", 1)
      SLEEP 2000   ' Take a breath
    ELSE
      SED_MsgBox(hWnd, "请设置编译器路径  ", %MB_OK OR %MB_ICONINFORMATION, "编译器未找到")
      EXIT SUB
    END IF
  END IF
  ' 再次枚举
  IF CpOpt.DefaultCompiler = 1 THEN
    EnumWindows CODEPTR(SearchPbWinIde), 0
  ELSEIF CpOpt.DefaultCompiler = 2 THEN
    EnumWindows CODEPTR(SearchPbccIde), 0
  END IF
  ' 如果没找到 PB IDE 窗口，则退出
  IF hPbIde = 0 THEN
    SED_MsgBox(hWnd, "编译器未找到  ", %MB_OK OR %MB_ICONERROR, "错误")
    EXIT SUB
  END IF
  ' 调用 PB IDE 到最前面
  SetForegroundWindow hPbIde
  ' 获取 PB IDE 菜单
  hPbMenu = GetMenu(hPbIde)
  IF hPbMenu = 0 THEN EXIT SUB
  ' 试图修改用于重打开的第一个文件菜单
  r = ModifyMenu(hPbMenu, 3081, %MF_BYCOMMAND, idPbRecentFiles, "&1 " & szPath)
  ' 如果失败，则表明没有用于重打开的文件列表
  ' 因此，追加当前文件及路径到主菜单的末尾
  IF r = 0 THEN
    r = AppendMenu(hPbMenu, %MF_ENABLED, idPbRecentFiles, "&1 " & szPath)
    IF r = 0 THEN
      SED_MsgBox(hWnd, "无法修改 IDE 菜单  ", %MB_OK OR %MB_ICONERROR, "错误")
      EXIT SUB
    END IF
  END IF
  ' 强制 IDE 加载当前文件
  SendMessage hPbIde, %WM_COMMAND, idPbRecentFiles, 0
  ' 强制 IDE 编译并调试
  ' (编译并调试选项的标识为 3115)
  SendMessage hPbIde, %WM_COMMAND, idPbCompileAndDebug, 0
END SUB
' *********************************************************************************************
' 运行编译后的程序，而不重编译
' *********************************************************************************************
SUB ExecuteProgram (BYVAL hWnd AS DWORD)

  LOCAL szPath           AS ASCIIZ * %MAX_PATH    ' // Path
  LOCAL szText           AS ASCIIZ * 256          ' // General purpose variable
  LOCAL szActiveFileName AS ASCIIZ * %MAX_PATH    ' // File path
  LOCAL strCurDir        AS ASCIIZ * %MAX_PATH    ' // Current directory
  LOCAL szDir            AS ASCIIZ * %MAX_PATH    ' // Folder
  LOCAL buffer           AS STRING                ' // General purpose buffer
  LOCAL hr               AS LONG                  ' // Result code
  LOCAL x                AS LONG                  ' // Starting position of a word
  LOCAL strExeFile       AS STRING                ' // Compiled file path
  LOCAL nLin             AS LONG                  ' // Number of line
  LOCAL nLen             AS LONG                  ' // Length
  LOCAL strCommand       AS STRING                ' // Command line
  LOCAL p                AS LONG                  ' // General purpose variable

  ' Store active child name to restore after compiling
  GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
  szActiveFileName = szPath
  szPath = ""
  strCurDir = ""
  ' SED_GetPrimarySourceFile returns the path of the primary source file (if any)
  ' or the path of the current active child window
  IF (INSTR(szActiveFileName, ANY ":\/") <= 0) AND _
          (LEFT$(UCASE$(GetFileName(szActiveFileName)), 8) = "UNTITLED") THEN
    szPath = szActiveFilename
    strCurDir = CURDIR$
    IF ISFALSE SaveUntitledFile(hWnd, szPath) THEN EXIT SUB
  ELSE
    szPath = GetPrimarySourceFile(%TRUE)
    IF szPath = "" THEN szPath = szActiveFileName
  END IF
  IF LEN(szPath) = 0 THEN EXIT SUB
  IF INSTR(UCASE$(szPath), ".BAS") = 0 THEN    ' If it is not a .BAS file
    MessageBox(hWnd, "这不是 .BAS 文件   ", _
            " 运行", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL)
  ELSE
    ' 改变目录
    szDir = GetFilePath(szPath)
    IF szDir <> CURDIR$ THEN CHDIR szDir
    IF ISTRUE GetEdit OR LEN(sProjectPrimary) THEN
      IF LEN(szPath) THEN
        ' 搜索 #COMPILE ，以确定是否有指定的名字
        buffer = ""
        SendMessage GetEdit, %SCI_SETTARGETSTART, 0, 0
        SendMessage GetEdit, %SCI_SETTARGETEND, SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0), 0
        szText = "#COMPILE "
        hr = SendMessage(GetEdit, %SCI_SEARCHINTARGET, LEN(szText), VARPTR(szText))
        ' 确认 #COMPILE 不在注释行中
        WHILE CINT(hr) <> -1
          nLin = SendMessage(GetEdit, %SCI_LINEFROMPOSITION, hr, 0)
          nLen = SendMessage(GetEdit, %SCI_LINELENGTH, nLin, 0)
          buffer = SPACE$(nLen)
          SendMessage GetEdit, %SCI_GETLINE, nLin, STRPTR(buffer)
          x = INSTR(UCASE$(buffer),"#COMPILE ")
          IF INSTR(MID$(buffer,1,x - 1), "'") <> 0  AND INSTR(MID$(buffer,1,x - 1), "REM") <> 0 THEN
            nLin = SendMessage(GetEdit, %SCI_POSITIONFROMLINE, nLin + 1, 0)
            SendMessage GetEdit, %SCI_SETTARGETSTART, nLin, 0
            SendMessage GetEdit, %SCI_SETTARGETEND, SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0), 0
            hr = SendMessage(GetEdit, %SCI_SEARCHINTARGET, LEN(szText), VARPTR(szText))
          ELSE
            buffer = UCASE$(TRIM$(buffer, ANY CHR$(32, 9, 13, 10)))
            buffer = TRIM$(REMOVE$(buffer, "#COMPILE "))
            IF LEFT$(buffer, 3) = "EXE" THEN
              buffer = TRIM$(MID$(buffer, 4))
              ' #COMPILE EXE 后可能会有注释
              p = INSTR(buffer, "'")
              IF p THEN
                buffer = LEFT$(buffer, p - 1)
                buffer = TRIM$(buffer)
              END IF
              strExeFile = REMOVE$(buffer, $DQ)
              IF LEN(strExeFile) AND RIGHT$(strExeFile, 4) <> ".EXE" THEN
                strExeFile = GetFilePath(szPath) + strExeFile + ".EXE"
              END IF
              IF strExeFile = "" THEN strExeFile = UCASE$(szPath)
              REPLACE ".BAS" WITH ".EXE" IN strExeFile
            ELSEIF LEFT$(buffer, 3) = "DLL" THEN
              buffer = TRIM$(MID$(buffer, 4))
              ' #COMPILE DLL 后可能有注释
              p = INSTR(buffer, "'")
              IF p THEN
                buffer = LEFT$(buffer, p - 1)
                buffer = TRIM$(buffer)
              END IF
              strExeFile = REMOVE$(buffer, $DQ)
              IF LEN(strExeFile) AND RIGHT$(strExeFile, 4) <> ".DLL" THEN
                strExeFile = GetFilePath(szPath) + strExeFile + ".DLL"
              END IF
              IF strExeFile = "" THEN strExeFile = UCASE$(szPath)
              REPLACE ".BAS" WITH ".DLL" IN strExeFile
            END IF
            hr = -1
          END IF
        WEND
        IF strExeFile = "" THEN
          strExeFile = UCASE$(szPath)
          REPLACE ".BAS" WITH ".EXE" IN strExeFile
        END IF
        ' 可执行文件
        IF RIGHT$(strExeFile, 4) <> ".EXE" THEN
          MessageBox(BYVAL hWnd, "不能运行 " & RIGHT$(strExeFile, 4) & " 文件   ", "", _
                  %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
        ELSE
          IF ISFALSE FileExist(strExeFile) THEN
            MessageBox(BYVAL hWnd, "未找到文件： " & strExeFile, "", _
                    %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
          ELSE
            strCommand = strExeFile
            IF LEN(CommandLine) THEN strCommand = strCommand & " " & CommandLine
            hr = SHELL(strCommand, 1)
          END IF
        END IF
      ELSE
        MessageBox(BYVAL hWnd, "无法获取要执行文件的路径   ", "", _
                  %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
      END IF
    ELSE
      MessageBox(BYVAL hWnd, "无文件可运行   ", "", _
              %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
    END IF
    SetFocus GetEdit
    ' 恢复原来的目录
    IF LEN(strCurDir) THEN CHDIR strCurDir
  END IF
END SUB
' *********************************************************************************************

' *********************************************************************************************
' 编译BAS文件
' *********************************************************************************************
SUB CompileBasFile (BYVAL hWnd AS DWORD, BYVAL cOption AS DWORD)
  LOCAL szPath           AS ASCIIZ * %MAX_PATH    ' // Path
  LOCAL szText           AS ASCIIZ * 256          ' // General purpose variable
  LOCAL szActiveFileName AS ASCIIZ * %MAX_PATH    ' // File path
  LOCAL strCurDir        AS ASCIIZ * %MAX_PATH    ' // Current directory
  LOCAL szDir            AS ASCIIZ * %MAX_PATH    ' // Folder
  LOCAL buffer           AS STRING                ' // General purpose buffer
  LOCAL p                AS LONG                  ' // General purpose variable
  LOCAL p1               AS LONG                  ' // General purpose variable
  LOCAL p2               AS LONG                  ' // General purpose variable
  LOCAL strTxt           AS STRING                ' // General purpose string variable
  LOCAL nLin             AS LONG                  ' // Number of line
  LOCAL nCol             AS LONG                  ' // Number of column
  LOCAL nLen             AS LONG                  ' // Length
  LOCAL x                AS LONG                  ' // Starting position of a word
  LOCAL CpOpt            AS CompilerOptionsType   ' // Compiler options
  LOCAL strCommand       AS STRING                ' // Command line
  LOCAL strExeFile       AS STRING                ' // Compiled file path
  LOCAL hLogFile         AS LONG                  ' // Handle of the log file
  LOCAL strLogFile       AS STRING                ' // Text of the log file
  LOCAL strCompilerPath  AS STRING                ' // Compiler path
  LOCAL hr               AS LONG                  ' // Result code
  LOCAL strFileName      AS STRING                ' // File name
  LOCAL szFilePath       AS ASCIIZ * %MAX_PATH    ' // File path
  LOCAL i                AS LONG                  ' // Loop counter
  LOCAL hMdi             AS DWORD                 ' // Handle of the MDI window
  LOCAL nTab             AS LONG                  ' // Tab number
  LOCAL endPos           AS LONG                  ' // Ending position
  LOCAL nPos             AS LONG                  ' // Position
  LOCAL strExeFilePath   AS STRING                ' // Exe file path

  ' Store active child name to restore after compiling
  GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
  IF INSTR(UCASE$(szPath), ".RC") THEN          ' If it is a resource file
    CompileResource(hWnd, szPath)
    EXIT SUB
  END IF
  szActiveFileName = szPath
  szPath = ""
  strCurDir = ""
  ' SED_GetPrimarySourceFile returns the path of the primary source file (if any)
  ' or the path of the current active child window
  IF (INSTR(szActiveFileName, ANY ":\/") = 0) AND (LEFT$(UCASE$(GetFileName(szActiveFileName)), 8) = "UNTITLED") THEN
    szPath = szActiveFilename
    strCurDir = CURDIR$
    IF ISFALSE SaveUntitledFile(hWnd, szPath) THEN EXIT SUB
  ELSE
    szPath = GetPrimarySourceFile(%TRUE)
    IF szPath = "" THEN szPath = szActiveFileName
  END IF
  IF LEN(szPath) = 0 THEN EXIT SUB
  IF INSTR(UCASE$(szPath), ".BAS") = 0 THEN    ' If it is not a .BAS file
    MessageBox(hWnd, "这不是 .BAS 文件   ", _
            " 编译", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL)
  ELSE
    ' Refesh the toolbar
    UpdateWindow g_hToolbar
    ' Hourglass mouse
    MOUSEPTR 11
    ' Change the directory
    szDir = GetFilePath(szPath)
    IF szDir <> CURDIR$ THEN CHDIR szDir
    ' Read the compiler options
    GetCompilerOptions CpOpt
    ' Check if the user has defined the compiler in the code
    nLen = SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
    Buffer = SPACE$(nLen + 1)
    SendMessage GetEdit, %SCI_GETTEXT, BYVAL LEN(Buffer), BYVAL STRPTR(Buffer)
    IF INSTR(Buffer, "SED_PBWIN") THEN
      IF CpOpt.DefaultCompiler <> 1 THEN
        CpOpt.DefaultCompiler = 1
        IniWrite g_zIni, "Compiler options", "DefaultCompiler", FORMAT$(CpOpt.DefaultCompiler)
        szText = "PBWIN"
        SendMessage g_hStatus, %SB_SETTEXT, 3, VARPTR(szText)
      END IF
    ELSEIF INSTR(Buffer, "SED_PBCC") THEN
      IF CpOpt.DefaultCompiler <> 2 THEN
        CpOpt.DefaultCompiler = 2
        IniWrite g_zIni, "Compiler options", "DefaultCompiler", FORMAT$(CpOpt.DefaultCompiler)
        szText = "PBCC"
        SendMessage g_hStatus, %SB_SETTEXT, 3, VARPTR(szText)
      END IF
    END IF

    ' 准备命令行并编译程序
    IF ISTRUE GetEdit OR LEN(sProjectPrimary) THEN
      IF LEN(szPath) THEN
        IF CpOpt.DefaultCompiler = 1 OR CpOpt.DefaultCompiler = 2 THEN
          strCommand = ""
          IF CpOpt.DefaultCompiler = 1 AND LEN(CpOpt.PBWINPath) > 0 THEN
            strCommand = CpOpt.PBWINPath
            strCompilerPath = CpOpt.PBWINPath
            IF LEN(CpOpt.PBWINIncPath) THEN
              ' 使用双引号以处理较长的文件名
              strTxt = CpOpt.PBWINIncPath
              REPLACE ";" WITH $DQ & ";" & $DQ IN strTxt
              strTxt = $DQ & strTxt
              IF RIGHT$(strTxt, 1) <> $DQ THEN strTxt = strTxt & $DQ
              strCommand = strCommand & " /I" & strTxt
            END IF
          ELSEIF CpOpt.DefaultCompiler = 2 AND LEN(CpOpt.PBCCPath) > 0 THEN
            strCommand = CpOpt.PBCCPath
            strCompilerPath = CpOpt.PBCCPath
            IF LEN(CpOpt.PBCCIncPath) THEN
              ' 使用双引号以处理较长的文件名
              strTxt = CpOpt.PBCCIncPath
              REPLACE ";" WITH $DQ & ";" & $DQ IN strTxt
              strTxt = $DQ & strTxt
              IF RIGHT$(strTxt, 1) <> $DQ THEN strTxt = strTxt & $DQ
              strCommand = strCommand & " /I" & strTxt
            END IF
          END IF
          IF LEN(strCommand) THEN
            ' 查找 #COMPILE 以确认是否指定了文件名
            buffer = ""
            SendMessage GetEdit, %SCI_SETTARGETSTART, 0, 0
            SendMessage GetEdit, %SCI_SETTARGETEND, SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0), 0
            szText = "#COMPILE "
            hr = SendMessage(GetEdit, %SCI_SEARCHINTARGET, LEN(szText), VARPTR(szText))
            WHILE CINT(hr) <> -1
              nLin = SendMessage(GetEdit, %SCI_LINEFROMPOSITION, hr, 0)
              nLen = SendMessage(GetEdit, %SCI_LINELENGTH, nLin, 0)
              buffer = SPACE$(nLen)
              SendMessage GetEdit, %SCI_GETLINE, nLin, STRPTR(buffer)
              x = INSTR(UCASE$(buffer),"#COMPILE ")
              IF INSTR(MID$(buffer, 1, x - 1), "'") <> 0  AND INSTR(MID$(buffer, 1, x - 1), "REM") <> 0 THEN
                nLin = SendMessage(GetEdit, %SCI_POSITIONFROMLINE, nLin + 1, 0)
                SendMessage GetEdit, %SCI_SETTARGETSTART, nLin, 0
                SendMessage GetEdit, %SCI_SETTARGETEND, SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0), 0
                hr = SendMessage(GetEdit, %SCI_SEARCHINTARGET, LEN(szText), VARPTR(szText))
              ELSE
                buffer = UCASE$(TRIM$(buffer, ANY CHR$(32, 9, 13, 10)))
                buffer = TRIM$(REMOVE$(buffer, "#COMPILE "))
                IF LEFT$(buffer, 3) = "EXE" THEN
                  buffer = TRIM$(MID$(buffer, 4))
                  ' #COMPILE EXE 后可能会有注释
                  p = INSTR(buffer, "'")
                  IF p THEN
                    buffer = LEFT$(buffer, p - 1)
                    buffer = TRIM$(buffer)
                  END IF
                  strExeFile = REMOVE$(buffer, $DQ)
                  IF LEN(strExeFile) AND RIGHT$(strExeFile, 4) <> ".EXE" THEN
                    strExeFile = GetFilePath(szPath) + strExeFile + ".EXE"
                  END IF
                  IF strExeFile = "" THEN strExeFile = UCASE$(szPath)
                  REPLACE ".BAS" WITH ".EXE" IN strExeFile
                ELSEIF LEFT$(buffer, 3) = "DLL" THEN
                  buffer = TRIM$(MID$(buffer, 4))
                  ' #COMPILE DLL 后可能会有注释
                  p = INSTR(buffer, "'")
                  IF p THEN
                    buffer = LEFT$(buffer, p - 1)
                    buffer = TRIM$(buffer)
                  END IF
                  strExeFile = REMOVE$(buffer, $DQ)
                  IF LEN(strExeFile) AND RIGHT$(strExeFile, 4) <> ".DLL" THEN
                    strExeFile = GetFilePath(szPath) + strExeFile + ".DLL"
                  END IF
                  IF strExeFile = "" THEN strExeFile = UCASE$(szPath)
                  REPLACE ".BAS" WITH ".DLL" IN strExeFile
                END IF
                hr = -1
              END IF
            WEND
            IF strExeFile = "" THEN
              strExeFile = UCASE$(szPath)
              REPLACE ".BAS" WITH ".EXE" IN strExeFile
            END IF
            ' 使用双引号以处理较长的文件名
            strExeFile = $DQ & strExeFile & $DQ
            ' 如果文件名中包含空格, 则头尾加双引号
            strFileName = GetFileName(szPath)
            IF INSTR(strFileName, " ") THEN
              strFileName = GetFilePath(szFilePath) & $DQ & strFileName & $DQ
            END IF
            ' 不显示结果 - 我们会显示我们自己的消息
            strCommand = strCommand & " /Q"
            ' 创建日志文件
            strCommand = strCommand & " /L " & strFileName
            ' 运行前更新客户区
            UpdateWindow GetEdit
            ' 调用编译器
            SED_Shell hWnd, strCommand
            ' 结束时哔一下
            IF CpOpt.BeepOnCompletion = %BST_CHECKED THEN BEEP
            ' 读日志文件
            strLogFile = UCASE$(szPath)
            REPLACE ".BAS" WITH ".LOG" IN strLogFile
            ERRCLEAR
            hLogFile = FREEFILE
            OPEN strLogFile FOR BINARY AS hLogFile
            IF ERR THEN
              MessageBox(hWnd, "打开日志文件时出错:" & STR$(ERR), _
                      " 日志文件", %MB_OK OR %MB_ICONERROR OR %MB_APPLMODAL)
            ELSE
              GET$ hLogFile, LOF(hLogFile), buffer
              CLOSE hLogFile
            END IF
            ' Search for errors
            strTxt = ""
            p1 = INSTR(buffer, $CRLF & "Error ")
            ' Display results
            IF ISFALSE p1 THEN
              FOR i = LEN(buffer) TO 1 STEP -1
                IF MID$(buffer, i, 1) <> $CR AND MID$(buffer, i, 1) <> $LF AND MID$(buffer, i, 1) <> " " THEN
                  buffer = LEFT$(buffer, i)
                  EXIT FOR
                END IF
              NEXT
              SetWindowText hCompilersTB, BYCOPY buffer
              SendMessage hCompilersTB, %EM_SETSEL, LEN(buffer), LEN(buffer)
              SendMessage hCompilersTB, %EM_SCROLLCARET, 0, 0
              IF CpOpt.DisplayResults = %BST_CHECKED THEN SED_MsgBox hWnd, buffer
              ' Restore active child to original
              hMdi = GetWindow(g_hWndClient, %GW_CHILD)                 ' first look at already opened docs
              WHILE hMdi
                GetWindowText hMdi, szPath, %MAX_PATH
                IF UCASE$(szPath) = UCASE$(szActiveFileName) THEN    ' if already opened
                  ARRAY SCAN gTabFilePaths(), = szPath, TO nTab
                  IF nTab THEN
                    SendMessage g_hTabMdi, %TCM_SETCURSEL, nTab - 1, 0  ' activate the tab associated with the window
                    SendMessage g_hWndClient, %WM_MDIACTIVATE, hMdi, 0 ' activate it
                  END IF
                END IF
                hMdi = GetWindow(hMdi, %GW_HWNDNEXT)
              WEND
            END IF
            ' 设置焦点到编辑器控件上
            SetFocus GetEdit
            ' 显示错误
            IF p1 THEN
              strTxt = MID$(buffer, p1)
              p1 = INSTR(strTxt, "(")
              IF p1 THEN
                p2 = INSTR(p1 + 1, strTxt, ")")
                IF p2 THEN
                  strTxt = MID$(strTxt, p1 + 1, p2 - 1)
                  IF LEN(strTxt) THEN
                    p1 = INSTR(strTxt, ":")
                    IF p1 THEN
                      nLin = VAL(LEFT$(strTxt, p1 - 1))
                      nCol = VAL(MID$(strTxt, p1 + 1))
                      ' Allow for Error 496 Destination File Write Error
                      IF nLin <> 0 OR nCol <> 0 THEN
                        ' If it is an untitled file it must be in the active window
                        IF (INSTR(szActiveFileName, ANY ":\/") <> 0) AND (LEFT$(UCASE$(GetFileName(szActiveFileName)), 8) <> "UNTITLED") THEN
                          ' Search for filename in error line
                          p1 = INSTR(buffer, $CRLF & "Error ")
                          strTxt = MID$(buffer, p1)
                          strTxt = PARSE$(strTxt, "(", 1)       ' Get left side
                          strTxt = PARSE$(strTxt, " in ", 2)    ' Get right side
                          IF INSTR(strTxt, ANY ":\/") = 0 THEN strTxt = CURDIR$ & "\" & strTxt
                          strTxt = UCASE$(strTxt)
                          ' Save Current Child Window Handle
                          p = 0
                          p1 = GetWindow(g_hWndClient, %GW_CHILD)
                          p2 = p1
                          ' Find matching child window
                          DO WHILE p2 <> 0
                            GetWindowText p2, szPath, SIZEOF(szPath)
                            IF strTxt = UCASE$(szPath) THEN
                              p1 = p2         ' Switch to child handle on match
                              p = %TRUE
                              EXIT DO         ' and exit the loop
                            END IF
                            p2 = GetWindow(p2, %GW_HWNDNEXT)
                          LOOP
                          IF ISTRUE p THEN
                            ARRAY SCAN gTabFilePaths(), = szPath, TO nTab
                            IF nTab THEN
                              SendMessage g_hTabMdi, %TCM_SETCURSEL, nTab - 1, 0  ' activate the tab associated with the window
                              MdiActivate g_hWndClient, p1
                            END IF
                          ELSE        ' They did not match
                            IF INSTR(strTxt, "\") THEN        ' Path Available?
                              OpenThisFile(strTxt)
                            ELSE        ' No path Available - Use current path
                              OpenThisFile(CURDIR$ + "\" + strTxt)
                            END IF
                          END IF
                        END IF
                        ' End position = length of the document
                        endPos = SendMessage(GetEdit, %SCI_GETTEXTLENGTH, 0, 0)
                        ' Set pointer the end position of document
                        SendMessage GetEdit, %SCI_GOTOLINE, endPos, 0
                        'Go to Error position
                        nPos = SendMessage(GetEdit, %SCI_POSITIONFROMLINE, nLin - 1, 0)
                        nPos = nPos + nCol - 1
                        SendMessage GetEdit, %SCI_GOTOPOS, nPos, 0
                      END IF
                      FOR i = LEN(buffer) TO 1 STEP -1
                        IF MID$(buffer, i, 1) <> $CR AND MID$(buffer, i, 1) <> $LF AND MID$(buffer, i, 1) <> " " THEN
                          buffer = LEFT$(buffer, i)
                          EXIT FOR
                        END IF
                      NEXT
                      SetWindowText hCompilersTB, BYCOPY buffer
                      SendMessage hCompilersTB, %EM_SETSEL, LEN(buffer), LEN(buffer)
                      SendMessage hCompilersTB, %EM_SCROLLCARET, 0, 0
                      SED_MsgBox hWnd, buffer
                    END IF
                  END IF
                END IF
              END IF
            END IF
            ' Delete the log file
            IF CpOpt.DeleteLogFile = %BST_CHECKED THEN KILL strLogFile
            IF LEN(strTxt) = 0 THEN    ' If no errors...
              ' Call Semen's processor to add special comments to a resource file,
              ' compile this file with RC.EXE and attach the .RES file to the .EXE
              IF ISTRUE VAL(IniRead(g_zIni, "Editor options", "AttachResources", "")) THEN
                strExeFilePath = REMOVE$(strExeFile, $DQ)  ' Remove quotes
                AttachResources BYCOPY szActiveFileName, strExeFilePath, GetFilePath(strCompilerPath)
              END IF
              IF ISTRUE VAL(IniRead(g_zIni, "Editor options", "CallPostProcessor", "")) THEN
                ' Call the post-processor
                U_01 BYCOPY szActiveFilename, BYCOPY strExeFilePath
              END IF
            END IF
            ' Exe file
            IF cOption = %IDM_COMPILERUN THEN
              IF LEN(strTxt) = 0 THEN    ' If no errors run the exe
                strCommand = strExeFile
                IF LEN(CommandLine) THEN strCommand = strCommand & " " & CommandLine
                hr = SHELL(strCommand, 1)
              END IF
            END IF
          ELSE
            MessageBox(BYVAL hWnd, "未指定编译器   ", "", _
                    %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
          END IF
        ELSE
          MessageBox(BYVAL hWnd, "未指定默认编译器   ", "", _
                  %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
        END IF
      ELSE
        MessageBox(BYVAL hWnd, "无法获得要编译的文件路径   ", "", _
                %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
      END IF
    ELSE
      MessageBox(BYVAL hWnd, "无文件可运行   ", "", _
              %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
    END IF
    SetFocus GetEdit
    ' Arrrow mouse
    MOUSEPTR 1
    ' Restore old directory
    IF LEN(strCurDir) THEN CHDIR strCurDir
  END IF
END SUB
' *********************************************************************************************
' Compiles a resource file
' *********************************************************************************************
SUB CompileResource (BYVAL hWnd AS DWORD, szPath AS ASCIIZ)

  LOCAL CpOpt        AS CompilerOptionsType     ' // Compiler options
  LOCAL CmdLine      AS ASCIIZ * %MAX_PATH      ' // Command line
  LOCAL Result       AS STRING                  ' // Result string
  LOCAL ExitCode     AS LONG                    ' // Exit code
  LOCAL sa           AS SECURITY_ATTRIBUTES     ' // Security attributes
  LOCAL hReadPipe    AS DWORD                   ' // Read pipe handle
  LOCAL hWritePipe   AS DWORD                   ' // Write pipe handle
  LOCAL SI           AS STARTUPINFO             ' // Startup info structure
  LOCAL pi           AS PROCESS_INFORMATION     ' // Process information structure
  LOCAL BytesRead    AS DWORD                   ' // Bytes read
  LOCAL BytesWritten AS DWORD                   ' // Bytes written
  LOCAL chBuf        AS ASCIIZ * 1024           ' // Character buffer
  LOCAL chEOF        AS ASCIIZ * 12             ' // End of file
  LOCAL szDir        AS ASCIIZ * %MAX_PATH      ' // Directory
  LOCAL p1           AS LONG                    ' // Starting position
  LOCAL p2           AS LONG                    ' // Ending position
  LOCAL nLin         AS LONG                    ' // Line number
  LOCAL strTxt       AS STRING                  ' // General purpose variab馿
  LOCAL strCommand   AS STRING                  ' // Command line

  IF LEN(szPath) = 0 THEN EXIT SUB

  strTxt = UCASE$(szPath)
  REPLACE ".RC" WITH ".RES" IN strTxt
  IF FileExist(strTxt) THEN KILL strTxt

  IF ISTRUE SendMessage(GetEdit, %SCI_GETMODIFY, 0, 0) THEN
    IF ISFALSE SAVEFILEs(hWnd, %FALSE) THEN
      SetFocus GetEdit
      EXIT SUB
    END IF
    IF ISTRUE SendMessage(g_hToolbar, %TB_ISBUTTONENABLED, %IDM_SAVE, 0) THEN _
            SendMessage g_hToolbar, %TB_ENABLEBUTTON, %IDM_SAVE, %FALSE
    UpdateWindow g_hToolbar
    ' Reread the caption (the path can be different)
    GetWindowText MdiGetActive(g_hWndClient), szPath, SIZEOF(szPath)
  END IF

  szDir = UCASE$(GetFilePath(szPath))
  IF szDir <> UCASE$(CURDIR$) THEN CHDIR szDir

  ' Read the compiler options
  GetCompilerOptions CpOpt

  IF LEN(CpOpt.RCPath) THEN
    CmdLine =  CpOpt.RCPath & " /R"
  ELSE
    MessageBox(BYVAL hWnd, "未指定资源编译器路径   ", "", _
         %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
    EXIT SUB
  END IF

  IF LEN(CpOpt.RCIncPath) THEN
    CmdLine = CmdLine & " /I " & $DQ & CpOpt.RCIncPath & $DQ
  ELSE
    MessageBox(BYVAL hWnd, "未指定 resource.h 文件路径", "", _
         %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
    EXIT SUB
  END IF

  IF LEN(CpOpt.PBRESPath) = 0 THEN
    MessageBox(BYVAL hWnd, "未指定 PBRES.EXE 文件路径  ", "", _
         %MB_OK OR %MB_ICONINFORMATION OR %MB_APPLMODAL)
    EXIT SUB
  END IF

  ' Update the client area of the control before calling CreateProcess
  UpdateWindow GetEdit

  sa.nLength = SIZEOF(SECURITY_ATTRIBUTES)
  sa.bInheritHandle = 1 ' TRUE

  IF CreatePipe(hReadPipe, hWritePipe, sa, BYVAL 0) = 0 THEN EXIT SUB

  SI.cb           = SIZEOF(STARTUPINFO)
  SI.dwFlags      = %STARTF_USESHOWWINDOW OR %STARTF_USESTDHANDLES
  SI.wShowWindow  = %SW_HIDE
  SI.hStdOutput   = hWritePipe
  SI.hStdError    = hWritePipe

  CmdLine = CmdLine & " " & $DQ & szPath & $DQ

  IF CreateProcess ("", CmdLine, BYVAL 0&, BYVAL 0&, 1,  %NORMAL_PRIORITY_CLASS, _
         BYVAL 0&, "", si, pi) = 0 THEN EXIT SUB

  WaitForSingleObject PI.hProcess, 15000
  GetExitCodeProcess PI.hProcess, ExitCode
  IF ExitCode = %STILL_ACTIVE THEN TerminateProcess pi.hProcess, 0

  chEOF = "End Of Pipe"
  Result = ""

  DO
    IF WriteFile (hWritePipe, chEOF, LEN(chEOF), BytesWritten, BYVAL 0) = 0 THEN  EXIT DO
    IF BytesWritten <> LEN(chEOF) THEN  EXIT DO
    DO
      IF ReadFile(hReadPipe, chBuf, SIZEOF(chBuf), BytesRead, BYVAL 0) = 0 THEN EXIT DO
      IF BytesRead = 0 THEN EXIT DO
      Result = Result + LEFT$(chBuf, BytesRead)
      IF RIGHT$(Result, LEN(chEOF)) = chEOF THEN Result$ = LEFT$(Result$, LEN(Result) - LEN(chEOF)): EXIT DO
    LOOP
    EXIT DO
  LOOP

  CloseHandle hReadPipe
  CloseHandle hWritePipe
  CloseHandle pi.hThread
  CloseHandle pi.hProcess

  UpdateWindow GetEdit
  IF LEN(Result) THEN
    SED_MsgBox hWnd, TRIM$(Result, ANY CHR$(32, 0, 13, 10))
    p1 = INSTR(Result, "(")
    IF p1 THEN
      p2 = INSTR(p1 + 1, Result, ")")
      nLin = VAL(MID$(Result, p1 + 1 , p2 - 1))
    END IF
    SendMessage GetEdit, %SCI_GOTOLINE, 0 , 0
    SendMessage GetEdit, %SCI_GOTOLINE, nLin - 1 , 0
  ELSE ' call PBRes.exe
    strTxt = UCASE$(szPath)
    REPLACE ".RC" WITH ".RES" IN strTxt
    IF FileExist(strTxt) THEN
      strCommand = CpOpt.PBRESPath & " " & strTxt
      ' Update the client area of the control before shelling
      UpdateWindow GetEdit
      SHELL strCommand, 0
      IF CpOpt.BeepOnCompletion = %BST_CHECKED THEN BEEP
      SED_MsgBox hWnd, "Resource file compiled"
    ELSE
      SED_MsgBox hWnd, "未找到 " & strTxt & $CRLF & " 编译失败"
    END IF
  END IF
END SUB
SUB SED_Shell (BYVAL hWnd AS DWORD, BYVAL CmdLine AS STRING)
  LOCAL Si AS STARTUPINFO
  LOCAL Pi AS PROCESS_INFORMATION
  LOCAL dwError AS DWORD
  LOCAL vi AS OSVERSIONINFO
  vi.dwOsVersionInfoSize = SIZEOF(vi)
  GetVersionEx vi
  IF (vi.dwPlatformId = %VER_PLATFORM_WIN32_NT) THEN
    Si.cb = SIZEOF(Si)
    IF CreateProcess("", BYVAL STRPTR (CmdLine), BYVAL %NULL, BYVAL %NULL, 0, _
             %NORMAL_PRIORITY_CLASS OR %CREATE_SEPARATE_WOW_VDM, _
             BYVAL %NULL, BYVAL %NULL, Si, Pi) THEN
      CALL WaitForSingleObject(pi.hProcess, %INFINITE)
      CALL CloseHandle(pi.hProcess)
      CALL CloseHandle(pi.hThread)
    ELSE
      'SED_Msgbox hWnd, "Can't start the compiler"
      dwError = GetLastError
      SED_Msgbox hWnd, "错误:" & STR$(dwError) & " [&H" & HEX$(dwError) & "] " _
             & SED_WinErrorMsg(dwError), 0, "SED_Shell"
    END IF
  ELSE
    SHELL CmdLine, 0
  END IF
END SUB
' *********************************************************************************************
' Returns the text for a given Windows error code
' (see GetLastError() in your Windows API documentation).
' *********************************************************************************************
FUNCTION SED_WinErrorMsg (BYVAL dError AS DWORD) AS STRING
  LOCAL pBuffer   AS ASCIIZ PTR
  LOCAL ncbBuffer AS DWORD
  LOCAL sText     AS STRING
  ncbBuffer = FormatMessage(%FORMAT_MESSAGE_ALLOCATE_BUFFER _
                   OR %FORMAT_MESSAGE_FROM_SYSTEM _
                   OR %FORMAT_MESSAGE_IGNORE_INSERTS, _
                   BYVAL %NULL, _
                   dError, _
                   BYVAL MAKELANGID(%LANG_NEUTRAL, %SUBLANG_DEFAULT), _
                   BYVAL VARPTR(pBuffer), _
                   0, _
                   BYVAL %NULL)
  IF ncbBuffer THEN
    sText = PEEK$(pBuffer, ncbBuffer)
    sText = REMOVE$ (sText, ANY CHR$(13) + CHR$(10))
    LocalFree pBuffer
  ELSE
    sText = "未知错误: "+ HEX$(dError, 8)
  END IF
  FUNCTION = sText
END FUNCTION


   '=================================================================
' Ansi转Unicode
MACRO FUNCTION AnsiUnicode (CodePage, AnsiTxt)
  MACROTEMP WideCharStr, l
  DIM l AS DWORD, WideCharStr AS STRING
  l = LEN(AnsiTxt) + 1
  WideCharStr = STRING$(l + l, 0)
  MultiByteToWideChar CodePage, 0, BYVAL STRPTR(AnsiTxt), l, _
          BYVAL STRPTR(WideCharStr), l
END MACRO = WideCharStr
' 读取exe文件内容到字符串szExeData中，返回exe类型
FUNCTION ReadExeFile (szExeFileName AS STRING, szExeData AS STRING, RsRcPresent AS DWORD) AS LONG
  LOCAL f                     AS DWORD
  LOCAL i                     AS DWORD
  LOCAL lpImageDosHeader      AS IMAGE_DOS_HEADER PTR
  LOCAL lpImageNtHeaders      AS IMAGE_NT_HEADERS PTR
  LOCAL lpImageSectionHeader  AS IMAGE_SECTION_HEADER PTR
  ERRCLEAR
  f = FREEFILE
  OPEN szExeFileName FOR BINARY LOCK WRITE AS #f
  IF ERR = 0 THEN
    GET$ #f, LOF(f), szExeData
  END IF
  CLOSE #f
  IF ERR THEN
    FUNCTION = 1
    EXIT FUNCTION
  END IF
  lpImageDosHeader = STRPTR(szExeData)
  IF LEN(szExeData) < SIZEOF(IMAGE_DOS_HEADER) THEN
    FUNCTION = 2
    EXIT FUNCTION
  END IF
  IF @lpImageDosHeader.e_magic <> %IMAGE_DOS_SIGNATURE THEN
    FUNCTION = 2
    EXIT FUNCTION
  END IF
  IF LEN(szExeData) < @lpImageDosHeader.e_lfanew + SIZEOF(IMAGE_NT_HEADERS) THEN
    FUNCTION = 2
    EXIT FUNCTION
  END IF
  lpImageNtHeaders = lpImageDosHeader + @lpImageDosHeader.e_lfanew
  IF @lpImageNtHeaders.Signature <> %IMAGE_NT_SIGNATURE THEN
    FUNCTION = 2
    EXIT FUNCTION
  END IF
  IF @lpImageNtHeaders.FileHeader.SizeOfOptionalHeader <> SIZEOF(@lpImageNtHeaders.OptionalHeader) OR _
          @lpImageNtHeaders.OptionalHeader.Magic <> %IMAGE_NT_OPTIONAL_HDR32_MAGIC THEN
    FUNCTION = 2
    EXIT FUNCTION
  END IF
  IF @lpImageNtHeaders.FileHeader.NumberOfSections < 1 THEN
    FUNCTION = 2
    EXIT FUNCTION
  END IF
  lpImageDosHeader     = STRPTR(szExeData)
  lpImageNtHeaders     = lpImageDosHeader + @lpImageDosHeader.e_lfanew
  lpImageSectionHeader = lpImageNtHeaders + SIZEOF(IMAGE_NT_HEADERS)

  IF @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_RESOURCE).VirtualAddress <> 0 THEN
    RsRcPresent = 1
  ELSE
    RsRcPresent = 0
  END IF
END FUNCTION

' 增加数据到Exe文件末尾
FUNCTION AddDataSectionToExe (szExeData AS STRING, szNewSectionName AS STRING, szNewSectionData AS STRING, BYVAL SectionTp AS DWORD) AS LONG
  DIM lpImageDosHeader          AS LOCAL IMAGE_DOS_HEADER PTR
  DIM lpImageNtHeaders          AS LOCAL IMAGE_NT_HEADERS PTR
  DIM lpImageSectionHeader      AS LOCAL IMAGE_SECTION_HEADER PTR
  DIM szSectionFileImage()      AS LOCAL STRING
  DIM nSections                 AS LOCAL DWORD
  DIM FileAlignment             AS LOCAL DWORD
  DIM SectionAlignment          AS LOCAL DWORD
  DIM SizeOfSectionOld          AS LOCAL DWORD
  DIM SizeOfSectionNew          AS LOCAL DWORD
  DIM i                         AS LOCAL DWORD
  DIM j                         AS LOCAL DWORD
  DIM k                         AS LOCAL DWORD
  DIM m                         AS LOCAL DWORD
  lpImageDosHeader     = STRPTR(szExeData)
  lpImageNtHeaders     = lpImageDosHeader + @lpImageDosHeader.e_lfanew
  lpImageSectionHeader = lpImageNtHeaders + SIZEOF(IMAGE_NT_HEADERS)

  nSections            = @lpImageNtHeaders.FileHeader.NumberOfSections
  FileAlignment        = @lpImageNtHeaders.OptionalHeader.FileAlignment
  SectionAlignment     = @lpImageNtHeaders.OptionalHeader.SectionAlignment

  IF (nSections < 1) OR (FileAlignment < 1) OR (SectionAlignment < 1) THEN FUNCTION = 1: EXIT FUNCTION

  REDIM szSectionFileImage(nSections + 1)

  FOR i = 0 TO nSections - 1
    FOR j = i + 1 TO nSections - 1
       IF @lpImageSectionHeader[j].VirtualAddress  < @lpImageSectionHeader[i].VirtualAddress THEN _
                SWAP @lpImageSectionHeader[j], @lpImageSectionHeader[i]
    NEXT
    SizeOfSectionOld = @lpImageSectionHeader[i].SizeOfRawData
    SizeOfSectionNew = FileAlignment * ((SizeOfSectionOld + FileAlignment - 1) \ FileAlignment)
    szSectionFileImage(i + 1) = MID$(szExeData, @lpImageSectionHeader[i].PointerToRawData + 1, SizeOfSectionOld) + _
       STRING$(SizeOfSectionNew - SizeOfSectionOld, 0)
    IF LEN(szSectionFileImage(i + 1)) <> SizeOfSectionNew THEN FUNCTION = 1: EXIT FUNCTION
  NEXT

  SizeOfSectionOld = LEN(szNewSectionData)
  SizeOfSectionNew = FileAlignment * ((SizeOfSectionOld + FileAlignment - 1) \ FileAlignment)
  szSectionFileImage(nSections + 1) = szNewSectionData + STRING$(SizeOfSectionNew - SizeOfSectionOld, 0)
  IF LEN(szSectionFileImage(nSections + 1)) <> SizeOfSectionNew THEN FUNCTION = 1: EXIT FUNCTION

  SizeOfSectionOld = lpImageSectionHeader - lpImageDosHeader + SIZEOF(IMAGE_SECTION_HEADER) * nSections
  SizeOfSectionNew = FileAlignment * ((SizeOfSectionOld + SIZEOF(IMAGE_SECTION_HEADER) + FileAlignment - 1) \ FileAlignment)

  szSectionFileImage(0) = LEFT$(szExeData, SizeOfSectionOld) + STRING$(SizeOfSectionNew - SizeOfSectionOld, 0)
  IF LEN(szSectionFileImage(0)) <> SizeOfSectionNew THEN FUNCTION = 1: EXIT FUNCTION

  INCR nSections
  k = 0
  FOR i = 0 TO nSections
    k = k + LEN(szSectionFileImage(i))
  NEXT i
  szExeData = STRING$(k, 0)
  k = STRPTR(szExeData)
  FOR i = 0 TO nSections
    MoveMemory BYVAL k, BYVAL STRPTR(szSectionFileImage(i)), LEN(szSectionFileImage(i))
    k = k + LEN(szSectionFileImage(i))
  NEXT

  lpImageDosHeader     = STRPTR(szExeData)
  lpImageNtHeaders     = lpImageDosHeader + @lpImageDosHeader.e_lfanew
  lpImageSectionHeader = lpImageNtHeaders + SIZEOF(IMAGE_NT_HEADERS)

  @lpImageSectionHeader[nSections - 1].NAME =  LEFT$(szNewSectionName + STRING$(8, 0), 8)
  @lpImageSectionHeader[nSections - 1].Misc.VirtualSize = LEN(szNewSectionData)
  @lpImageSectionHeader[nSections - 1].Characteristics   = &H40000040???

  k = LEN(szSectionFileImage(0))
  @lpImageNtHeaders.FileHeader.NumberOfSections = nSections
  @lpImageNtHeaders.OptionalHeader.SizeOfHeaders = k
  m = k
  FOR i = 0 TO nSections - 1
    @lpImageSectionHeader[i].PointerToRawData = k
    j = LEN(szSectionFileImage(i + 1))
    @lpImageSectionHeader[i].SizeOfRawData = j
    k = k + j
    m = SectionAlignment * ((m + SectionAlignment - 1) \ SectionAlignment)
    IF i = nSections - 1 THEN
      @lpImageSectionHeader[i].VirtualAddress = m
      IF SectionTp = 1 THEN
        @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_RESOURCE).VirtualAddress = @lpImageSectionHeader[i].VirtualAddress
        @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_RESOURCE).nSize = @lpImageSectionHeader[i].Misc.VirtualSize
      END IF
    END IF
    IF m > @lpImageSectionHeader[i].VirtualAddress THEN FUNCTION = 1: EXIT FUNCTION
    IF (@lpImageSectionHeader[i].VirtualAddress MOD SectionAlignment) THEN FUNCTION = 1: EXIT FUNCTION
    m = @lpImageSectionHeader[i].VirtualAddress + @lpImageSectionHeader[i].Misc.VirtualSize
  NEXT
  @lpImageNtHeaders.OptionalHeader.SizeOfImage = SectionAlignment * ((m + SectionAlignment - 1) \ SectionAlignment)
  @lpImageNtHeaders.OptionalHeader.SizeOfInitializedData = @lpImageNtHeaders.OptionalHeader.SizeOfInitializedData + _
     @lpImageSectionHeader[nSections - 1].Misc.VirtualSize
END FUNCTION
' 增加空res部分
FUNCTION AddDummyResSection (szExeFileNameIn AS STRING, szExeFileNameOut AS STRING) AS LONG
  DIM szExeData                 AS LOCAL STRING
  DIM szSectionData             AS LOCAL STRING
  DIM RsRcPresent               AS LOCAL DWORD
  DIM f                         AS LOCAL DWORD
  IF ReadExeFile (szExeFileNameIn, szExeData, RsRcPresent) THEN FUNCTION = -1: EXIT FUNCTION

  IF RsRcPresent = 0 THEN
    szSectionData = STRING$(16, 0)
    IF AddDataSectionToExe (szExeData, ".rsrc", szSectionData, 1) THEN FUNCTION = -2: EXIT FUNCTION
    ERR = 0: f = FREEFILE: OPEN szExeFileNameOut FOR OUTPUT AS #f
    IF ERR = 0 THEN PRINT #f, szExeData;
    CLOSE #f: IF ERR THEN FUNCTION = -4: EXIT FUNCTION
  ELSEIF szExeFileNameIn <> szExeFileNameOut THEN
    IF CopyFile (BYVAL STRPTR(szExeFileNameIn), BYVAL STRPTR(szExeFileNameOut), 0) = 0 THEN _
       FUNCTION = -4: EXIT FUNCTION
  END IF
END FUNCTION
' 增加res到exe文件
FUNCTION AddResToExe (szExeFileNameIn AS STRING, szExeFileNameOut AS STRING, szResFileName AS STRING, BYVAL DeleteExisting AS DWORD, szUnicowsDirectory AS STRING) AS LONG
  DIM hLib                      AS STATIC DWORD
  DIM hBeginUpdateResourceW     AS STATIC DWORD
  DIM hUpdateResourceW          AS STATIC DWORD
  DIM hEndUpdateResourceW       AS STATIC DWORD
  DIM szTmp                     AS LOCAL STRING
  DIM szTmpU                    AS LOCAL STRING
  DIM hUpdateRes                AS LOCAL DWORD
  DIM IniPos                    AS LOCAL DWORD
  DIM CurPos                    AS LOCAL DWORD PTR
  DIM EndPos                    AS LOCAL DWORD
  DIM lpResType                 AS LOCAL WORD PTR
  DIM lpResName                 AS LOCAL WORD PTR
  DIM lpDataVersion             AS LOCAL DWORD PTR
  DIM lpCodePage                AS LOCAL WORD PTR
  DIM ResTypeIdAddr             AS LOCAL DWORD
  DIM ResNameIdAddr             AS LOCAL DWORD
  DIM f                         AS LOCAL DWORD
  DIM i                         AS LOCAL DWORD
  '------------------------------------------------------------------------------
  IF hLib = 0 THEN
    hLib = GetModuleHandle("Unicows.dll")
    IF hLib = 0 THEN hLib = LoadLibrary("Unicows.dll")
    IF hLib = 0 THEN hLib = GetModuleHandle("Kernel32.dll")
    IF (hLib = 0) AND (szUnicowsDirectory <> "") THEN
      IF RIGHT$(szUnicowsDirectory, 1) = "\" THEN hLib = LoadLibrary(szUnicowsDirectory + "Unicows.dll") ELSE _
         hLib = LoadLibrary(szUnicowsDirectory + "\Unicows.dll")
    END IF
  END IF
  IF hBeginUpdateResourceW = 0 THEN hBeginUpdateResourceW = GetProcAddress(hLib, "BeginUpdateResourceW")
  IF hUpdateResourceW      = 0 THEN hUpdateResourceW      = GetProcAddress(hLib, "UpdateResourceW")
  IF hEndUpdateResourceW   = 0 THEN hEndUpdateResourceW   = GetProcAddress(hLib, "EndUpdateResourceW")
  '------------------------------------------------------------------------------
  IF AddDummyResSection (szExeFileNameIn, szExeFileNameOut) THEN FUNCTION = -1: EXIT FUNCTION
  f = FREEFILE: ERRCLEAR
  OPEN szResFileName FOR BINARY SHARED AS #f
  IF ERR = 0 THEN GET$ #f, LOF(f), szTmp
  CLOSE #f
  IF LEFT$(szTmp, 32) <> CHR$(0, 0, 0, 0, 32, 0, 0, 0, 255, 255, 0, 0, 255, 255, 0, 0) + STRING$(16, 0) THEN _
     FUNCTION = 1: EXIT FUNCTION ' Resource file is not found
  szTmpU = AnsiUnicode(GetACP, szExeFileNameOut)
  IF hBeginUpdateResourceW = 0 THEN hUpdateRes = 0 ELSE _
     CALL DWORD hBeginUpdateResourceW USING BeginUpdateResource (BYVAL STRPTR(szTmpU), DeleteExisting) TO hUpdateRes
  IF hUpdateRes = 0 THEN FUNCTION = 1: EXIT FUNCTION ' Could not open file for writing
  IniPos = STRPTR(szTmp)
  CurPos = IniPos + 32
  EndPos = IniPos + LEN(szTmp)
  DO
    IF CurPos >= EndPos THEN EXIT DO
    lpResType = CurPos + 8
    IF @lpResType = &HFFFF?? THEN
      ResTypeIdAddr = @lpResType[1]
      INCR lpResType
    ELSE
      ResTypeIdAddr = lpResType
      WHILE @lpResType <> 0: INCR lpResType: WEND
    END IF
    lpResName = lpResType + 2
    IF @lpResName = &HFFFF?? THEN
      ResNameIdAddr = @lpResName[1]
      INCR lpResName
    ELSE
      ResNameIdAddr = lpResName
      WHILE @lpResName <> 0: INCR lpResName: WEND
    END IF
    lpDataVersion = lpResName + 2
    IF ((lpDataVersion - CurPos) MOD 4) <> 0 THEN INCR lpDataVersion
    lpCodePage = lpDataVersion + 6
    IF hUpdateResourceW = 0 THEN i = 0 ELSE _
      CALL DWORD hUpdateResourceW USING UpdateResource(hUpdateRes, BYVAL ResTypeIdAddr, BYVAL ResNameIdAddr, _
         @lpCodePage, BYVAL CurPos + @CurPos[1], BYVAL @CurPos) TO i
    IF i = 0 THEN FUNCTION = 1: EXIT FUNCTION ' Could not add resource
    @CurPos = FIX((@CurPos + 3) / 4) * 4
    CurPos = CurPos + @CurPos[1] + @CurPos
  LOOP
  IF hEndUpdateResourceW = 0 THEN i = 0 ELSE _
    CALL DWORD hEndUpdateResourceW USING EndUpdateResource (hUpdateRes, 0) TO i
  IF i = 0 THEN FUNCTION = 1: EXIT FUNCTION ' Could not write changes to file
END FUNCTION

MACRO ReadPeFile
  f = FREEFILE: ERR = 0
  OPEN szPePath FOR BINARY SHARED AS #f
  IF ERR = 0 THEN GET$ #f, LOF(f), szPeFileData
  CLOSE #f: IF ERR THEN EXIT DO

  ' Is this 'correct' PE ?  (assumes 16 sections in optional PE-header)
  lpImageDosHeader = STRPTR(szPeFileData)
  IF LEN(szPeFileData) < SIZEOF(IMAGE_DOS_HEADER) THEN ERR = 1: EXIT DO
  IF @lpImageDosHeader.e_magic <> %IMAGE_DOS_SIGNATURE THEN ERR = 1: EXIT DO
  IF LEN(szPeFileData) < @lpImageDosHeader.e_lfanew + SIZEOF(IMAGE_NT_HEADERS) THEN ERR = 1: EXIT DO
  lpImageNtHeaders = lpImageDosHeader + @lpImageDosHeader.e_lfanew
  IF @lpImageNtHeaders.Signature <> %IMAGE_NT_SIGNATURE THEN ERR = 1: EXIT DO
  IF @lpImageNtHeaders.FileHeader.SizeOfOptionalHeader <> SIZEOF(@lpImageNtHeaders.OptionalHeader) OR _
     @lpImageNtHeaders.OptionalHeader.Magic <> %IMAGE_NT_OPTIONAL_HDR32_MAGIC THEN ERR = 2: EXIT DO
END MACRO

MACRO WritePeFile
  f = FREEFILE
  OPEN szPePath FOR OUTPUT AS f
  IF ERR = 0 THEN PRINT #f, szPeFileData;
  CLOSE #f
END MACRO

'=================================================================
FUNCTION MarkAsConsole (szPePath AS STRING, ConvertToConsole AS DWORD) AS DWORD
  DIM szPeFileData                AS LOCAL STRING
  DIM lpImageDosHeader            AS LOCAL IMAGE_DOS_HEADER PTR
  DIM lpImageNtHeaders            AS LOCAL IMAGE_NT_HEADERS PTR
  DIM f                           AS LOCAL DWORD
  DIM i                           AS LOCAL DWORD
  DO
    IF ConvertToConsole = 0 THEN EXIT FUNCTION
    ReadPeFile
    @lpImageNtHeaders.OptionalHeader.SubSystem = ConvertToConsole: EXIT DO
  LOOP
  IF ERR THEN MSGBOX "Can't process SUBSYSTEM CONSOLE for " + szPePath, _
          %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle: EXIT FUNCTION
  WritePeFile
  IF ERR THEN MSGBOX "Can't save " + szPePath, %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle: EXIT FUNCTION
  ConvertToConsole = 0: FUNCTION = 1 ' successful
END FUNCTION

'========================= DELAYLOAD ==============================

'   TYPE IMAGE_IMPORT_DESCRIPTOR
'      OriginalFirstThunk              AS DWORD
'      TimeDateStamp                   AS DWORD
'      ForwarderChain                  AS DWORD
'      pName                           AS DWORD
'      FirstThunk                      AS DWORD
'   END TYPE

FUNCTION DelayLoad (szPePath AS STRING, szDelayLoadDlls() AS STRING, BYVAL nDelayLoadDlls AS DWORD, ConvertToConsole AS DWORD) AS DWORD
  DIM szPeFileData                AS LOCAL STRING
  DIM lpImageDosHeader            AS LOCAL IMAGE_DOS_HEADER PTR
  DIM lpImageNtHeaders            AS LOCAL IMAGE_NT_HEADERS PTR
  DIM lpImageSectionHeader        AS LOCAL IMAGE_SECTION_HEADER PTR
  DIM lpImageImportDescriptor     AS LOCAL IMAGE_IMPORT_DESCRIPTOR PTR
  DIM lpImageImportDescriptorNew  AS LOCAL IMAGE_IMPORT_DESCRIPTOR PTR
  DIM lpDllName                   AS LOCAL ASCIIZ PTR
  DIM szDllName                   AS LOCAL STRING
  DIM lpFuncNameRef               AS LOCAL DWORD PTR
  DIM lpFuncAddr                  AS LOCAL DWORD PTR
  DIM VirtualAddressCorrection    AS LOCAL DWORD
  DIM f                           AS LOCAL DWORD
  DIM i                           AS LOCAL DWORD
  DO
    IF nDelayLoadDlls = 0 THEN EXIT FUNCTION
    ReadPeFile
    lpImageSectionHeader = lpImageNtHeaders + SIZEOF(IMAGE_NT_HEADERS)
    @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_COPYRIGHT).VirtualAddress = _
            @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress
    lpImageImportDescriptor = @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress
    IF ConvertToConsole THEN @lpImageNtHeaders.OptionalHeader.SubSystem = ConvertToConsole

    FOR i = 0 TO @lpImageNtHeaders.FileHeader.NumberOfSections - 1
      IF (lpImageImportDescriptor >= @lpImageSectionHeader[i].VirtualAddress) AND _
                 (lpImageImportDescriptor <= (@lpImageSectionHeader[i].VirtualAddress + @lpImageSectionHeader[i].SizeOfRawData)) THEN _
                 VirtualAddressCorrection = lpImageDosHeader + @lpImageSectionHeader[i].PointerToRawData - @lpImageSectionHeader[i].VirtualAddress: EXIT FOR
    NEXT
    IF VirtualAddressCorrection = 0 THEN ERR = 2: EXIT DO
    lpImageImportDescriptor = lpImageImportDescriptor + VirtualAddressCorrection
    lpImageImportDescriptorNew = lpImageImportDescriptor
    WHILE @lpImageImportDescriptor.OriginalFirstThunk <> 0 ' Dll
      lpDllName = @lpImageImportDescriptor.Name + VirtualAddressCorrection
      szDllName = UCASE$(@lpDllName)
      FOR i = 1 TO nDelayLoadDlls
        IF szDelayLoadDlls(i - 1) = szDllName THEN
          ' MsgBox szDllName + " was found and will be hidden", %MB_ICONINFORMATION Or %MB_TASKMODAL, $MsgTitle
          SWAP @lpImageImportDescriptor, @lpImageImportDescriptorNew
          INCR lpImageImportDescriptorNew
          @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress = _
             @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress + SIZEOF(IMAGE_IMPORT_DESCRIPTOR)
          @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_IMPORT).nSize = _
             @lpImageNtHeaders.OptionalHeader.DataDirectory(%IMAGE_DIRECTORY_ENTRY_IMPORT).nSize - SIZEOF(IMAGE_IMPORT_DESCRIPTOR)
          EXIT FOR
        END IF
      NEXT
      INCR lpImageImportDescriptor
    WEND
    EXIT DO
  LOOP
  IF ERR THEN MSGBOX "Can't process DELAY LOAD for " + szPePath, %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle: EXIT FUNCTION
  WritePeFile
  IF ERR THEN MSGBOX "Can't save " + szPePath, %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle: EXIT FUNCTION
  ConvertToConsole = 0: FUNCTION = 1 ' successful
END FUNCTION
'===================================================================================
FUNCTION AttachFileAfterEof(szPePath AS STRING, szAttachFileLegend() AS STRING, szAttachFilePath() AS STRING, BYVAL nAttachFiles AS DWORD) AS DWORD
  DIM szPeFileData                AS LOCAL STRING
  DIM lpImageDosHeader            AS LOCAL IMAGE_DOS_HEADER PTR
  DIM lpImageNtHeaders            AS LOCAL IMAGE_NT_HEADERS PTR
  DIM lpImageSectionHeader        AS LOCAL IMAGE_SECTION_HEADER PTR
  DIM szAttachFileData()          AS LOCAL STRING
  DIM szLegends                   AS LOCAL STRING
  DIM OffsetsSizes()              AS LOCAL DWORD

  DIM f                           AS LOCAL DWORD
  DIM i                           AS LOCAL DWORD
  DIM j                           AS LOCAL DWORD
  DIM k                           AS LOCAL DWORD
  DO
    IF nAttachFiles = 0 THEN EXIT FUNCTION

    ReadPeFile
    lpImageSectionHeader = lpImageNtHeaders + SIZEOF(IMAGE_NT_HEADERS)
    j = lpImageSectionHeader - lpImageDosHeader + SIZEOF(IMAGE_SECTION_HEADER) * @lpImageNtHeaders.FileHeader.NumberOfSections
    FOR i = 0 TO @lpImageNtHeaders.FileHeader.NumberOfSections - 1
      j = MAX(j, @lpImageSectionHeader[i].PointerToRawData + @lpImageSectionHeader[i].SizeOfRawData)
    NEXT
    IF LEN(szPeFileData) < j THEN ERR = 1: EXIT DO
    REDIM szAttachFileData(nAttachFiles - 1)
    REDIM OffsetsSizes(3 + nAttachFiles + nAttachFiles)
    k = j + (UBOUND(OffsetsSizes()) + 1) * 4
    OffsetsSizes(3) = k
    szLegends = ""
    FOR i = 1 TO nAttachFiles
      k = k + LEN(szAttachFileLegend(i - 1)) + 1
      szLegends =  szLegends + szAttachFileLegend(i - 1) + CHR$(0)
      OffsetsSizes(i + 3) = k
    NEXT
    FOR i = 1 TO nAttachFiles
      f = FREEFILE: ERR = 0
      OPEN szAttachFilePath(i - 1) FOR BINARY SHARED AS #f
      IF ERR = 0 THEN GET$ #f, LOF(f), szAttachFileData(i - 1)
      CLOSE #f
      IF ERR THEN MSGBOX "Can't load " + szAttachFileLegend(i - 1), %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle: EXIT DO
      k = k + LEN(szAttachFileData(i - 1))
      OffsetsSizes(nAttachFiles + i + 3) = k
    NEXT
    OffsetsSizes(0) = 19900821 ' Magic for attach
    OffsetsSizes(1) = k - j
    OffsetsSizes(2) = nAttachFiles

    f = FREEFILE: ERR = 0
    OPEN szPePath FOR OUTPUT AS f
    IF ERR = 0 THEN PRINT #f, LEFT$(szPeFileData, j);
    IF ERR = 0 THEN PRINT #f, PEEK$(VARPTR(OffsetsSizes(0)), 4 * (UBOUND(OffsetsSizes()) + 1)) + szLegends;
    FOR i = 1 TO nAttachFiles
      IF ERR = 0 THEN PRINT #f, szAttachFileData(i - 1);
    NEXT
    CLOSE #f

    IF ERR THEN MSGBOX "Can't save " + szPePath, %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle
    EXIT DO
  LOOP
END FUNCTION
'================================================================================

'   FUNCTION Compile ALIAS "Compile" (szFileBas AS STRING, szPePath AS STRING, szCompilerDirectory AS STRING) EXPORT AS LONG
FUNCTION AttachResources ALIAS "AttachResources" (szFileBas AS STRING, szPePath AS STRING, szCompilerDirectory AS STRING) EXPORT AS LONG
  DIM szResource             AS LOCAL STRING
  DIM DontCompileRc          AS LOCAL DWORD
  DIM nAttachFiles           AS LOCAL DWORD
  DIM mAttachFiles           AS LOCAL DWORD
  DIM szAttachFileLegend()   AS LOCAL STRING
  DIM szAttachFilePath()     AS LOCAL STRING
  DIM nDelayLoadDlls         AS LOCAL DWORD
  DIM mDelayLoadDlls         AS LOCAL DWORD
  DIM szDelayLoadDll()       AS LOCAL STRING
  DIM ConvertToConsole       AS DWORD
  DIM Er                     AS LOCAL DWORD
  DIM nLineBas               AS LOCAL DWORD
  DIM PosInLine              AS LOCAL DWORD
  DIM LenOfWord              AS LOCAL DWORD
  DIM Symbol                 AS LOCAL DWORD
  DIM InsideDq               AS LOCAL DWORD
  DIM EndOfWord              AS LOCAL DWORD
  DIM nWord                  AS LOCAL DWORD
  DIM kServiceWord           AS LOCAL DWORD
  DIM azBasLine              AS LOCAL ASCIIZ * 256
  DIM bBasLine()             AS LOCAL BYTE
  DIM szTmp                  AS LOCAL STRING
  DIM Continue               AS LOCAL DWORD
  DIM i                      AS LOCAL DWORD
  DIM f                      AS LOCAL DWORD

  DO
    f = FREEFILE: nLineBas = 0: ERR = 0
    OPEN szFileBas FOR INPUT SHARED AS #f LEN = 32768: IF ERR THEN Er = 2: EXIT DO
    REDIM bBasLine(255) AT VARPTR(azBasLine)
    WHILE NOT EOF(f)
      LINE INPUT #f, azBasLine: IF ERR THEN CLOSE #f: Er = 2: EXIT LOOP
      INCR nLineBas:  bBasLine(255) = 0
      PosInLine = 0: LenOfWord = 0: InsideDq = 0: nWord = 0: kServiceWord = 0: Continue = 0
      DO
        Symbol = bBasLine(PosInLine)
        IF Symbol >= 97 THEN IF Symbol <= 122 THEN Symbol = Symbol - 32
        IF Symbol = 0 THEN
          EndOfWord = 1
        ELSEIF InsideDq THEN
          IF Symbol = 34 THEN
            IF bBasLine(PosInLine + 1) = 34 THEN
              bBasLine(LenOfWord) = Symbol: INCR LenOfWord: INCR PosInLine
            ELSE
              EndOfWord = 1
            END IF
          ELSE
            bBasLine(LenOfWord) = Symbol: INCR LenOfWord
          END IF
        ELSEIF Symbol = 39 THEN
          IF (nWord = 0) AND (LenOfWord = 0) AND (bBasLine(PosInLine + 1) = 37) THEN
            nWord = 1: kServiceWord = 1: INCR PosInLine: Continue = 1 ' '%
          ELSE
            EndOfWord = 1: Symbol = 0
          END IF
        ELSEIF Symbol = 34 THEN
          InsideDq = 1
        ELSEIF Symbol = 32 THEN
          IF LenOfWord THEN EndOfWord = 1
        ELSE
          bBasLine(LenOfWord) = Symbol: INCR LenOfWord
        END IF
        IF EndOfWord THEN
          IF (LenOfWord) OR (InsideDq) THEN
            IF Continue = 2 THEN Er = 3: EXIT DO
            INCR nWord: bBasLine(LenOfWord) = 0
            SELECT CASE AS LONG kServiceWord
              CASE 0
                EXIT DO
              CASE 1 ' '%
                SELECT CASE azBasLine
                  CASE "ATTACH"    : kServiceWord = 11
                  CASE "DELAY"     : kServiceWord = 12
                  CASE "RC"        : kServiceWord = 13
                  CASE "SUBSYSTEM" : kServiceWord = 14
                  CASE ELSE        : Er = 3: EXIT DO
                END SELECT
              CASE 11                : ' ATTACH
                                       kServiceWord = 111: INCR nAttachFiles
                                       IF nAttachFiles > mAttachFiles THEN
                                         IF mAttachFiles = 0 THEN mAttachFiles = 100 ELSE mAttachFiles = mAttachFiles + mAttachFiles
                                         REDIM PRESERVE szAttachFileLegend(mAttachFiles - 1)
                                         REDIM PRESERVE szAttachFilePath(mAttachFiles - 1)
                                       END IF
                                       szAttachFilePath(nAttachFiles - 1) = azBasLine

              CASE 111               : IF azBasLine = "ALIAS" THEN kServiceWord = 1111 ELSE Er = 3: EXIT DO
              CASE 1111              : szAttachFileLegend(nAttachFiles - 1) = azBasLine: Continue = 2
              CASE 12                : IF azBasLine = "LOAD" THEN kServiceWord = 121 ELSE Er = 3: EXIT DO ' DELAY
              CASE 121               : kServiceWord = 1211: INCR nDelayLoadDlls
                                       IF nDelayLoadDlls > mDelayLoadDlls THEN
                                         IF mDelayLoadDlls = 0 THEN mDelayLoadDlls = 100 ELSE mDelayLoadDlls = mDelayLoadDlls + mDelayLoadDlls
                                         REDIM PRESERVE szDelayLoadDll(mDelayLoadDlls - 1)
                                       END IF
                                       szDelayLoadDll(nDelayLoadDlls - 1) = azBasLine: Continue = 2
              CASE 13                : ' RC
                                       SELECT CASE azBasLine
                                         CASE "COMPILE"  : kServiceWord = 131
                                         CASE "TEXT"     : IF Symbol <> 0 THEN bBasLine(LenOfWord) = 32: _
                                                                szResource = szResource + RTRIM$(MID$(azBasLine, PosInLine + 2)) + $CRLF
                                                           Continue = 2: Symbol = 0
                                         CASE ELSE       : Er = 3: EXIT DO
                                       END SELECT
              CASE 131                : ' RC COMPILE
                                        SELECT CASE azBasLine
                                          CASE "ON"      : DontCompileRc = 0: Continue = 2
                                          CASE "OFF"     : DontCompileRc = 1: Continue = 2
                                          CASE ELSE      : Er = 3: EXIT DO
                                        END SELECT
              CASE 14                 : ' SUBSYSTEM
                                        SELECT CASE azBasLine
                                          CASE "CONSOLE" : ConvertToConsole = 3: Continue = 2
                                          CASE "GUI"     : ConvertToConsole = 2: Continue = 2
                                          CASE ELSE      : Er = 3: EXIT DO
                                        END SELECT
            END SELECT
          END IF
          LenOfWord = 0: InsideDq = 0: EndOfWord = 0
        END IF
        IF Symbol = 0 THEN
          IF Continue = 1 THEN Er = 3
          EXIT DO
        END IF
        INCR PosInLine
      LOOP
      IF Er THEN EXIT DO
    WEND
    CLOSE #f
    EXIT DO
  LOOP
  SELECT CASE AS LONG Er
    CASE 1: MSGBOX "Unexpected error", %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle
    CASE 2: MSGBOX "Can't process " + szFileBas, %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle
    CASE 3: MSGBOX szFileBas + " : error in line" + STR$(nLineBas), %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle
  END SELECT
  IF Er THEN EXIT FUNCTION
  DelayLoad szPePath, szDelayLoadDll(), nDelayLoadDlls, ConvertToConsole
  IF (DontCompileRc = 0) AND (szResource <> "") THEN
    f = FREEFILE: ERR = 0
    OPEN "~Tmp.Rc" FOR OUTPUT AS #f LEN = 32768
    IF ERR = 0 THEN PRINT #f, szResource;
    CLOSE #f
    IF ERR THEN
      MSGBOX "Can't create ~Tmp.Rc", %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle
    ELSE
      SHELL $DQ + szCompilerDirectory + "Rc.Exe" + $DQ + $SPC + "~Tmp", 0
      IF AddResToExe (szPePath, szPePath, "~Tmp.Res", 0, szCompilerDirectory) THEN MSGBOX _
              "Can't attach resources", %MB_ICONEXCLAMATION OR %MB_TASKMODAL, $MsgTitle
'            KILL "~Tmp.Rc": KILL "~Tmp.Res"
    END IF
  END IF
  AttachFileAfterEof  szPePath, szAttachFileLegend(), szAttachFilePath(), nAttachFiles
  MarkAsConsole szPePath, ConvertToConsole
END FUNCTION

FUNCTION U_09(BYREF a$,BYREF b$) AS LONG
  DIM f AS LOCAL DWORD
  DIM i AS LOCAL DWORD
  DIM j AS LOCAL DWORD
  DIM lpImageDosHeader AS LOCAL IMAGE_DOS_HEADER PTR
  DIM lpImageNtHeaders AS LOCAL IMAGE_NT_HEADERS PTR
  DIM lpImageSectionHeader AS LOCAL IMAGE_SECTION_HEADER PTR
  DIM szExeData AS LOCAL STRING
  DIM szSectionFileImage() AS LOCAL STRING
  DIM nSections AS LOCAL DWORD
  ERRCLEAR: f = FREEFILE: OPEN a$ FOR BINARY LOCK WRITE AS #f
  IF ERR = 0 THEN GET$ #f, LOF(f), szExeData
  CLOSE #f: IF ERR THEN FUNCTION = 1: EXIT FUNCTION
  lpImageDosHeader = STRPTR(szExeData)
  lpImageNtHeaders = lpImageDosHeader + @lpImageDosHeader.e_lfanew
  lpImageSectionHeader = lpImageNtHeaders + SIZEOF(IMAGE_NT_HEADERS)
  nSections = @lpImageNtHeaders.FileHeader.NumberOfSections
  REDIM szSectionFileImage(nSections)
  FOR i = 0 TO nSections - 1
    szSectionFileImage(i + 1) = MID$(szExeData, @lpImageSectionHeader[i].PointerToRawData + 1, @lpImageSectionHeader[i].SizeOfRawData)
  NEXT
  szSectionFileImage(0) = LEFT$(szExeData, @lpImageNtHeaders.OptionalHeader.SizeOfHeaders)
  j = 0: FOR i = 0 TO nSections: j = j + LEN(szSectionFileImage(i)): NEXT
  szExeData = STRING$(j, 0): j = STRPTR(szExeData)
  FOR i = 0 TO nSections
    MoveMemory BYVAL j, BYVAL STRPTR(szSectionFileImage(i)), LEN(szSectionFileImage(i))
    j = j + LEN(szSectionFileImage(i))
  NEXT
  lpImageDosHeader = STRPTR(szExeData)
  lpImageNtHeaders = lpImageDosHeader + @lpImageDosHeader.e_lfanew
  lpImageSectionHeader = lpImageNtHeaders + SIZEOF(IMAGE_NT_HEADERS)
  j = LEN(szSectionFileImage(0))
  FOR i = 0 TO nSections - 1
    @lpImageSectionHeader[i].PointerToRawData = j
    j = j + LEN(szSectionFileImage(i + 1))
  NEXT
  ERR = 0: f = FREEFILE: OPEN b$ FOR OUTPUT AS #f
  IF ERR = 0 THEN PRINT #f, szExeData;
  CLOSE #f
  IF ERR THEN FUNCTION = -4: EXIT FUNCTION
  KILL a$
  NAME  b$ AS a$
END FUNCTION

'----------------------------------------------------------------------------------------------------------------------------
' Is_File()
FUNCTION F_AF(BYVAL a$)AS LONG
  LOCAL x&
  x& = GETATTR(a$)
  FUNCTION = (ERRCLEAR = 0)
END FUNCTION

'----------------------------------------------------------------------------------------------------------------------------
' S_Part,
FUNCTION S_Part(BYREF a AS STRING,BYVAL b&,BYVAL c&) AS STRING
  LOCAL e$
  IF c&=0 THEN c&=LEN(a)
  e$=MID$(a,b&,c&-b&+1)
  FUNCTION = e$
END FUNCTION
'---------------------------------------------------------------------------------------------------
SUB DoEventsAPI
  ON ERROR RESUME NEXT
  STATIC Msg AS tagMsg
  IF PeekMessage(Msg, %NULL, 0, 0, %PM_REMOVE) THEN
    TranslateMessage Msg
    DispatchMessage Msg
  END IF
END SUB
'----------------------------------------------------------------------------------------------------------------------------
SUB U_10(BYVAL a$)
  DO
    SLEEP 50:DOeventsAPI
  LOOP UNTIL F_AF(a$)
  SLEEP 100
END SUB
'----------------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------
FUNCTION WinExecute (s AS STRING) AS LONG
  LOCAL i AS LONG
  LOCAL zExe AS ASCIIZ * %MAX_PATH
  LOCAL zCmd AS ASCIIZ * %MAX_PATH
  LOCAL tStartupInfo AS STARTUPINFO
  LOCAL tProcessInfo AS PROCESS_INFORMATION
  tStartupInfo.cb = LEN(STARTUPINFO)
  zExe = LTRIM$(s)
  zCmd = zExe
  i = INSTR(zExe, " ")
  IF i THEN
    zExe = LEFT$(zExe, i - 1)
  END IF
  i = CreateProcess(zExe, _ ' pointer to name of executable module
       zCmd, _ ' pointer to command line string
       BYVAL %NULL, _ ' pointer to process security attributes
       BYVAL %NULL, _ ' pointer to thread security attributes
       %NULL, _ ' handle inheritance flag
       %NORMAL_PRIORITY_CLASS OR %CREATE_SEPARATE_WOW_VDM, _ ' creation flags %DETACHED_PROCESS ???
       BYVAL %NULL, _ ' pointer to new environment block
       BYVAL %NULL, _ ' pointer to current directory name
       tStartupInfo, _ ' pointer to STARTUPINFO
       tProcessInfo) ' pointer to PROCESS_INFORMATION
  IF i THEN
    i = tProcessInfo.dwProcessId
  ELSE
    MSGBOX "CreateProcess GetLastError:" + STR$(GetLastError)
  END IF
  FUNCTION = i
END FUNCTION
'-------------------------------------------------------------------------------
FUNCTION ExecAndWait (cmd AS STRING) AS LONG
  LOCAL i AS LONG
  LOCAL ProcessID AS LONG
  LOCAL hProcess AS LONG
  LOCAL excode AS LONG
  'fPrint 1, "Executing: " + cmd
  ProcessID = WinExecute(cmd)
  IF ProcessID = 0 THEN
    MSGBOX "Problem with execution: " + cmd
    excode = 1
  ELSE
    hProcess = OpenProcess(%PROCESS_QUERY_INFORMATION, 0, ProcessID)
    IF hProcess = 0 THEN
      MSGBOX "Problem with process: " + cmd
      excode = 2
    ELSE
      DO
        i = GetExitCodeProcess(hProcess, excode)
        IF excode <> %STILL_ACTIVE THEN
          EXIT DO
        END IF
        DoEventsAPI
      LOOP
      CloseHandle hProcess
    END IF
  END IF
  IF excode THEN
    MSGBOX "Problem" + STR$(excode) + " with execution: " + cmd
  END IF
  SLEEP 10
  FUNCTION = excode
END FUNCTION

'----------------------------------------------------------------------------------------------------------------------------
' GetFN    FNO$
FUNCTION U_07(BYVAL a$) AS STRING
  DIM e AS STRING, al AS LONG
  al = LEN(a$)
  IF al > 0 THEN
     e = RIGHT$(a$, LEN(a$)-INSTR(-1,a$, "\"))
  ELSE: e = ""
  END IF
  FUNCTION = e
END FUNCTION
'----------------------------------------------------------------------------------------------------------------------------
' NOEXT
FUNCTION U_08(BYVAL a$) AS STRING
  FUNCTION=LEFT$(a$,INSTR(-1,a$,".")-1)
END FUNCTION
'----------------------------------------------------------------------------------------------------------------------------
' GetFile(a$)
FUNCTION F_AI(BYVAL a$,BYREF e AS STRING) AS LONG
  LOCAL b AS INTEGER, f AS LONG
  f=0
  TRY
    b = FREEFILE
    OPEN a$ FOR BINARY ACCESS READ AS #b LEN=16384
    GET$ #b, LOF(b), e: CLOSE #b:f=1
  CATCH
    MSGBOX "File <"+a$+"> could not be loaded."+$CRLF+"Error was: "+STR$(ERR)+"."
  END TRY
  FUNCTION=f
END FUNCTION
'----------------------------------------------------------------------------------------------------------------------------
' BeforeLastBackslash
FUNCTION U_05(BYVAL a$) AS STRING
  LOCAL e AS STRING,c&
  c&=INSTR(-1,a$,"\"): e=S_Part(a$,1,c&)
  FUNCTION=e
END FUNCTION
'---------------------------------------------------------------------------------------------------
'  Get App-Path
FUNCTION U_03() AS STRING
  #REGISTER NONE
  LOCAL poz AS LONG
  LOCAL buffer AS ASCIIZ * 512
  GetModuleFileName 0,BYVAL VARPTR(buffer),512
  ! lea esi, buffer
  ! xor edx, edx
  ! xor ecx, ecx
  ! cld
  gapSt:
  ! lodsb
  ! inc ecx ; count bytes
  ! cmp al, 0
  ! je gapEnd
  ! cmp al, "\"
  ! jne gapSt
  ! mov edx, ecx ; PUT count of each "\" IN edx
  ! jmp gapSt
  gapEnd:
  ! mov poz, edx
  FUNCTION = LEFT$(buffer,poz)
END FUNCTION
'----------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------
' NOEXT
FUNCTION U_06(BYVAL a$) AS STRING
  DIM e$,b&
  b&=INSTR(-1,a$,"."):e$=LEFT$(a$,b&-1)
  FUNCTION=e$
END FUNCTION
'---------------------------------------------------------------------------------------------------
' Get Windows Temp-Dir  SFN
FUNCTION U_04() AS STRING
  LOCAL lResult AS LONG
  LOCAL buff AS ASCIIZ * %MAX_PATH
  lResult = GetTempPath(BYVAL SIZEOF(buff), Buff)
  FUNCTION = TRIM$(buff)
END FUNCTION
'----------------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------
' App-Executer
' Parameter 1 is complete line icl. Special folder
' Parameter 2 is a Dword - nowait (2) or (1) - wait or omitted then default=1
FUNCTION U_02(BYVAL a$,b AS DWORD) AS LONG
  LOCAL c AS LONG, d AS STRING, e AS LONG
  SELECT CASE b
    CASE 2
      c = SHELL(a$)
    CASE 3: d = U_06(a$) + ".tmp"
      e = U_09(a$, d)
    CASE 4: U_10 a$
    CASE ELSE
      ExecAndWait a$
  END SELECT
  SLEEP 200
END FUNCTION
'----------------------------------------------------------------------------------------------------------------------------
' a$=Filename of .bas-file
' Returns:
' - 1 if a #COMPIL.E was found (EXE-Name in b$)  or
' - 2 if a "SED-Block" was found (Block in c$),
' - 3 if both
FUNCTION U_00(BYREF a$, BYREF b$, BYREF c$) AS LONG
  LOCAL d AS LONG, e AS STRING, f AS LONG
  LOCAL block$, g AS DWORD,h AS LONG, j AS STRING, fn$, m$
  d=F_AI(a$,e)
  ' File could not be loaded, we go. ELSE File is in e
  IF d=0 THEN GOTO erx
  ' Scan for Block
  block$=PARSE$(e$,$U01_SEDB,2)
  ' Remove the "Start" etc.
  g=INSTR(block$,$CRLF):block$=TRIM$(MID$(block$,g),$CRLF):g=LEN(block$)
  IF g>0 THEN c$=TRIM$(block$,$CRLF):f=(f OR 2)
  ' Scan for #COMPIL
  m$="#COMPILE":j=UCASE$(e):g=INSTR(j,m$)
  ' If no #COMPIL.E found
  IF g=0 THEN GOTO enx
  ' Search Line-End ->h
  h=INSTR(g,e,$CRLF)
  IF h=0 THEN h=LEN(e)
  ' Get Temp-String (#compil-line ->j
  j=TRIM$(S_Part(e,g,h))
  ' j contains Line with #COMPIL.E
  IF LEN(j)>256 THEN GOTO erx
  fn$=TRIM$(PARSE$(j,CHR$(34),2))
  IF (LEN(fn$)>256) OR (LEN(fn$)<3) THEN GOTO erx
  f=(f OR 1):b$=fn$
enx:
  FUNCTION=f
  EXIT FUNCTION
erx:
  f=0:b$="":block$="":c$=""
  FUNCTION=f
END FUNCTION
'----------------------------------------------------------------------------------------------------------------------------
' Call ME:
' a$=Filename of ".bas"-File, b$ if given overrides the output-file
SUB U_01(BYREF a$,BYVAL b$)

  LOCAL c AS LONG, d$, e$,f$,g AS LONG,h AS STRING, i AS DWORD,k AS LONG,m AS STRING, n AS LONG,p AS STRING
  c=U_00(a$,e$,f$) ' Read File get all Infos
  IF LEN(b$)>0 THEN d$=b$:GOTO weit
  IF (c AND 1)=1 THEN d$=e$
weit:
  IF (c AND 2)<>2 THEN EXIT SUB
  ' Execute line by line
  g=PARSECOUNT(f$,$CRLF)
  FOR i=1 TO g
    h=TRIM$(PARSE$(f$,$CRLF,i),$CRLF)
    IF LEN(h)<3 THEN ITERATE FOR
    IF LEFT$(h,2)="'#" THEN h=TRIM$(MID$(h,3))
    p=LCASE$(LEFT$(h,3)):h=TRIM$(MID$(h,4))
    SELECT CASE p
      CASE "?w ":k=1 'Start upx (etc.) and wait
      CASE "?s ":k=3 'Call Semens PE-Procesor
      CASE "?f ":k=4 ' Shall be WaitForFile
      CASE ELSE
        k=2
    END SELECT
    m=U_05(a$)' Get App-Dir
    'MSGBOX "a$="+a$+$CRLF+"h="+h
    REPLACE "?app\" WITH U_03() IN h
    REPLACE "?tmp\" WITH U_04() IN h
    REPLACE "?prodir\" WITH m IN h
    REPLACE "?file.bas" WITH a$ IN h
    REPLACE "?file.exe" WITH d$ IN h
    REPLACE "?file" WITH U_06(a$) IN h
    REPLACE "?fnoext" WITH U_08(U_07(a$)) IN h
    ' call execute_line
    'MSGBOX "Calling Execute with:"+$CRLF+"k="+STR$(k)+$CRLF+"h="+h
    n=U_02(h,k)
  NEXT
END SUB
