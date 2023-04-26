'==============================================================================
'
' INITCTRL.INC - Initialize COMCTL32.DLL and return the version number of
' COMCTL32.DLL at the same time.
'
' Copyright (c) 2001 PowerBASIC, Inc.
' All Rights Reserved.
'
' Last updated 30 May 2002
'
' Here we use LoadLibrary and GetProcAddress to avoid crash on failure when
' older Win95 systems are running without the latter IE3+ COMCTL32 versions.
'
' Example use:  ComCtlVersion = InitComCtl32(%ICC_BAR_CLASSES)
' If the InitComCtl32() parameter is zero, it defaults to initialization
' with %ICC_WIN95_CLASSES only.
'
' Additional %ICC_ equates are defined in COMMCTRL.INC
'
'==============================================================================

#IF NOT %DEF(%INITCTRL_INC)

%INITCTRL_INC = 1

#IF NOT %DEF(%COMMCTRL_INC)
    #INCLUDE "COMMCTRL.INC"
#ENDIF

DECLARE FUNCTION zInitCommonControlsEx (iccex AS INIT_COMMON_CONTROLSEX) AS LONG


'------------------------------------------------------------------------------
' Retrieve a OS system file's version number, for example, "COMCTL32.DLL"
'------------------------------------------------------------------------------
FUNCTION GetSysFileVer (BYVAL Fname AS STRING) EXPORT AS CURRENCYX

    LOCAL dwLen     AS LONG
    LOCAL dwHandle  AS DWORD
    LOCAL sBuf      AS STRING
    LOCAL tWFD      AS WIN32_FIND_DATA
    LOCAL szFile    AS ASCIIZ * %MAX_PATH
    LOCAL pFI       AS VS_FIXEDFILEINFO PTR

    ' Get System directory/folder path
    GetSystemDirectory szFile, %MAX_PATH
    IF RIGHT$(szFile, 1) <> "\" THEN
        szFile = szFile + "\"
    END IF
    szFile = szFile + Fname

    ' Attempt to obtain version info from the file
    dwLen = GetFileVersionInfoSize(szFile, dwHandle)
    IF dwLen THEN
        sBuf     = SPACE$(dwLen)                                    ' Allocate space
        GetFileVersionInfo szFile, dwLen, dwLen, BYVAL STRPTR(sBuf) ' Get version info
        VerQueryValue BYVAL STRPTR(sBuf), "\", pFI, dwLen           ' Get VS_FIXEDFILEINFO info

        ' Format the version number
        FUNCTION = HIWRD(@pFI.dwFileVersionMS) +  _     ' Major version number
                   LOWRD(@pFI.dwFileVersionMS) / 100    ' Minor version number
    END IF

END FUNCTION



'------------------------------------------------------------------------------
' Safely initialize the Common Control Library, and return the version number
' of the DLL as a CURRENCYX/CUX data type, ie, 4.71.
'------------------------------------------------------------------------------
#IF %DEF(%NOBYVALPTR)
FUNCTION InitComCtl32 ([BYVAL iccClasses AS DWORD]) EXPORT AS CURRENCYX
#ELSE
FUNCTION InitComCtl32 (OPTIONAL BYVAL iccClasses AS DWORD) EXPORT AS CURRENCYX
#ENDIF

    LOCAL hLib     AS DWORD
    LOCAL dwProc   AS DWORD
    LOCAL iccex    AS INIT_COMMON_CONTROLSEX

    ' Use dynamic (run-time) linking to initialize the common control library
    ' unless only a very old version is installed.  Implicit (load-time) linking
    ' should be avoided, since this will stop the application from starting up
    ' if the "InitCommonControlsEx" API is not available in COMCTL32.DLL (Win95a).

    hLib = LoadLibrary("COMCTL32.DLL")
    IF hLib THEN
        dwProc = GetProcAddress(hLib, "InitCommonControlsEx")
        IF dwProc THEN
            ' Fill the ICCEX structure
            iccex.dwSize = SIZEOF(iccex)

            ' Set the control classes to initiate
            IF iccClasses THEN
                iccex.dwICC = iccClasses
            ELSE
                iccex.dwICC = %ICC_WIN95_CLASSES
            END IF

            ' Directly call the API function
            CALL DWORD dwProc USING zInitCommonControlsEx(iccex)
        ELSE
            ' Most likely a Windows 95 initialization
            InitCommonControls
        END IF

        ' Call FreeLibrary now, as the common control library has been initialized
        FreeLibrary hLib

        ' Return the version number of COMCTL32.DLL
        FUNCTION = GetSysFileVer("COMCTL32.DLL")
    END IF

END FUNCTION



#ENDIF    ' #IF NOT DEF(%INITCTRL_INC)
