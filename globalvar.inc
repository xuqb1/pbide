'------------------------------------------------------------------------------
TYPE SED_SavePos
   Position AS LONG                ' ����λ��
   FileName AS ASCIIZ * 255 '%MAX_PATH  ' �ļ���
END TYPE
GLOBAL g_hInst        AS DWORD      '�������ʵ��
GLOBAL g_hMenu        AS DWORD      '������˵�
GLOBAL g_hMenuFile    AS DWORD      '�ļ��˵�
GLOBAL g_hMenuReopen  AS DWORD      '�ش��ļ�
GLOBAL g_hMenuEdit    AS DWORD      '�༭�˵�
GLOBAL g_hMenuSearch  AS DWORD      '���Ҳ˵�
GLOBAL g_hMenuRun     AS DWORD      '���в˵�
GLOBAL g_hMenuView    AS DWORD      '��ͼ�˵�
GLOBAL g_hMenuTool    AS DWORD      '���߲˵�
GLOBAL g_hMenuProject AS DWORD      '��Ŀ�˵�
GLOBAL g_hMenuOptions AS DWORD      '���ò˵�
GLOBAL g_hMenuCoop    AS DWORD      'Э���˵�
GLOBAL g_hMenuWindow  AS DWORD      '���ڲ˵�
GLOBAL g_hMenuHelp    AS DWORD      '�����˵�
GLOBAL g_hStatus      AS DWORD      '״̬�����
GLOBAL g_hWndMain     AS DWORD      '�����ھ��
GLOBAL g_hWndClient   AS DWORD      '�ͻ������
GLOBAL g_hGuiderWin   AS DWORD      '�򵼴��ھ��
GLOBAL g_hGuiderTab() AS DWORD      '�򵼴����е�Tab�������
GLOBAL g_hOptionWin   AS DWORD      'ѡ��ھ��
GLOBAL g_fClosed      AS LONG       '���һ���༭���ڹرձ�ʶ
GLOBAL g_FreezeMenu   AS LONG       '����˵�
GLOBAL g_hTabMdi      AS DWORD
GLOBAL g_NewComCtl    AS CURRENCYX  '
GLOBAL g_NewDocNum    AS LONG       '
GLOBAL g_zIni         AS ASCIIZ * 500 '%MAX_PATH '
GLOBAL ghRebar    AS DWORD         ' ReBar�ؼ����.
GLOBAL ghToolBar  AS DWORD         ' �������ؼ����.
GLOBAL ghWindowBar AS DWORD        ' ���ڹ��������
GLOBAL ghWndSize AS DWORD          ' ���ڹ�����������ڱ༭����λ�ô�С��ʾ����
GLOBAL ghComboBox AS DWORD         ' ComboBox�ؼ����.
GLOBAL ghDebugBar AS DWORD         ' ������Թ��������
GLOBAL ghCodeEditBar AS DWORD      '����༭���������
GLOBAL ghWinArragBar AS DWORD      ' �������й��������
GLOBAL g_FileType     AS DWORD      '��ʶ�½�ʱ���ļ�����:0,bas;1,inc;2,rc

GLOBAL ghRebarMenu AS DWORD        'Rebar�ϵ��Ҽ������˵����
GLOBAL g_hRebar   AS DWORD
GLOBAL g_hToolbar     AS DWORD      '���������

GLOBAL pbtempfile     AS STRING
GLOBAL language       AS STRING
GLOBAL canLog         AS INTEGER
GLOBAL ghFont         AS LONG     '�趨�����б����ֵ�����ľ��
GLOBAL CodeTipsFile   AS STRING
GLOBAL hCodetipsFile  AS LONG
GLOBAL hCodeKeepFile  AS LONG
GLOBAL hFont          AS LONG
GLOBAL hCurSplit      AS LONG
GLOBAL hCurSplitV     AS LONG
GLOBAL hCurNorm       AS LONG
GLOBAL strPBKeyWords  AS STRING
GLOBAL fMaximize      AS LONG
GLOBAL ConstructAutocompletion  AS LONG
GLOBAL AutoCompleteTypes        AS LONG
GLOBAL TrimTrailingBlanks       AS LONG
GLOBAL ShowProcedureName        AS LONG
GLOBAL sProjectName             AS STRING
GLOBAL sAutoCompletionTypes     AS STRING
GLOBAL bMsgWindow               AS LONG
GLOBAL g_hTabMdi                AS DWORD
GLOBAL hTabImageList            AS DWORD
'GLOBAL hCodeFinder              AS DWORD  '��ghComboBox���
GLOBAL cUntitledFiles           AS LONG
GLOBAL gTabFilePaths()          AS STRING
GLOBAL ghEditStr                AS STRING
GLOBAL strContextMenu           AS STRING
GLOBAL SED_LastPosition         AS SED_SavePos
GLOBAL fClosed              AS LONG
GLOBAL sProjectPrimary      AS STRING
GLOBAL ghTreeView           AS DWORD
GLOBAL hPTROOT              AS DWORD
GLOBAL PrimarySourceFile    AS STRING
GLOBAL hHandles()           AS LONG         'Gripper handles
GLOBAL gZorder()            AS LONG         '��������ά�������Լ���Z���б�
GLOBAL hPopupClient         AS LONG         '���༭���ڵ����˵����
GLOBAL hPopupCtl            AS LONG         '���༭�����ϵ�ǰ�ؼ����Ҽ������˵����
GLOBAL hWndMbuilder         AS DWORD
GLOBAL hWndEdit             AS DWORD
GLOBAL hWndMDIEdit          AS DWORD
GLOBAL IDhitValue           AS LONG         '����,��������ָʾ
' ��Ŀ���
GLOBAL hProjectTV           AS DWORD '��Ŀ�����е�����ͼ
GLOBAL hPTRoot              AS DWORD         ' �������
GLOBAL hPTRCParent          AS DWORD
GLOBAL hPTBASParent         AS DWORD
GLOBAL hPTINCParent         AS DWORD
GLOBAL ghProjectTV          AS DWORD
GLOBAL RecentProjects()     AS STRING

TYPE PROJECT_PATH_NAMES
  sFName     AS STRING * 25
  sFPathName AS STRING * 255 '%MAX_PATH
END TYPE
GLOBAL udtProjectData()     AS PROJECT_PATH_NAMES
GLOBAL szProjectDir         AS STRING

GLOBAL hFind                AS DWORD '����/�����滻���ھ��

GLOBAL strCompilerTitle        AS STRING
GLOBAL CommandLine             AS STRING                 ' �������г���ʱ�����õ������в���
GLOBAL hPbIde               AS DWORD
$PBWINTITLE     = "PowerBASIC Compiler for Windows IDE"
$PBWIN80TITLE   = "PB/Win IDE"
$PBCCTITLE      = "PowerBASIC Console Compiler IDE"
$PBCC80TITLE    = "PB/CC IDE"
'global ghMsgWindow          as dword '��hCompilersTB���
$MsgTitle = "Pre-processor"
$U01_SEDB="'# SED_Block"+":"

GLOBAL RecentFiles()           AS STRING                 ' // Recent file names