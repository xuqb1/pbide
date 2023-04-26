$PROGRAMCLASSNAME       = "VisPowerBASIC"  ' ��������
$INIFILENAME            = "config.ini"    ' �����ļ�
$HELPFILE               = "PBWin.hlp"      ' �����ļ� (�ɸ�����Ҫ�޸�����)
$DATABASENAME           = "data.db"
$CONFIGTABLENAME        = "t_config"
$DEFAULTCONFIGTABLENAME = "t_config_default"
$DATASETTABLENAME       = "t_dataset"
$STRINGLANGTABLENAME    = "t_string_lang"
$CONTROLSTYLETABLENAME  = "t_control_style"
$SCI_CODETIPSDB         = "SCI_CDTIP.TRM"
$SCI_CODEKEEPDB         = "SCI_CODKP.TRM"
$LYNXWINDOW             = "GTI_LYNX"
$OldWndProc             = "OldWndProc"
$MainEdit               = "MAINFORM"
$MainMDIEdit            = "MAINMDIFORM"
'------------------------------------------------------------------------------
'%MAX_PATH = 500
'���봰��ID
%ID_CLIENTWINDOW  = 100  '%WM_USERֵ��1024�������ڿͻ���ʵ�崰��ID
%ID_STATUSBAR     = 101  ' ״̬��
%IDC_TABMDI       = 102  ' ����Tab��

%ID_PROPERTY      = 110   '���Դ���ID
%ID_CONTROLS      = 111   '�ؼ�����ID
%ID_PROJECT       = 112   '��Ŀ����ID
%ID_COMPILERS     = 113   '����������ID
%ID_FINDRS        = 114   '���ҽ������ID

%ID_FTOOLBAR      = 120  ' �����䴰���еĹ�����

'���봰�ڶԲ����������������֯�󸳸�byside
%BYLEFT       = 8
%BYTOP        = 4
%BYRIGHT      = 2
%BYBOTTOM     = 1
'���������ĸ��ͻ����Ĳ������������
%TOLEFT       = 4
%TOTOP        = 4
%TOCENTER     = 2 '�κ�һ������������0,��˵�������ڣ����ڵ���2�������(ʵ��ʹ�ô���0)
%TORIGHT      = 1
%TOBOTTOM     = 1
'���뵽��λ��
%MLEFT        = 1 '�����ڵ��������£�1,2,3,4
%MTOP         = 2
%MRIGHT       = 3
%MBOTTOM      = 4
%DLEFT        = 5 'ĳ�����봰�ڵ��������£�5,6,7,8
%DTOP         = 6
%DRIGHT       = 7
%DBOTTOM      = 8



%IDM_NEWTEMP          = 130  ' ��ģ���½��ļ� +20
' �ļ��˵�
%IDM_NEWBAS           = 160  ' �½�bas�ļ�
%IDM_NEWINC           = 161  ' �½�inc�ļ�
%IDM_NEWRC            = 162  ' �½�rc�ļ�
%IDM_OPEN             = 163  ' ���ļ�
%IDM_OPENPROJECT      = 164  ' ����Ŀ�ļ�
%IDM_INSERTFILE       = 165  ' �����ļ�
%IDM_NEW              = 166  ' �½��˵�
%IDM_REOPEN           = 167  ' �ش򿪲˵�
%IDM_REFRESH          = 168  ' ˢ�²˵�
'-
%IDM_SAVE             = 170  ' ����
%IDM_SAVEAS           = 171  ' ����Ϊ
%IDM_SAVEALL          = 172  ' ȫ������
%IDM_SAVEPROJECT      = 173  ' ������Ŀ
'-
%IDM_PRINTSETTING     = 174  ' ��ӡ����
%IDM_PRINTPREVIEW     = 175  ' ��ӡԤ��
%IDM_PRINT            = 176  ' ��ӡ...
'-
%IDM_CLOSE            = 177  '�ر�
%IDM_CLOSEALL         = 178  '�ر�ȫ��
%IDM_CLOSEOTHERS      = 179  '�ر�����
'-
%IDM_RECENTFILES      = 180
%IDM_RECENT1          = 181  ' ������ļ� 1
%IDM_RECENT2          = 182  ' ������ļ� 2
%IDM_RECENT3          = 183  ' ������ļ� 3
%IDM_RECENT4          = 184  ' ������ļ� 4
%IDM_RECENT5          = 185  ' ������ļ� 5
%IDM_RECENT6          = 186  ' ������ļ� 6
%IDM_RECENT7          = 187  ' ������ļ� 7
%IDM_RECENT8          = 188  ' ������ļ� 8
'%IDM_RECENT9          = 189  ' ������ļ� 9
'%IDM_RECENT10         = 190  ' ������ļ� 10
'-
%IDM_OPENCMD          = 190  ' ��������
%IDM_EXIT             = 191  ' �˳�

' �༭�˵�
%IDM_UNDO             = 200   ' ����
%IDM_REDO             = 201   ' ����
'-
%IDM_CUT              = 202   ' ����
%IDM_COPY             = 203   ' ����
%IDM_PASTE            = 204   ' ճ��
%IDM_PASTEIE          = 205   ' ճ��
%IDM_SELALL           = 206   ' ȫѡ
%IDM_DELETE           = 207   ' ɾ��
%IDM_LINEDELETE       = 208   ' ɾ��
%IDM_CLEARALL         = 209   ' ɾ��
'-
%IDM_COMMENT          = 210   ' ע��
%IDM_UNCOMMENT        = 211   ' ��ע��
%IDM_INDENT           = 212   ' ����
%IDM_OUTDENT          = 213   ' ������
%IDM_FORMATREGION     = 214   ' ��ʽ��
%IDM_TABULATEREGION   = 215   ' �б���
'-
%IDM_SELTOUPPERCASE   = 216   ' ��д��
%IDM_SELTOLOWERCASE   = 217   ' Сд��
%IDM_SELTOMIXEDCASE   = 218   ' �շ廯
'-
%IDM_TEMPLATES        = 219   ' ģ�����
%IDM_HTMLCODE         = 220   ' תΪ��ҳ
%IDM_GUID             = 221   ' ����GUID
'%IDM_BLOCK            = 2039&   ' ��
'%IDM_SPACEINDENT      = 2044&   ' TAB������
'%IDM_SPACEOUTDENT     = 2045&   ' TAB������
'-

'�����˵�
%IDM_FIND             = 230   ' ����
%IDM_FINDNEXT         = 231   ' ������һ��
%IDM_FINDBACKWARDS    = 232
%IDM_REPLACE          = 233   ' �滻
%IDM_GOTOLINE         = 234   ' ת����
'-
%IDM_TOGGLEBOOKMARK   = 235
%IDM_NEXTBOOKMARK     = 236
%IDM_PREVIOUSBOOKMARK = 237
%IDM_DELETEBOOKMARKS  = 238
'%IDM_GOTOBOOKMARK     = 2054&   ' ת����ǩ
%IDM_FINDINFILES      = 239
%IDM_EXPLORER         = 240
%IDM_WINDOWSFIND      = 241
%IDM_FILEFIND         = 242

'���в˵�
%IDM_COMPILE          = 250   ' ����
%IDM_COMPILERUN       = 251   ' ���벢����
%IDM_COMPILEDEBUG     = 252   ' ���벢����
%IDM_EXECUTE          = 253   ' ����
'-
%IDM_SETPRIMARY       = 254   ' ������Դ�ļ�
%IDM_CMDPARAMETER     = 255   ' ���������в���
%IDM_DEBUGTOOL        = 256   ' ���Թ���

'��ͼ�˵�
%IDM_PROPERTY         = 260
%IDM_CONTROLS         = 261
%IDM_PROJECT          = 262
%IDM_COMPILERS        = 263
%IDM_FINDRS           = 264
'-
%IDM_REBAR            = 265
'-
%IDM_TOGGLE           = 266
%IDM_TOGGLEALL        = 267
%IDM_FOLDALL          = 268
%IDM_EXPANDALL        = 269
'-
%IDM_ZOOMIN           = 270
%IDM_ZOOMOUT          = 271
'-
%IDM_USETABS          = 272
%IDM_AUTOINDENT       = 273
%IDM_SHOWLINENUM      = 274
%IDM_SHOWMARGIN       = 275
%IDM_SHOWINDENT       = 276
%IDM_SHOWSPACES       = 277
%IDM_SHOWEOL          = 278
%IDM_SHOWEDGE         = 279
%IDM_SHOWPROCNAME     = 280
'-
%IDM_CVEOLTOCRLF      = 281
%IDM_CVEOLTOCR        = 282
%IDM_CVEOLTOLF        = 283
'-
%IDM_REPLSPCWITHTABS  = 284
%IDM_REPLTABSWITHSPC  = 285
'-
%IDM_FILEPROPERTIES   = 286
%IDM_SYSINFO          = 287


' ���ڲ˵�
%IDM_CASCADE          = 290   ' �������
%IDM_TILEH            = 291   ' ˮƽ����
%IDM_TILEV            = 292   ' ��ֱ����
'-
%IDM_RESTOREWSIZE     = 293   ' ���������ڳߴ�
%IDM_SWITCHWINDOW     = 294   ' ת������
%IDM_ARRANGE          = 295   ' ͼ������
'-
%IDM_CLOSEWIN         = 296   ' ȫ���ر�
%IDM_MAXWIN           = 297   ' ��󻯴���

'��Ŀ�˵�
%IDM_NEWPROJECT       = 300   ' �½���Ŀ
'%IDM_OPENPROJECT      = 164  ' ����Ŀ�ļ�
%IDM_LOADPROJECT      = 301   ' ������Ŀ
'%IDM_SAVEPROJECT     = 2111&   ' ������Ŀ,ʹ���ļ��˵��е�%IDM_SAVEPROJECT
%IDM_SAVEPROJECTas    = 302   ' ������Ŀ
%IDM_CLOSEPROJECT     = 303   ' �ر���Ŀ
'-
%IDM_REGPROJECTEXT    = 304   ' ������Ŀ�ļ�
'-
%IDM_ACHIVE           = 305   ' ��Ŀ�鵵
%IDM_NEWVISION        = 306   ' ��Ŀ�°汾
%IDM_FILEVISIONSETTING= 307   ' ��Ŀ���ļ��汾����
%IDM_VISIONPACKAGE    = 308   ' �汾���
'-
%IDM_RECENTPRJ1       = 310   ' �����Ŀ1
%IDM_RECENTPRJ2       = 311   ' �����Ŀ2
%IDM_RECENTPRJ3       = 312   ' �����Ŀ3
%IDM_RECENTPRJ4       = 313   ' �����Ŀ4
%IDM_RECENTPRJ5       = 314   ' �����Ŀ5
%IDM_RECENTPRJ6       = 315   ' �����Ŀ6
%IDM_RECENTPRJ7       = 316   ' �����Ŀ7
%IDM_RECENTPRJ8       = 317   ' �����Ŀ8

'���ò˵�
%IDM_INTERFACEOPT     = 320
%IDM_FILEEXTOPT       = 321
%IDM_EDITOROPT        = 322
%IDM_COLORSOPT        = 323
%IDM_FORMOROPT        = 324
%IDM_COMPILEROPT      = 325
%IDM_TOOLSOPT         = 326
%IDM_FOLDINGOPT       = 327
%IDM_PRINTOPT         = 328
%IDM_SHORTKEYOPT      = 329

'����
%IDM_MAKEMENU         = 340   ' �����˵�
%IDM_MAKETOOLBAR      = 341   ' ����TOOLBAR
%IDM_GOTCHA           = 342
%IDM_FINDMULTIDIR     = 343   ' �ڶ��Ŀ¼�в���
%IDM_REPLACEMULTIDIR  = 344   ' �ڶ��Ŀ¼���滻
'-
%IDM_FORMATCODE       = 345   ' ��ʽ������
%IDM_CODEC            = 346   ' ��ʽ������
%IDM_CODECOUNT        = 347   ' ����ͳ��
%IDM_CODEMANAGER      = 348   ' ���������
%IDM_CODEKEEPER       = 349   ' ���뱣��
%IDM_KBDMACROS        = 350   ' �����ݼ�
%IDM_INCLEAN          = 351   ' �������
%IDM_WINSPY           = 352   ' ��Ŀ�괰�����ɴ���
'-
%IDM_CHARMAP          = 353
%IDM_MSGBOXDESIGNER   = 354
%IDM_PBFORMS          = 355
%IDM_PBCOMBR          = 356
%IDM_TYPELIBBR        = 357
%IDM_IMGEDITOR        = 358
%IDM_RCEDITOR         = 359
%IDM_CODETIPSBUILDER  = 360
%IDM_CODETYPEBUILDER  = 361
%IDM_DLGEDITOR        = 362
%IDM_POFFS            = 363
%IDM_COPYCAT          = 364
%IDM_CALCULATOR       = 365
%IDM_MORETOOLS        = 366
'-
%IDM_CUSTOMTOOL1      = 370   ' �Զ��幤��1
%IDM_CUSTOMTOOL2      = 371   ' �Զ��幤��2
%IDM_CUSTOMTOOL3      = 372   ' �Զ��幤��3
%IDM_CUSTOMTOOL4      = 373   ' �Զ��幤��4
%IDM_CUSTOMTOOL5      = 374   ' �Զ��幤��5
%IDM_CUSTOMTOOL6      = 375   ' �Զ��幤��6
%IDM_CUSTOMTOOL7      = 376   ' �Զ��幤��7
%IDM_CUSTOMTOOL8      = 377   ' �Զ��幤��8
'-
%IDM_CUSTOMTOOLSETTING= 2100&   ' �Զ��幤������

'Э��
%IDM_COOPSETTING      = 2130&   ' Э������
%IDM_SENDINVITE       = 2131&   ' ��������
%IDM_SENDVISIT        = 2132&   ' ���ͷ�������
%IDM_DOWNLOADFILE     = 2133&   ' �����ļ�
%IDM_UPLOADFILE       = 2134&   ' �����ļ�
%IDM_CHAT             = 2135&   ' ��ͨ

'-
%IDM_OPTION           = 2149&   ' ѡ��...

' ����
%IDM_IDEHELP          = 2160&   ' IDE����
%IDM_PBWINCONTENT     = 2161&   ' PB/WIN��������
%IDM_PBWININDEX       = 2162&   ' PB/WIN��������
%IDM_PBWINSEARCH      = 2163&   ' PB/WIN��������
%IDM_WINDOWSSDK       = 2164&   ' WINDOWS SDK����
'-
%IDM_ABOUT            = 2166&   ' ����
%IDM_UPDATE           = 2167&   ' ����

'���ڱ༭��������ť�����
%IDM_WIN              = 2200    ' �½�����
%ID_TOOL_MENU         = 2201    ' �½��˵�
%ID_TOOL_TOOLBAR      = 2202    ' �½�������
%ID_BARONLY_STICKY    = 2203    ' ճ��
%ID_BARONLY_SNAP      = 2204    ' ���뵽����
%ID_BARONLY_GRID      = 2205    ' ��ʾ/��������
%ID_SETV_CENTER       = 2206    ' ��ֱ���ж���
%ID_SETH_CENTER       = 2207    ' ˮƽ���ж���
%ID_SETV_EQUATE       = 2208    ' ��ֱ�Ⱦ�ֲ�
%ID_SETH_EQUATE       = 2209    ' ˮƽ�Ⱦ�ֲ�
%ID_TOOL_COMPILE      = 2210    ' ת��DDT��SDK����
%ID_TOOL_POS          = 2211    'λ�ð�ť
%ID_TOOL_WH           = 2212    '���߰�ť
%ID_SETH_LEFT         = 2213    'ˮƽ����밴ť
%ID_SETH_RIGHT        = 2214    'ˮƽ�Ҷ��밴ť
%ID_SETV_TOP          = 2215    '��ֱ�϶��밴ť
%ID_SETV_BOTTOM       = 2216    '��ֱ�¶��밴ť
%MN_F_DDTCODE         = 2217    'DDT����
%MN_F_SDKCODE         = 2218    'SDK����

'���룬���Բ˵�/������
%IDM_SETBRKPO         = 2230    '���öϵ�
%IDM_PRVBRKPO         = 2231    '��һ�ϵ�
%IDM_NEXBRKPO         = 2232    '��һ�ϵ�
%IDM_DELBRKPO         = 2233    'ɾ���ϵ�
%IDM_GOBRKPO          = 2234    'ת���ϵ�
'%IDM_COMPILE          = 2235    '����
%IDM_CMPRUN           = 2236    '��һ�ϵ�
%IDM_CMPDEBUG         = 2237    '��һ�ϵ�
'%IDM_PRVBRKPO         = 2238    '��һ�ϵ�
'%IDM_PRVBRKPO         = 2239    '��һ�ϵ�


'����༭�˵�/������
%IDM_BLOCKCOMMIT      = 2250    '��ע��
%IDM_REBLOCKCOMMIT    = 2251    '����ע��
%IDM_BLOCKINDENT      = 2252    '������
%IDM_REBLOCKINDENT    = 2253    '��������
%IDM_CAPWORD          = 2254    'ת��д
%IDM_SMALLWORD        = 2255    'תСд
%IDM_FIRSTCAPWORD     = 2256    '����ĸ��д
'%IDM_FIND             = 2257    '����
'%IDM_FINDNEXT         = 2258    '������һ��
%IDM_FINDREPLACE      = 2259    '���Ҳ��滻
%IDM_GOLINE           = 2260    'ת����

%IDM_PBTEMPLATES    = 9000      ' ����ָʾ����PBģ��
%IDM_WINTEMPLATES   = 8000      ' ����ģ��

%ID_GUIDERTAB                   = 2270    '�򵼴������ID
%IDC_GUILDNEWLIST               = 2271    '�򵼴����½�TAB�е��б���ͼ
%IDC_GUILDEXISTFINDLABEL        = 2272    '�򵼴����ִ�TAB�е�"���ҷ�Χ:"��ǩ
%IDC_GUILDEXISTFINDCOMBOBOX     = 2273    '�򵼴����ִ�TAB�е�"���ҷ�Χ:"��Ͽ�
%IDC_GUILDEXISTFINDLIST         = 2274    '�򵼴����ִ�TAB�е�"���ҷ�Χ:"�б���ͼ
%IDC_GUILDEXISTFILENAMELABEL    = 2275    '�򵼴����ִ�TAB�е�"�ļ���:"��ǩ
%IDC_GUILDEXISTFILENAMETEXTBOX  = 2276    '�򵼴����ִ�TAB�е�"�ļ���:"�ı���
%IDC_GUILDEXISTFILETYPELABEL    = 2277    '�򵼴����ִ�TAB�е�"�ļ�����:"��ǩ
%IDC_GUILDEXISTFILETYPECOMBOBOX = 2278    '�򵼴����ִ�TAB�е�"�ļ�����:"��Ͽ�
%IDC_GUILDRECENTLIST            = 2279    '�򵼴�������TAB�е��б���ͼ
%IDC_GUILDHELPBUTTON            = 2280    '�򵼴��ڰ�����ť
%IDC_GUILDSHOWCHECKBOX          = 2281    '�򵼴���"������ʾ����Ի���"����


%IDC_EDIT1    = 2310
%ID_ProcEnd    = 2311

%SPLITSIZE                  = 4
%IDM_INITASNEW      = 3303      ' ��ʼ��Ϊ�µ�
%IDM_CLEAR          = 3304      ' ���
'%IDM_CLEARALL       = 3305      ' �������
'%IDM_SELECTALL      = 3309      ' ȫѡ
'%IDM_PASTEIE        = 3312      ' ճ�� HTML
'%IDM_LINEDELETE     = 3313      ' ɾ����
%IDM_AUTOCOMPLETE   = 3314      ' �Զ����
'%IDM_TEMPLATES      = 3315      ' ģ��
'%IDM_HTMLCODE       = 3316      ' ת��Ϊ Html
'%IDM_SELTOUPPERCASE = 3317      ' ת����ѡ�ı�Ϊ��д
'%IDM_SELTOLOWERCASE = 3318      ' ת����ѡ�ı�ΪСд
'%IDM_SELTOMIXEDCASE = 3319      ' ת����ѡ�ı�Ϊ��ϴ�Сд
'%IDM_FORMATREGION   = 3320      ' ��ʽ����ѡ�ı�
'%IDM_TABULATEREGION = 3321      ' �б�ѡ���ı�
'%IDM_BLOCKINDENT    = 3322      ' ������
%IDM_BLOCKUNINDENT  = 3323      ' ��ȡ������
%IDM_GOTOBEGINPROC  = 3324      ' ��ת����������Ŀ�ͷ
%IDM_GOTOENDPROC    = 3325      ' ��ת�����������ĩβ
%IDM_GOTOBEGINTHISPROC = 3326       ' ��ת���˳�������Ŀ�ͷ
%IDM_GOTOENDTHISPROC = 3327         ' ��ת���˳��������ĩβ

%IDK_COMMENT        = 3331      ' ��/��ע��
%IDK_UNCOMMENT      = 3332      ' ��/��ȡ��ע��
%IDM_HELP           = 3551

'%IDM_REGPROJECTEXT  = 3609      ' ע���ļ���չ��
%IDM_INSERTINPROJECT= 3610      ' �ڵ�ǰ��Ŀ�в����ļ�
%IDM_RECENTPROJECTS = 3650      ' �����Ŀ
'%IDM_RECENTPRJ1     = 3651      ' �������Ŀ1
'%IDM_RECENTPRJ2     = 3652      ' �������Ŀ2
'%IDM_RECENTPRJ3     = 3653      ' �������Ŀ3
'%IDM_RECENTPRJ4     = 3654      ' �������Ŀ4

%IDM_GOTOSELFILE    = 3801      ' ��ת����ع��������ͬ�ļ������ļ�
%IDM_GOTOSELPROC    = 3802      ' ��ת�����������ͬ���ֵĺ���/����
%IDM_GOTOLASTPOSITION = 3803        ' ��ת���������ĵ�
%IDM_RECENTPROJECTSSEPARATOR = 3655         ' �ָ���

%ID_LYNX_FIND       = 110

%IDM_FORM_PRINT     = 3217      '��ӡ����

'* Popup Contextmenu WM_RBUTTONDOWN on a Control���༭�����ϵĵ����˵�,�������ؼ���WM_RBUTTONDOWN
%MN_UNDO            =40200      '����
%MN_REDO            =40201      '����
%MN_CUT             =40202      '���иÿؼ�Cut this control
%MN_COPY            =40203      '���Ƹÿؼ�Copy this control
%MN_PASTE           =40204      'ճ���ÿؼ�Paste this control
%MN_DELETE          =40205      'ɾ���ÿؼ�Delete this control
%MN_UPLAYER         =40206      '����һ��
%MN_DOWNLAYER       =40207      '����һ��
%MN_TOPLAYER        =40208      '�Ƶ�������
%MN_BOTTOMLAYER     =40209      '�Ƶ�������
%MN_CTLPROP         =40210          '�ؼ�����
'* Popup Contextmenu WM_RBUTTONDOWN on the Client���༭�����ϵĵ����˵�,���������ڵ�WM_RBUTTONDOWN
%MN_F_UNDO          = 40220     '����
%MN_F_REDO          = 40221     '����
'%MN_F_DDTCODE       = 40222     'DDT����
'%MN_F_SDKCODE       = 40223     'SDK����
%MN_F_PROP          = 40224     '��������
%keyLen = 300
%dataLen = 1000
%use_HANDLESMALL = 1
%use_HANDLEMEDIUM = 2
%use_HANDLELARGE = 3
%use_backColor = 1
%use_backBrush = 2
%WM_POSITIONSIZE       = %WM_USER +136
%TVS_NOSCROLL = &H00002000     'OR %TVS_NOSCROLL
''' tokens

%undoDepth      = 100

%moveWindow     = 1
%CreateWindowEx = 2
%DestroyWindow  = 3
%SetWindowLong  = 4
%SetWindowText  = 5
%WS_MINIMAL = %WS_CHILD OR %WS_VISIBLE OR %WS_CLIPSIBLINGS

%IMAGE_PRJ  = 0
%IMAGE_FOLD = 1
%IMAGE_BAS  = 2
%IMAGE_INC  = 3
%IMAGE_RC   = 4
%IMAGE_PRI  = 5

' Keyboard Macros - Do not change unless you also modify the SED_InsertMacroText procedure
%IDM_ALT0              = 3700
%IDM_ALT1              = 3701
%IDM_ALT2              = 3702
%IDM_ALT3              = 3703
%IDM_ALT4              = 3704
%IDM_ALT5              = 3705
%IDM_ALT6              = 3706
%IDM_ALT7              = 3707
%IDM_ALT8              = 3708
%IDM_ALT9              = 3709
%IDM_CTRL0             = 3710
%IDM_CTRL1             = 3711
%IDM_CTRL2             = 3712
%IDM_CTRL3             = 3713
%IDM_CTRL4             = 3714
%IDM_CTRL5             = 3715
%IDM_CTRL6             = 3716
%IDM_CTRL7             = 3717
%IDM_CTRL8             = 3718
%IDM_CTRL9             = 3719
' ��ӡ���ô���
%IDCPS_CBPAPERSIZE     = 6020
%IDCPS_CBPAPERBIN      = 6021
%IDCPS_PORTRAIT        = 6022
%IDCPS_LANDSCAPE       = 6023
%IDCPS_CBFONTNAME      = 6024
%IDCPS_CBFONTSIZE      = 6025
%IDCPS_CKFONTBOLD      = 6026
%IDCPS_CKFONTITALIC    = 6027
%IDCPS_CKFONTUNDERLINE = 6028
%IDCPS_MARGINLEFT      = 6029
%IDCPS_MARGINRIGHT     = 6030
%IDCPS_MARGINTOP       = 6031
%IDCPS_MARGINBOTTOM    = 6032
%IDCPS_LINEHEIGHT      = 6033