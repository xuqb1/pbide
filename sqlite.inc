'=============================
#if not %def(%SQLITE_INC)
%SQLITE_INC = 1
    '''
    ''' save as - SQLite.inc
    ''' ver 1.03
    '''
    ''' only thing needed is SQLite3 DLL
    ''' http://www.sqlite.org/download.html 
    ''' Precompiled Binaries For Windows
    ''' sqlitedll-3_2_7.zip or latest
    '''
    ''' Thanks to Terence McDonnell & Don Dickinson
    ''' just wraps Terences's & Don's work into a simple INC
    '''
    ''' SQLite 3 databases
    '''
    ''' public domain
    '''
    ''' the SQLite result is moved to a PB RecordSet array(row,col)
    '''     Zero based array
    '''     first row = Column names
    '''
    '''     redim rs(-1,-1)
    '''     X = sqlQuery( hDB, sSql, rs() )
    '''     X < 0 'error - ? sqlErrMsg(hDB)
    '''     X = 0 'zero rows returned or execute only statement
    '''     else X = result row count
    '''
    '''     doesn't hold a lock on the SQLite Table
    '''

'internal use - sqlite3.dll
declare function sqlite3_open cdecl lib "sqlite3.dll" alias "sqlite3_open" (zFilename as asciiz, hDB as long) as long
declare sub      sqlite3_close cdecl lib "sqlite3.dll" alias "sqlite3_close" (byval hDB as long)
declare function sqlite_get_table cdecl lib "sqlite3.dll" alias "sqlite3_get_table" (byval hDB as long, szSql as asciiz, lpTable as long, nRow as long, nColumn as long, lpErrMsg as long) as long
declare function sqlite_exec cdecl lib "sqlite3.dll" alias "sqlite3_exec" (byval hDB as long, szSql as asciiz, callfun as dword, firstarg as variant, lpErrMsg as long) as long
declare function sqlite_free_table cdecl lib "sqlite3.dll" alias "sqlite3_free_table" (byval lpTable as long) as long
declare function sqlite3_errmsg cdecl lib "sqlite3.dll" alias "sqlite3_errmsg" (byval hDB as long) as long
'internal use - sqlite3.dll
'
declare function sqlOpen( sFileSpec as string, byref hDB as long ) as long
declare sub sqlClose( byval hDB as long )
declare function sqlQuery( hDB as long, sSql as string, saRecSetZ() as string ) as long
declare function sqlExe( hDB as long, sSql as string ) as long
declare function sqlErrMsg( byval hDB as long ) as string
' you can delete from here - down
' these aren't needed
declare function sqlParam( byval sSql as string , saParam() as string ) as string
declare function sqlFix( byval sString as string) as string
declare function sqlDateBasicToSql( byval sDate as string ) as string
'
'
' --------------------------------------------------
function sqlOpen( sFileSpec as string, byref hDB as long ) as long
    local szFileSpec as asciiz * 500
'
    ''' create or open database
    ''' set database handle - hDB
    ''' if function doesn't =0, there was an error
'
    szFileSpec = sFileSpec
    function = sqlite3_open( szFileSpec, hDB )
end function
' --------------------------------------------------
' --------------------------------------------------
sub sqlClose( byval hDB as long )
    sqlite3_close(hDB)
end sub
' --------------------------------------------------
' --------------------------------------------------
function sqlQuery( hDB as long, sSql as string, saRecSetZ() as string ) as long
    'thanks to Terence McDonnell
    register i as long
    local lpTable    as long ptr     ' Array of fields returned from get_table (starts with column names)
    local nRows      as long         ' Number of rows in returned recordset
    local nColumns   as long         ' Number of columns in returned recordset
    local lpErrorSz  as long         ' Error msg
    local zField     as asciiz ptr   ' Field return from get_table (element in lpTable array)
    local iFields    as long         ' Number of fields returned by get table
    local iRow       as long
    local iCol       as long
    local iResult    as long         ' Number of rows or error returned by my function
'
    ''' execute query - results in saRecSetZ( RowNo&, ColNo& ) - ZERO based array
    ''' first row 0, result column names
    ''' data starts at row 1

    redim saRecSetZ()
    if sqlite_get_table(hDB, byval strptr(sSQL), lpTable, nRows, nColumns, lpErrorSz) = 0 then 'Succeeded
        if nRows = 0 then 
          function=0
          goto Bail
        end if
'
        iFields = ((nRows+1) * nColumns)-1
        redim saRecSetZ(nRows, nColumns-1)

        for i = 0 to iFields
            zField = @lpTable[i]
            saRecSetZ(iRow, iCol) = @zField
            incr iCol : if iCol = nColumns then iCol = 0
            if (i+1) mod nColumns = 0 then incr iRow
        next i
    else
        function = -1
        goto Bail
    end if
'
    ''' 0 = no rows
    ''' <0 = error
    ''' ? sqlErrMsg(hDB)
    ''' if sqlQuery(hDB, sSql, saRecSet())<0 then ? sqlErrMsg(hDB)
    function = nRows
'
Bail:
    sqlite_free_table lpTable
end function
' --------------------------------------------------
' --------------------------------------------------
function sqlExe( hDB as long, sSql as string ) as long
    local lpTable    as long ptr
    local nRows      as long
    local nColumns   as long
    local lpErrorSz  as long
    local iRow       as long
    local iCol       as long
    local saRecSetZ() as string : redim saRecSetZ()
'
    ''' execute no return SQL statement
    ''' <0 = error
    ''' ? sqlErrMsg(hDB)

    'function = sqlite_get_table(hDB, byval strptr(sSQL), lpTable, nRows, nColumns, lpErrorSz)
    function=sqlite_exec(hDB,byval strptr(sSql),0,byval 0,lpErrorSz)
    sqlite_free_table lpTable
end function
' --------------------------------------------------
' --------------------------------------------------
function sqlErrMsg( byval hDB as long ) as string
    'Thanks to Don Dickinson
    dim pzErr as asciiz ptr
'
    ''' get the SQLite error message
    ''' ? sqlErrMsg(hDB)
'
    pzErr = sqlite3_errmsg(hDB)
    if pzErr <> 0 then function = @pzErr
end function
' --------------------------------------------------
' --------------------------------------------------
    '''
    '''
    ''' you can delete from here - down
    ''' these aren't needed
    '''
    '''
' --------------------------------------------------
' --------------------------------------------------
function sqlParam( byval sSql as string , saParam() as string ) as string
    local i as long
    local saSql() as string
    local s, sParam as string
'
    ''' FAKE parameters
    ''' it escapes all single quotes
    ''' makes it easy to construct a statement
    ''' need to put '?' single quote around text column values
'
    ''' redim saParam(1 to 4)
    ''' array assign saParam() = "Bob's Best Burgers", "123 First St", "NoWhere, USA", "27"
    ''' s = sqlParam( "update tabel set AAA='?', BBB='?', CCC='?' where rowid=?", saParam() )
'
    ''' redim saParam(1 to 7)
    ''' array assign saParam() = "MyTable", "AAA", "Bob's Best Burgers", "BBB", "123 First St", "CCC", "NoWhere, USA", "27"
    ''' s = sqlParam( "update ? set ?='?', ?='?', ?='?' where rowid=?", saParam() )
'
    if instr(sSql,"?")<1 then
        ? "sqlParam() SQL statement missing '?'"
        exit function
    end if
'
    redim saSql( 1 to parsecount(sSql,"?") )
    parse  sSql, saSql(), "?"
    if ubound(saSql) <> ubound(saParam)+1 then
        ? "sqlParam() parameter missmatch"
        exit function
    end if
'
    s = saSql(1)
    for i=1 to ubound(saParam)
        sParam = saParam(i)
        replace "'" with "''" in sParam
        s = s + sParam + saSql(i+1)
    next i
'
    function = s
end function
' --------------------------------------------------
' --------------------------------------------------
function sqlFix( byval sString as string) as string
    ''' fix single quotes
    replace "'" with "''" in sString
    function = sString
end function
' --------------------------------------------------
' --------------------------------------------------
function sqlDateBasicToSql( byval sDate as string ) as string
    local s, y, m, d as string
'
    y = parse$(sDate,"-",3)
    m = parse$(sDate,"-",1)
    d = parse$(sDate,"-",2)
    s = y +"-"+ m +"-"+ d
    function = s
end function
' --------------------------------------------------

#endif
'=============================