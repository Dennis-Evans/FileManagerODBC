
  member()
  
  include('odbcBaseCl.inc'),once
  include('odbcSqlStrCl.inc'),once
  include('dynstr.inc'),once

  map 
    module('odbc32')
      SQLExecDirect(SQLHSTMT StatementHandle, odbcWideStr StatementText, SQLINTEGER TextLength),sqlReturn,pascal,raw,name('SQLExecDirectW')
      SQLCancel(SQLHSTMT StatementHandle),sqlReturn,pascal,proc
      SQLFetch(SQLHSTMT hs),sqlReturn,pascal
      SQLMoreResults(SQLHSTMT StatementHandle),sqlReturn,pascal
      SQLFreeStmt(SQLHSTMT StatementHandle, SQLUSMALLINT Option),sqlReturn,pascal
      SQLGetData(SQLHSTMT StatementHandle, SQLUSMALLINT Col_or_Param_Num, SQLSMALLINT TargetType, *cstring TargetValuePtr, SQLLEN BufferLength,  *SQLLEN StrLen_or_IndPtr),pascal,sqlReturn,raw,name('SQLGetData');
      SQLGetDataBin(SQLHSTMT StatementHandle, SQLUSMALLINT Col_or_Param_Num, SQLSMALLINT TargetType, long TargetValuePtr, SQLLEN BufferLength,  *SQLLEN StrLen_or_IndPtr),pascal,sqlReturn,raw,name('SQLGetData');
      SQLSetConnectAttr(SQLHDBC handle, SQLINTEGER Attribute, SQLPOINTER valuePtr, SQLINTEGER stringLength),sqlReturn,pascal,name('SQLSetConnectAttrW')
      SQLEndTran(SQLSMALLINT HandleType, SQLHANDLE Handle,  SQLSMALLINT   CompletionType),sqlReturn,pascal
    end
  end

! ---------------------------------------------------------------------------
! Init 
! sets up the instance for use.  
!  allocates and init's the class to handle the sql statement or string. 
!  alllocates the error class used t oread the error message from the driver
! ---------------------------------------------------------------------------  
odbcBaseClType.init procedure(*ODBCErrorClType e)   

retv     byte(level:benign)

  code 
  
  self.sqlStr &= new(sqlStrClType)
  if (self.sqlStr &= null)
    retv = level:notify
  else 
    self.sqlStr.init()
  end  
     
  self.errs &= e

  return retv 
! end Init
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! frees the memory used
! ----------------------------------------------------------------------
odbcBaseClType.kill procedure() !,virtual  

  code 

  if (~self.sqlStr &= null)
    self.sqlStr.kill() 
    self.sqlStr &= null
  end  
  
  return
! end kill
! ----------------------------------------------------------------------
 
! ----------------------------------------------------------------------
! default destructor, calls the kill method
! ---------------------------------------------------------------------- 
odbcBaseClType.destruct procedure()  !virtual

  code 

  self.kill()
  
  return
! end destructor
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! unbinds any columns that are currently bound
! typical use is when a call that returns multiple result sets is in use.
! the first result set is processed and then call this to unbind the columns,
! then bind the caloums for the second result set, 
! repeat as needed
! ----------------------------------------------------------------------
odbcBaseClType.unBindColums procedure(SQLHSTMT hStmt)

retv sqlReturn
h  long

  code

  ! if there is a statment handle then unbind
  ! if not then nothing to do
  !if (self.conn.getHStmt() > 0) 
  !  self.conn.clearStmthandle()
  !end

  return
! end unBindColums
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! virtual place holder
! use this function to format the fields or columns read prior to the display
! ----------------------------------------------------------------------
odbcBaseClType.formatRow procedure() !,virtual  

  code
  
  ! format queue elements for display in the derived object 
  
  return
! end formatRow 
! ----------------------------------------------------------------------

! -----------------------------------------------------------------------------
! Local worker overloaded function to assign the sql str (the actual sql statement) used in this 
! call to the class member
! -----------------------------------------------------------------------------  
odbcBaseClType.setSqlCommand procedure(string s) ! sqlReturn,protected

  code 
  
  ! make sure there is one
  if (len(clip(s)) = 0)
    return sql_Error
  end 
  
  self.sqlStr.init(s)

  return sql_Success
! end setSqlCommand
! ----------------------------------------------------------------------  

! ----------------------------------------------------------------------  
! checks for the next result set, if any, and moves to the next result set
! returns true if there is more and false if not
! ----------------------------------------------------------------------  
odbcBaseClType.nextResultSet procedure(SQLHSTMT hStmt) 

retv bool(false)
res  sqlReturn,auto

  code 
 
  res = SQLMoreResults(hStmt) 
  case res 
    ! only two info messages and they will seldom happen,
    ! but read the meesage if it does happen
    of SQL_SUCCESS_WITH_INFO
      self.getError(hStmt)
      retv = true
    of sql_success
      retv = true
    else   
      self.getError(hStmt)
      retv = false;
  end
   
  return retv;
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! reads the next result set by calling the fetch function
! ------------------------------------------------------------------------------
odbcBaseClType.readnextResult procedure(SQLHSTMT hStmt,  *queue q) !,sqlReturn,virtual

retv   sqlReturn,auto

  code

  retv = self.fetch(hStmt, q)
  if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO)
    self.getError(hStmt)
  end 

  return retv
! end  readnextResult
! ------------------------------------------------------------------------------

! ----------------------------------------------------------------------
! execute a query that does not return a result set and does not use any 
! parameters
! converts the string input to a wide string and then calls execute direct
! ----------------------------------------------------------------------
odbcBaseClType.executeDirect procedure(SQLHSTMT hStmt, *IDynStr sqlCode) !,sqlReturn,virtual

retv    sqlReturn,auto
wideStr CWideStr

  code 

  if (wideStr.init(sqlCode.cstr()) <= 0)
    return sql_Error
  end 

  retv = SQLExecDirect(hStmt, wideStr.GetWideStr(), SQL_NTS)
  ! store any error or infomration messages
  if (retv <> SQL_SUCCESS)
    self.getError(hStmt)
  end 

  return retv
! --------------------------------------------------------------------

! ------------------------------------------------------------------------------
! fetch with out a result set.  
! ------------------------------------------------------------------------------
odbcBaseClType.fetch procedure(SQLHSTMT hStmt) !sqlReturn,virtual

retv   sqlReturn,auto

  code 
  
  retv = SQLFetch(hStmt)
  ! store any error or infomration messages
  if (retv <> SQL_SUCCESS)
    self.getError(hStmt)
  end 

  return retv
! end fetch
! -----------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! fetch
! reads the result set, one row at a time and places the data into the queue fields.
! Queue fields are already bound to the columns so all that is needed here is an add(q)
!
! Note, the queue fields must be bound before this method is called.
! ------------------------------------------------------------------------------
odbcBaseClType.fetch procedure(SQLHSTMT hStmt, *queue q) !sqlReturn,virtual

retv   sqlReturn,auto

  code 
  
  ! start loop and keep looping until an error or no_data is returned  
  
  loop
    retv = SQLFetch(hStmt)
    case retv 
    of SQL_NO_DATA
      ! set back to success, no_data is expected (end of result set), 
      ! but caller is going to check for success
      retv = Sql_Success    
      break
    of Sql_Success
      ! format the queue elements for display, if needed, and add the element to the queue
      self.formatRow()
      add(q)
    of Sql_Success_with_info
      self.getError(hStmt)
      ! format the queue elements for display, if needed, and add the element to the queue
      self.formatRow()
      add(q)
    else 
      ! dump the queue, something went wrong and 
      ! the code should not return a partial result set
      free(q)
      self.getError(hStmt)
      break    
    end  ! case
  end ! loop
 
  return retv
! end fetch
! -----------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! fetch
! reads the result set, one row at a time and places the data into the queue fields.
! Queue fields are already bound to the columns so all that is needed here is an add(q)
! this over load will also handle the unbound large columns, if there are any
!
! Note, the queue fields must be bound before this method is called.
! ------------------------------------------------------------------------------
odbcBaseClType.fetch procedure(SQLHSTMT hStmt, *queue q, *columnsClass cols) !sqlReturn,virtual

retv   sqlReturn,auto

  code 
  
  ! if the queue is empty then there are no large columns in the read
  ! this fetch function  does a little less work because it does not call 
  ! the processLargeOpbjects function
  if (records(cols.colb) <= 0) 
    retv = self.fetch(hStmt, q)
  else 
    loop
      clear(q)
      retv = SQLFetch(hStmt)
      case retv 
      of SQL_NO_DATA
        ! set back to success, no_data is expected (end of result set), 
        ! but caller is going to check for success
        retv = Sql_Success    
        break
      of Sql_Success
        self.processLargeObjects(hStmt, cols)  
        self.ProcessRead(q, cols)
      of Sql_Success_with_info
        self.getError(hStmt)  
        self.processLargeObjects(hStmt, cols)  
        self.ProcessRead(q, cols)
      else 
        ! dump the queue, something went wrong and 
        ! the code should not return a partial result set
        free(q)
        self.getError(hStmt)
        break    
      end  ! case
    end ! loop
  end ! if records
  
  return retv
! end fetch
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! Binds the columns from the queue to the columns in the result set
! then calls fetch to read the result set
! -----------------------------------------------------------------------------
odbcBaseClType.fillResult procedure(SQLHSTMT hStmt, *columnsClass cols, *queue q, long setId = 1) !,sqlReturn,private

retv   sqlReturn,auto

  code 
 
  ! bind the columns just before the fetch, not needed for the execute query calls 
  ! so do it here, 
  retv = cols.bindColumns(Hstmt)
  case retv 
    of Sql_Success 
      retv = self.fetch(hStmt, q, cols)
    of SQL_SUCCESS_WITH_INFO
      ! bind columns returned with an info message
      self.getError(hStmt)
      retv = self.fetch(hStmt, q, cols)
    else 
      ! we went south 
      self.getError(hStmt)
  end  

  return retv
! end fillResult
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! reads the large objects from the result set, if there are any, often there 
! will not be any large vzrchar(x) or varbinary(x) columns in the reads
! but o the occasion they are present deal with them here 
! and in the calling function, the values assigned here must be pushed back to the 
! using queue.
! -----------------------------------------------------------------------------
odbcBaseClType.processLargeObjects procedure(SQLHSTMT hStmt, *columnsClass cols) 

x  long,auto

  code

  loop x = 1 to records(cols.colb) 
    get(cols.colb, x)
    case cols.colb.ColType
      of SQL_LONGVARCHAR
        self.readLargeChar(hStmt, cols)
      of  SQL_LONGVARBINARY
        self.ReadLargeBinary(hStmt, cols)
    end ! end case
  end ! loop

  return
! end processLargeObjects
! -------------------------------------------------------------------------

! -------------------------------------------------------------------------
! reads a large character column. 
! typically this will be varchar(n) where n is a a number larger than the normal 
! character column size.  
! -------------------------------------------------------------------------
odbcBaseClType.readLargeChar procedure(SQLHSTMT hStmt, *columnsClass cols) 

retv     sqlReturn,auto
colLen   SQLLEN(0)

  code

  cols.colb.colSize = self.getCharLength(hStmt, cols)
 
  ! if greater than zero then some data is in the buffer, it may be 1 byte or it may be many bytes
  ! allocate the storage and read it 
  if (cols.colb.colSize > 0) 
    cols.colb.charHolder  &= new(cstring(cols.colb.colSize + 1))
    retv = SQLGetData(hStmt, cols.colB.colId, SQL_C_CHAR, cols.colb.charHolder, cols.colb.colSize + 1,  colLen)
    case retv 
      of SQL_SUCCESS
        put(cols.colb)
      of SQL_SUCCESS_WITH_INFO
        self.getError(hStmt)
        put(cols.colb)
      else 
        self.getError(hStmt)
    end
  end 

  return
! end ReadLargeChar
! -------------------------------------------------------------------------

! -------------------------------------------------------------------------
! gets the length, in the number of bytes, for a large character column. 
! typically this will be varchar(n) where n is a a number larger than the normal 
! character column size.  will almost always be used for varchar(max)
! other sizes use will vary
! -------------------------------------------------------------------------
odbcBaseClType.getCharLength procedure(SQLHSTMT hStmt, *columnsClass cols)

retv   sqlReturn,auto
outStr &cstring     ! used t oread the size
colLen SQLLEN,auto  ! output parameter that contains the size needed for the buffer

  code 

  ! get the number of bytes for the column.  the outStr parameter is 
  ! null so the call just sends back the size of the string
  ! note the out value of colLen can be the length, SQL_NULL_DATA or SQL_NO_TOTAL
  retv = SQLGetData(hStmt, cols.colb.colId, SQL_C_CHAR, outStr, 0,  colLen)
  if (retv <> SQL_SUCCESS) 
    self.getError(hStmt)
  end  

  return colLen
! getCharLength
! -------------------------------------------------------------------------

! -------------------------------------------------------------------------
! reads a large binary column from the result set.  this will be a binary(x) 
! or a varbinary(x) on the back end.
! -------------------------------------------------------------------------
odbcBaseClType.ReadLargeBinary procedure(SQLHSTMT hStmt, *columnsClass cols)

retv     sqlReturn,auto
colLen   SQLLEN(0)

  code

  cols.colb.colSize = self.getBinaryLength(hStmt, cols)
 
  ! if greater than zero then some data is in the buffer, it will 1 to n bytes
  ! allocate the storage and read it 
  if (cols.colb.colSize > 0) 
    cols.colb.binaryHolder &= new(string(cols.colb.colSize))
    retv  = SQLGetDataBin(hStmt, cols.colB.colId, SQL_C_BINARY, address(cols.colb.binaryHolder), cols.colb.colSize,  colLen)
    case retv 
      of SQL_SUCCESS
        put(cols.colb)
      of SQL_SUCCESS_WITH_INFO
        self.getError(hStmt)
        put(cols.colb)
      else 
        self.getError(hStmt)
    end
  end 

  return
! end ReadLargeBinary
! -------------------------------------------------------------------------

! -------------------------------------------------------------------------
! gets the length, in the number of bytes, for a blob or binary column. 
! typically this will be varbinary(n) where n is the size of the array
! typical use will be for varbinary(max) but can be used for fixed size binary(x)
! columns 
! -------------------------------------------------------------------------
odbcBaseClType.getBinaryLength procedure(SQLHSTMT hStmt, *columnsClass cols)

retv   sqlReturn,auto
outByte &byte
colLen SQLLEN,auto  ! output parameter that contains the size needed for the buffer

  code 
  
  ! get the number of bytes for the column.  the outByte parameter is 
  ! null so the call just sends back the size of the array
  ! note the out value of colLen can be the length, SQL_NULL_DATA or SQL_NO_TOTAL
  retv = SQLGetDataBin(hStmt, cols.colb.colId, SQL_C_BINARY, address(outByte), 0,  colLen)
  if (retv <> SQL_SUCCESS) 
    self.getError(hStmt)
  end  

  return colLen
! getBinaryLength
! -------------------------------------------------------------------------

odbcBaseClType.ProcessRead procedure(*queue q, *columnsClass cols) ! virtual,protected

  code

  if (cols.getallowNulls() = true)
    !cols.setDefaultNullValue(q)
  end
  ! format the queue elements for display, if needed, and add the element to the queue
  self.formatRow()
  add(q)

  return
! end ProcessRead ------------------------------------------------------------

! -----------------------------------------------------------------------------
! call the error class to read the error information
! -----------------------------------------------------------------------------
odbcBaseClType.getError procedure(SQLHSTMT hStmt)  ! protected

  code 
  
  self.errs.getError(SQL_HANDLE_STMT, hStmt)

  return
! end getError  
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! call the error class to read the error information for a connection
! -----------------------------------------------------------------------------
odbcBaseClType.getConnectionError procedure(SQLHDBC hDbc)  ! protected

  code 
  
  self.errs.getError(SQL_HANDLE_DBC, hDbc)

  return
! end getError  
! -----------------------------------------------------------------------------

