
  member()
  
  include('odbcConn.inc'),once 
  include'odbcTypes.inc'),once
  include('svcom.inc'),once

  map 
    module('odbc32')
      SQLConnect(SQLHDBC ConnectionHandle, *SQLCHAR ServerName, SQLSMALLINT NameLength1, long UserName, SQLSMALLINT NameLength2, long Authentication, SQLSMALLINT NameLength3),sqlReturn,pascal,raw
      SQLDriverConnect(SQLHDBC ConnectionHandle, SQLHWND WindowHandle, long InConnectionString, SQLSMALLINT StringLength1, long  OutConnectionString, SQLSMALLINT BufferLength, *SQLSMALLINT StringLength2Ptr, SQLUSMALLINT DriverCompletion),sqlReturn,pascal,raw,Name('SQLDriverConnectW')
      SQLDisconnect(SQLHDBC ConnectionHandle),sqlReturn,pascal   
      SQLGetInfo(SQLHDBC hDbc, long attrib, *cstring valuePtr, long buffLength, *long strLenPtr),long,pascal,raw
      SQLSetEnvAttr(SQLHENV EnvironmentHandle, SQLINTEGER Attribute,  SQLPOINTER Value, SQLINTEGER StringLength),sqlReturn,pascal
      SQLSetStmtAttr(SQLHSTMT hStmt, SQLINTEGER attribute, SQLPOINTER valuePtr, SQLINTEGER id),long,pascal,Name('SQLSetStmtAttrW')
      SQLSetConnectAttr(SQLHDBC Handle, SQLINTEGER Attribute, SQLPOINTER ValuePtr, SQLINTEGER StringLength),sqlReturn,pascal,name('SQLSetConnectAttrW')
      SQLGetConnectAttr(SQLHDBC Handle, SQLINTEGER Attribute, *SQLPOINTER ValuePtr, SQLINTEGER BufferLength, SQLINTEGER StringLengthPtr),sqlReturn,pascal,name('SQLGetConnectAttrW')
    end 
  end

eNoWindow  equate(0)

ODBCConnectionClType.init procedure() !,sqlReturn

retv    sqlReturn,auto

  code 
  
  self.hEnv &= new(EnvHandle) 
  if (self.hEnv &= null) 
    return sql_Error
  end
      
  self.hDbc &= new(OdbcHandleClType)
  if (self.hDbc &= null) 
    return sql_Error
  end   

  self.hStmt &= new(OdbcStmtHandleClType)
  if (self.hStmt &= null) 
    return sql_Error
  end   
  
  return Sql_Success
! end init
! -------------------------------------------------------------------------

ODBCConnectionClType.init procedure(baseConnStrClType connString) !,sqlReturn

retv    sqlReturn,auto

  code 

  if (connString &= null) 
    return sql_Error
  end    

  self.connStr &= connString

  retv = self.Init()
    
  return retv
! end init
! -------------------------------------------------------------------------
    
ODBCConnectionClType.kill procedure()

  code 
  
  if (~self.hDbc &= null)
    dispose(self.hdbc)
    self.hDbc &= null
  end 
  
  if (~self.hEnv &= null)
    dispose(self.hEnv)
    self.hEnv &= null
  end 
  
  self.connStr &= null
  if (~self.hStmt &= null)
    dispose(self.hStmt)
    self.hStmt &= null
  end

  return 
! end kill 
! -------------------------------------------------------------------------  

ODBCConnectionClType.gethEnv procedure() !,SQLHEnv

  code 
  return self.hEnv.gethandle()
! end gethEnv -----------------------------------------------------

ODBCConnectionClType.gethDbc procedure() !,SQLHDBC

dbHandle    SQLHDBC,auto

  code 

  if (~self.hdbc &= null)
    dbHandle = self.hDbc.getHandle()
  else 
    dbHandle = SQL_NO_HANDLE  
  end 
  
  return dbHandle
 ! end gethDbc -----------------------------------------------------

ODBCConnectionClType.gethStmt procedure() !,SQLHStmt

  code 
  return self.hStmt.gethandle()
 ! end gethStmt -----------------------------------------------------

! ---------------------------------------------------------------------------
! checks to see if the connection is dead or active.  the call to get the 
! connection attribute sets the status field to false if the connection is active
! and to true if the connection is dead, so the return value is set to false 
! at the start and if the connection is not dead then it is set to true
! ---------------------------------------------------------------------------
ODBCConnectionClType.isConnected procedure(SQLHDBC dbHandle = 0) !,bool

res      sqlReturn,auto
retv     bool(false)    ! assume the connection is dead or not active at the start
status   long,auto  

  code

   if (dbHandle <= 0) 
    ! get the current connection handle, check for a value.  
    dbHandle = self.gethDbc()
  end  
  ! if > 0 then it has been connected
  if (dbHandle > 0) 
    ! get the staus, if the res is not good then just assume the connection is not active
    ! note this does not make a trip to the server, the code queries the driver maanger
    ! and retuns the status of the last call, good or bad.  
    res = SQLGetConnectAttr(dbHandle, SQL_ATTR_CONNECTION_DEAD, status, SQL_IS_POINTER, 0)
    if (res = Sql_Success) 
      if (status = SQL_CD_FALSE)
        ! connection is active so return true
        retv = true
      end
    end  ! if (res = Sql_Success) 
  end ! if (dbHandle > 0) 

  return retv
! ---------------------------------------------------------------------------

ODBCConnectionClType.setOdbcVersion procedure(long verId) 

retv            sqlReturn

  code

  retv  = SQLSetEnvAttr(self.gethEnv(), SQL_ATTR_ODBC_VERSION, verId, SQL_IS_INTEGER)

  return retv
! --------------------------------------------------------------------------

! --------------------------------------------------------------------------
! connect to the database 
! 
! if the statement parameter is true, the default, a statement handle is allocated 
! if false then the statement handle will need t obe allocated from the calling code
!
! typically better t ouse the default and just let this block allocate the statement handle
! one less step for the using code
! --------------------------------------------------------------------------
ODBCConnectionClType.connect procedure()

retv       sqlReturn,auto
outLength  sqlsmallint,auto

! 1,024 is the size shown in the doc's some what arbitray
! but we don't use it, so no one cares
outConnStr cstring(1024)

wideStr    Cwidestr

  code 
  
  ! if the handle has not been allocated then allocate one
  ! if it has been allocated then do not allocate
  if (self.hdbc.getHandle() = SQL_NO_HANDLE)
    retv = self.hDbc.allocateHandle(SQL_HANDLE_DBC, self.hEnv.getHandle())
  else 
    retv = SQL_SUCCESS
  end
  
  if (retv = sql_Success) or (retv = sql_success_with_info)
    ! make the connection string a wide string for the ODBC 
    if (wideStr.init(self.connStr.ConnectionString()) = false) 
      return sql_Error
    end
     retv = SQLDriverConnect(self.hDbc.getHandle(), eNoWindow, widestr.GetWideStr(), SQL_NTS, 0, size(outConnStr), outLength, SQL_DRIVER_NOPROMPT)
  end   

  return retv 
! end connect 
! ----------------------------------------------------------------------

! ------------------------------------------------------
! Allocates a statement handle for the connection
! called by the connect function or called 
! by the using code to create a statement handle
! note, the default is to create a statement handle with the 
! hDbc, most actions will only need a single hStmt
! and will close the connection when the action completes.
! see the over loaded function if multiple handle are needed.
! ------------------------------------------------------
ODBCConnectionClType.AllocateStmtHandle procedure() !,protected,virtual

retv sqlReturn,auto

  code

  retv = self.hStmt.AllocateHandle(SQL_HANDLE_STMT, self.hDbc.getHandle())

  return retv
! end  AllocateStmtHandle ----------------------------------------------

! ------------------------------------------------------
! Allocates a statement handle for the connection
! called by user to allocate a more than the default statement handle.
! this function is not used by the connect call.
! 
! note, the caller should check the value of the output hStmt
! if zero or less the handle was not allocated
! ------------------------------------------------------
ODBCConnectionClType.AllocateStmtHandle procedure(*SQLHSTMT hStmt) !,protected,virtual

retv sqlReturn,auto

  code

  retv = self.hStmt.AllocateHandle(self.hDbc.getHandle(), hStmt)
  if (retv <> sql_Success) and (retv <> Sql_Success_With_Info)
    hStmt = SQL_NO_HANDLE
  end   

  return retv
! end  AllocateStmtHandle ----------------------------------------------

! ------------------------------------------------------
! frees the statement handle.  if this is called then 
! the connection no longer has a hStmt and one must be 
! allocated before the connection is used again.
! ------------------------------------------------------
ODBCConnectionClType.freeStmthandle procedure(SQLHSTMT hStmt = 0) 

  code

  self.hStmt.freeHandle(hStmt)

  return
! end  closeStmthandle -------------------------------------------------

! ------------------------------------------------------
! clears the statement handle or any binding, result sets, ... 
! call this when the statement handle is going to be used 
! for multiple calls.  this allows the connection to perform 
! more than one action and not need to allocate a hStmt for each,
! saves a little overhead
!
! Note, some actions can be executed in sequence and not need 
! the hStmt cleared, some cannot be.  Good practice would be to 
! call the clear function between any calls, but dpending 
! on the actual action executed it may not be required.
!
! if the parameter input is the default then the 
! instance hStmt is used. if not the default the input 
! hStmt will be used.
! ------------------------------------------------------
ODBCConnectionClType.clearStmthandle procedure(SQLHSTMT hStmt = 0) 

  code
 
  self.hstmt.freeBindings(hStmt)

  return 
! end clearStmthandle --------------------------------------------------

! ----------------------------------------------------------------------
! disconnect from the database, freeing all hStmt handles in use
! on the connection.
! if the dbc handle is not valid then nothing is done
! ----------------------------------------------------------------------
ODBCConnectionClType.Disconnect procedure()

! assume success at the start
retv      sqlReturn(sql_Success)
h         SQLHDBC,auto

  code 

  h = self.hDbc.getHandle()
  if (h > 0)
    retv = SQLDisconnect(h)
  end 

  return retv
! end Disconnect 
! ----------------------------------------------------------------------

ODBCConnectionClType.asyncOn procedure()

retv   sqlReturn,auto

  code 

  retv = SQLSetConnectAttr(self.hDbc.getHandle(), SQL_ATTR_ASYNC_DBC_FUNCTIONS_ENABLE, SQL_ASYNC_DBC_ENABLE_ON, SQL_IS_INTEGER)

  return retv