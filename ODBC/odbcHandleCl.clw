
  member()
  
  include('odbcHandleCl.inc'),once 
  include('odbcTypes.inc'),once 

  map 
    module('odbc32')
      SQLAllocHandle(SQLSMALLINT HandleType, SQLHANDLE InputHandle, *SQLHANDLE OutputHandlePtr),SqlReturn,pascal
      SQLFreeHandle(SqlSmallInt hType, SqlHandle h),long,pascal
      SQLFreeStmt(SQLHSTMT StatementHandle, SQLSMALLINT opt),sqlReturn,pascal,proc
      SQLSetEnvAttr(SQLHENV EnvironmentHandle, SQLINTEGER Attribute,  SQLPOINTER Value, SQLINTEGER StringLength),sqlReturn,pascal
    end 
  end

OdbcHandleClType.kill procedure() !,virtual

retv   sqlReturn,auto

  code 
  
  retv = self.freeHandle()    
  if (retv = sql_Success) 
    self.handle = 0  
  end 
  
  return 
! end kill 
! ----------------------------------------------------------------------------
    
OdbcHandleClType.destruct procedure()

  code 

  self.kill() 
  
  return

OdbcHandleClType.allocateHandle procedure(SQL_HANDLE_TYPE hType, SQL_HANDLE_TYPE pType) !,sqlReturn,proc

!err ODBCErrorClType

retv   sqlReturn,auto

verId  ulong(SQL_OV_ODBC3_80)

  code
  
  self.handleType = hType
  retv = SQLAllocHandle(hType, pType, self.handle)
  
  if (retv <> Sql_Success) 
    !err.getError(pType, self.handle)
  else 
    ! set the version attribute if this is an SQL_HANDLE_ENV  
    if (hType = SQL_HANDLE_ENV) 
      retv  = SQLSetEnvAttr(self.handle, SQL_ATTR_ODBC_VERSION, verId, SQL_IS_INTEGER);
    end 
  end

  return retv
! end allocateHandle 
! ----------------------------------------------------------------------------
  
OdbcHandleClType.freeHandle procedure(long handleType) !,sqlReturn,proc,virtual,protected

retv   sqlReturn,auto

  code
  
  retv = SqlFreeHandle(handleType, self.handle)    
  
  if (retv <> Sql_Success) 
    !self.getError(SQL_HANDLE_STMT, self.hStmt)
  else 
    self.handle = SQL_NO_HANDLE
  end 
  
  return retv
! end freeHandle 
! ----------------------------------------------------------------------------

OdbcHandleClType.freeHandle procedure() !,sqlReturn,proc

retv   sqlReturn,auto

  code
  
  retv = SqlFreeHandle(self.handleType, self.handle)    
  
  if (retv <> Sql_Success) 
    !self.getError(SQL_HANDLE_STMT, self.hStmt)
  else 
    self.handle = SQL_NO_HANDLE
  end 
  
  return retv
! end freeHandle 
! ----------------------------------------------------------------------------
  
OdbcHandleClType.getHandle procedure() !,SQLHANDLE

  code
  
  return self.handle
! end getHandle 
! ----------------------------------------------------------------------------  

! ------------------------------------------------------------
! allocates a statment handle for the hDbc input.  
! use this function if multiple statment handles are needed for 
! for the connection. 
! note, the caller must track the handle output and make any 
! calls to the clear the handle or free the handle.
! if the hDbc will be closed then the driver 
! will manage any clean up.
! ------------------------------------------------------------
OdbcStmtHandleClType.allocateStmtHandle procedure(SQLHDBC hDbc, *SQLHSTMT hStmt) !,sqlReturn,proc

retv   sqlReturn,auto

  code
  
  retv = SQLAllocHandle(SQL_HANDLE_STMT, hDbc, hStmt)

  if (retv <> Sql_Success) 
    hStmt = 0
  end

  return retv

! ------------------------------------------------------------
! unbinds the columns and the parameters for the statement 
! handle.  the statment handle at this stime could be reused 
! if needed/wanted
! if the parameter input is the default then the 
! instance hStmt is used. if not the default the input 
! hStmt will be used.
! ------------------------------------------------------------
OdbcStmtHandleClType.freeBindings  procedure(SQLHSTMT hStmt = 0) !,sqlReturn,proc,virtual

retv       sqlReturn,auto
tempHandle &SQLHSTMT

  code 
   
  if (hStmt <= 0) 
    tempHandle &= self.handle
  else 
    tempHandle &= hStmt  
  end 

  ! these will not fail, and not sure what to if they do fail
  ! I suppose they could be called on an invalid handle but then 
  ! no harm is actually done.
  retv = SQLFreeStmt(tempHandle, SQL_UNBIND)
  retv = SQLFreeStmt(tempHandle, SQL_RESET_PARAMS)
  retV = SQLFreeStmt(tempHandle, SQL_CLOSE);  

  return retv
! ----------------------------------------------------------------------------

! ------------------------------------------------------------
! frees the statement handle 
! if the parameter input is the default then the 
! instance hStmt is used. if not the default the input 
! hStmt will be used.
! ------------------------------------------------------------
OdbcStmtHandleClType.freeHandle procedure(SQLHSTMT hStmt = 0) !,sqlReturn,proc,virtual

retv       sqlReturn,auto
tempHandle &SQLHSTMT

  code 
 
  if (hStmt <= 0) 
    tempHandle &= self.handle
  else 
    tempHandle &= hStmt  
  end 

  if (tempHandle = SQL_NO_HANDLE) 
    return sql_Success
  end
  
  self.freeBindings(tempHandle)

  retv = SqlFreeHandle(self.handleType, tempHandle)    

  tempHandle = SQL_NO_HANDLE

  return retv  
! end freehhandle ----------------------------------------------  
  