
  
  member()

  include('EnvHandle.inc'),once 
  include('cwsynchc.inc'),once
  
  map 
    module('odbc32')
      SQLAllocHandle(SQLSMALLINT HandleType, SQLHANDLE InputHandle, *SQLHANDLE OutputHandlePtr),SqlReturn,pascal
      SQLSetEnvAttr(SQLHENV EnvironmentHandle, SQLINTEGER Attribute,  SQLPOINTER Value, SQLINTEGER StringLength),sqlReturn,pascal
      SQLFreeHandle(SqlSmallInt hType, SqlHandle h),sqlReturn,pascal
    end 
  end

! file managers may be threaded and we don't 
! want different threads in the code 
critSection       &CriticalSection

! --------------------------------------------------
! allocates the env handle using the module level variable
! and sets the class properties to use that handle value and
! the module level counter
! --------------------------------------------------
DbcHandle.construct  procedure()

  code
  
  ! if null then make one
  if (critSection &= null)
    critSection &= new(CriticalSection) 
  end
    
  critSection.wait()

  self.hdbcQue &= new(HandleQueue)

  critSection.release()

  return
! end construct ------------------------------------------------------------------

! --------------------------------------------------
! frees the env handle when the reference count hits zero
! if the count is greater than than zero does not free
! --------------------------------------------------
DbcHandle.destruct procedure()
  
  code 

  critSection.wait()

  ! decrement the count
  self.refCount -= 1

  ! if there is an instance still using the handle do not free
  if (self.refCount <= 0) 
    if (SqlFreeHandle(SQL_HANDLE_ENV, self.handle) <> Sql_Success)
    !self.getError(SQL_HANDLE_STMT, self.hStmt)
    else 
      self.handle = SQL_NO_HANDLE
    end 
  end  ! if (self.refCount <= 0) 
  
  critSection.release()

  return
! end destruct --------------------------------------------------------------------

! --------------------------------------------------
! allocates the hDbc  handle
! --------------------------------------------------
DbcHandle.allocateHandle procedure(SQLHENV hEnv) !,sqlReturn,protected

retv   sqlReturn,auto

  code
  
  self.hdbcQue.Label = 'Default'

  if (self.refCount > 0) 
    return SQL_SUCCESS
  end
  

  ! first time then allocate the handle
  retv = SQLAllocHandle(SQL_HANDLE_DBC, 0, self.handle)
      
  if (retv = SQL_SUCCESS) 
    retv = self.setVersion()
  end 

  if (retv = SQL_SUCCESS) 
    retv = self.EnablePooling();
  end

  return retv
! end allocateHandle 
! ----------------------------------------------------------------------------

! --------------------------------------------------
! returns the env handle 
! --------------------------------------------------  
DbcHandle.getHandle procedure() !,SQLHANDLE

  code  
  return self.handle
! end getHandle 
! ----------------------------------------------------------------------------  

