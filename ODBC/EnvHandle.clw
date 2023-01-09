
  
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

! allow one enviorment handle, if a use case shows the need for more than one
! then this could be changed.  the need for multiple env handles for the file managers
! is going to be farly uncommon. 
EnvironmentHandle SQLHENV

! module level valriable for the reference count
! all instances will use this value
handleCount       long

! file managers may be threaded and we don't 
! want different threads in the code 
critSection       CriticalSection
! --------------------------------------------------
! allocates the env handle using the module level variable
! and sets the class properties to use that handle value and
! the module level counter
! --------------------------------------------------
EnvHandle.construct  procedure()

  code
  
  critSection.wait()

  self.handle &= EnvironmentHandle  
  self.refCount &= handleCount
 
  ! if the call returns success incrment the refCount
  if (self.allocateHandle() = SQL_SUCCESS) 
    self.refCount += 1
  end
  
  critSection.release()

  return
! end construct ------------------------------------------------------------------

! --------------------------------------------------
! frees the env handle when the reference count hits zero
! if the count is greater than than zero does not free
! --------------------------------------------------
EnvHandle.destruct procedure()
  
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
! allocates the env handle and sets the 
! ODBC version.  version will be set to 3.8
! --------------------------------------------------
EnvHandle.allocateHandle procedure() !,sqlReturn,private

retv   sqlReturn,auto

  code
  
  if (self.refCount > 0) 
    return SQL_SUCCESS
  end

  ! first time then allocate the handle
  retv = SQLAllocHandle(SQL_HANDLE_ENV, 0, self.handle)
      
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
! set the version for the driver. 
! the ODBC version will be set to 3.8.  no reason 
! to use older versions.
! -------------------------------------------------- 
EnvHandle.SetVersion procedure() 

retv    sqlReturn,auto

  code

  retv  = SQLSetEnvAttr(self.handle, SQL_ATTR_ODBC_VERSION, SQL_OV_ODBC3_80, SQL_IS_INTEGER);

  if (retv <> SQL_SUCCESS) and (retv <> SQL_SUCCESS_WITH_INFO)
    halt(1, 'Valid ODBC Driver is not installed.')
  end 

  return retv
! end setVersion -------------------------------------------------------------

! --------------------------------------------------
! enable connection pooling.   the default is to use 
! pooling.  
! --------------------------------------------------
EnvHandle.EnablePooling procedure() !private,sqlReturn

retv    sqlReturn,auto

  code

  retv = SQLSetEnvAttr(self.handle, SQL_ATTR_CONNECTION_POOLING, SQL_CP_DRIVER_AWARE, SQL_IS_INTEGER)
  if (retv <> SQL_SUCCESS) and (retv <> SQL_SUCCESS_WITH_INFO)
    stop('pooling failed')
  end 

  return retv
! end EnablePooling ----------------------------------------------------------

! --------------------------------------------------
! returns the env handle 
! --------------------------------------------------  
EnvHandle.getHandle procedure() !,SQLHANDLE

  code  
  return self.handle
! end getHandle 
! ----------------------------------------------------------------------------  

! --------------------------------------------------
! returns the reference count for the objects
! --------------------------------------------------
EnvHandle.getRefCount procedure() !,long

  code 

  return self.refCount
! end getRefCount -------------------------------------------------------  