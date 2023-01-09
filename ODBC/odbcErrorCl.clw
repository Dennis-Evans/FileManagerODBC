
  member()
  
  include('odbcErrorCl.inc'),once 

  map 
    module('odbc32')
      SQLGetDiagField(SQLSMALLINT HandleType, SQLHANDLE Handle, SQLSMALLINT RecNumber, SQLSMALLINT DiagIdentifier, *SQLPOINTER DiagInfoPtr, SQLSMALLINT BufferLength, *SQLSMALLINT StringLengthPtr),sqlReturn,pascal,name('SQLGetDiagFieldW')
      SQLGetDiagRec(SQLSMALLINT HandleType, SQLHANDLE Handle, SQLSMALLINT RecNumber, odbcWideStr errState, *SQLINTEGER NativeErrorPtr, odbcWideStr MessageText, SQLSMALLINT BufferLength, *SQLSMALLINT TextLengthPtr),sqlReturn,pascal,raw,name('SQLGetDiagRecW')
    end 
  end

! ---------------------------------------------------------------------------

OdbcErrorClType.init procedure() !,sqlReturn

  code 

  if (self.makeObjects() <> sql_Success)
    return sql_error
  end 
  
  return sql_Success
! end init 
! ----------------------------------------------------------------------------
  
OdbcErrorClType.kill procedure() !,virtual

  code 
  
  self.destroyObjects()
  
  return
! end kill
! ----------------------------------------------------------------------------
  
OdbcErrorClType.destruct procedure()

  code 
  
  self.kill()
  
  return
! end destruct 
! ----------------------------------------------------------------------------  
  
OdbcErrorClType.getNumberMsg  procedure() !,byte

  code 

  return self.errorCount
! end getErrorCount ---------------------------------------------------------

OdbcErrorClType.getDataBaseError procedure(ODBCConnectionClType conn) !,sqlReturn,proc

retv   sqlReturn,auto

  code 

  retv = self.getError(SQL_HANDLE_DBC, conn.gethDbc())
  
  return retv 
! end getDatabaseError
! ----------------------------------------------------------------------------  

OdbcErrorClType.getErrorGroup  procedure(long ndx, *string stateText, *string msgText)

  code

  get(self.errorMsgQ, ndx)
  if (errorcode() = 0)
    stateText = self.errorMsgQ.sqlState
    msgText = self.errorMsgQ.messagetext
  else 
    stateText = '01000'
    msgText = 'No Message'
  end   
  
  return 
! end getErrorGroup ----------------------------------------------------------

! ----------------------------------------------------------------------------  
! reads the error and information messages from the list using the 
! SQLGetDiagRec function.  each message is added to the queue for display, if needed
! this function clears the queue of messages and then finds the number 
! of messages for the handle type input.  if one or mor then the code loops 
! over the list and places the wide string output values into the queue
! after they are converted to ANSI strings
! ----------------------------------------------------------------------------       
OdbcErrorClType.getError procedure(SQLSMALLINT HandleType, SQLHANDLE Handle)  

retv      sqlReturn,auto

count     long,auto

! function inputs
claStateMsg  cstring(12)    ! sql state is a static size
claErrMsg    cstring(2001)  ! this is large enough
! wide string instances 
stateMsg  CWideStr
errMsg    CWideStr
! used to convert a wide string to a ansi string
outState  CStr
outErr    CStr

tempholder bool

  code 
  
  self.freeErrorMsgQ()
  self.getDiagRecCount(handleType, handle)

  loop count = 1 to self.errorCount
    ! allocate the inpusts
    claStateMsg = all(' ')
    tempholder = statemsg.Init(claStateMsg)
    claErrMsg = all(' ')
    tempholder = errMsg.Init(claErrMsg)

    ! call the function with the errMsg parameter set to null and the call will output the number of bytes 
    ! for the length of the message, allocate and call a second time
    retv = SQLGetDiagRec(handleType, handle, count, stateMsg.getWideStr(), self.errorMsgQ.NativeErrorPtr, errMsg.getWideStr(), 2000, self.errorMsgQ.textLengthPtr)
    ! if call worked thne from wide string to ansi and place i nthe queue 

    if (retv = sql_Success) or (retv = sql_success_with_info)
      tempholder = outState.Init(stateMsg)
      self.errorMsgQ.sqlState = outState.getCStr()
      tempholder = outErr.Init(errMsg)
      self.errorMsgQ.MessageText &= new(cstring(tempholder + 2))
      self.errorMsgQ.MessageText = outErr.getCStr()
      add(self.errorMsgQ)
    end  ! if
  end ! loop

  return self.errorCount
! end getError
! ----------------------------------------------------------------------
  
! ----------------------------------------------------------------------  
! calls the SQLGetDiagField function to get the number of messages 
! they maybe error messages or just information messages
! called internally by the getError function
! this call can return other information but they are not currently used
!
! the call can return the number of rows affected by an insert, update or delete action
! and some other types of information
! ----------------------------------------------------------------------    
OdbcErrorClType.getDiagRecCount procedure(SQLSMALLINT HandleType, SQLHANDLE Handle) !,long,private  

retv            sqlReturn,auto
StringLengthPtr short,auto
 
  code 

  ! just find the number of errors in the list
  retv = SQLGetDiagField(HandleType, Handle, 0, SQL_DIAG_NUMBER, self.errorCount, 0, StringLengthPtr)
  
  case retv 
  !of SQL_SUCCESS  
    ! valid result, nothing to do
  !of SQL_SUCCESS_WITH_INFO 
    ! not currently using this value
  of SQL_INVALID_HANDLE
    ! this is a programing error not a run time error
    ! just show a message for testing
    message('Call to getDiagField with an invalid handle.', 'Invalid Handle', icon:exclamation)
    self.freeErrorMsgQ()
  of SQL_ERROR 
    ! this is a programing error not a run time error
    message('Call to getDiagField returned SQL_ERROR, verify the handle type used is the correct type.', 'SQL error', icon:exclamation)
    self.freeErrorMsgQ()
  end 
    
  return retv
! end getDiagRecCount
! ----------------------------------------------------------------------  

! ----------------------------------------------------------------------
! displays an erro on the screen using a simple message call
! shows all the errors and information messages from the call for the handle type
! ----------------------------------------------------------------------
OdbcErrorClType.showError procedure()

count     long,auto

  code 

  loop count = 1 to self.errorCount
    get(self.errorMsgQ, count)
    message('ODBC Error State  -> ' & self.errorMsgQ.sqlState & '|System Error Code -> ' & self.errorMsgQ.NativeErrorPtr &  '|Error Message  -> ' & clip(self.errorMsgQ.MessageText), 'Database Error', icon:exclamation)
  end  
  
  return   
! end showError
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! free the queue of error messages
! this is called each time the getError function is called.  
! the messages from the most recent error are deleted, 
! ----------------------------------------------------------------------  
OdbcErrorClType.freeErrorMsgQ procedure() !,private

x  long,auto

  code 

  self.errorCount = 0

  loop x = 1 to records(self.errorMsgQ)
    get(self.errorMsgQ, x)
    dispose(self.errorMsgQ.MessageText)
  end  

  free(self.errorMsgQ)
  clear(self.errorMsgQ)
  
  return   
! end freeErrorMsgQ
! ----------------------------------------------------------------------  
  
! ----------------------------------------------------------------------
! allocates the message queue 
! ----------------------------------------------------------------------  
OdbcErrorClType.makeObjects procedure() !,sqlReturn,private

  code 

  self.errorMsgQ &= new(OdbcErrorQueue)
  if (self.errorMsgQ &= null) 
    return sql_Error
  end 
    
  return sql_Success
! end makeObjects 
! -------------------------------------------------------------------------
  
! ----------------------------------------------------------------------
! does the clean up, called by the kill method
! ----------------------------------------------------------------------  
OdbcErrorClType.destroyObjects procedure() !,sqlReturn,private

  code 
  
  if (~self.errorMsgQ &= null) 
    self.freeErrorMsgQ()
    dispose(self.errorMsgQ) 
    self.errorMsgQ &= null
  end 
    
  return
! end destroyObjects 
! -------------------------------------------------------------------------