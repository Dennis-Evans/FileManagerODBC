
  member()
  
  include('odbcTransCl.inc'),once

  map 
    module('odbc32')
      SQLSetConnectAttr(SQLHDBC handle, SQLINTEGER Attribute, SQLPOINTER valuePtr, SQLINTEGER stringLength),sqlReturn,pascal,name('SQLSetConnectAttrW')
      SQLEndTran(SQLSMALLINT HandleType, SQLHANDLE Handle,  SQLSMALLINT   CompletionType),sqlReturn,pascal
    end
  end

! ---------------------------------------------------------------------------
! Init 
! sets up the instance for use.  
! ---------------------------------------------------------------------------  
odbcTransactionClType.init procedure(SQLHDBC hDbc)   

retv     byte(level:benign)

  code 
  
  self.hDbc = hDbc
  ! set the default isolation level, typically it is best to 
  ! leave as read committed but it can be set to what ever is needed
  self.defaultIsolationLvl = SQL_TRANSACTION_READ_COMMITTED

  return retv 
! end Init
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! sets the handle to be invalid 
! ----------------------------------------------------------------------
odbcTransactionClType.kill procedure() !,virtual  

  code 

  self.hDbc = SQL_NO_HANDLE

  return
! end kill
! ----------------------------------------------------------------------
 
! ----------------------------------------------------------------------
! default destructor, calls the kill method
! ---------------------------------------------------------------------- 
odbcTransactionClType.destruct procedure()  !virtual

  code 

  self.kill()
  
  return
! end destructor
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! sets the default value for the isolation level
! read committed is the normal default and is set in the constructor.
! read committed is typically the best option for the isolation level.
! 
! There may be use cases where some other default value is needed/wanted.
! 
! However, if altering the default be sure you understand the implications 
! of using the other types of isolation levels.
! ----------------------------------------------------------------------
odbcTransactionClType.setDefaultIsolationLevel procedure(SQLINTEGER level)

  code 

  self.defaultIsolationLvl = level

  return
! end  setDefaultIsolationLevel
! ---------------------------------------------------------------------- 

! ----------------------------------------------------------------------
! sets the isolation level for the hDbc input.  the connection must not 
! have any open transaction when this is called.  the connection can be opend 
! or closed. 
! ----------------------------------------------------------------------
odbcTransactionClType.setIsolationLevel procedure(long level) !,sqlReturn,protected

retv  sqlReturn,auto

  code

  retv = SQLSetConnectAttr(self.hDbc, SQL_ATTR_TXN_ISOLATION, level, SQL_IS_INTEGER) 
  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO) 
    self.currentIsolationLvl = level
  end 

  return retv
! end  setIsolationLevel
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! the next set of functions sets the isolation level for the hDbc input to 
! the level indicated by the function name.  typically the default value of 
! read committed is all that will be used, there are common use cases for 
! serializable transaction and some other, less common use cases, 
! for the other tpyes.
! ----------------------------------------------------------------------
odbcTransactionClType.setIsolationSerializable procedure() !,sqlReturn 

retv sqlReturn,auto 

  code 

  retv = self.setIsolationLevel(SQL_TRANSACTION_SERIALIZABLE) 

  return retv;
! end setIsolationSerializable
! --------------------------------------------------------------

odbcTransactionClType.setIsolationReadUncommitted procedure() !,sqlReturn

retv sqlReturn,auto 

  code 

  retv = self.setIsolationLevel(SQL_TRANSACTION_READ_UNCOMMITTED) 

  return retv;
! end setIsolationReadUncommitted
! --------------------------------------------------------------

odbcTransactionClType.setIsolationReadCommitted procedure() !,sqlReturn

retv sqlReturn,auto 

  code 

  retv = self.setIsolationLevel(SQL_TRANSACTION_READ_COMMITTED) 

  return retv;
! end setIsolationReadCommitted
! --------------------------------------------------------------

odbcTransactionClType.setIsolationRepeatabelread procedure() !,sqlReturn

retv sqlReturn,auto 

  code 

  retv = self.setIsolationLevel(SQL_TRANSACTION_REPEATABLE_READ) 

  return retv;
! end setIsolationRepeatabelread
! --------------------------------------------------------------

! ----------------------------------------------------------------------
! begins a transaction for the connection handle input.  
! this actually turns off auto-commit mode.
! 
! note, if this is called then you must call the commit or roll back functions
! when the work is completd.  failing to end a transaction will cause 
! bad things to happen
! ----------------------------------------------------------------------
odbcTransactionClType.beginTrans procedure()

autoOff   SQLINTEGER(SQL_AUTOCOMMIT_OFF)
retv      sqlReturn,auto

  code

  retv = SQLSetConnectAttr(self.hDbc, SQL_ATTR_AUTOCOMMIT, autoOff, SQL_IS_INTEGER)  

  return retv
! end beginTransaction
! -----------------------------------------------------------------------

! ----------------------------------------------------------------------
! commits a transaction for the connection handle input.  
! ----------------------------------------------------------------------
odbcTransactionClType.Commit procedure()

retv      sqlReturn,auto

  code

  retv = self.EndTrans(SQL_COMMIT)

  return retv
! end CommitTransaction 
! ----------------------------------------------------------------------  

! ----------------------------------------------------------------------
! rolls back a transaction for the connection handle input.  
! ----------------------------------------------------------------------
odbcTransactionClType.Rollback procedure()

retv      sqlReturn,auto

  code

  retv = self.EndTrans(SQL_ROLLBACK)

  return retv
! end RollbackTransaction 
! ----------------------------------------------------------------------  

! ----------------------------------------------------------------------
! ends a transaction for the connection handle input.  
! called from the commit or rollback functions.
! ----------------------------------------------------------------------
odbcTransactionClType.EndTrans procedure(long committRollBack) !sqlReturn,private

autoOn    SQLINTEGER(SQL_AUTOCOMMIT_ON)
retv      sqlReturn,auto

  code

  retv = SQLEndTran(SQL_HANDLE_DBC, self.hDbc, committRollback)
  ! if it was ended then reset to the default
  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO) 
    ! if the current value is not the default set it back to the default
    ! if it is the default then there is nothing to do
    if (self.currentIsolationLvl <> self.defaultIsolationLvl)
      retv = SQLSetConnectAttr(self.hDbc, SQL_ATTR_TXN_ISOLATION, self.defaultIsolationLvl, SQL_IS_INTEGER)
      self.currentIsolationLvl = self.defaultIsolationLvl
    end  
  end ! 
    
  ! turn it back on for the next call
  retv = SQLSetConnectAttr(self.hDbc, SQL_ATTR_AUTOCOMMIT, autoOn, SQL_IS_INTEGER)  

  return retv
! end EndTransaction 
! ---------------------------------------------------------------------- 