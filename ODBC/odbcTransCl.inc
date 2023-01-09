!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCTransClassPresent_)
_ODBCTransClassPresent_   EQUATE(1)

  include('odbcTypes.inc'),once

! -------------------------------------------------------------------------------
! object to handle the transaction needs for the various calls.
! the code will begin a transaction, commit or roll back a transaction
! and the various isolation levels can be set for selected transactions.
! 
! the object, by default, uses the read committed isolation level.  the other 
! levels can be used if needed/wanted.  if a level other than the default is used 
! the code will set the isoaltion level back to the default when the transaction
! end, with either a commit or a roll back.
! -------------------------------------------------------------------------------
odbcTransactionClType class,type,module('odbcTransCl.clw'),link('odbcTransCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! ------------------------------------------------------
! the connection handle this instance will use for 
! the various calls.  
! ------------------------------------------------------
hDbc     SQLHDBC,private

! -------------------------------------------------------------------
! the isolation level set by the user for 
! a transaction. used to reset to the default when a transaction ends
! -------------------------------------------------------------------
currentIsolationLvl SQLINTEGER

! ------------------------------------------------------------------
! the default isolation level that will be used by the transactions
! set this to what ever is needed/wnated 
!
! Note, unless there are very specific reasons to change the default, 
! and you understand the implications of the different isolation levels
! the default value of read committed, set in the constructor and 
! is the default level used by Sql Server, should be used.
! ------------------------------------------------------------------
defaultIsolationLvl SQLINTEGER,private 

! set up and tear down
init             procedure(SQLHDBC hDbc),byte,proc
kill             procedure(),virtual
destruct         procedure(),virtual

! ----------------------------------------------------------------------
! sets the default value for the isolation level
! read committed is the normal default and is set in the constructor
! there may be use cases where some other default value is needed/wanted.
! However, if altering the default be sure you understand the implications 
! of using the other types of isolation levels.
! ----------------------------------------------------------------------
setDefaultIsolationLevel procedure(SQLINTEGER level)

! ----------------------------------------------------------------------
! sets the isolation level for the hDbc input.  the connection must not 
! have any open transactions when this is called.  the connection can be open 
! or closed.  as a general rule the connection would be open but that is 
! not required.  
! Note, if there are open transactions when this is called bad things will happen 
! ----------------------------------------------------------------------
setIsolationLevel procedure(long level),sqlReturn,protected 

! ----------------------------------------------------------------------
! sets the isolation level for the hDbc input to the level indicated 
! by the function name.  typically the default value of read committed is all that 
! will be used, there are use case for serializable transaction and in some 
! other instances the two types are used.
! ----------------------------------------------------------------------
setIsolationSerializable procedure(),sqlReturn 
setIsolationReadUncommitted procedure(),sqlReturn
setIsolationReadCommitted procedure(),sqlReturn
setIsolationRepeatabelread procedure(),sqlReturn

! ----------------------------------------------------------------------
! begins a transaction for the connection handle input.  
! this actually turns off auto-commit mode.
! 
! note, if this is called then you must call the commit or roll back functions
! when the work is completd.  failing to end a transaction will cause 
! bad things to happen
! ----------------------------------------------------------------------
beginTrans procedure(),sqlReturn

! ----------------------------------------------------------------------
! commits a transaction for the connection handle input.  
! ----------------------------------------------------------------------
Commit procedure(),sqlReturn

! ----------------------------------------------------------------------
! rolls back a transaction for the connection handle input.  
! ----------------------------------------------------------------------
Rollback procedure(),sqlReturn

! ----------------------------------------------------------------------
! ends a transaction for the connection handle input.  
! called from the commit or rollback functions.
! ----------------------------------------------------------------------
EndTrans procedure(long committRollBack),sqlReturn,private

            end

  _EndOfInclude_
