
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCErrorClPresent_)
_ODBCErrorClPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcConn.inc'),once
  include('svcom.inc'),once

    
eLengthSqlState    equate(6)
eLengthErrText     equate(2000)
  
OdbcErrorGroup group,type
sqlState         cstring(eLengthSqlState)  ! the state code is a fixed size
NativeErrorPtr   long
MessageText      &cstring                  ! variable length
textLengthPtr    SQLSMALLINT
               end 

OdbcErrorQueue queue(OdbcErrorGroup)
               end 
               
ODBCErrorClType class,type,module('odbcErrorCl.clw'),link('odbcErrorCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! ------------------------------------------------
! queue to hold the error and information messages
! ------------------------------------------------
errorMsgQ         &OdbcErrorQueue

displayError      bool,private    ! display an error using the message function when it happens

! ------------------------------------------------
! number of error records retruned by the most recent error for the handle type
! ------------------------------------------------
errorCount        long,private    

! ----------------------------------------------------------------
! setup and clean up
! ----------------------------------------------------------------
init              procedure(),sqlReturn,proc
kill              procedure(),virtual
destruct          procedure(),virtual
!

getNumberMsg      procedure(),byte

getDataBaseError  procedure(ODBCConnectionClType conn),sqlReturn,proc

getError          procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),byte,proc
getErrorGroup     procedure(long ndx, *string stateText, *string msgText)
!----------------------------------------------------------------------
! calls the SQLGetDiagField function to get the number of messages 
! they may be error messages or just information messages
! called internally by the getError function
! this call can return other information but they are not currently used
!
! the call can return the number of rows affected by an insert, update or delete action
! and some other types of information
! ----------------------------------------------------------------------    
getDiagRecCount   procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),long,private,proc

! ----------------------------------------------------------------------
! displays an erro on the screen using a simple message call
! shows all the errors and information messages from the call for the handle type
! ----------------------------------------------------------------------
showError         procedure()

! ----------------------------------------------------------------------
! free the queue of error messages
! this is called each time the getError function is called.  
! the messages from the most recent error are deleted, 
! ----------------------------------------------------------------------
freeErrorMsgQ     procedure(),private

! ----------------------------------------------------------------------
! allocates the message queue 
! ----------------------------------------------------------------------
makeObjects       procedure(),sqlReturn,private

! ----------------------------------------------------------------------
! does the clean up, called by the kill method
! ----------------------------------------------------------------------
destroyObjects    procedure(),private
                end ! class 


  _EndOfInclude_
