
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCConnectionPresent_)
_ODBCConnectionPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once
  
  include('EnvHandle.inc'),once

  include('odbcConnStrCl.inc'),once
  include('OdbcHandleCl.inc'),once
                       
  itemize(0),pre(Connection)
Opened equate
CallerOpened  equate
failed equate
  end
  
ODBCConnectionClType class,type,module('odbcConn.clw'),link('odbcConn.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! ODBC handles 
hEnv                   &EnvHandle
hDbc                   &OdbcHandleClType,private ! db connection handle
hStmt                  &OdbcStmtHandleClType,private! statement handle

connStr                &baseConnStrClType ! instance to hold the connection string

! -------------------------------------------------------
! some getters, no setters.  code will need to query these 
! values and use locally, but they will never be modifed
! ------------------------------------------------------
gethDbc                procedure(),SQLHDBC
gethStmt               procedure(),SQLHSTMT
gethEnv                procedure(),SQLHENV

! ------------------------------------------------------
! checks to see if the connection handle input or the 
! instance connection handle is active or dead
! ------------------------------------------------------
isConnected            procedure(SQLHDBC dbHandle = 0),bool

! set up and clean up
init                   procedure(),sqlReturn
init                   procedure(baseConnStrClType  connString),sqlReturn,proc
kill                   procedure()

! ------------------------------------------------------
! sets the ODBC version in use 
! ------------------------------------------------------
setOdbcVersion         procedure(long verId),sqlReturn,proc 

! ------------------------------------------------------
! connect to the database
! if withStatement is true, the default.  a statment handle is 
! allocated 
! ------------------------------------------------------
Connect                procedure(),sqlReturn

! ------------------------------------------------------
! disconnect from the database
! calls free handle for the connection
! freeing the hDbc will also free and all hStmt on the connection
! ------------------------------------------------------
Disconnect             procedure(),sqlReturn,proc

! ------------------------------------------------------
! Allocates a statement handle for the connection
! called by user to allocate a more than the default statement handles.
! this function is not used by the connect call.
! 
! note, the caller should check the value of the output hStmt
! if zero or less the handle was not allocated
! ------------------------------------------------------
AllocateStmtHandle procedure(*SQLHSTMT hStmt),sqlReturn,protected,virtual

! ------------------------------------------------------
! Allocates a statement handle for the connection
! called by the connect function or called 
! by the using code to create a statement handle
! note, the default is to create a statement handle with the 
! hDbc, most actions will only need a single hStmt handle 
! and will close the connection when the action completes.
! see the over loaded function if multiple handle are needed.
! ------------------------------------------------------
AllocateStmtHandle     procedure(),sqlReturn,virtual

! ------------------------------------------------------
! frees the statement handle.  if this is called then 
! the connection no longer has a hStmt and one must be 
! allocated before the connection is used again.
!
! if the parameter input is the default then the 
! instance hStmt is used. if not the default the input 
! hStmt will be used.
! ------------------------------------------------------
FreeStmthandle         procedure(SQLHSTMT hStmt = 0) 

! ------------------------------------------------------
! clears the statement handle or any binding, result sets, ... 
! call this when the stamenet handle is going to be used 
! for multiple calls.  this allows the connection to perform 
! more than one action and not need t oallocate a hStmt for each,
! saves a little overhead
! Note, some actions can be executed in sequence and not need 
! the hStmt cleared, some cannot be.  Good practice would be to 
! call the clear function between any calls, but dpending 
! on the actual action executed it is not always required.
!
! if the parameter input is the default then the 
! instance hStmt is used. if not the default the input 
! hStmt will be used.
! ------------------------------------------------------
clearStmthandle        procedure(SQLHSTMT hStmt = 0)

asyncOn procedure(),sqlReturn

                     end ! class 


  _EndOfInclude_
