

!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCCallPresent_)
_ODBCCallPresent_   EQUATE(1)

  include('dynStr.inc'),once
  include('odbcConn.inc'),once
  include('odbcBaseCl.inc'),once
  include('odbcTypes.inc'),once
  include('odbcParamsCl.inc'),once 
  include('odbcSqlStrCl.inc'),once
  include('odbcColumnsCl.inc'),once 

! ----------------------------------------------------------------------------------
! class to handle the stored procedure and scalar function calls.
! there is not much difference between calling a stored procedure or executing a query
! that is passed in as a string but there some minor differences. 
! splitting the the calls into two objects, both derived from the same base object
! gives us a little cleaner interface and makes adding or changing parts of the 
! classes much simpler.
! ----------------------------------------------------------------------------------
odbcCallType class(odbcBaseClType),type,module('odbcCall.clw'),link('odbcCall.clw',_ABCLinkMode_),dll(_ABCDllMode_)
! set up and tear down

! ----------------------------------------------------------------------
! initilizes the object 
! ----------------------------------------------------------------------
init             procedure(*ODBCErrorClType e),byte,proc

! ----------------------------------------------------------------------
! does any needed clean up and calls the parent kill
! ----------------------------------------------------------------------
destruct         procedure(),virtual

! ----------------------------------------------------------------------
! virtual place holder
! use this function to format the fields or columns read prior to the display
! ----------------------------------------------------------------------
formatRow        procedure(),virtual

! ------------------------------------------------------------------
! call stored procedure with one or more table valued parameters.
! this is a virtual place holder and needs to be overloaded in function
! or in some code that can hold the array's until the write completes.
! ------------------------------------------------------------------
execTableSp    procedure(SQLHSTMT hStmt, string spName, *ParametersClass params, long numberRows),sqlReturn,virtual

! ------------------------------------------------------------------
! main worker function that does the actual calls to the back end.
! the other execSp functions do some setup and formatting and 
! thne call this function.
! ------------------------------------------------------------------
execSp         procedure(SQLHSTMT hStmt, *IDynStr sqlCode),protected,virtual,sqlReturn

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that does not return a value or a result set. 
! ------------------------------------------------------------------------------  
execSp         procedure(SQLHSTMT hStmt, string spName),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that does not return a value or a result set. 
! binds any parameters 
! ------------------------------------------------------------------------------  
execSp         procedure(SQLHSTMT hStmt, string spName, *ParametersClass params),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that returns a result set, the 
! queue parameter is bound to the resutls, 
! sp does not expect any parameters
! ------------------------------------------------------------------------------  
execSp         procedure(SQLHSTMT hStmt, string spName, *columnsClass cols, *queue q),sqlReturn,virtual

! -----------------------------------------------------------------------------
! execSp
! call an stored procedure that returns a result set, the 
! queue parameter is bound to the results, 
! binds any parameters and calls execSp/0 
! ------------------------------------------------------------------------------  
execSp         procedure(SQLHSTMT hStmt, string spName, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn,virtual

! ----------------------------------------------------------------------
! calls a sclar function and puts the returned value in the bound parameter
! ----------------------------------------------------------------------  
callScalar     procedure(SQLHSTMT hStmt, string spName, *ParametersClass params),sqlReturn,virtual

! ----------------------------------------------------------------------
! sets up a call, just formats the string with the {call spname()}
! this one is used for a stored procedure with no parameters
! ----------------------------------------------------------------------
setupSpCall   procedure(string spName),sqlReturn,virtual

! ----------------------------------------------------------------------
! sets up a call, just formats the string with the {call spname(?, ...)}
! adds a place holder for each parameter
! ----------------------------------------------------------------------  
setupSpCall   procedure(string spName, *ParametersClass params),sqlReturn,virtual

            end

  _EndOfInclude_
