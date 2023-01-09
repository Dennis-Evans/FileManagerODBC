

!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCExecClassPresent_)
_ODBCExecClassPresent_   EQUATE(1)

  include('dynStr.inc'),once
  include('odbcConn.inc'),once
  include('odbcBaseCl.inc'),once
  include('odbcTypes.inc'),once
  include('odbcParamsCl.inc'),once 
  include('odbcSqlStrCl.inc'),once
  include('odbcColumnsCl.inc'),once 

! ----------------------------------------------------------------------
! class that is used to execute a query from the file manage rinstance. 
! the query is passed in a string and then formated and called.
! 
! result sets, if any, are processed.  input and output parameters 
! can be used. 
!
! Note, the term exec in the name does not imply that the query is called using the 
! exec 'some statement' format.  All these calls use the call syntax.
! the call syntax does less work than the exec style and allows dfor output parameters.
! ----------------------------------------------------------------------
odbcExecType class(odbcBaseClType),type,module('odbcExecCl.clw'),link('odbcExecCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! set up and tear down
! ----------------------------------------------------------------------
! initilizes the object 
! ----------------------------------------------------------------------
init             procedure(*ODBCErrorClType e),byte,proc

destruct         procedure(),virtual

! ----------------------------------------------------------------------
! virtual place holder
! use this function to format the fields or columns read prior to the display
! ----------------------------------------------------------------------
formatRow        procedure(),virtual

! ----------------------------------------------------------------------
! execute a query that does not return a result set and does not use any 
! parameters
! ----------------------------------------------------------------------
execQuery       procedure(SQLHSTMT hStmt, *IDynStr sqlCode),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! execute the statement then fill the queue or buffers
!
! this method does not accept the parameters class instance so use this one for queries that 
! do not have parameters.
! ------------------------------------------------------------------------------   
execQuery       procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *columnsClass cols, *queue q),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that does not return a result set.  
! but has input parameters and may have output parameters.
! this can be used for normal reads (select field from ...) or it can be used for 
! scalar type calls (select ? = column from ...)
!
! it does not handle calls that return result sets
! ------------------------------------------------------------------------------      
execQuery       procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *ParametersClass params),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
! ------------------------------------------------------------------------------    
execQuery       procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn,virtual

            end

  _EndOfInclude_
