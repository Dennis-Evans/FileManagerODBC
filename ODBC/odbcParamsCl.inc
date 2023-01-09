
!ABCIncludeFIle

  OMIT('_EndOfInclude_',_ODBCParamPresent_)
_ODBCParamPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once
  include('odbcSqlStrCl.inc'),once
    
! ----------------------------------------------------------------------    
! queue to store information for each parameter used in a sp call
! or in a query that accepts parameters
! add each parameter to the queue using the various functions
! the code will bind the parameters before the actual call to the database
! ----------------------------------------------------------------------
ParametersQueue queue,type
paramId           long     ! ordinal of the paremter
InOutType         SQLSMALLINT   ! type of parameter. in or out 
valueType         SQLSMALLINT   ! C type of the data in the parameter
paramType         SQLSMALLINT   ! ODBC type of the parameter 
paraSize          SQLULEN       ! len or size of the parameter
DecimalDigits     SQLSMALLINT   ! number of decimal digits, if needed
ParamPtr          SQLPOINTER    ! pointer to the data in the code
paramLength       SQLLEN        ! length
tableName         cstring(128)  ! name of the table type, in the server, if this is a table valued parameter
                end ! queue 
                
!
ParametersClass class,type,module('odbcParamsCl.clw'),link('odbcParamsCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
setupFailed           bool,private
paramQ                &ParametersQueue

alreadyBound          bool !,private
! ---------------------------------------------------------------------------
! default constructor for the parameters
! ---------------------------------------------------------------------------
construct procedure()

! ---------------------------------------------------------------------------
! sets up the object for use allocates a queue and sets a flag
! ---------------------------------------------------------------------------  
init procedure(),byte,proc

! ---------------------------------------------------------------------------
! does the standard clean up when the object goes out of scope 
! ---------------------------------------------------------------------------    
kill procedure()

! ---------------------------------------------------------------------------
! destructor, calls the Kill method
! ---------------------------------------------------------------------------        
destruct procedure(),virtual

! ---------------------------------------------------------------------------
! flag used for the table valued parameters.  the table or tables are  bound 
! and then the columns are bound, bu the columns to not increase the parameter count.
! this is set when a table is added and cleared when the queue is freed.
! ---------------------------------------------------------------------------        
AlreadyBound procedure(),bool

! ---------------------------------------------------------------------------------
! free the queue of parameters, this does not remove the bindings, if any
! ---------------------------------------------------------------------------------
clearQ procedure()

! -----------------------------------------------------------------------------
! Bind a table valued parameter for the statement.  
! 
! note the number of rows parameter.  the system writes back to this parameter
! amd it must remain in scope during the call. 
!
! once the table is bound use the two focus functions t obind the table columns
! 
! parameters for the ODBC api call 
! hStmt   = handle to the ODBC statement
! paramId = ord value of the parmaeter, 1, 2, 3 ... the ordinal position
! InOutType = type of parmeter in/out/inout
! value type = the C data type of the parameter, an equate value from the API
! param type = the sql data type of the parameter, an equate value from the API
! param size = the size, in bytes of the column for this parameter typicall the same as the paraLength
! decimal digits = number of decimal of the column 
! ParamPtr = pointer to the data for the parameter
! paramLength = the size of the paramPtr buffer
! colInd = pointer to a buffer for the size of the parameter.   not used and null in this example 
! -----------------------------------------------------------------------------  
bindParameters procedure(SQLHSTMT hStmt, *long numberRows),sqlReturn

! ---------------------------------------------------------------------------------
! set the focus of the driver to the table, once the focus is set 
! bind the columns of the table
! note, this must be called when using a table valued parameter
! ---------------------------------------------------------------------------------
focusTableParameter procedure(SQLHSTMT hStmt, long ordinal),sqlReturn,proc  

! ---------------------------------------------------------------------------------
! remove  the focus of the driver from the table, call this after the table columns
! are bound 
! note, this must be called when using a table valued parameter
! ---------------------------------------------------------------------------------
unfocusTableParameter procedure(SQLHSTMT hStmt),sqlReturn,proc

! -----------------------------------------------------------------------------
! bindParameters
! Bind parameters for the statement.  A parameter in a sql statement is marked with a ? 
! this process binds the actual parameter to that place holder so the back end knows 
! what is going on.   The order the parameters are added to the queue MUST match the order 
! ? marks appear in the sql statement.  There is no mapping by name, yet. 
!
! If a parameter appears in the statement twice or multiple times it must also be in the 
! queue twice or how ever many times. 
! 
! parameters for the ODBC api call 
! hStmt   = handle to the ODBC statement
! paramId = ord value of the parmaeter, 1, 2, 3 ... the ordinal position
! InOutType = type of parmeter in/out/inout
! value type = the C data type of the parameter, an equate value from the API
! param type = the sql data type of the parameter, an equate value from the API
! param size = the size, in bytes of the column for this parameter typicall the same as the paraLength
! decimal digits = number of decimal of the column 
! ParamPtr = pointer to the data for the parameter
! paramLength = the size of the paramPtr buffer
! colInd = pointer to a buffer for the size of the parameter.   not used and null in this example 
! -----------------------------------------------------------------------------  
bindParameters procedure(SQLHSTMT hStmt),sqlReturn

! ----------------------------------------------------------
! function the check the rows in the queue
! returns true if there is at least one row and false is no  rows
! ----------------------------------------------------------
HasParameters procedure(),bool

! ---------------------------------------------------------------------------------
! add a table parameter to the call
! number of rows is the number of rows in the arrays used as the source
! table name is the name of the table type on the server
! 
! if the type name is not found the call to the backend  will fail
! ---------------------------------------------------------------------------------
AddTableParameter procedure(long numberRows, *cstring tableName),sqlReturn,proc

! ----------------------------------------------------------------------------------
! add a parameter to the queue.  typically called internally but can be used 
! from anywhere if needed desired.
! ----------------------------------------------------------------------------------
addParameter procedure(SQLSMALLINT InOutType, SQLSMALLINT ValueType, SQLSMALLINT ParameterType, SQLULEN ColumnSize, SQLSMALLINT DecimalDigits, SQLPOINTER varPtr, SQLLEN BufferLength),byte,proc

! ----------------------------------------------------------
! add the various input parameters
! ----------------------------------------------------------

! ----------------------------------------------------
! integer types 
! ----------------------------------------------------
AddInParameter procedure(*byte varPtr),sqlReturn,proc
AddInParameter procedure(*short varPtr),sqlReturn,proc
! ----------------------------------------------------------
! add a long parameter.  clarion uses an equate for booleans 
! this will also be used when working with clarion's bool
! ----------------------------------------------------------
AddInParameter procedure(*long colPtr),sqlReturn,proc

! ----------------------------------------------------
! string types 
! ----------------------------------------------------
AddInParameter procedure(*cstring varPtr),sqlReturn,proc     
AddInParameter procedure(*string varPtr),sqlReturn,proc     

! ----------------------------------------------------
! floating point types
! ----------------------------------------------------
AddInParameter procedure(*sreal varPtr),sqlReturn,proc
AddInParameter procedure(*real varPtr),sqlReturn,proc

! ----------------------------------------------------
! date types 
! ----------------------------------------------------
AddInParameter procedure(*date varPtr),sqlReturn,proc
AddInParameter procedure(*time varPtr),sqlReturn,proc
AddInParameter procedure(*TIMESTAMP_STRUCT varPtr),sqlReturn,proc

! ----------------------------------------------------------
! add the different types of arrays for the table valued parameters
! these pararmeters are always inputs, no output parameters allowed
!
! the array parameter is the address of the first element in the array
! so call like this AddLongArray(address(someArray[1]))
! ----------------------------------------------------------
AddlongArray procedure(long array),sqlReturn,proc
AddRealArray procedure(long array),sqlReturn,proc
! ------------------------------------------------------------------------------------
! same as the other over loads but note the elementSize parameter.  this will be the 
! string length the array elements,  for input parameter.  required so the driver knows 
! how many bytes to copy 
! ------------------------------------------------------------------------------------
AddCStringArray procedure(long array, long elementSize),sqlReturn,proc

! ----------------------------------------------------------------
! add the output types 
! ----------------------------------------------------------------
AddOutParameter procedure(*byte varPtr),sqlReturn,proc
AddOutParameter procedure(*short varPtr),sqlReturn,proc
AddOutParameter procedure(*long varPtr),sqlReturn,proc
AddOutParameter procedure(*cstring varPtr),sqlReturn,proc
AddOutParameter procedure(*real varPtr),sqlReturn,proc
AddOutParameter procedure(*sreal varPtr),sqlReturn,proc

! --------------------------------------------------------------------------------
! worker function called by the sql string object to add the place holders to a 
! call.  there will be one ? added for each parameter
! --------------------------------------------------------------------------------
FillPlaceHolders procedure(sqlStrClType sqlCode, long startPos = 1),long

                    end

  _EndOfInclude_
