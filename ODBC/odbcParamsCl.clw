
  member()
  
  include('odbcParamsCl.inc'),once 
  include('odbcSqlStrCl.inc'),once
  
  map 
    module('odbc32')
      SQLBindParameter(SQLHSTMT StatementHandle, SQLUSMALLINT ParameterNumber, SQLSMALLINT InputOutputType, SQLSMALLINT ValueType, SQLSMALLINT ParameterType, SQLULEN ColumnSize, SQLSMALLINT DecimalDigits, SQLPOINTER ParameterValuePtr, SQLLEN BufferLength, *SQLLEN StrLen_or_IndPtr)sqlReturn,pascal      
      SQLSetStmtAttr(SQLHSTMT StatementHandle, SQLINTEGER Attribute, SQLPOINTER ValuePtr, SQLINTEGER StringLength),sqlReturn,pascal,name('SQLSetStmtAttrW')
    end 
    module('c6')
      memmove(long dest, long src, long num),long,proc,name('_memmove')
      strLen(long cstr),long,name('_strlen')
    end 
  end

! ---------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! default constructor for the parameters
! ---------------------------------------------------------------------------  
ParametersClass.construct procedure()

  code 

  self.init() 
      
  return
! end construct
! ----------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! sets up the object for use allocates a queue and sets a flag
! ---------------------------------------------------------------------------  
ParametersClass.Init procedure()

retv     byte,auto 

  code 
  
  if (self.paramQ &= null) 
    self.paramQ &= new(ParametersQueue)
    if (self.paramQ &= null) 
      retv = level:notify  
      self.setupFailed = true
    else 
      self.setupFailed = false  
    end 
  end   
  
  return retv
! end init 
! ------------------------------------------------------------------------------------
  
! ---------------------------------------------------------------------------
! does the standard clean up when the object goes out of scope 
! ---------------------------------------------------------------------------    
ParametersClass.kill procedure()

  code
    
  if (self.setupFailed = false) 
    free(self.paramQ)
    dispose(self.paramQ)
    self.paramQ &= null
    self.setupFailed = true
  end  
  
  return 
! end kill 
! -----------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! destructor, calls the Kill method
! ---------------------------------------------------------------------------        
ParametersClass.destruct procedure()

  code 
  
  self.kill() 
  
  return
! destruct 
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! gets the status  ofr the flag for use in the stored procedure calls
! -----------------------------------------------------------------------------
ParametersClass.AlreadyBound procedure()

  code

  return self.alreadyBound
! end AlreadyBound -----------------------------------------------------------

! ---------------------------------------------------------------------------------
! free the queue of parameters, this does not remove the bindings, if any
! ---------------------------------------------------------------------------------
ParametersClass.clearQ procedure()

  code 
  
  free(self.paramQ)
  ! done with this group so clear the flag
  self.alreadyBound = false
 
  return
! end clear
! ---------------------------------------------------------------------------------

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
ParametersClass.bindParameters procedure(SQLHSTMT hStmt, *long numberRows) !,sqlReturn  

retv        sqlReturn(SQL_SUCCESS)   
count       long,auto

wideStr     CWideStr    ! used to convert the ansi string to a wide string
numberBytes long,auto   ! returned number of bytes

  code 

  self.alreadyBound = true

  ! once for each parameter
  loop count = 1 to records(self.paramQ)
    get(self.paramQ, count)
    ! convert the table type name
    numberBytes = wideStr.Init(self.paramQ.tableName)
    retv = SQLBindParameter(hStmt, self.paramQ.ParamId, SQL_PARAM_INPUT, SQL_C_DEFAULT, SQL_SS_TABLE, self.paramQ.paraSize, 0, wideStr.GetWideStr(), SQL_NTS, numberRows)
    ! if check for errors and return failure 
    if (retv <> sql_Success) and (retv <> Sql_Success_with_info)
      break
    end  
  end    

  return retv
! end bindParameters 
! ---------------------------------------------------------------------------------

! ---------------------------------------------------------------------------------
! set the focus of the driver to the table, once the focus is set 
! bind the columns of the table
! note, this must be called when using a table valued parameter
! ---------------------------------------------------------------------------------
ParametersClass.focusTableParameter procedure(SQLHSTMT hStmt, long ordinal)  

retv       sqlReturn

  code
                               
  retv = SQLSetStmtAttr(hStmt, SQL_SOPT_SS_PARAM_FOCUS, ordinal, SQL_IS_INTEGER)

  return retv
! end focusTableParameter
! -------------------------------------------------------------------------------

! ---------------------------------------------------------------------------------
! remove  the focus of the driver from the table, call this after the table columns
! are bound 
! note, this must be called when using a table valued parameter
! ---------------------------------------------------------------------------------
ParametersClass.unfocusTableParameter procedure(SQLHSTMT hStmt)  

retv       sqlReturn

  code

  retv = SQLSetStmtAttr(hStmt, SQL_SOPT_SS_PARAM_FOCUS, 0, SQL_IS_INTEGER)

  return retv
! end unfocusTableParameter
! -------------------------------------------------------------------------------

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
ParametersClass.bindParameters procedure(SQLHSTMT hStmt) !,sqlReturn  

colInd   &long       ! null pointer, param not used in this example
retv     sqlReturn   
count    long,auto

  code 

  if (self.setupFailed = true) 
    return sql_error
  end   

  ! once for each parameter
  loop count = 1 to records(self.paramQ)
    get(self.paramQ, count)
    retv = SQLBindParameter(hStmt, self.paramQ.ParamId, self.paramQ.InOutType, self.paramQ.valueType, self.paramQ.ParamType, self.paramQ.paraSize, self.paramQ.DecimalDigits, self.paramQ.ParamPtr, self.paramQ.paramLength, colInd)
    ! if not a good call then get out, if one is missing the rest do not matter
    if (retv <> sql_Success) and (retv <> Sql_Success_with_info)
      break
    end  
  end    

  ! reset for the caller
  if (retv = Sql_Success_with_info)
    retv = Sql_Success
  end

  return retv
! end bindParameters 
! ---------------------------------------------------------------------------------

! ----------------------------------------------------------
! function the check the rows in the queue
! returns true if there is at least one row and false is no  rows
! ----------------------------------------------------------
ParametersClass.HasParameters procedure() !,bool

retv  bool(false)

  code

  if (records(self.paramQ) > 0) 
    retv = true
  end

  return retv
! end HasParameters --------------------------------------------------------------

! ---------------------------------------------------------------------------------
! add a table parameter to the call
! number of rows is the number of rows in the arrays used as the source
! table name is the name of the table type on the server
! 
! if the type name is not found the call to the backend  will fail
! ---------------------------------------------------------------------------------
ParametersClass.AddTableParameter procedure(long numberRows, *cstring tableName)  !,sqlReturn,proc

retv sqlReturn

  code

  if (self.setupFailed = true) 
    return sql_error
  end   
  
  self.paramQ.paramId = records(self.ParamQ) + 1
  self.paramQ.InOutType = SQL_PARAM_INPUT  ! tables parameters are always inputs, cannot be out type parameters
  self.paramQ.valueType = SQL_C_DEFAULT    ! default for the table type
  self.paramQ.paramType = SQL_SS_TABLE     ! type
  self.paramQ.paraSize = numberRows        
  self.paramQ.DecimalDigits = 0            ! always zero for the table type
  self.paramQ.tableName = tableName        ! this is the type name on the server
  self.paramQ.paramLength = SQL_NTS         
  self.paramQ.ParamPtr = 0                 ! always zero or a null pointer

  add(self.paramQ)
  
  if (errorcode() > 0) 
    retv = level:notify
  end 

  return retv
! end AddTableParameter  
! ----------------------------------------------------------------------------------

! ----------------------------------------------------------------------------------
! add a parameter to the queue.  typically called internally but can be used 
! from anywhere if needed desired.
! ----------------------------------------------------------------------------------
ParametersClass.addParameter procedure(SQLSMALLINT InOutType, SQLSMALLINT ValueType, | 
                                       SQLSMALLINT ParameterType, SQLULEN ColumnSize, | 
                                       SQLSMALLINT DecimalDigits, SQLPOINTER varPtr, | 
                                       SQLLEN BufferLength) !,byte,proc,private

retv     byte(level:benign)

  code 
  
  if (self.setupFailed = true) 
    return sql_error
  end   
  
  self.paramQ.paramId = records(self.ParamQ) + 1
  self.paramQ.InOutType = inOutType
  self.paramQ.valueType = ValueType
  self.paramQ.paramType = ParameterType
  self.paramQ.paraSize = ColumnSize
  self.paramQ.DecimalDigits = DecimalDigits
  self.paramQ.ParamPtr = varPtr
  self.paramQ.paramLength = BufferLength

  add(self.paramQ)
  
  if (errorcode() > 0) 
    retv = level:notify
  end 
      
  return retv
! end addParameter
! ------------------------------------------------------------------------------------

! ------------------------------------------------------------------------------------
! add the various input parameters
! ------------------------------------------------------------------------------------

ParametersClass.AddInParameter procedure(*byte varPtr)  !,sqlReturn,proc

  code 

  self.addParameter(SQL_PARAM_INPUT, SQL_C_TINYINT, SQL_TINYINT, eSizeByte, 0, address(varPtr), eSizeByte)

  return 0
! end AddInParameter --------------------------------------------------------

ParametersClass.AddInParameter procedure(*short varPtr) !,sqlReturn,proc
  
  code 

  self.addParameter(SQL_PARAM_INPUT, SQL_C_SHORT, SQL_SMALLINT, eSizeShort, 0, address(varPtr), eSizeShort)

  return 0
! end AddInParameter --------------------------------------------------------

ParametersClass.AddInParameter procedure(*long varPtr)

retv   sqlReturn

  code

  self.addParameter(SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, eSizeLong, 0, address(varPtr), eSizeLong)

  return retv  
! end AddInParameter
! --------------------------------------------------------------------------------

ParametersClass.AddInParameter procedure(*sreal varPtr) !,sqlReturn,proc

retv   sqlReturn

  code

  self.addParameter(SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_REAL, eSizeSReal, 0, address(varPtr), eSizeSReal)

  return retv  
! end AddInParameter
! --------------------------------------------------------------------------------

ParametersClass.AddInParameter procedure(*real varPtr) !,sqlReturn,proc

retv   sqlReturn

  code

  self.addParameter(SQL_PARAM_INPUT, SQL_C_DOUBLE, SQL_FLOAT, eSizeReal, 0, address(varPtr), eSizeReal)

  return retv  
! end AddInParameter
! --------------------------------------------------------------------------------

ParametersClass.AddInParameter procedure(*cstring varPtr) !,sqlReturn,proc

retv   sqlReturn

  code

  self.addParameter(SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, size(varPtr), 0, address(varPtr), size(varPtr))
  
  return retv  
! end AddInParameter
! --------------------------------------------------------------------------------

ParametersClass.AddInParameter procedure(*string varPtr) !,sqlReturn,proc

retv   sqlReturn

  code

  self.addParameter(SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, size(varPtr), 0, address(varPtr), size(varPtr))
  
  return retv  
! end AddInParameter
! --------------------------------------------------------------------------------

ParametersClass.AddInParameter procedure(*date varPtr)  !,sqlReturn,proc

retv   sqlReturn

  code

  self.addParameter(SQL_PARAM_INPUT, SQL_C_DATE, SQL_DATE, size(SQL_C_DATE), 0, address(varPtr), size(SQL_C_DATE))

  return retv    
! end AddInParameter
! --------------------------------------------------------------------------------  

ParametersClass.AddInParameter procedure(*time varPtr)  !,sqlReturn,proc

retv   sqlReturn

  code

  self.addParameter(SQL_PARAM_INPUT, SQL_C_TIME, SQL_TIME, size(SQL_C_TIME), 0, address(varPtr), size(SQL_C_TIME))

  return retv    
! end AddInParameter
! --------------------------------------------------------------------------------  

ParametersClass.AddInParameter procedure(*TIMESTAMP_STRUCT varPtr) !,sqlReturn,proc  

retv   sqlReturn

  code

  self.addParameter(SQL_PARAM_INPUT, SQL_C_TYPE_TIMESTAMP, SQL_TYPE_TIMESTAMP, size(SQL_C_TYPE_TIMESTAMP), 0, address(varPtr), size(SQL_C_TYPE_TIMESTAMP))

  return retv    
! end AddInParameter
! --------------------------------------------------------------------------------  

! ----------------------------------------------------------
! add the different types of arrays for the table valued parameters
! these pararmeters are always inputs, no output parameters allowed
!
! the array parameter is the address of the first element in the array
! so call like this AddLongArray(address(someArray[1]))
! ----------------------------------------------------------
ParametersClass.AddLongArray procedure(long array) !,sqlReturn,proc

retv sqlReturn

  code
  
  ! add the array, this will always be an input
  retv = self.addParameter(SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, eSizeLong, 0, array, eSizeLong)

  return retv
! end AddlongArray
! ------------------------------------------------------------------------------------

ParametersClass.AddRealArray procedure(long array) !,sqlReturn,proc

retv sqlReturn

  code
  
  ! add the array, this will always be an input
  retv = self.addParameter(SQL_PARAM_INPUT, SQL_C_DOUBLE, SQL_FLOAT, eSizeReal, 0, array, eSizeReal)

  return retv
! end AddRealArray
! ------------------------------------------------------------------------------------

! ------------------------------------------------------------------------------------
! same as the other over loads but note the elementSize parameter.  this will be the 
! string length of the input parameter.  required so the driver knows how many 
! bytes to copy 
! ------------------------------------------------------------------------------------
ParametersClass.AddCStringArray procedure(long array, long elementSize) !,sqlReturn,proc

retv sqlReturn

  code

  retv = self.addParameter(SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, elementSize, 0, array, elementSize)

  return retv
! end AddCStringArray
! --------------------------------------------------------------------------------------  

! ------------------------------------------------------------------------------------
! add the various output parameters
! ------------------------------------------------------------------------------------
ParametersClass.AddOutParameter procedure(*byte varPtr) !,sqlReturn,proc

retv   sqlReturn,auto

  code
 
  retv = self.addParameter(SQL_PARAM_OUTPUT, SQL_C_TINYINT, SQL_TINYINT, eSizeLong, 0, address(varPtr), eSizeLong)

  return retv
! end AddOutParameter
! --------------------------------------------------------------------------------

ParametersClass.AddOutParameter procedure(*short varPtr) !,sqlReturn,proc

retv   sqlReturn,auto

  code
 
  retv = self.addParameter(SQL_PARAM_OUTPUT, SQL_C_SHORT, SQL_SMALLINT, eSizeLong, 0, address(varPtr), eSizeLong)

  return retv
! end AddOutParameter
! --------------------------------------------------------------------------------

ParametersClass.AddOutParameter procedure(*long varPtr) !,sqlReturn,proc

retv   sqlReturn,auto

  code
 
  retv = self.addParameter(SQL_PARAM_OUTPUT, SQL_C_SLONG, SQL_INTEGER, eSizeLong, 0, address(varPtr), eSizeLong)

  return retv
! end AddOutParameter
! --------------------------------------------------------------------------------

ParametersClass.AddOutParameter procedure(*cstring varPtr) !,sqlReturn,proc

retv   sqlReturn,auto

  code
 
  retv = self.addParameter(SQL_PARAM_OUTPUT, SQL_C_CHAR, SQL_CHAR, size(varPtr), 0, address(varPtr), size(varPtr))

  return retv
! end AddOutParameter
! --------------------------------------------------------------------------------

ParametersClass.AddOutParameter procedure(*sreal varPtr) !,sqlReturn,proc

retv   sqlReturn,auto

  code
 
  retv = self.addParameter(SQL_PARAM_OUTPUT, SQL_C_FLOAT, SQL_REAL, eSizeReal, 0, address(varPtr), eSizeReal)

  return retv
! end AddOutParameter
! --------------------------------------------------------------------------------

ParametersClass.AddOutParameter procedure(*real varPtr) !,sqlReturn,proc

retv   sqlReturn,auto

  code
 
  retv = self.addParameter(SQL_PARAM_OUTPUT, SQL_C_DOUBLE, SQL_FLOAT, eSizeReal, 0, address(varPtr), eSizeReal)

  return retv
! end AddOutParameter
! --------------------------------------------------------------------------------

! --------------------------------------------------------------------------------
! worker function called by the sql string object to add the place holders to a 
! call.  there will be one ? added for each parameter
! --------------------------------------------------------------------------------
ParametersClass.FillPlaceHolders procedure(sqlStrClType sqlCode, long startPos = 1) 

count    long
recCount long

  code 
  
  recCount = records(self.paramQ)
  loop count = startpos to recCount 
    get(self.paramQ, count)
    if (count < recCount)
      sqlCode.cat(eMarkerComma)
    else 
      sqlCode.cat(eFinalMarker)
    end 
  end       

  return recCount
! end FillPlaceHolders
! --------------------------------------------------------------------------------  

