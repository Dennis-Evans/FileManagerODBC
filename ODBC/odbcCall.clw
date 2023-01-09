  member()
  
  include('odbcCall.inc'),once
  include('odbcSqlStrCl.inc'),once

  map 
  end

! ----------------------------------------------------------------------
! initilizes the object 
! ----------------------------------------------------------------------
odbcCallType.init procedure(*ODBCErrorClType e)   

retv     byte(level:benign)

  code 
  
  retv = parent.init(e)
     
  return retv 
! end Init
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! does any needed clean up and calls the parent kill
! ---------------------------------------------------------------------- 
odbcCallType.destruct procedure()  !virtual

  code 

  self.kill()
  
  return
! end destructor
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! virtual place holder
! use this function to format the fields or columns read prior to the display
! ----------------------------------------------------------------------
odbcCallType.formatRow procedure() !,virtual  

  code
  
  ! format queue elements for display in the derived object 
  
  return
! end formatRow 
! ----------------------------------------------------------------------

! ------------------------------------------------------------------
! call stored procedure with one or more table valued parameters.
! this is a virtual place holder and needs to be overloaded in function
! or in some code that can hold the array's until the write completes.
! ------------------------------------------------------------------
odbcCallType.execTableSp procedure(SQLHSTMT hStmt, string spName, *ParametersClass params, long numberRows) !,sqlReturn,virtual

  code
 
  ! this function must be overloaded in a derived class

  return sql_success
! -----------------------------------------------------------------------

! ---------------------------------------------------------------
! stored procedure and scalar function calls
! ---------------------------------------------------------------

! ------------------------------------------------------------------
! main worker function that does the actual calls to the back end.
! the other execSp functions do some setup and formatting and 
! thne call this function.
!
! Note, the call escape syntax is used not the exec. 
! -----------------------------------------------------------------------------
odbcCallType.execSp procedure(SQLHSTMT hStmt, *IDynStr sqlCode) !protected,virtual,sqlReturn

retv     sqlReturn,auto

  code

  retv = self.executeDirect(hStmt, sqlCode)
      
  return retv
! end execSp
! ---------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that does not return a value or a result set. 
! ------------------------------------------------------------------------------  
odbcCallType.execSp procedure(SQLHSTMT hStmt, string spName) !,sqlReturn

params  &ParametersClass
retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName, params) <> sql_Success)
    return sql_Error
  end 
  stop(self.sqlStr.sqlStr.cstr())
  retv = self.execSp(hStmt, self.sqlStr.sqlStr)

  return retv 
! end execSp
! ----------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that does not return a value or a result set. 
! binds any parameters and calls execSp/0
! ------------------------------------------------------------------------------  
odbcCallType.execSp procedure(SQLHSTMT hStmt, string spName, *ParametersClass params) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName, params) <> sql_Success)
    return sql_Error
  end 
  
  if (~params &= null) 
    ! check the status, if a table or tables are used as paremter then 
    ! don't bibd again, they have apready been boud to the hStmt
    if (params.AlreadyBound() = false)
      retv = params.bindParameters(hStmt)
    end
  end
  
  if (retv = sql_Success) 
    retv = self.execSp(hStmt, self.sqlStr.sqlStr)
  end  

  return retv 
! end execSp
! ----------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that returns a result set, the 
! queue parameter is bound to the resutls, 
! sp does not expect any parameters
! ------------------------------------------------------------------------------  
odbcCallType.execSp procedure(SQLHSTMT hStmt, string spName, columnsClass cols, *queue q) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName) <> sql_Success)
    return sql_Error
  end 
  
  retv = self.execSp(hStmt, self.sqlStr.sqlStr)
  if (retv = sql_Success) or (retv = Sql_Success_with_info)
    retv = self.fillResult(hStmt, cols, q)
  end   
  
  if (retv <> Sql_Success) and (retv <> Sql_Success_with_info)
    self.getError(hStmt)
  end

  return retv
! end execSp
! ----------------------------------------------------------------------

! -----------------------------------------------------------------------------
! execSp
! call an stored procedure that returns a result set, the 
! queue parameter is bound to the results, 
! binds any parameters and calls execSp/0 
! ------------------------------------------------------------------------------  
odbcCallType.execSp procedure(SQLHSTMT hStmt, string spName, columnsClass cols, *ParametersClass params, *queue q) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName, params) <> sql_Success)
    return sql_Error
  end 
    
  retv = params.bindParameters(hStmt)
    
  if (retv = sql_Success) or (retv = Sql_Success_with_info)
    retv = self.execSp(hStmt, self.sqlStr.sqlStr)
    if (retv = sql_Success) or (retv = Sql_Success_with_info)
      retv = self.fillResult(hStmt, cols, q)
      if (retv <> Sql_Success) and (retv <> Sql_Success_with_info)
        self.getError(hStmt)
      end
    end   
  end  

  return retv
! end execSp
! ----------------------------------------------------------------------
  
! ----------------------------------------------------------------------
! calls a sclar function and puts the returned value in the bound parameter
! ----------------------------------------------------------------------  
odbcCallType.callScalar procedure(SQLHSTMT hStmt, string spName, *ParametersClass params) 

retv    sqlReturn

  code 
  
  self.sqlStr.formatScalarCall(spName, params)
    
  retv = params.bindParameters(hStmt)
    
  if (retv = sql_Success) or (retv = Sql_Success_with_info)
    retv = self.execSp(hStmt, self.sqlStr.sqlStr)
  end  

  return retv
! end callScalar
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! sets up a call, just formats the string with the {call spname()}
! this one is used for a stored procedure with no parameters
! ----------------------------------------------------------------------
odbcCallType.setupSpCall procedure(string spName) 

retv     sqlReturn,auto
params   &ParametersClass

  code 
  
  retv = self.setupSpCall(spName, params)
  
  return retv 
! end setupSpCall
! ----------------------------------------------------------------------
  
! ----------------------------------------------------------------------
! sets up a call, just formats the string with the {call spname(?, ...)}
! adds a place holder for each parameter
! ----------------------------------------------------------------------  
odbcCallType.setupSpCall procedure(string spName, *ParametersClass params) ! sqlReturn

retv    sqlReturn 

  code 
  
  if (spName = '') 
    return sql_error
  end 

  if ((params &= null) or (params.HasParameters() = false))
    self.sqlStr.formatSpCall(spName)
  else   
    self.sqlStr.formatSpCall(spName, params)
  end   
    
  return retv
! end setupSpCall
! ----------------------------------------------------------------------  