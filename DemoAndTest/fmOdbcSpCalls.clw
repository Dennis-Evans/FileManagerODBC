   member('fmOdbcDemo')

   map
   end

! --------------------------------------------------
! fills the queue from a stored procedure 
! this function does an explicit open and close for the 
! connection.  
! --------------------------------------------------
fillSp procedure(fileMgrODBC fmOdbc)

retv       sqlReturn,auto
openedHere byte,auto

  code

  writeLine(logFile, 'begin Call stored procedure with connect')

  fm.columns.AddColumn(demoQueue.sysId)
  fm.columns.AddColumn(demoQueue.Label)
  fm.columns.AddColumn(demoQueue.amount)

  retv = conn.connect()

  case retv 
    of SQL_SUCCESS 
      retv = fm.conn.AllocateStmthandle()
    of SQL_SUCCESS_WITH_INFO
      fm.getConnectionError()
      retv = fm.conn.AllocateStmthandle()
    else 
      retv = fm.conn.AllocateStmthandle()
  end

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    retv = fm.callSp('dbo.ReadLabelDemo', demoQueue)
    fm.ClearInputs()
    conn.disconnect()
  end

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    writeLine(logFile, 'Call stored procedure with connect passed ')   
  else 
    writeLine(logFile, 'Call stored procedure with connect failed') 
    AllTestsPassed = false  
  end  

  writeLine(logFile, 'end Call stored procedure with connect')

  return 
! end fillSp ---------------------------------------------------

! --------------------------------------------------
! fills the queue from a stored procedure 
! this procedure uses the default connection 
! --------------------------------------------------
fillSpNoOpen procedure(fileMgrODBC fmOdbc)

retv   byte,auto

  code

  writeLine(logFile, 'begin Call stored procedure, default open')

  fm.columns.AddColumn(demoQueue.sysId)
  fm.columns.AddColumn(demoQueue.Label)
  fm.columns.AddColumn(demoQueue.amount)

  retv = fm.callSp('dbo.ReadLabelDemo', demoQueue)

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    writeLine(logFile, 'Call stored procedure, default open passed ')   
  else 
    writeLine(logFile, 'Call stored procedure, default open failed')  
    AllTestsPassed = false 
  end  

  writeLine(logFile, 'end Call stored procedure, default open')

  return
! end fillSpNoOpen ---------------------------------------------------

! --------------------------------------------------
! fills the queue from a stored procedure 
! this procedure uses an input parameter that will 
! filter the result set down to the rows that match 
! --------------------------------------------------
fillSpWithParam procedure(fileMgrODBC fmOdbc)

inLabel string('Willma')
retv    byte,auto

  code

  writeLine(logFile, 'begin Call stored procedure with a parameter')

  fmOdbc.columns.AddColumn(demoQueue.sysId)
  fmOdbc.columns.AddColumn(demoQueue.Label)
  fmOdbc.columns.AddColumn(demoQueue.amount)

  fmOdbc.Parameters.AddInParameter(inLabel)

  retv = fmOdbc.CallSp('dbo.ReadLabelDemoByLabel', demoQueue)

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    writeLine(logFile, 'Call stored procedure with a parameter passed ')   
  else 
    writeLine(logFile, 'Call stored procedure with a parameter failed')   
    AllTestsPassed = false
  end  

  fm.ClearInputs()

  writeLine(logFile, 'end Call stored procedure with a parameter')

  return
! end fillSpWithParam --------------------------------------------------

! --------------------------------------------------
! calls a scalar function and gets the returned value
! --------------------------------------------------
callScalar procedure(fileMgrODBC fmOdbc)

retv      byte,auto
inLabel   cstring('Willma')
outParam  long,auto

  code

  writeLine(logFile, 'begin Call scalar function')

  fmOdbc.parameters.AddOutParameter(outParam)
  fmOdbc.parameters.AddInParameter(inLabel)

  retv = fmOdbc.callScalar('dbo.getId')

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    writeLine(logFile, 'Call Scalar function passed')
    writeLine(logFile, 'Label Used as a filter was ' & inLabel & ', the ID value returned ' &  outParam)
  else 
    writeLine(logFile, 'Call Scalar function failed')  
    AllTestsPassed = false
  end  

  fm.ClearInputs()

  writeLine(logFile, 'end Call scalar function')

  return
! end callScalar ---------------------------------------------------

! --------------------------------------------------
! calls a stored procedure that returns three result sets 
! the first one is prcessed normally, the second and third
! sets use the readNextResult function
! --------------------------------------------------
callMulti procedure(fileMgrODBC odbcFm)

openedHere byte,auto

inAmount   real(20.00)  ! used as a filter value in the second of the three result sets

retv       byte,auto

! instances to bind the second and third results
secondCols ColumnsClass
thirdCols  ColumnsClass

  code

  writeLine(logFile, 'begin Call multiple result sets, stored procedure')

  ! add the input parameter
  ! this is added before the sp call and is used by the sp
  ! once the sp call retruns the parameter can be thrown away
  ! it is not need to process the resutl set
  ! the parameter is used in the second result set
  odbcFm.Parameters.AddInParameter(inAmount)

  ! first result set just use the default columns instance from the object
  ! the first result set actually returns three columns, but the third column 
  ! is not bound so it will be ignored
  odbcFm.columns.AddColumn(demoQueue.sysId)
  odbcFm.columns.AddColumn(demoQueue.Label)

  ! becasue there are morethan one result sets the connection is opened here
  ! if it was not then the second and third sets would be lost when 
  ! the connection was closed 
  openedHere = odbcFm.openConnection()
  
  if (openedHere <> Connection:Failed)
    retv = odbcFm.CallSp('dbo.readTwo', demoQueue)
    ! this will clear the bound columns and bound parameters
    odbcFm.ClearInputs()

    if (retv = SQL_SUCCESS)
      ! bind the sceond set of columns
      ! and fill the queue from the second result set
      secondCols.init()
      secondCols.AddColumn(secondDemoQueue.Amount)
      secondCols.AddColumn(secondDemoQueue.Label)
      retv = odbcFm.readNextResult(secondDemoQueue, secondCols)
      ! and do some cleanup
      odbcFm.ClearInputs()
    end

    if (retv = SQL_SUCCESS)
      ! bind the columns used by the third result set
      thirdCols.init()
      thirdCols.AddColumn(thirdDemoQueue.name)
      thirdCols.AddColumn(thirdDemoQueue.department)
      retv = odbcFm.readNextResult(thirdDemoQueue, thirdCols)
      odbcFm.ClearInputs()
    end
  end

  odbcFm.closeConnection(openedHere)

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    writeLine(logFile, 'Call multiple result sets, stored procedure passed')
  else 
    writeLine(logFile, 'Call multiple result sets, stored procedure failed')  
    AllTestsPassed = false
  end  

  if (records(secondDemoQueue) <= 0)
    writeLine(logFile, 'The second result set was not read.')
    AllTestsPassed = false
  end 

  if (records(thirdDemoQueue) <= 0)
    writeLine(logFile, 'The third result set was not read.')
    AllTestsPassed = false
  end 

  writeLine(logFile, 'end Call multiple result sets, stored procedure')

  return
! end callSpMulti ----------------------------------------------