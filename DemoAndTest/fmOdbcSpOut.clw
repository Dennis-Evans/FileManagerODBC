   member('fmOdbcDemo')

   map
   end

! --------------------------------------------------
! call a stored procedure with an out parameter
! while this example uses a single parameter a 
! procedure can have more than one output, just 
! add the fields or locals as parameters
! --------------------------------------------------
spWithOut procedure(fileMgrODBC fmOdbc)

retv     byte,auto
rowCount long,auto

  code

  writeLine(logFile, 'begin Stored Procedure with output parameter')

  fmOdbc.parameters.addOutParameter(rowCount)

  retv = fmOdbc.callSp('dbo.CountDemoLabels')

  fmOdbc.ClearInputs()

  totalRows = rowCount

  if (retv = SQL_SUCCESS)
    writeLine(logFile, 'Stored Procedure with output parameter passed, number out was ' & rowCount & '.')
  else 
    writeLine(logFile, 'Stored Procedure with output parameter, Failed')  
    AllTestsPassed = false
  end 


  writeLine(logFile, 'End Stored Procedure with output parameter')

  return
! end spWithOut ---------------------------------------------------

spResutSetWithOut procedure(fileMgrODBC fmOdbc)

retv       byte,auto
rowCount   long,auto
openedHere byte,auto

  code

  writeLine(logFile, 'Begin  Stored Procedure with result set and output parameter')

  ! set to a value outside the range of the count 
  ! function, just so we can see the value was actually set
  rowCount = -1

  fmOdbc.parameters.addOutParameter(rowCount)

  fmOdbc.columns.AddColumn(demoQueue.sysId)
  fmOdbc.columns.AddColumn(demoQueue.Label)
  fmOdbc.columns.AddColumn(demoQueue.amount)

  retv = conn.connect()

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    retv = fmOdbc.callSp('dbo.ReadLabelDemoWithCount', demoQueue)

    if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
      retv = fmOdbc.odbcCall.nextResultSet(fmOdbc.conn.getHstmt())
      ! out parameter is now filled
    else 
      writeLine(logFile, 'Initial read of result set failed.')  
      AllTestsPassed = false
    end
    conn.Disconnect()
  end

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    writeLine(logFile, 'Stored Procedure with result set and output parameter passed')
    writeLine(logFile, 'Outputparameter value was ' & rowCount & '.')
  else 
    writeLine(logFile, 'Stored Procedure with result set and output parameter failed')
  end 
 
  fmOdbc.ClearInputs()
 
  writeLine(logFile, 'End Stored Procedure with result set and output parameter')

  return
! end spWithOut ---------------------------------------------------

