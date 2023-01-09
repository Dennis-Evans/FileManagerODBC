   member('fmOdbcDemo')

   map
     readCurrentCount(fileMgrODBC fmOdbc),long
     deletenewRow(fileMgrODBC fmOdbc)
   end
   
   include('odbcTranscl.inc'),once

! ----------------------------------------------------------------------
! insert a row into a table using a stored procedure.  Read the value 
! of the identity column fomr an output parameter
! ----------------------------------------------------------------------
insertRow procedure(fileMgrODBC fmOdbc)

newLabel  cstring('Will Scarlet')
newAmount real(87.41)

identValue      long,auto
startRowCount   long,auto
endRowCount     long,auto
retv            sqlReturn,auto

  code

  writeLine(logFile, 'begin Insert Row.')

  startRowCount = readCurrentCount(fmOdbc)
  
  fmOdbc.parameters.AddInParameter(newLabel)
  fmOdbc.parameters.AddInParameter(newAmount)
  fmOdbc.parameters.AddOutParameter(identValue)

  retv = fmOdbc.callSp('dbo.addLabelRow')

  fmOdbc.clearInputs()
  endRowCount = readCurrentCount(fmOdbc)
  
  if (startRowCount + 1 = endRowCount)
    writeLine(logFile, 'Insert Row passed, number out was ' & endRowCount & '.')
  else 
    writeLine(logFile, 'Insert Row failed, number out was ' & endRowCount & '.')
    allTestsPassed = false
  end 
  
  deletenewRow(fmOdbc)
  startRowCount = readCurrentCount(fmOdbc)
  if (startRowCount <> 6) 
    writeLine(logFile, 'The new row was not deleted, test failed.')
    allTestsPassed = false
  end 

  writeLine(logFile, 'End Insert Row.')

  return
! end InsertRow ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! insert a row into a table using a query in the code.  Read the value 
! of the identity column from the second query
! note the differences when using a two sql statements
! ----------------------------------------------------------------------
insertRowQuery procedure(fileMgrODBC fmOdbc)

startRowCount   long,auto
endRowCount     long,auto

dynStr     &IDynStr
! values for the new row
newLabel   cstring('Hank smith')
newAmount  real(33.12)
! id value out
identValue long,auto
retv       sqlReturn,auto

  code

  writeLine(logFile, 'begin Insert Row from query.')

  startRowCount = readCurrentCount(fmOdbc)
  

  dynStr &= newDynStr()
  dynStr.cat('insert into dbo.LabelDemo(label, amount) ' & |
     'values(?, ?);' & |
     'select ? = scope_identity();')

  ! add the inputs and the output
  fmOdbc.parameters.AddInParameter(newLabel)
  fmOdbc.parameters.AddInParameter(newAmount)
  ! 
  fmOdbc.parameters.AddOutParameter(identValue)

  retv = fmOdbc.ExecuteNonQueryOut(dynStr)
  fmOdbc.clearInputs()
  endRowCount = readCurrentCount(fmOdbc)
  
  if (startRowCount + 1 = endRowCount)
    writeLine(logFile, 'Insert Row passed, number out was ' & endRowCount & '.')
  else 
    writeLine(logFile, 'Insert Row failed, number out was ' & endRowCount & '.')
    allTestsPassed = false
  end 
  
  deletenewRow(fmOdbc)
  startRowCount = readCurrentCount(fmOdbc)
  if (startRowCount <> 6) 
    writeLine(logFile, 'The new row was not deleted, test failed.')
    allTestsPassed = false
  end 

  writeLine(logFile, 'End  Insert Row from query.')

  return
! end InsertRowQuery ----------------------------------------------------------------------

insertTvpNoTrans procedure(fileMgrODBC fmOdbc, long rows)

startRowCount   long,auto
endRowCount     long,auto

  code 

  writeLine(logFile, 'begin insert TVP No Transaction.')

  startRowCount = readCurrentCount(fmOdbc)
  
  insertTvp(fmOdbc, rows, false)
 
  endRowCount = readCurrentCount(fmOdbc)
  
  if (endRowCount <> startRowCount + rows) 
    writeLine(logFile, 'Row count after Insert TVP in not correct, test failed.')
    allTestsPassed = false
  end   

  deletenewRow(fmOdbc)
  startRowCount = readCurrentCount(fmOdbc)
  if (startRowCount <> 6) 
    writeLine(logFile, 'The new row was not deleted, test failed.')
    allTestsPassed = false
  end 

  writeLine(logFile, 'End Insert with TVP No transaction.')

  return

! ------------------------------------------------------------------------------
! insert some number of rows into the datbase using a table valued parameter TVP
! the demo inserts a 1,000 rows
! ------------------------------------------------------------------------------
insertTvp  procedure(fileMgrODBC fmOdbc, long rows, bool withTrans)

!sqlStr        sqlStrClType 
retv          sqlReturn,auto

LabelArray     cstring(60),dim(Rows),auto
AmountArray    real,dim(Rows),auto
sysIdArray     long,dim(Rows),auto
RowActionArray long,dim(Rows),auto

x              long,auto
t              long,auto
hStmt          SQLHSTMT,auto

openedhere     byte,auto
parameters     ParametersClass
tablevalues    ParametersClass
typeName       cstring('LabelDemoType')
trans           odbcTransactionClType

  code

  loop x = 1 to Rows
    get(demoQueue, x)
    sysIdArray[x] = demoQueue.sysId
    LabelArray[x] = demoQueue.label
    AmountArray[x] = demoQueue.Amount
    RowActionArray[x] = 1
  end 

  writeLine(logFile, 'begin Call Insert TVP')
  if (withTrans = true) 
    writeLine(logFile, 'Insert TVP with Transaction')
  end 

  t = clock()  

  openedhere = fmOdbc.openConnection()
  if (withTrans = true) 
    trans.init(fmOdbc.conn.gethdbc())
    trans.beginTrans()
  end 

  hStmt = fmOdbc.conn.gethStmt()

  parameters.Init()

  tablevalues.init()
  
  Parameters.AddTableParameter(rows, typeName)
  retv = Parameters.bindParameters(hStmt, rows)
  
  tablevalues.focusTableParameter(hStmt, 1)   
  ! add the arrays and bind 
  tablevalues.AddlongArray(address(sysIdArray))  
  tablevalues.AddCStringArray(address(labelArray), size(labelArray[1]))
  tablevalues.addrealArray(address(amountArray))
  tablevalues.AddlongArray(address(rowActionArray))
  retv = tablevalues.bindParameters(hStmt) 
    ! remove the focus  and execute
  tablevalues.unfocusTableParameter(hStmt)

  retv = fmOdbc.odbccall.execSp(hStmt, 'dbo.InsertaTable', Parameters)

  ! if with trans is true roll it back
  if (withTrans = true) 
    trans.rollback()
  end 

  fmOdbc.closeConnection(openedhere)

  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
     writeLine(logFile, 'Insert TVP passed')
     writeLine(logFile, 'There were ' & rows & ' rows inserted in ' & format(clock() - t, @t4) & ' clock tics ' & clock() - t)
  else 
    fmOdbc.ShowErrors()
    writeLine(logFile, 'Insert TVP failed')
    allTestsPassed = false
  end 

  ! if with trans is true roll it back
  ! do a count to be sure the inserted rows were removed
  if (withTrans = true) 
    if (readCurrentCount(fmOdbc) <> 6)
      writeLine(logFile, 'The rollback failed.')
      allTestsPassed = false
    end 
  end 

  writeLine(logFile, 'end Call Insert TVP')
  if (withTrans = true) 
    writeLine(logFile, 'Insert TVP with Transaction')
  end 

  return
! end insertTvp --------------------------------------------------------------------------------   

readCurrentCount procedure(fileMgrODBC fmOdbc)

dynStr    &IDynStr
retv      long,auto
x         long,auto
!trans     odbcTransactionClType
outParam  long,auto

  code
 
  dynStr &= newDynStr()  
  dynStr.cat('select ? = count(*) from dbo.LabelDemo ld;')

  ! note the order of the bindings, the out parameter and 
  ! then the in parameter, the oder is important for the palce holders
  ! in the query 
  fmOdbc.parameters.AddOutParameter(outParam)

  retv = fmOdbc.ExecuteScalar(dynStr)
  
  if (retv = SQL_SUCCESS) or (retv = SQL_SUCCESS_WITH_INFO)
    retv = outparam
  end 

  dynStr.kill()
  fmOdbc.clearInputs()
  freeQueues()

  return retv
! ------------------------------------------------------------------------------------------------------  

deletenewRow procedure(fileMgrODBC fmOdbc)

dynStr    &IDynStr
retv      byte,auto
x         long,auto
!trans     odbcTransactionClType
outParam  long,auto

  code
 
  dynStr &= newDynStr()  
  dynStr.cat('delete from dbo.LabelDemo where sysId > 6;')

  ! note the order of the bindings, the out parameter and 
  ! then the in parameter, the oder is important for the palce holders
  ! in the query 
  fmOdbc.parameters.AddOutParameter(outParam)

  retv = fmOdbc.ExecuteScalar(dynStr)
    
  dynStr.kill()
  fmOdbc.clearInputs()
  
  return
! ------------------------------------------------------------------------------------------------------  
