   member('fmOdbcDemo')

   map
   end

   include('odbcTranscl.inc'),once

expected   long
actual     long

! -------------------------------------------------------------------
! execute a single query and place the result set into a queue
! the defaults for the table is six rows. 
! -------------------------------------------------------------------
executeQuery procedure(fileMgrODBC fmOdbc)

dynStr    &IDynStr
retv      byte,auto
x         long,auto
trans     odbcTransactionClType

  code

  writeLine(logFile, 'Begin Execute Query new')

  dynStr &= newDynStr()
  dynStr.cat('select ld.SysId, ld.Label, ld.amount ' & |
             'from dbo.LabelDemo ld ' & |
             ' order by ld.SysId;')

  ! add the colums of the queue that will be read by the query
  fmOdbc.columns.AddColumn(demoQueue.SysId)
  fmOdbc.columns.AddColumn(demoQueue.Label)
  fmOdbc.columns.AddColumn(demoQueue.amount)

  ! do the actual read
  retv = fmOdbc.ExecuteQuery(dynStr, demoQueue)

  fmOdbc.clearInputs()
  dynStr.kill()

  if (retv = SQL_SUCCESS)
    writeLine(logFile, 'Execute Query, passed and returned ' & records(demoQueue) & ' Rows.')
  else 
    writeLine(logFile, 'Execute Query, Failed') 
    AllTestsPassed = false 
  end 

  writeLine(logFile, 'end Execute Query')

  return
! end execureQuery -----------------------------------------------------------

! -----------------------------------------------------------------
! executes a scalar style query that returns one row and one column
! the query filters out some rows and retuens a count of the 
! remaining rows.
! -----------------------------------------------------------------
execScalar procedure(fileMgrODBC fmOdbc, *cstring fltLabel)

dynStr    &IDynStr
retv      byte,auto
outParam  long,auto

  code

  writeLine(logFile, 'begin Execute Scalar Query')

  dynStr &= newDynStr()
  dynStr.cat('select ? = count(*) from dbo.LabelDemo ld where ld.Label <> ?;')

  ! note the order of the bindings, the out parameter and 
  ! then the in parameter, the oder is important for the palce holders
  ! in the query 
  fmOdbc.parameters.AddOutParameter(outParam)
  fmOdbc.parameters.AddInParameter(fltLabel)

  retv = fmOdbc.ExecuteScalar(dynStr)

  writeLine(logFile, 'Label Used as a filter was ' & fltLabel & ', the count of rows is ' &  outParam & '. One row was removed by the filter.')

  dynStr.kill()

  if (retv = SQL_SUCCESS)
    writeLine(logFile, 'Execute Scalar Query, passed')
  else 
    writeLine(logFile, 'Execute Scalar Query, Failed')  
    AllTestsPassed = false
  end 

  writeLine(logFile, 'end Execute Scalar Query')

  return
! end execScalar ---------------------------------------------------

! -------------------------------------------------------------------
! execute a query with a single join clause and place the result set into 
! a queue,
! the department table contains four rows for the default
! -------------------------------------------------------------------
executeQueryTwo procedure(fileMgrODBC fmOdbc)

dynStr    &IDynStr
retv      byte,auto

  code

  writeLine(logFile, 'begin ExecuteQueryTwo')

  dynStr &= newDynStr()
  dynStr.cat('select ld.SysId, ld.Label, ld.amount, d.Label ' & |
             'from dbo.labelDemo ld ' & |
             'inner join dbo.Department d on ' & |
               'd.ldSysId = ld.sysId ' & |
             'order by d.Label desc, ld.label asc;')
  
  fmOdbc.columns.AddColumn(demoQueue.sysId)
  fmOdbc.columns.AddColumn(demoQueue.Label)
  fmOdbc.columns.AddColumn(demoQueue.amount)
  fmOdbc.columns.AddColumn(demoQueue.department)

  retv = fmOdbc.ExecuteQuery(dynStr, demoQueue)

  dynStr.kill()
  fmOdbc.clearInputs()

  if (retv = SQL_SUCCESS)
    writeLine(logFile, 'ExecuteQueryTwo, a simple join, passed and retuned ' & records(demoQueue) & ' rows.')
  else 
    writeLine(logFile, 'ExecuteQueryTwo, a simple join, Failed')  
    AllTestsPassed = false
  end 

  writeLine(logFile, 'end ExecuteQueryTwo')

  return
! end execureQury -----------------------------------------------------------