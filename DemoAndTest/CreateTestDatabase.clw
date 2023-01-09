
   member('fmOdbcDemo')

   include('odbcExecCl.inc'),once
   include('odbcConn.inc'),once
   include('odbcParamsCl.inc'),once
   include('odbcTypes.inc'),once
   include('dynstr.inc'),once

   map
     NewDatabaseCalls(),bool
     DatabaseObjects(),bool
     DoesTestDbExist(),bool
     DropTestDb(),bool
     CreateTestDb(),bool
     setup()
     createDatabaseObjects(),bool
     processSqlStatements(*cstring sqlCode), bool
     findDelimitPos(string sqlCode),long
     parseOutOneStatement(string sqlCode, *IdynStr block, long endPos)
     executeCurrentStatement(*IDynStr statement),bool
   end

createConnStr &MSConnStrClType
createConn    &ODBCConnectionClType
createOdbc    &odbcExecType

statementDelimit  equate('GO')

! -------------------------------------------------------------------
! -------------------------------------------------------------------
createTestDatabase procedure()

retv  bool,auto
res   sqlReturn,auto

  code

  writeLine(logFile, 'Begin Create Database')
  
  setup()
  
  retv = NewDatabaseCalls() 

  ! original connection was to the master database, 
  ! test database was dropped (if needed) and created so 
  ! change the connection to the test database 
  ! for adding the tables and other objects
  if (retv = true) 
    createConnStr.setDbName('fmOdbctest')
  end  

  retv = DatabaseObjects()

  if (retv = true)
    writeLine(logFile, 'Create Database Success')
  else 
    writeLine(logFile, 'Create Database Failed')
  end

  writeLine(logFile, 'End Create Database')

  return
! end createTestDatabase ---------------------------------------------

! --------------------------------------------------------------------
!  checks for the test database and if found drops the database.
!  creates a new database using the label fmOdbcTest
!  Note, there are no permission checks done here.  
!  if the user does not have the correct permissions then 
!  these steps will fail and the tests will not be ran.
! --------------------------------------------------------------------
NewDatabaseCalls procedure()

retv    bool,auto
res     sqlReturn,auto

  code 

  writeLine(logFile, 'Begin New Database')

  res = createConn.connect()

  if (res <> SQL_SUCCESS) and (res <> SQL_SUCCESS_WITH_INFO)
    return false
  else 
    res = createConn.AllocateStmtHandle()  
  end 

  ! does the databsae exists now
  if (DoestestDbExist() = true) 
    retv = DropTestDb()
  else 
    retv = true  
  end 

  retv = CreateTestDb()

  createConn.Disconnect()

  if (retv = true)
    writeLine(logFile, 'New Database Success')
  else 
    writeLine(logFile, 'New Database Failed')
  end

  writeLine(logFile, 'End New Database')

  return retv
! end  NewDatabaseCalls ---------------------------------------------

DatabaseObjects procedure() ! bool 

retv    bool,auto
res     sqlReturn,auto

  code 

  writeLine(logFile, 'Begin Database Objects')

  res = createConn.connect()

  if (res <> SQL_SUCCESS) and (res <> SQL_SUCCESS_WITH_INFO)
    return false
  end

  res = createConn.AllocateStmtHandle()  

  retv = createDatabaseObjects()
  if (retv = true)
    writeLine(logFile, 'Database Objects Success')
  else 
    writeLine(logFile, 'Database Objects Failed')
  end

  createConn.Disconnect()
  
  writeLine(logFile, 'End Database Objects')

  return retv
! end  DatabaseObjects -----------------------------------------------

setup procedure() 

e ODBCErrorClType

  code

  createconnStr &= new(MSConnStrClType)
  createconnStr.Init('dennishyperv\dev', 'master')

  createconn &= new(ODBCConnectionClType)
  createconn.Init(createConnstr)
  createodbc &= new(odbcExecType)
  createodbc.Init(e)

  return
 ! end  setup ----------------------------------------------------------

! ---------------------------------------------------
! creates the test database using the system defaults for size, 
! file location etc...
! returns true if the test database was created
! false if not created.
! ---------------------------------------------------
CreateTestDb procedure() !,bool

dynStr    &IDynStr
retv      bool,auto
res       sqlReturn 

  code

  dynStr &= newDynStr()
  dynStr.cat('create database fmOdbcTest;')

  res = createOdbc.ExecQuery(createConn.gethStmt(), dynStr)

  dynStr.kill()

  if (res = SQL_SUCCESS) or (res = SQL_SUCCESS_WITH_INFO)
    if (DoesTestDbExist() = true) 
      retv = true
    end   
  end  
  
  return retv
! end CreateTestDb  ------------------------------------------------  

! ---------------------------------------------------
! drops the test database 
! ---------------------------------------------------
DropTestDb procedure() !,bool

dynStr    &IDynStr
retv      bool,auto
res       sqlReturn,auto

  code

  writeLine(logFile, 'Begin drop database')

  dynStr &= newDynStr()
  dynStr.cat('alter database fmOdbcTest set single_user with rollback immediate; drop database fmOdbcTest;')
  
  res = createOdbc.ExecQuery(createConn.gethStmt(), dynStr)

  dynStr.kill()

  if (res = SQL_SUCCESS) or (res = SQL_SUCCESS_WITH_INFO)
    writeLine(logFile, 'Drop Database Success')
    retv = true
  else 
    writeLine(logFile, 'Drop Database Failed')
    retv = false
  end 

  writeLine(logFile, 'End drop database')

  return retv
! end DropTestDb  -----------------------------------

! ---------------------------------------------------
! checks to see if the test database exists, returns true 
! if does and false if does not
! ---------------------------------------------------
DoesTestDbExist procedure() !,bool

dynStr    &IDynStr
dbId      long,auto
retv      bool,auto
parameters ParametersClass
res       sqlReturn,auto

  code

  dynStr &= newDynStr()
  dynStr.cat('select ? = isnull(db_id(''fmOdbcTest''), 0);')
  
  parameters.Init()
  parameters.AddOutParameter(dbId)

  res = createOdbc.execQuery(createConn.gethStmt(), dynStr, parameters)
  
   dynStr.kill()
  
  if (res = SQL_SUCCESS) or (res = SQL_SUCCESS_WITH_INFO)
    if (dbId > 0) 
      retv = true
    else 
      retv = false
    end
  else 
    retv = false
  end  
  
  return retv
! end DoestestDbExist ------------------------------------------------  

! ---------------------------------------------------
! create the various database objects and inserts 
! the default test data for the test database.
! ---------------------------------------------------
createDatabaseObjects  procedure() ! bool

retv      bool(true)
sqlCode   &cstring
fileName  cstring('sqlStatements.sql')
fileSize  long,auto
f         handle,auto

  code

  writeLine(logFile, 'Create Objects')

  f = openFile(fileName)
 
  if (f > 0) 
    fileSize = findFileSize(f) 
    if (fileSize > 0)
      sqlCode &= new(cstring(fileSize + 2))
      readFileToEnd(fileName, f, sqlCode, fileSize)  
    else 
      retv = false  
    end 
    CloseFile(f)
  else 
    retv = false
  end    

  if (retv = true) 
    retv = processSqlStatements(sqlCode)
  end 
   
  writeLine(logFile, 'End Objects')

  return retv
! end createDatabaseObjects --------------------------------------------------------  

! ---------------------------------------------------
! process the sql statements form the scrip file.
! each statement or statements is processed
! ---------------------------------------------------
processSqlStatements  procedure(*cstring sqlCode) ! bool

retv         bool(true)
statement    &IDynStr
delimitPos   long,auto

  code

  statement &= newDynStr()

  loop     
    delimitPos = findDelimitPos(sqlCode)
    if (delimitPos > 0)
      parseOutOneStatement(sqlCode, statement, delimitPos)
      ! if the execution fails return an error
      if (executeCurrentStatement(statement) = false) 
        retv = false
        break
      end 

      sqlCode = sub(sqlCode, delimitPos + 2, len(sqlCode))
      if (sqlCode = '') 
        break
      end  
    end
  end 

  return retv
! end processSqlStatements  --------------------------------------------------------  

! ---------------------------------------------------
! executes the statement input and clears the hStmt 
! for the next run
! ---------------------------------------------------
executeCurrentStatement procedure(*IDynStr statement) 

retv  bool,auto
res  sqlReturn,auto

  code

  res = createOdbc.ExecQuery(createConn.gethStmt(), statement)
  if (res = SQL_SUCCESS) or (res = SQL_SUCCESS_WITH_INFO)
    retv = true
  else 
    retv = false  
  end  
  createConn.clearStmthandle() 

  return retv
! end executeCurrentStatement ------------------------------------------------------

! ---------------------------------------------------
! finds the position of the next delimter string in the text
! buffer input
! ---------------------------------------------------
findDelimitPos procedure(string sqlCode)  

pos   long,auto

  code

  pos = instring(statementDelimit, sqlCode, 1)  

  return pos
! end findDelimitPos -----------------------------------------------------------

parseOutOneStatement procedure(string sqlCode, *IdynStr block, long endPos)

  code 

  block.Kill()
  block.cat(sub(sqlCode, 1, endPos - 1))

  return