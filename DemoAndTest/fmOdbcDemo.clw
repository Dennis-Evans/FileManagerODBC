    program

   include('aberror.inc'),once
   include('fileMgrOdbc.inc'),once
   include('odbcExecCl.inc'),once
   include('odbcConn.inc'),once

   map
     main()
     odbcSetup()
     fileManagerSetup()
     freeQueues()
     makeString(),string
     fillQueue(long numberRows)

     module('fmOdbcSpCalls')
       fillSp(fileMgrODBC fmOdbc)
       fillSpNoOpen(fileMgrODBC fmOdbc)
       fillSpWithParam(fileMgrODBC fmOdbc)
       callScalar(fileMgrODBC fmOdbc)
       callMulti(fileMgrODBC fmOdbc)
     end

     module('fmOdbcSpOut')
       spWithOut(fileMgrODBC fmOdbc)
       spResutSetWithOut(fileMgrODBC fmOdbc)
     end

     module('fmOdbcFileMgr')
       fileFill(fileMgrODBC fmOdbc)
       propSqlFill()
       viewFill(fileMgrODBC fmOdbc)
     end

     module('fmodbcExeQuery')
      executeQuery(fileMgrODBC fmOdbc)
      execScalar(fileMgrODBC fmOdbc, *cstring fltLabel)
      executeQueryTwo(fileMgrODBC fmOdbc)
     end

     module('fmOdbcPageLoad')
       pageLoad(filemgrODBC odbcFm, long currentRow)
     end

     module('fmOdbcInserts')
       insertRow(fileMgrODBC fmOdbc)
       insertRowQuery(fileMgrODBC fmOdbc)
       insertTvpNoTrans(fileMgrODBC fmOdbc, long rows)
       insertTvp(fileMgrODBC fmOdbc, long rows, bool withTrans)
     end

     module('fmOdbcAsyn')
       callSpAsync(fileMgrODBC odbcFm)
    end

    module('osFile')
       openFile(*cstring filename), handle
       Write(handle f, string text)
       WriteLine(handle f, string text)
       readFileToEnd(*cstring fileName, handle f, *cstring sqlCode, long fSize) 
       makeFile(*cstring filePath),handle
       CloseFile(handle f) 
       findFileSize(handle f),long
     end
  
    module('fmOdbcBcp')
      fmOdbcBcpUpdate(fileMgrODBC fmOdbc)  
    end 

    module('createTestDatabase')
      createTestDatabase()
    end 

   end ! map

! number of rows to be inserted using TVP
! adjust as needed, typically around a 1,000 is the high end
! MS recommends if more than 1,000 rows use the BCP
numberTvpRows  equate(1000)

! number of rows in the initial page size for the page loading example 
pageLoadSize       equate(8)

defaultStrLength   equate(60)

! -------------------------------------------------------------------------------------
! define the connection string(s)
! -------------------------------------------------------------------------------------
!Connstr     string('Driver={{SQL Server Native Client 11.0};server=dennishyperv\dev;Database=default_test;trusted_connection=yes;App=lmno;')

! ----------------------------
! change set up function also
! ----------------------------
!Connstr     string('dennishyperv\dev,default_test,,Driver={{ODBC Driver 13 for SQL Server};App=lmno')
!Connstr     string('dennishyperv\dev,default_test,,Driver={{SQL Server Native Client 11.0};App=lmno')

!Connstr     string('Driver={{ODBC Driver 13 for SQL Server};server=dennishyperv\dev;Database=default_test;trusted_connection=yes;App=lmno')
!Connstr     string('Driver={{ODBC Driver 13 for SQL Server};server=dennishyperv\dev;Database=default_test;trusted_connection=yes;in_valid=demo;App=lmno')

Connstr     string('Driver={{ODBC Driver 17 for SQL Server};server=dennishyperv\dev;Database=default_test;trusted_connection=yes;App=lmno')
! -------------------------------------------------------------------------------------

! -------------------------------------------------------------------------------------
! define a simple file for the demo
! -------------------------------------------------------------------------------------
LabelDemo   file,driver('ODBC'),owner(ConnStr),name('dbo.labeldemo')
!LabelDemo   file,driver('MSSQL'),owner(ConnStr),name('dbo.labeldemo')
Record        record,pre()
SysId           long
Label           string(defaultStrLength)
Amount          real
              end ! record
            end ! file
! -------------------------------------------------------------------------------------

demoView  view(labelDemo)
            project(labelDemo.SysId)
            project(labelDemo.Label)
            project(labelDemo.Amount)
          end
! -------------------------------------------------------------------------------------
! queues used by the demo
! all queues are global
! -------------------------------------------------------------------------------------
demoQueue   queue
SysId         long
Label         string(defaultStrLength)
Amount        real
Department    string(defaultStrLength)

            end
! ---------------------------------
secondDemoQueue queue
Amount            real
Label             string(defaultStrLength)
                end
! ---------------------------------
thirdDemoQueue queue
name             string(defaultStrLength)
department       string(defaultStrLength)
               end
! -------------------------------------------------------------------------------------

! error class used by the file manager
errors         &errorclass
ErrorStatus    ErrorStatusClass

! -------------------------------------------------------------------------------------
! connection object for the ODBC file manager
! -------------------------------------------------------------------------------------
conn          &ODBCConnectionClType

! -------------------------------------------------------------------------------------
! connection string object for the ODBC file manager
! -------------------------------------------------------------------------------------
odbcConnStr   &MSConnStrClType

! -------------------------------------------------------------
! derive the file manager object for the overloaded init method
! -------------------------------------------------------------
localFm       class(fileMgrODBC),type
init            procedure(),virtual
              end
! define an instance
fm            &localFm

totalRows     long

logFile       handle
logFileName   cstring('testResults.log')

demoFlt       cstring('Willma')

AllTestsPassed  bool
alwaysShowLog   bool
! --------------------------------------------------------------------------
! program entry point 
! --------------------------------------------------------------------------
  code

  AllTestsPassed = true
  alwaysShowLog = true

  logFile = makeFile(logFileName)
  
  WriteLine(logFile, 'Begin Tests for the File Manager ODBC.')

  ! set things up
  odbcSetup()
  fileManagerSetup()

  !createTestDatabase()

  freeQueues()
  fmOdbcBcpUpdate(fm)  
  stop('kill it')
  executeQuery(fm)
  freeQueues()
  
  execScalar(fm, demoFlt)
  freeQueues()
 
  executeQueryTwo(fm)
  freeQueues()

  fillSp(fm)
  freeQueues()

  fillSpNoOpen(fm)
  freeQueues()

  fillSpWithParam(fm)
  freeQueues()

  callScalar(fm)
  freeQueues()

  callMulti(fm)
  freeQueues()
  
  spWithOut(fm)
  freeQueues()

  spResutSetWithOut(fm)
  freeQueues()

  insertRow(fm)
  freeQueues()

  insertRowQuery(fm)
  freeQueues()

  insertTvpNoTrans(fm, numberTvpRows)
  freeQueues()

  insertTvp(fm, numberTvpRows, true)
  freeQueues()

  if (AllTestsPassed = false) 
    writeLine(logFile, 'One or more tests failed.')
  else 
    writeLine(logFile, 'All test Passed.')
  end 

  CloseFile(logFile)
  
  if (AllTestsPassed = false) or (alwaysShowLog = true)
    run('notepad.exe C:\git_repo\FileManager_ODBC\DemoAndTest\testresults.log')
  end

  ! and go
  !main()

  return
! end program ------------------------------------------------------------

main procedure()

currentRow long(-1)
filterStr  cstring('Willma')
retv       sqlReturn,auto

Window WINDOW('Demo'),AT(,,622,286),FONT('MS Sans Serif',8,,FONT:regular),GRAY
       BUTTON('Call Stored Procedure (Connect)'),AT(9,9,109,14),USE(?btnSpCall)
       BUTTON('Call Stored Procedure'),AT(123,9,124,14),USE(?btnSpNoConnect)
       BUTTON('Call Scalar Function'),AT(258,9,99,14),USE(?btnCallScalar)
       BUTTON('Execute a Query'),AT(9,31,109,14),USE(?btnExecQuery)
       BUTTON('Execute Scalar'),AT(123,31,124,14),USE(?btnExecScalar)
       BUTTON('Execute Query Two Tables'),AT(258,31,109,14),USE(?btnExecQueryTwo)
       BUTTON('Call Stored Procedure W/Multi Results'),AT(373,31,135,14),USE(?btnMultiResults)
       BUTTON('Stored Procedure W/Parameter'),AT(9,50,109,14),USE(?btnSpWithParam)
       BUTTON('Stored Procedure w/out parameter'),AT(123,50,124,14),USE(?spWithOut)
       BUTTON('Stored Procedure Result and Out Parameter'),AT(258,50,156,14),USE(?btnResultWithOut)
       BUTTON('Insert Row Query'),AT(9,72,98,14),USE(?btnInsertRowQuery)
       BUTTON('Insert Rows using a TVP'),AT(123,72,109,14),USE(?btnInsertTvp)
       BUTTON('Page Load, Next'),AT(258,72,109,14),USE(?btnNextPage)
       BUTTON('Page Load, Previous'),AT(373,72,109,14),USE(?btnPrevPage)
       BUTTON('File Manager Loop'),AT(9,92,109,14),USE(?btnFileManager)
       BUTTON('Prop Sql'),AT(123,92,74,14),USE(?btnPropSql)
       BUTTON('Fill from a View'),AT(258,92,109,14),USE(?btnViewFill)
       LIST,AT(15,114,363,83),USE(?demoList),FORMAT('71L(2)|M~System Id~@N20@125L(2)|M~Label~59L(2)|M~Amount~@N20.2@40L(2)|M~Departme' &|
           'nt~@s60@'),FROM(demoQueue)
       LIST,AT(391,114,214,85),USE(?List2),FORMAT('53L(2)|M~Amount~@N10.2@50L(2)|M~Label~@s50@'),FROM(secondDemoQueue)
       LIST,AT(15,206,263,52),USE(?List3),FORMAT('115L(2)|M~Name~50L(2)|M~Department~'),FROM(thirdDemoQueue)
       BUTTON('&Done'),AT(163,262,36,14),USE(?btnCancel)
       BUTTON('Clear Queues'),AT(15,264,71,14),USE(?btnClearQ)
     END

  code

  open(window)
  accept
    case event()
      of Event:Accepted
      case field()
        of ?btnExecQuery
          freeQueues()
          executeQuery(fm)
        of ?btnExecScalar
          freeQueues()
          execScalar(fm, filterStr)
        of ?btnExecQueryTwo
          freeQueues()        
          executeQueryTwo(fm)

        of ?btnSpCall
          freeQueues()        
          fillSp(fm)
        of ?btnSpNoConnect
          freeQueues()        
          fillSpNoOpen(fm)
        of ?btnCallScalar
          freeQueues()        
          callScalar(fm)
        of ?btnSpWithParam
          freeQueues()        
          fillSpWithParam(fm)

        of ?spWithOut
          freeQueues()
          spWithOut(fm)

        of ?btnResultWithOut
          freeQueues()        
          spResutSetWithOut(fm)

        of ?btnMultiResults
          freeQueues()        
          callMulti(fm)

        of ?btnInsertRowQuery
          freeQueues()        
          InsertRowQuery(fm)
        of ?btnInsertTvp
          fillQueue(numberTvpRows)
          insertTvp(fm, numberTvpRows, false)
          freeQueues()
          
        of ?btnNextPage
          freeQueues()
          if (currentRow < 0)
            currentRow = 0
          else
            currentRow += pageLoadSize
          end
          if (currentRow >= totalRows)
            currentRow = totalRows - pageLoadSize
          end 
          pageLoad(fm, currentRow)
        of ?btnPrevPage
          freeQueues()
          currentRow -= pageLoadSize
          if (currentRow < 0)
            currentRow = 0
          end
          pageLoad(fm, currentRow)
           
        of ?btnFileManager
          freeQueues()        
          fileFill(fm)

        of ?btnPropSql
          freeQueues()        
          propSqlFill()

        of ?btnViewFill
          freeQueues()
          viewFill(fm)

        of ?btnClearQ
          freeQueues()

        of ?btnCancel
          break
      end ! case field
    end ! case event

  end

  close(window)

  return
! end main -------------------------------------------------

! --------------------------------------------------
! sets up the connection and connection string instances
! --------------------------------------------------
odbcSetup procedure()

  code

  odbcConnStr &= new(MSConnStrClType)
  odbcConnStr.Init(connStr)
  !odbcConnStr.Init('dennisHyperv\dev', 'default_test')

  conn &= new(ODBCConnectionClType)
  conn.Init(odbcConnStr)

  return
! end odbcSetup -------------------------------------------------------------

! --------------------------------------------------
! allocates the file manager and the error class
! does some default set up
! --------------------------------------------------
fileManagerSetup procedure()

  code

  errors &= new(errorclass)
  errors.Init(ErrorStatus)

  fm &= new(localFm)
  
  fm.init()
  fm.init(labelDemo, errors)
  
  fm.SetEnvironment(conn)
  
  return
! end fileMangerSetup -------------------------------------------------------------

! --------------------------------------------------
! overloaded init method so the buffer and some other defaults can be set
! --------------------------------------------------
localFm.Init PROCEDURE

  code

  self.Initialized = False
  self.Buffer &= labelDemo.Record
  self.FileNameValue = 'LabelDemo'
  self.SetErrors(Errors)
  self.File &= LabelDemo
  parent.Init()

  return
! end init --------------------------------------------

freeQueues procedure()

  code

  clear(demoQueue)
  free(demoQueue)
  clear(secondDemoQueue)
  free(secondDemoQueue)
  clear(thirdDemoQUeue)
  free(thirdDemoQUeue)

  return
! end freeQueues --------------------------------------

makeString procedure() ! string

s       string(60),auto
lenStr  long,auto
x       long,auto
  code

  lenStr = random(1, 55)
  loop x = 1 to lenStr 
    s[x] = chr(random(65, 97))
  end 

  return s
! end makeString ----------------------------------------  

fillQueue procedure(long numberRows) 

x   long,auto

  code

  loop x = 1 to numberRows
    demoQueue.SysId = 0
    demoQueue.Label = makeString()
    demoQueue.Amount = random(10, 5000)
    add(demoQueue)
  end 
  
  return