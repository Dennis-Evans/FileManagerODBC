
  member()
  
  include('odbcConnStrCl.inc'),once 
  include('cwsynchc.inc'),once

  map 
    include('clib.clw')
  end

eTrustedConnTextOn   equate('Trusted_Connection=yes;')
eUseMarsConnTextOn  equate('MARS_Connections=yes;')

eDefaultRegKey      equate('software\ConnString')

! adjust this as needed
eDefaultRegPath     equate('software\ConnString')

eDefaultDriverName equate('ODBC Driver 17 for SQL Server')

!eDriverLabel         equate('Driver={{')
eDriverLabel         equate('Driver=')
eServerLabel         equate('server=')
eDbLabel             equate('Database=')
eUserLabel           equate('UserId=')
ePasswordLabel       equate('Password=')

eConnDelimit      equate(';')

critSection       CriticalSection

connStrLabelQueue  queue,type
label                string(100)
                   end

! -------------------------------------------------------------------------------------
! Init 
! -------------------------------------------------------------------------------------
baseConnStrClType.init procedure() !,virtual,byte

  code 
  
  self.connStr &= newDynStr()

  return level:benign
! end init 
! -----------------------------------------------------------------

! -----------------------------------------------------------------
! kill
! dispose the dyn string 
! -----------------------------------------------------------------
baseConnStrClType.kill procedure()

  code 
  
  disposeDynStr(self.connStr)
  self.connStr &= null
  
  return
! end kill
! -----------------------------------------------------------------

! -----------------------------------------------------------------
! virtual place holder returns null
! this function should lawqays be overloaded in a derived object.
! -----------------------------------------------------------------
baseConnStrClType.ConnectionString procedure() !,*cstring,virtual

  code 
    
  return null
! end ConnectionString
! ------------------------------------------------------------------------------

! -----------------------------------------------------------------
! Setters for the instance
! -----------------------------------------------------------------
baseConnStrClType.setDriverName procedure(string driverName)

  code 
  
  self.driverName = eDriverLabel & clip(driverName)  & eConnDelimit

  return
! end setDbName
! ------------------------------------------------------------------------------
  
baseConnStrClType.setDbName procedure(string dbname)

  code 
  
  self.dbName = eDbLabel & clip(dbName) & eConnDelimit
  
  return
! end setDriverName
! ------------------------------------------------------------------------------
  
baseConnStrClType.setSrvName procedure(string srvName)

  code 
  
  self.srvName = eServerLabel & clip(srvName) & eConnDelimit
  
  return
! end setServerName
! ------------------------------------------------------------------------------
  
baseConnStrClType.setUserName procedure(string  user)

  code 
  
  self.userName = 'User ID=' & clip(user) & eConnDelimit
  
  return
! end setUserName
! ------------------------------------------------------------------------------
  
baseConnStrClType.setPassword procedure(string pw)

  code 
  
  self.password = 'Password=' & clip(pw) & eConnDelimit
  
  return
! end setpassword
! ------------------------------------------------------------------------------

baseConnStrClType.setAppname procedure(string an) 

  code 

  self.appname = 'App=' & clip(an) & eConnDelimit

  return
! end setappName 
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! finds the various tokens in the connection string tah was input and 
! places each token in the queue input
!
! this is only called by the init(fullConnStr) function
! ------------------------------------------------------------------------------
baseConnStrClType.findTokens procedure(string s, connStrLabelQueue q) 

delimiter  cstring(';')
nullStr    &cstring
workstr    cstring(1000)
resultStr  cstring(500)

  code

  workstr = clip(s)
  resultStr = strtok(workStr, delimiter)
  loop 
    if (resultStr = '')
      break
    end  
    q.label = lower(resultStr)
    add(q)  
    resultStr = strtok(nullStr, delimiter)
    if (resultStr = '')
      break
    end  
  end 

  return
! end split ----------------------------------------------------------------  

! end base class 
! begin MS overload

! --------------------------------------------------------------------------------
! init the object and complete connection string input
! --------------------------------------------------------------------------------
MSConnStrClType.Init procedure(string fullConnStr) !,virtual

  code 

  critSection.wait()

  parent.init()

  self.parseConnStr(fullConnStr)

  critSection.release()

  return level:benign
! ------------------------------------------------------------------------------

! --------------------------------------------------------------------------------
! init the object and uses the server name and database name input. the default 
! driver label is used. 
! --------------------------------------------------------------------------------
MSConnStrClType.init procedure(string srvName, string dbName) !,virtual

  code
  
  critSection.wait()

  parent.init()

  ! assume the default driver in the equate label
  self.setDriverName(eDefaultDriverName)

  self.setSrvName(srvName)
  self.setDbName(dbName)
  ! trusted connection is on
  self.setTrustedConn(true)

  critSection.release()

  return level:benign
! ------------------------------------------------------------------------------
  
! --------------------------------------------------------------------------------
! init the object and uses the server name, database name, user name and 
! password input. the default driver label is used. 
! --------------------------------------------------------------------------------  
MSConnStrClType.init procedure(string srvName, string dbName, string user, string pw) !,virtual

  code

  critSection.wait()

  self.Init(srvName, dbname)
  self.userName = user
  self.password = pw
  ! trusted connection is off
  
  self.setTrustedConn(false)  

  critSection.release()

  return level:benign
! end init 
! -----------------------------------------------------------------

! --------------------------------------------------------------------------------
! init the object and reads the connection string fields from the registry
! returns level:benign for succes and notify if any of the required fields 
! are missing
! --------------------------------------------------------------------------------
MSConnStrClType.init procedure() !,virtual

retv      byte(level:benign)
onOff     string(3)
testStr   string(100)

  code

  critSection.wait()

  parent.init()

  testStr = self.readRegValue('Driver')
  if (testStr <> '') 
    self.SetDriverName(testStr)
  else 
    retv = level:notify
  end 

  testStr = self.readRegValue('server')
  if (testStr <> '') 
    self.SetSrvName(testStr)
  else 
    retv = level:notify
  end 
  
  testStr = self.readRegValue('database')
  if (testStr <> '') 
    self.SetDbName(testStr)
  else 
    retv = level:notify
  end 

  testStr = self.readRegValue('appname')
  ! if not found this is not an error, assume it is not in use
  if (testStr <> '') 
    self.SetAppName(testStr)
  end     

  if (self.readRegValueyesNo('trustedConnection') = 'yes')
    self.setTrustedConn(true)
  else 
    ! if sql auth then user id and password must be present
    testStr = self.readRegValue('user')
    if (testStr <> '') 
      self.setUserName(testStr)
    else 
      retv = level:notify
    end
    testStr = self.readRegValue('password')
    if (testStr <> '') 
      self.setPassword(testStr)
    else       
      retv = level:notify
    end
  end  

  if (self.readRegValueYesNo('mars') = 'yes')
    self.setUseMars(true)
  end 

  critSection.release()

  return retv
! end init 
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------
! reads a registry value that will contain the various connection string 
! options.  the returned string is clipped
! ------------------------------------------------------------------------ 
MSConnStrClType.readRegValue procedure(string nodeLabel) !,private,string

retv string(128)

  code
  
  retv = getreg(REG_CURRENT_USER, eDefaultRegKey, nodeLabel)

  return clip(retv)
! end readRegValue -------------------------------------------------------------  

! ------------------------------------------------------------------------
! reads a registry value thatwill contain yes/no or an empty string
! returns the string value read.  the returned string is in lower case
! ------------------------------------------------------------------------
MSConnStrClType.readRegValueYesNo procedure(string nodeLabel) !,private,string

retv string(3)

  code
  
  retv = getreg(REG_CURRENT_USER, eDefaultRegKey, nodeLabel)
  
  return lower(retv)
! end readRegValue -------------------------------------------------------------  

! ------------------------------------------------------------------------
! sets the trusted connection option for the conenction.
! this is done in the connection string so it will on or off for the 
! duration of the connection.
! 
! Note, if off then user and password are required for sql auth
! ------------------------------------------------------------------------ 
MSConnStrClType.setTrustedConn procedure(bool onOff)

  code 

  if (onOff = true) 
    self.trustedConn = eTrustedConnTextOn
  else 
    self.trustedConn = ''
  end
  
  return
! end setTrustedConn
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------
! turns mars on or of for the conenction.
! this is done in the connection string so it will on or off for the 
! duration of the connection 
! ------------------------------------------------------------------------ 
MSConnStrClType.setUseMars procedure(bool onOff)

  code 
  
  if (onOff = true) 
    self.useMars = eTrustedConnTextOn
  else 
    self.useMars = ''
  end
  
  return 
! end setUseMars
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------
! returns the connection, as a cstring, the conenction string returned is 
! built from the various memeber fields. 
! ------------------------------------------------------------------------ 
MSConnStrClType.ConnectionString procedure() !,*cstring,virtual

  code 
    
  critSection.wait()

  ! clear it 
  self.connStr.Kill()

  ! if the driver was not added use the default value
  if (self.driverName = '') 
    self.driverName = eDriverLabel & clip(eDefaultDriverName)  & eConnDelimit
  end 

  ! add the defaults that are always used  
  self.connStr.cat(self.driverName & self.SrvName &  self.dbName)
  ! if not empty the add 
  if (self.trustedConn <> '') 
    self.connStr.cat(self.trustedConn)
  else 
    ! if it was empty then add sql auth
    self.connStr.cat(self.userName)
    self.connStr.cat(self.password)
  end  
  ! add these, if they are empty then nothing is added t othe string
  self.connStr.cat(self.appname)
  self.connStr.cat(self.useMars)

  critSection.release()

  return self.connStr.cstr()
! end ConnectionString
! ------------------------------------------------------------------------------

MSConnStrClType.parseConnStr procedure(string cs) 

startPos   long,auto
!endPos     long,auto
labels     connStrLabelQueue
x          long,auto

  code

  critSection.wait()
 
  self.findTokens(cs, labels)

  loop x = 1 to records(labels)
    get(labels, x)    
    startPos = instring('=', labels.label)
    ! if equal sign is missing, skip this one, it cannot be valid
    if (startPos <= 0) 
      cycle
    end  
    case sub(labels.label, 1, startPos - 1) 
    of 'driver'
      self.DriverName = clip(labels.label) & eConnDelimit
    of 'server'
      self.srvName = clip(labels.label) & eConnDelimit
    of 'database'
      self.dbname = clip(labels.label) & eConnDelimit
    of 'trusted_connection'  
      if (instring('yes', labels.label, 1) > 0)
        self.TrustedConn = clip(labels.label) & eConnDelimit
      else 
        self.TrustedConn = ''
      end 
    of 'uid'
      self.userName = clip(labels.label) & eConnDelimit
      self.trustedConn = ''
    of 'pwd' 
      self.password = clip(labels.label) & eConnDelimit
      self.trustedConn = ''
    of 'mars_connection'
      if (instring('yes', labels.label, 1) > 0)
        self.useMars = clip(labels.label) & eConnDelimit
      else 
        self.useMars = ''
      end 
    of 'app'
      self.appname = clip(labels.label) & eConnDelimit
    end ! case 

  end ! loop

  critSection.release()

  return
! end parseConnStr ---------------------------------------------------------- 
