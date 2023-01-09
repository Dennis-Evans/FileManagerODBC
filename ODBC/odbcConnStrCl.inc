
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCConnectionStrPresent_)
_ODBCConnectionStrPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once

! some default string lengths of the connections   
eDriverNameLength equate(100)  
eDefaultLength    equate(128)

! ---------------------------------------------------------------------------
! baseConnStrClType, class to handle the standard connection strings.  Each 
! vendor will use this as the starting point, derive the class for any specifics. 
!
! 
! ---------------------------------------------------------------------------
baseConnStrClType  class,type,module('odbcConnStrCl.clw'),link('odbcConnStrCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! -----------------------------------------
! instance to store the built connection string
! -----------------------------------------
connStr              &IDynStr,private

! -------------------------------------------------------------
! connection string fields, access is through the interface setters
! getters could be added but there is seldom any need to query these fields
! -------------------------------------------------------------
srvName              cstring(eDefaultLength),private
dbName               cstring(eDefaultLength),private
driverName           cstring(eDriverNameLength),private
userName             cstring(eDefaultLength),private
password             cstring(eDefaultLength),private
appname              cstring(eDefaultLength),private

trustedConn          cstring(eDefaultLength),private
useMars              cstring(eDefaultLength),private

! ------------------------------------------------------
! setters for the conncetion string fields
! ------------------------------------------------------
setDriverName        procedure(string driverName)
setDbName            procedure(string dbname)
setSrvName           procedure(string srvName)
setUserName          procedure(string  user)
setPassword          procedure(string pw)
setAppname           procedure(string appName)

! ---------------------------------------------------------------
! gets the connection, the string is built from the various parts
! that are in use.
! 
! Note, always implement this method, the actual format of the connection 
! string varies significantly between vendors
! ---------------------------------------------------------
ConnectionString    procedure(),*cstring,virtual

! setup and clean up, 
init                 procedure(),virtual,byte,proc
kill                 procedure(),virtual

findTokens           procedure(string s, connStrLabelQueue q),private
                   end
! ---------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! connection string class for MS. 
! 
! ---------------------------------------------------------------------------               
MSConnStrClType  class(baseConnStrClType),type,module('odbcConnStrCl.clw'),link('odbcConnStrCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! --------------------------------------------------------------------------------
! init the object and reads the connection string fields from the registry
! returns level:benign for succes and notify if any of the required fields 
! are missing
! --------------------------------------------------------------------------------
init               procedure(),virtual,byte

! --------------------------------------------------------------------------------
! init the object and uses the server name and database name input. the default 
! driver label is used. 
! --------------------------------------------------------------------------------
init               procedure(string srvName, string dbName),virtual,byte,proc

! --------------------------------------------------------------------------------
! init the object and uses the server name, database name, user name and 
! password input. the default driver label is used. 
! --------------------------------------------------------------------------------
init               procedure(string srvName, string dbName, string user, string pw),virtual,byte,proc

! --------------------------------------------------------------------------------
! init the object and complete connection string input
! --------------------------------------------------------------------------------
Init               procedure(string fullConnStr),virtual,byte,proc

! ------------------------------------------------------------------------
! sets the trusted connection option for the conenction.
! this is done in the connection string so it will on or off for the 
! duration of the connection.
! 
! Note, if off then user and password are required for sql auth
! ------------------------------------------------------------------------ 
setTrustedConn     procedure(bool onOff)

! ------------------------------------------------------------------------
! turns mars on or of for the conenction.
! this is done in the connection string so it will on or off for the 
! duration of the connection 
! ------------------------------------------------------------------------ 
setUseMars         procedure(bool onOff)

! ------------------------------------------------------------------------
! returns the connection, as a cstring, the conenction string returned is 
! built from the various memeber fields. 
! ------------------------------------------------------------------------ 
ConnectionString   procedure(),*cstring,virtual
 
! ------------------------------------------------------------------------
! reads a registry value that will contain the various connection string 
! options.  the returned string is clipped
! ------------------------------------------------------------------------ 
readRegValue       procedure(string nodeLabel),private,string

! ------------------------------------------------------------------------
! reads a registry value thatwill contain yes/no or an empty string
! returns the string value read.  the returned string is in lower case
! ------------------------------------------------------------------------
readRegValueYesNo  procedure(string nodeLabel),private,string

! ------------------------------------------------------------------------
! called when the init method that acceptes a single string (the fullConnstr) 
! is used.  the function will parse out the indivdual elements and assign 
! them to the objects fields.
! if the init(fullConnStr) is not used then this function is not called.
! ------------------------------------------------------------------------
parseConnStr       procedure(string cs) 
                 end 
! ---------------------------------------------------------------------------

   _EndOfInclude_                   
