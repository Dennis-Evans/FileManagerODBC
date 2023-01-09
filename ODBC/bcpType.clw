  member()
  
  include('bcpType.inc'),once
  include('odbcTypes.inc'),once

  map 
    module('bcp')
       ! called once t odo some set up, returns an env. handle store the handle locally
       ClaBcpInit(),long,c
       ! called once to establish a connection to the database
       ! and set some connection attributes for BCP
       ! returns a connection hanlde that is used bythe BCP API calls, store locally.
       claBcpConnect(SQLHENV hEnv, *cstring connStr),long,c,raw
       ! shuts down the connection and closes the two handles
       ClaBcpKill(SQLHENV hEnv, SQLHDBC hDbc),LONG,PROC,c
       ! sets up the table input for the insert
       ! use the tablename like this, schemaName.TableNAme
       init_Bcp(SQLHDBC hDbc, *cstring tName),long,c,raw
       ! sends a row
       sendRow_Bcp(SQLHDBC hDbc),long,c
       ! commit a batch of rows to the database
       batch_Bcp(SQLHDBC hDbc),long,c
       ! commit all rows to the database and does some clean up on the server.
       ! this function must be called when the process is complete.
       done_Bcp(SQLHDBC hDbc),long,c
       ! bind a long variable to a table column.
       ! input the connection handle, the local variable and the ordinal position of the
       ! column in the table.  this is one based not zero based.
       bindLong(SQLHDBC hDbc, *long colv, long colOrd),long,c,name('bind_Bcpl')
       ! bind a real variable
       bindReal(SQLHDBC hDbc, *real colv, long colOrd),long,c,name('bind_bcpf')
       ! bind a string variable.  note sLen parameter. this should be the size of the
       ! clarion string, use size(string) do not use len(clip(string))
       ! internally, in the C dll, each bind call sends the server the size of the data type
       ! strings can vary in size so the extra parameter is needed.  for most
       ! of the data types the size is known.
       ! use this for the char(x) columns
       bindString(SQLHDBC hDbc, *string colv, long colOrd, long slen),ushort,c,raw,name('bind_bcps')
       ! bind a clarion cstring.  Note the size parameter is not used for the cstring.
       ! the system will find the length to insert.
       ! use this for the varchar(x) columns
       bindCStr(SQLHDBC hDbc, *cstring colv, long colOrd),long,c,raw,name('bind_bcpcs')
       ! bind a boolean variable
       bindBool(SQLHDBC hDbc, *bool colv, long colOrd),long,c,name('bind_bcpb')
       ! bind a date variable
       bindDate(SQLHDBC hDbc, *date_Struct colv, long colOrd),long,c,raw,name('bind_bcpd')
       ! bind a date time variable,
       ! used a string in the standard ODBC format for date and times
       ! be sure the dates are formatted correctly, yyyy-mm-dd, use leading zeros
       bindDateTime(SQLHDBC hDbc, *dateTimeString colv, long colOrd),long,c,raw,name('bind_bcpdt')
       ! bind a byte variable
       bindByte(SQLHDBC hDbc, *byte colv, long colOrd),long,c,name('bind_Bcpby')
       ! bind a short variable
       bindshort(SQLHDBC hDbc, *short colv, long colOrd),long,c,name('bind_BcpSh')
       ! bind a sreal variable
       bindSReal(SQLHDBC hDbc, *sreal colv, long colOrd),long,c,name('bind_Bcpsf')
       ! used a cstring for the time.  watch the time(n) closely
       ! if it over flows the row will not be inserted
       ! be sure the times are formatted correctly, hh:mm:ss.fraction, use leading zeros
       bindTime(SQLHDBC hDbc, *timeString t, long colOrd),long,c,raw,name('bind_Bcpt')
    end
  end

! ----------------------------------------------------------------------
! initilizes the object and creates a seperate hEnv for use by 
! the BCP operation.
! ----------------------------------------------------------------------
bcpType.init_bcp procedure()   

retv     byte(level:benign)

  code 
  
  self.hEnv = ClaBcpInit() 
  if (self.hEnv <= 0) 
    retv = level:notify
  end   
     
  self.rowsSent = 0
  
  return retv 
! end Init
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! gets a connection to the databse and set the connection attributes 
! for the BCP
! ----------------------------------------------------------------------
bcpType.connect procedure(*cstring connStr)

retv    byte(Level:benign)

  code

  self.hDbc = ClaBcpConnect(self.hEnv, connStr)
  if (self.hDbc <= 0) 
    retv = Level:notify
  end 

  return retv
! end connect ----------------------------------------------------------

! ----------------------------------------------------------------------
! frees the hEnv and hDbc used by the BCP 
! ----------------------------------------------------------------------
bcpType.disconnect procedure()

  code 

  ClaBcpKill(self.hEnv, self.hDbc)

  return
! end disconnect ------------------------------------------------------

! ----------------------------------------------------------------------
! sets up the table for the insert 
! ----------------------------------------------------------------------
bcpType.init_Bcp procedure(*cstring tName)

retv  bool,auto

  code 

  retv = init_Bcp(self.hDbc, tName)

  return retv 
! end init_bcp ----------------------------------------------------------

! ----------------------------------------------------------------------
! sends a row to the server and checks the batch size. 
! if the number of rows sent is greater than or equal to the batch size 
! the data is written
! ----------------------------------------------------------------------
bcpType.sendRow procedure()

retv       bool,auto
rows       long,auto

  code 

  if (sendRow_Bcp(self.hDbc) = bcp_fail) 
    retv = false
  else 
    retv = true
    self.rowsSent += 1
    ! if the batch size has been set then 
    ! check for a batch,  
    if (self.batchSize > 0) 
      if (self.rowsSent >= self.batchSize)
        rows  = batch_Bcp(self.hDbc)
        self.rowsSent = 0
      end 
    end
  end 

  return retv 
! end send_Row ------------------------------------------------------------

! ----------------------------------------------------------------------
! writes a batch of rows to the server 
! returns the number of rows written
! ----------------------------------------------------------------------
bcpType.batch_Bcp procedure()

retv  long,auto

  code 

  retv = batch_Bcp(self.hDbc)

  return retv 
! end batch_bcp --------------------------------------------------------

! ----------------------------------------------------------------------
! writes the data to the server, if there is any, and shuts down the BCP 
! operations on the server.   
! returns the number of rows written
! this function MUST be called when the insert is completed
! ----------------------------------------------------------------------
bcpType.done_Bcp procedure()

retv  long

  code 

  retv = done_Bcp(self.hDbc)

  return retv 
! end done_bcp --------------------------------------------------------

! ----------------------------------------------------------------------
! sets the number of rows to be used for a the batch during the bcp operations
! the default is zero, no batch size, set as needed after testing for 
! best performance
! ----------------------------------------------------------------------
bcpType.setBcpBatchSize procedure(long rows)

  code 

  self.batchSize = rows
  
  return
! end setBcpBatchSize -------------------------------------------------

! ---------------------------------------------------------------------
! add a column to the bcp of the insert
! each of these functions take field (from a queue, tps file, what ever) 
! and bind the field to the bcp layer.  the colOrd is the ordinal position of the 
! of the field in the table.  starting at 1.
! 
! all of the functions do the same thing, 
! note the comments on addColumnDateTime, addColumnTime and addColumnBool
! --------------------------------------------------------------------
bcpType.addColumn procedure(*byte colv, long colOrd)

retv bool 

  code

  retv = bindByte(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

bcpType.addColumn procedure(*short colv, long colOrd)

retv bool 

  code

  retv = bindshort(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

bcpType.addColumn procedure(*long colv, long colOrd)

retv bool 

  code

  retv = bindLong(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

!  -----------------------------------------------------------
! function to add a boolean to the insert
! a bool is equated to a long so the function label is specific
!  -----------------------------------------------------------
bcpType.addColumnBool procedure(*bool colv, long colOrd)

retv bool 

  code

  retv = bindBool(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

bcpType.addColumn procedure(*DATE_STRUCT colv, long colOrd)

retv bool 

  code

  retv = bindDate(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

bcpType.addColumn procedure(*real colv, long colOrd)

retv bool 

  code

  retv = bindReal(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

bcpType.addColumn procedure(*string colv, long colOrd)

retv bool 
strLen  long,auto

  code

  strLen = size(colv)
  retv = bindString(self.hDbc, colv, colOrd, strLen)

  return retv
! end addColumn -----------------------------------------------------------

bcpType.addColumn procedure(*cstring colv, long colOrd)

retv bool 

  code
  stop(colv)
  retv = bindCStr(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

bcpType.addColumn procedure(*sreal colv, long colOrd)

retv bool 

  code

  retv = bindSReal(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

!  -----------------------------------------------------------
! function to add a datetime value, in a string format to the insert
! the string must be formatted as a  valid datetime string
!  -----------------------------------------------------------
bcpType.addColumnDate procedure(*dateTimeString colv, long colOrd)

retv bool 

  code

  retv = bindDateTime(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------

!  -----------------------------------------------------------
! function to add a time value, in a string format to the insert
! the string must be formatted as a  valid time string
!  -----------------------------------------------------------
bcpType.addColumnTime procedure(*timeString colv, long colOrd)

retv bool 

  code

  retv = bindTime(self.hDbc, colv, colOrd)

  return retv
! end addColumn -----------------------------------------------------------  