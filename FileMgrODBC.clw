  member

  include('FileMgrODBC.inc'),once
  include('odbcTypes.inc'),once
  include('odbcConn.inc'),once
  include('odbcErrorCl.inc'),once

  map 
  end

  itemize(900),pre(fmErr)
ConnectionInformation equate
ConnectionError       equate
StatementInformation  equate
StatementError        equate
  end

! indicates the message is an information message from the driver
informationIdCode equate('01')

fmErrors GROUP
Number USHORT(4)
       USHORT(fmErr:ConnectionInformation)
       BYTE(Level:Notify)
       PSTRING('Connection Information Message')
       PSTRING('SQL State -> %File.|%Message.')
       USHORT(fmErr:ConnectionError)
       BYTE(Level:Notify)
       PSTRING('Connection Error')
       PSTRING('SQL State -> %File.|%Message.')
       USHORT(fmErr:StatementInformation)
       BYTE(Level:Notify)
       PSTRING('Statment Information Message')
       PSTRING('SQL State -> %File.|%Message.')
       USHORT(fmErr:StatementError)
       BYTE(Level:Notify)
       PSTRING('Statment Error')
       PSTRING('SQL State -> %File.|%Message.')
  end       
ErrorMessagedAdded  long(0)

! ---------------------------------------------------------------------
! overloaded file manager init method, 
! used to provide a location to do some set up for this object
! ---------------------------------------------------------------------
FileMgrODBC.Init procedure() !,virtual 

  code 

  ! call the base object
  parent.Init()
  if (ErrorMessagedAdded <= 0) 
    !PROCEDURE(USHORT Id,*STRING Message,*STRING Title,BYTE Fatality,ASTRING Category),VIRTUAL
    self.errors.addErrors(fmErrors)
    ErrorMessagedAdded += 1
  end 

  self.errs &= new (ODBCErrorClType)
  self.errs.init()

  self.odbcExec &= new(odbcExecType)
  self.odbcExec.init(self.errs)

  self.odbcCall &= new(odbcCallType)
  self.odbcCall.Init(self.errs)

  ! allocate the column and parameter objects
  self.columns &= new(columnsClass)
  self.columns.Init()
  
  self.Parameters &= new(ParametersClass)
  self.Parameters.Init()

  return 
! end Init -------------------------------------------------------------

FileMgrODBC.SetInformationMessages procedure(bool onOff)

  code 

  self.InformationMessages = onOff

  return
! end SetInformationMessages -------------------------------------------  

! ----------------------------------------------------------
! sets the instance connection member to the one input
! that was created by the caller.   allocates the odbc instance for the object.
! set the default of false for the information messages
! ---------------------------------------------------------
FileMgrODBC.SetEnvironment  procedure(*ODBCConnectionClType  conn)

  code 

  self.conn &= conn

  ! no info messages is the default
  self.SetInformationMessages(false)

  return
! end SetEnviorment ----------------------------------------------------------  

! ----------------------------------------------------------
! clears any columns in the queue.  typically called 
! before after some operation so the columns are cleared 
! for the next binding call.  this can be called when ever 
! it is needed, before an operation or after.  
! do not call duing an operation or the bound columns will be lost
! ---------------------------------------------------------
FileMgrODBC.ClearColumns   procedure()

  code

  self.columns.clearQ()

  return
! end ClearColumns ------------------------------------------------------------

! ----------------------------------------------------------
! clears any parameters  in the queue.  typically called 
! after some operation to clear the queue for the next pass
! 
! can be called any time, before or after
! if calling after be sure the operation has completed
! or the bound parameters will be lost
! ---------------------------------------------------------
FileMgrODBC.ClearParameters procedure()

  code

  self.Parameters.clearQ()

  return
! end ClearParameters ---------------------------------------------------------

! ----------------------------------------------------------
! clears any columns and parameters used.  
! this is called from the disconnect function if the connection was 
! opened here, if not then the caller will need to do any clean up
!
! main issue with clearing is the use of multiple resutl sets.
! clearing the parameters would be fine, becasue the call has completed
! the columns may need to be cleared using a different queue in the second and 
! following reqult sets and in some cases they may 
! not need to be cleared, using the same queue for a like result set,  
! best to leave to the developer for the specific instance.
! ---------------------------------------------------------
FileMgrODBC.ClearInputs procedure() !,virtual

  code

  self.ClearColumns()
  self.ClearParameters()

  return
! end ClaerInputs ---------------------------------------------------------   

! --------------------------------------------------------------------------
! execute the sql statement input 
! this call does not return a result set but may use parameters
!
! this can be used to execute queries that do DML statements or 
! DDL statements, 
! 
! sqlStatement contains the code to be sent to the server. may be any 
! valid sql statment.  
!  
! example DML,
! insert into schema.Table(col_one, colTwo) values(?, ?);
!
! example DDL,
! alter database <database name> set single_user with rollback immediate;
! --------------------------------------------------------------------------
FileMgrODBC.ExecuteNonQuery procedure(*IDynStr sqlStatement) !,virual,SQLRETURN

retv        sqlReturn(SQL_SUCCESS)
openedHere  byte,auto

  code
 
  openedHere = self.OpenConnection()

  if (openedHere <> Connection:Failed) 
    retv = self.odbcExec.execQuery(self.conn.getHStmt(), sqlStatement)    
  end 

  self.closeConnection(openedHere)

  return retv
! end ExecuteNonQuery -------------------------------------------------------------

! --------------------------------------------------------------------------
! execute the sql statment input 
! and get the value of an output parameter.  
! this call does not return a result set
! --------------------------------------------------------------------------
FileMgrODBC.ExecuteNonQueryOut  procedure(*IDynStr sqlStatement) !,virtual,sqlreturn

retv        sqlReturn(SQL_ERROR)
openedHere  byte,auto
rows        short,auto

  code
 
  openedHere = self.OpenConnection()

  if (openedHere <> Connection:Failed) 
    retv = self.odbcExec.execQuery(self.conn.getHStmt(), sqlStatement, self.Parameters)
    if (retv = SQL_SUCCESS)
      if (self.odbcExec.nextResultSet(self.conn.getHStmt()) = true)
        retv = self.odbcExec.fetch(self.conn.getHStmt())
      end
    end 
  end

  self.closeConnection(openedHere)

  return retv
! end ExecuteQueryOut ------------------------------------------------------------- 

! --------------------------------------------------------------------------
! execute the sql statment input 
! this call will fill the queue with a result set
! --------------------------------------------------------------------------
FileMgrODBC.ExecuteQuery  procedure(*IDynStr sqlStatement, *queue q) !,virtual,sqlreturn

retv        sqlReturn(SQL_ERROR)
openedHere  byte,auto

  code
 
  openedHere = self.OpenConnection()
  
  if (openedHere <> Connection:Failed)
    retv = self.odbcExec.execQuery(self.conn.getHStmt(), sqlStatement, self.columns, self.Parameters, q)
  end 

  self.closeConnection(openedHere)

  if (retv <> SQL_SUCCESS) 
    self.ShowErrors()
  end 
  
  return retv
! end ExecuteQuery -------------------------------------------------------------

! --------------------------------------------------------------------------
! execute the sql statment input 
! and get the value of one or more output parameters.  
! this call also returns a result set
! 
! note, the result set is processed first then the out parameters are set
! --------------------------------------------------------------------------
FileMgrODBC.ExecuteQueryOut  procedure(*IDynStr sqlStatement, *queue q) !,virtual,sqlreturn

retv        sqlReturn(SQL_ERROR)
openedHere  byte,auto

  code
 
  openedHere = self.OpenConnection()

  if (openedHere <> Connection:Failed) 
    retv = self.odbcExec.execQuery(self.conn.getHStmt(), sqlStatement, self.columns, self.Parameters, q)
    if (retv = SQL_SUCCESS)
      if (self.odbcExec.nextResultSet(self.conn.getHStmt()) = true)

      end
    end 
  end
  
  self.closeConnection(openedHere)

  return retv
! end ExecuteQueryOut ------------------------------------------------------------- 

! --------------------------------------------------------------------------
! calls a scalar function and retruns the returned value
! --------------------------------------------------------------------------
FileMgrODBC.ExecuteScalar procedure(*IDynStr scalarQuery) !,long,virtual

retv        long,auto
openedHere  byte,auto

  code

  openedHere = self.OpenConnection()
  
  if (openedHere <> Connection:Failed) 
    retv = self.odbcExec.ExecQuery(self.conn.getHStmt(), scalarQuery, self.Parameters)
  end 

  self.closeConnection(openedHere)
  
  if (retv <> SQL_SUCCESS) 
    self.ShowErrors()
  end 

  return retv
! end ExcuteSclar --------------------------------------------------------

! --------------------------------------------------------------------------
! execute the stored procedure input in the string 
! this stored procedure does not return a result set but 
! may have output parameters
! --------------------------------------------------------------------------
FileMgrODBC.callSp procedure(string spName) !,virtual,sqlreturn

retv        sqlReturn(SQL_ERROR)
openedHere  byte,auto

  code
 
  openedHere = self.OpenConnection()

  if (openedHere <> Connection:Failed) 
    retv = self.odbcCall.ExecSp(self.conn.getHStmt(), spName, self.Parameters)
  end 
  
  self.closeConnection(openedHere)

  return retv
! end ExecuteSpOut ----------------------------------------------------------

! --------------------------------------------------------------------------
! execute the stored procedure input in the string 
! this stored procedure returns a result set
! --------------------------------------------------------------------------
FileMgrODBC.callSp procedure(string spName, *queue q) !,virtual,sqlreturn

retv       sqlReturn(SQL_ERROR)
openedHere byte,auto

  code

  openedHere = self.OpenConnection()

  if (openedHere <> Connection:Failed) 
    retv = self.odbcCall.execSp(self.conn.getHStmt(), spName, self.columns, self.Parameters, q)
  end

  self.closeConnection(openedHere)
  
  if (retv <> SQL_SUCCESS) 
    self.ShowErrors()
  end 

  return retv
! end ExecuteSp ------------------------------------------------------------

! --------------------------------------------------------------------------
! execute the stored procedure input in the string 
! this stored procedure returns multiple result sets
! 
! this one needs to be overloaded in a derived instance
! there will be more than one buffer is use and those buffers
! may change for each result set.  the queue parameter will be used 
! for the first result set.  the remaining result sets will have a buffer 
! defined in the over loaded function
! --------------------------------------------------------------------------
FileMgrODBC.callSpMulti procedure(string spName, *queue q) !,virtual,sqlreturn

  code
  return SQL_ERROR
! end callSpMulti -----------------------------------------------------------  

! --------------------------------------------------------------------------
! calls a scalar function and retruns the returned value
! --------------------------------------------------------------------------
FileMgrODBC.callScalar procedure(string  fxName) !,long,virtual

retv        long(SQL_ERROR)
openedHere  byte,auto

  code

  openedHere = self.OpenConnection()
  
  if (openedHere <> Connection:Failed) 
    retv = self.odbcCall.callScalar(self.conn.getHStmt(), fxName, self.Parameters)
  end 

  self.closeConnection(openedHere)
  
  return retv
! end ExcuteSclar --------------------------------------------------------

! --------------------------------------------------------------------------
! read the second, third, ... resutls sets.  this can be from a query 
! or a stored procedure.  
! --------------------------------------------------------------------------
FileMgrODBC.readNextResult procedure(*queue q, *ColumnsClass cols) !,sqlreturn

retv  sqlReturn,auto

  code

  ! move to the next result set
  if (self.odbcExec.nextResultSet(self.conn.getHStmt()) = true)
     if (cols.bindColumns(self.conn.getHStmt()) = SQL_SUCCESS)
       retv = self.odbcExec.readNextResult(self.conn.getHStmt(), q)
     end
  else 
    ! no more results
    retv = SQL_NO_DATA   
  end

  return retv
! end readNextResult -------------------------------------------------------

! --------------------------------------------------------------------------
! opens a connection for a call to the server, if the connection is not 
! currently open.  
! returns Connection:Opened if the connection was opened,
! Connection:CallerOpened if the connection is already opened
! and Connection:Failed if it was not open and the open attempt failed.
!
! the input withStatement defaults to true, if it is true then the 
! hStmt is also allocated. if false the hStmt is not allocated and 
! if the connectin fails the hStmt is not allocated
! 
! as a some what general rule, the file manager will open a connection
! and get a hStmt, do some work and close the connection.  the default is to 
! allocate the hStmt, in cases where the handle needs to allcated seperatly 
! call with false for the parmeter.
! --------------------------------------------------------------------------
FileMgrODBC.OpenConnection procedure(bool statement = withStatement) ! ,byte

res  sqlreturn,auto
retv byte,auto

  code 
 
  ! if not open then open it
  if (self.conn.isConnected() = false)
    res = self.conn.connect();
    case res 
      of SQL_SUCCESS
        retv = Connection:Opened
      of SQL_SUCCESS_WITH_INFO
      self.getConnectionError()
      retv = Connection:Opened
    else
      retv = Connection:Failed      
      self.getConnectionError()
    end  
  else 
    retv = Connection:CallerOpened
  end

  ! if a hStmt is wanted and the connection did not fail
  ! in some cases the withStatement will be false so don't allocate
  if (withStatement = true) and (retv <> Connection:Failed)
    ! this allocates the statement handle, the allocate handle function
    ! only returns one information message and I have never seen the MS driver return that value
    ! for a statment handle, the call can fail and return an error so check for both just to be sure
    res = self.conn.AllocateStmtHandle()
    if (res <> SQL_SUCCESS) and (res <> SQL_SUCCESS_WITH_INFO)
      self.getConnectionError()
    end
  end 

  return retv
! end OpenConnection -------------------------------------------------------

! --------------------------------------------------------------------------
! closes  a connection
! the input value should be from the returned value of the OpenConnection call
! if the input is Connectio:Opened then connection is closed.  
! if any other value the connection is not closed because it was opened by the caller 
! or the open attempt failed
!
! note if the connection was opened here then the columns and parameters are cleared.
! if opened by some other calling code they are not cleared and the user must 
! do any needed clean up.  
! --------------------------------------------------------------------------
FileMgrODBC.CloseConnection procedure(byte openedHere) 

  code

  if (openedHere = Connection:opened)
    !self.ClearInputs()
    self.conn.disconnect()
  end 

  return
! end closeConnection ------------------------------------------------------


fileMgrOdbc.getConnectionError procedure() 
 
   code

   self.errs.getError(SQL_HANDLE_DBC, self.conn.gethDbc())

   self.ShowErrors()

   return

FileMgrODBC.ShowErrors procedure() 

numErrs  byte,auto
x        long,auto

stateText string(5)
msgtext   string(3000)

  code

  numErrs = records(self.errs.errorMsgQ)

  loop x = 1 to numErrs
    get(self.errs.errorMsgQ, x)
    ! check the first two characters of the SQL State code
    ! and skip if information messages are not wanted
    if (sub(self.errs.errorMsgQ.sqlState, 1, 2) = informationIdCode) 
      if (self.InformationMessages = true) 
        self.errors.setFile(self.errs.errorMsgQ.sqlState)
        self.errors.throwMessage(fmErr:ConnectionInformation, self.errs.errorMsgQ.messageText)
      end  
    else 
      self.errors.setFile(self.errs.errorMsgQ.sqlState)
      self.errors.throwMessage(fmErr:ConnectionError, self.errs.errorMsgQ.messageText)
    end
  end  

  return
! end ShowErrors ----------------------------------------------------------

! ----------------------------------------------------------------------
! allocates the bcp instance for the file manager and 
! calls the init_bcp function to do some set up
! ----------------------------------------------------------------------
fileMgrOdbc.init_Bcp    procedure() !,byte

retv    byte,auto

  code 

  self.bcp &= new(bcpType)

  retv = self.bcp.init_Bcp()

  return retv
! end init_Bcp ---------------------------------------------------------

! ----------------------------------------------------------------------
! gets a connection to the database and sets the required connection
! attributes used by the BCP
! ----------------------------------------------------------------------
FileMgrODBC.connectBcp procedure() !,byte

retv   byte,auto

  code

  retv = self.bcp.connect(self.conn.connStr.connectionString())

  return retv  
! end connectBcp ---------------------------------------------------

! ----------------------------------------------------------------------
! callst he bcp disconnect function to free the handles and 
! disposes the instance 
! ----------------------------------------------------------------------
FileMgrODBC.disconnectBcp procedure() !,long

retv   long

  code

  self.bcp.disconnect()
  dispose(self.bcp) 
  self.bcp &= null

  return
! end disconnectBcp ------------------------------------------------

! ----------------------------------------------------------------------
! sets the number of rows to be used for a the batch during the bcp operations
! the default is zero, no batch size, set as needed after testing for 
! best performance
! ----------------------------------------------------------------------
FileMgrODBC.setBcpBatchSize procedure(long rows)

  code 

  if (~self.bcp &= null) 
    self.bcp.setBcpBatchSize(rows)
  end 

  return
! end  setBcpBatchSize ------------------------------------------------ 