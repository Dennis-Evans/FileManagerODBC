
  member()
  
  include('odbcColumnsCl.inc'),once 

  map 
    module('odbc32')
      SQLBindCol(SQLHSTMT StatementHandle, SQLUSMALLINT ColumnNumber, SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength, *SQLLEN  StrLen_or_Ind),sqlReturn,pascal
    end
  end
! ---------------------------------------------------------------------------

defaultFloating equate(0.0)
defaultString   equate('')
defaultInteger  equate(0)
defaultBoolean  equate(0)

sizeOfLong    equate(4)
sizeOfDate    equate(4)
sizeOfReal    equate(8)

! ---------------------------------------------------------------------------
!  default constructor, calls the init function for the set up
! ---------------------------------------------------------------------------
columnsClass.construct procedure()  

  code 
  
  self.init()
  
  return 
! end construct
! ------------------------------------------------------------------------------
  
! ---------------------------------------------------------------------------
!  allocates the queue 
! ---------------------------------------------------------------------------  
columnsClass.init procedure()

retv      byte(level:benign)

  code 

  self.colq &= new(columnsQueue)
  if (self.colQ &= null)
    return level:notify
  end 
    
  self.colB &= new(ColumnsLarge)
  if (self.colb &= null)
    return level:notify
  end 

  return retv
! end init 
! ------------------------------------------------------------------------------
  
! ------------------------------------------------------------------------------
! disposes the queue and the dyn str
! ------------------------------------------------------------------------------
columnsClass.kill procedure()

x  long,auto

  code 

  self.clearQ()

  dispose(self.colQ)
  self.colQ &= null
  dispose(self.colB)
  self.colb &= null
  
  return
! end kill
! ------------------------------------------------------------------------------

! ---------------------------------------------------------------------------
!  default destructor, calls the kill function for the clean up
! ---------------------------------------------------------------------------
columnsClass.destruct procedure() ! virtual

  code 

  self.kill()  
  
  return 
! end destruct
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! free the queue and the dyn str
! ------------------------------------------------------------------------------
columnsClass.clearQ procedure()

x    long,auto

  code 
  
  clear(self.colQ)
  free(self.colQ)
  self.allowNulls = false

  clear(self.colb)
  free(self.colb)

  return
! end clearQ
! ------------------------------------------------------------------------------

! bindCols
! Bind the queue, group or seperate fields to the columns in the result set.
! column order is typically the  same order as the select statment,
! 
! parameters for the ODBC api call 
! hStmt   = handle to the ODBC statement
! colId = ord value of the parmaeter, 1, 2, 3 ... the ordinal position
! colType = the C data type of the column 
! ColBuffer = pointer to the buffer field 
! colSize = the size of the buffer or the queue field 
! colInd = pointer to a buffer for the size of the parameter. not used and null in this example 
! -----------------------------------------------------------------------------    
columnsClass.bindColumns procedure(SQLHSTMT hStmt) ! sqlReturn

retv      sqlReturn(SQL_SUCCESS)  ! set ot success at the start, some calls will have an empty queue
x         long,auto         

  code 
  
  ! iterate over the list, if any fail return an error
  loop x = 1 to records(self.colq)
    get(self.colQ, x)
    retv = SQLBindCol(hStmt, self.colQ.colId, self.Colq.colType, self.colQ.ColBuffer, self.Colq.colSize, self.ArrayPtr[x])
    if (retv <> sql_Success) and (retv <> Sql_Success_With_Info) 
      break
    end  
  end
  
  return retv 
! end bindColumns
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! the various addColumn functions are called by the using code and are used for the
! specific data types.  each calls the AddColumn/3 function to actually 
! add a columns
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*byte colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code

  self.addColumn(SQL_C_TINYINT, address(colPtr), size(colPtr), allowNulls, actualQueuePos)

  return
! end AddColumn ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*short colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code

  self.addColumn(SQL_C_SHORT, address(colPtr), size(colPtr), allowNulls, actualQueuePos)

  return
! end AddColumn ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*long colPtr, bool allowNulls = false, long actualQueuePos = -1) 

  code 
  
  self.addColumn(SQL_C_SLONG, address(colPtr), size(colPtr), allowNulls, actualQueuePos)
   
  return
! end AddColumn ------------------------------------------------------------------------------
 
columnsClass.AddColumn procedure(*string colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 

  self.addColumn(SQL_C_CHAR, address(colPtr), len(colPtr), allowNulls, actualQueuePos)  
   
  return
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*cstring colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 

  ! note the size(colPtr) for cstrings, don't use len() here
  self.addColumn(SQL_C_CHAR, address(colPtr), size(colPtr), allowNulls, actualQueuePos)  
   
  return
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddLargeColumn procedure(SQLSMALLINT  colType, long colId, bool allowNulls = true)

  code 
  
  if (coltype <> SQL_LONGVARCHAR) and (coltype <> SQL_LONGVARBINARY)
    return level:notify
  end  

  if (colId <= records(self.colQ))
    return level:notify
  end 
    
  self.colb.ColId = colId  
  self.colb.ColType = colType
  if (allowNulls = true) 
    self.allowNulls = true
  end 

  add(self.colb)

   return level:benign
! end AddLargeColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*real colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 
  
  self.addColumn(SQL_C_DOUBLE, address(colPtr), size(colPtr), allowNulls, actualQueuePos)
     
  return
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*sreal colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 
  
  self.addColumn(SQL_C_FLOAT, address(colPtr), size(colPtr), allowNulls, actualQueuePos)
     
  return
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*decimal colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 
  
  self.addColumn(SQL_C_DECIMAL, address(colPtr), size(colPtr), allowNulls, actualQueuePos)
     
  return
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddBooleanColumn procedure(*bool colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 
  
  self.addColumn(SQL_C_BIT, address(colPtr), size(colPtr), allowNulls, actualQueuePos)

  return
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*TIMESTAMP_STRUCT colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 
  
  self.addColumn(SQL_C_TYPE_TIMESTAMP, address(colPtr), size(TIMESTAMP_STRUCT), allowNulls, actualQueuePos)
   
  return
! end AddColumn
! ------------------------------------------------------------------------------
  
columnsClass.AddColumn procedure(*Date colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 
  
  self.addColumn(SQL_C_DATE, address(colPtr), sizeOfDate, allowNulls, actualQueuePos)
   
  return
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*time colPtr, bool allowNulls = false, long actualQueuePos = -1)

  code 
  
  self.addColumn(SQL_C_TIME, address(colPtr), size(colPtr), allowNulls, actualQueuePos)
   
  return
! end AddColumn
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! add a column to the queue.  each column added here will be bound to the 
! statment handle.  The colums are bound when the execute function is called.
! this is typically called by the various addColumn/1 functions but can be called directly
! in a derived instance.
! the columns are boud in the sequence they are added.
! ------------------------------------------------------------------------------  
columnsClass.AddColumn procedure(SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength, bool allowNulls = false, long actualQueuePos = -1) 

  code 
  
  ! order is the order the columns are added
  self.colQ.ColId = records(self.colq) + 1
  self.colq.ColType = targetType
  self.colQ.ColBuffer = TargetValuePtr
  self.colQ.ColSize = BufferLength

  if (allowNulls = true) 
    self.allowNulls = true
  end 

  self.colq.allowNulls = allowNulls
  if (actualQueuePos = -1) 
    self.colq.actualQuePos = self.colq.colId
  else 
    self.colq.actualQuePos = actualQueuePos
  end

  add(self.Colq)
   
  return
! end AddColumn
! -----------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! sets the default value for the queue elment if the read returns a null from the 
! database.  Most columns will not be null but null's must be handled some way
! ------------------------------------------------------------------------------
columnsCLass.setDefaultNullValue procedure(*queue q) 

x  long,auto
v  any

  code

  loop x = 1 to records(self.colQ)
    get(self.colQ, x) 
    if (self.colq.allowNulls = true) 
      self.AssignDefaultValue(q)
    end ! if allow nulls
  end  ! loop

  return
! setDefaultNullValue ----------------------------------------------------------

! -----------------------------------------------------------------------------
! assigns the default value for a type to the buffer field when the back end 
! is a null.  
! -----------------------------------------------------------------------------
columnsClass.AssignDefaultValue procedure(*queue q) !,virtual,protected

v    any

  code 

  ! if the array at the column index indicates null data
  ! assign to an any and set the default for the type
  if (self.ArrayPtr[self.colQ.colId] = SQL_NULL_DATA) 
    v &= what(q, self.colQ.actualQuePos)
    case self.colq.colType 
      of SQL_C_FLOAT
      orof SQL_C_DOUBLE
        v = defaultFloating
      of SQL_C_CHAR 
        v = defaultString
      of SQL_C_TINYINT
      orof SQL_C_SHORT
      orof SQL_C_SLONG
        v = defaultInteger
      of SQL_C_BIT
        v = defaultBoolean
      end   ! case 
    end ! if arrayptr


  return
! end AssignDefaultValue -------------------------------------------------------

! ------------------------------------------------------------------------------
! returns the value of the instances allowNulls member. 
! this is called when the result set is being read so the additional 
! work for a null column can be handled
! ------------------------------------------------------------------------------
columnsClass.getAllowNulls procedure() !,bool

  code 

  return self.allowNulls
! end getAllowNulls -------------------------------------------------------------  