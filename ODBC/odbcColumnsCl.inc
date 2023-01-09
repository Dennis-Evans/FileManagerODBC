
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCColumssPresent_)
_ODBCColumssPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcHandleCl.inc'),once
  include('dynstr.inc'),once

! -------------------------------------------------------------------------  
! this is an arbitray value that is used to create an array of long values.
! the array is used when binding the columns and the read will write back information 
! to the array elements.  if the column is null the element will be set to 
! SQL_NULL_DATA.
! the value of 100 is just a number, most queries will contain less than 100 columns
! if a larger number is needed adjust, but don't get stupid.  this waste space
! with 100 most of the time   clarion does not support dynamic arrays
! So we are going t o cheat and use a static array and waste a small amount of space.  
! -------------------------------------------------------------------------
numberCols  equate(100)

ColumnsLarge queue,type
binaryHolder   &string       ! if this is a varbinary column this will hold the data that was read 
charHolder     &cstring       ! if this is a varchar column this will hold the data that was read
ColSize        long          ! size of the data that was read
ColId          long
ColType        SQLSMALLINT    ! type of the col, odbc const this will always be one of the two 
                              ! SQL_LONGVARBINARY or SQL_LONGVARCHAR
             end  

! -----------------------------------------------------------  
! queue to store the columns used in a query
! -----------------------------------------------------------  
ColumnsQueue queue,type
ColId          long           ! ord of the column in the select statement
ColType        SQLSMALLINT    ! type of the col, odbc const
ColBuffer      SQLPOINTER     ! pointer to the buffer for the result
ColSize        SQLLEN         ! size of the column in bytes
colLenStrPtr   long
allowNulls     bool           ! true indicates this column allows nulls in the back end
ActualQuePos   long           ! 
             end 

! --------------------------------------------------------------------- 
! class to manage the colums used in a query or a stored procedure call
! each column used is added to the queue and will be mapped or bound 
! to the destination.  the destenation can be any buffer needed/wanted,
! a queue, a group, an individual field, ...
! --------------------------------------------------------------------  
columnsClass class,type,module('odbcColumnsCl.clw'),link('odbcColumnsCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! ---------------------------------------------------------------------------
! queue of the columns that will be used by the database call
! ---------------------------------------------------------------------------
colQ           &ColumnsQueue

colB           &ColumnsLarge
! ---------------------------------------------------------------------------
! array of long for use by the bind column call
! this is used for the nullvalues on the back end
! and if the column is null it will be set to SQL_NULL_DATA
! ---------------------------------------------------------------------------
ArrayPtr       long,dim(numberCols)

! ---------------------------------------------------------------------------
! indicaates that at least one column in the result set allows null values
AllowNulls     bool,private
! ---------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! flag indicates this set has a large column, character or binary
! the type will be found in the column type field
! ---------------------------------------------------------------------------
hasLargeColumn  bool,private

! ---------------------------------------------------------------------------
!  default constructor, calls the init function for the set up
! ---------------------------------------------------------------------------
construct procedure()

! ---------------------------------------------------------------------------
!  allocates the queue 
! ---------------------------------------------------------------------------  
init procedure(),byte,proc

! ------------------------------------------------------------------------------
! disposes the queue and the dyn str
! ------------------------------------------------------------------------------
kill procedure()

! ---------------------------------------------------------------------------
!  default destructor, calls the kill function for the clean up
! ---------------------------------------------------------------------------
destruct procedure(),virtual

! ------------------------------------------------------------------------------
! free the queue and the dyn str that contains the fileds names
! ------------------------------------------------------------------------------
clearQ procedure()

! ------------------------------------------------------------------------------
! Bind the queue, group or seperate fields to the columns in the result set.
! column order is typically the  same order as the select statment,
! 
! parameters for the ODBC api call 
! hStmt   = handle to the ODBC statement
! colId = ord value of the parmaeter, 1, 2, 3 ... the ordinal position
! colType = the C data type of the column 
! colValue = pointer to the buffer field 
! colSize = the size of the buffer or the queue field 
! colInd = pointer to a buffer for the size of the parameter. not used and null in this example 
! -----------------------------------------------------------------------------    
bindColumns procedure(SQLHSTMT  hStmt),sqlReturn,virtual

! ------------------------------------------------------------------------------
! the various addColumn functions are called by the using code and are used for the
! specific data types.  each calls the AddColumn/3 function to actually 
! add a columns
! ------------------------------------------------------------------------------
!

! --------------------------------------------------
! AddColumn, each of these functions add a column that
! will be filled from a result set.  
! inputs :
! colPtr, a pointer (an address) to the buffer field,
!         the filed can be from a queue, a group, a class 
!         or an individual variable.
! allowNulls, default is false which indicates the column in the 
!             database does not allow null values. 
!             true indicates the column does allow nulls
! actualQueuePos, an integer that indicates the location of the field 
!                 in sequence when using a queue.  the value is 
!                 used when a queue is the buffer that will displayed in 
!                 a list and the displayed position does not match 
!                 the ordinal of the queue  
! --------------------------------------------------

! --------------------------------------------------
! the first three are for the typical numeric types used
! the unsigned type can be added if needed/wanted
! --------------------------------------------------
AddColumn procedure(*byte colPtr, bool allowNulls = false, long actualQueuePos = -1)
AddColumn procedure(*short colPtr, bool allowNulls = false, long actualQueuePos = -1)
AddColumn procedure(*long colPtr, bool allowNulls = false, long actualQueuePos = -1)
!
! --------------------------------------------------
! the string types 
! these will also be used for the guid or unique identifier 
! data type
! --------------------------------------------------
AddColumn procedure(*string colPtr, bool allowNulls = false, long actualQueuePos = -1)
AddColumn procedure(*cstring colPtr, bool allowNulls = false, long actualQueuePos = -1)
!
AddLargeColumn procedure(SQLSMALLINT colType, long colId, bool allowNulls = true),byte

! --------------------------------------------------
! the floating point  types 
! --------------------------------------------------
AddColumn procedure(*real colPtr, bool allowNulls = false, long actualQueuePos = -1)
AddColumn procedure(*sreal colPtr, bool allowNulls = false, long actualQueuePos = -1)
!
! --------------------------------------------------
! decimal type 
! --------------------------------------------------
AddColumn procedure(*decimal colPtr, bool allowNulls = false, long actualQueuePos = -1)
!
! --------------------------------------------------
! boolean type, this is actually long in Clarion 
! and using bool as the parameter type causes compiler issues.
! --------------------------------------------------
AddBooleanColumn procedure(*bool colPtr, bool allowNulls = false, long actualQueuePos = -1)
!
! --------------------------------------------------
! date, time and datetime 
! --------------------------------------------------
AddColumn procedure(*TIMESTAMP_STRUCT colPtr, bool allowNulls = false, long actualQueuePos = -1)
AddColumn procedure(*Date colPtr, bool allowNulls = false, long actualQueuePos = -1)
AddColumn procedure(*time colPtr, bool allowNulls = false, long actualQueuePos = -1)

! ------------------------------------------------------------------------------
! add a column to the queue.  each column added here will be bound to the 
! statment handle.  The colums are bound when the execute function is called.
! this is typically called by the various addColumn/3 functions but can be called directly
! in a derived instance.
!
! the columns are boud in the sequence they are added.
! ------------------------------------------------------------------------------  
AddColumn procedure(SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength, bool allowNulls = false, long actualQueuePos = -1),protected

! ------------------------------------------------------------------------------
! sets the default value for the queue elment if the read returns a null from the 
! database.  Most columns will not be null but null's must be handled in a resonale manner
! ------------------------------------------------------------------------------
setDefaultNullValue procedure(*queue q) 

! -----------------------------------------------------------------------------
! assigns the default value for a type to the buffer field when the back end 
! is a null.  
! -----------------------------------------------------------------------------
AssignDefaultValue procedure(*queue q),virtual,protected

! ------------------------------------------------------------------------------
! returns the value of the instances allowNulls member. 
! this is called when the result set is being read so the additional 
! work for a null column can be handled
! ------------------------------------------------------------------------------
getAllowNulls procedure(),bool

            end 

  _EndOfInclude_
