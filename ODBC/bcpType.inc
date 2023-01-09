

!ABCIncludeFile

  OMIT('_EndOfInclude_',_BCPPresent_)
_BCPPresent_   EQUATE(1)

  include('dynStr.inc'),once
  include('odbcTypes.inc'),once

! ----------------------------------------------------------------------------------
! ----------------------------------------------------------------------------------
bcpType class,type,module('bcpType.clw'),link('bcpType.clw',_ABCLinkMode_),dll(_ABCDllMode_)
! set up and tear down

! handles for the bcp 
hEnv        SQLHENV
hDbc        SQLHDBC

! -----------------------------------------------------------------
! number of rows to be inserted for a batch
! if zero or less then the rows are not batched
! -----------------------------------------------------------------
batchSize   long,private

! -----------------------------------------------------------------
! number of rows that have been sent to the server for a batch
! -----------------------------------------------------------------
rowsSent    long,private

! ----------------------------------------------------------------------
! initilizes the object and creates a seperate hEnv for use by 
! the BCP operation.
! ----------------------------------------------------------------------
init_Bcp             procedure(),byte

! ----------------------------------------------------------------------
! gets a connection to the databse and set the connection attributes 
! for the BCP
! ----------------------------------------------------------------------
connect          procedure(*cstring connStr),byte

! ----------------------------------------------------------------------
! frees the hEnv and hDbc used by the BCP 
! ----------------------------------------------------------------------
disconnect       procedure()

! ----------------------------------------------------------------------
! sets up the table for the insert 
! called once for each table in the process
! ----------------------------------------------------------------------
init_Bcp         procedure(*cstring tName),bool

! ----------------------------------------------------------------------
! sets the number of rows to be used for a the batch during the bcp operations
! the default is zero, no batch size, set as needed after testing for 
! best performance
! ----------------------------------------------------------------------
setBcpBatchSize procedure(long rows)

! ----------------------------------------------------------------------
! sends a row to the server and checks the batch size. 
! if the number of rows sent is greater than or equal to the batch size 
! the data is written
! ----------------------------------------------------------------------
sendRow          procedure(),bool

! ----------------------------------------------------------------------
! writes a batch of rows to the server 
! returns the number of rows written
! ----------------------------------------------------------------------
batch_Bcp        procedure(),long

! ----------------------------------------------------------------------
! writes the data to the server, if there is any, and shuts down the BCP 
! operations on the server.   
! returns the number of rows written
! this function MUST be called when the insert is completed
! ----------------------------------------------------------------------
done_Bcp         procedure(),long

! ---------------------------------------------------------------------
! add a column to the bcp of the insert
! each of these functions take field (from a queue, tps file, what ever) 
! and bind the field to the bcp layer.  the colOrd is the ordinal position of the 
! of the field in the table.  starting at 1.
! 
! all of the functions do the same thing, 
! note the comments on addColumnDateTime, addColumnTime and addColumnBool
! --------------------------------------------------------------------
addColumn        procedure(*byte colv, long colOrd),bool
addColumn        procedure(*short colv, long colOrd),bool
addColumn        procedure(*long colv, long colOrd),bool

!  -----------------------------------------------------------
! function to add a boolean to the insert
! a bool is equated to a long so the function label is specific
!  -----------------------------------------------------------
addColumnBool    procedure(*bool colv, long colOrd),bool

addColumn        procedure(*DATE_STRUCT colv, long colOrd),bool
addColumn        procedure(*real colv, long colOrd),bool
addColumn        procedure(*string colv, long colOrd),bool
addColumn        procedure(*cstring colv, long colOrd),bool
addColumn        procedure(*sreal colv, long colOrd),bool

!  -----------------------------------------------------------
! function to add a datetime value, in a string format to the insert
! the string must be formatted as a  valid datetime string
!  -----------------------------------------------------------
addColumnDate    procedure(*dateTimeString colv, long colOrd),bool

!  -----------------------------------------------------------
! function to add a time value, in a string format to the insert
! the string must be formatted as a  valid time string
!  -----------------------------------------------------------
addColumnTime    procedure(*timeString colv, long colOrd),bool

            end

  _EndOfInclude_
