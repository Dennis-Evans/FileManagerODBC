
!ABCIncludeFile

  OMIT('_EndOfInclude_',_DbcHandlePresent_)
_DbcHandlePresent_   equate(1)

  include('odbcTypes.inc'),once
  
! ------------------------------------------------------------------
! object to manage the Database connection handle. 
! ------------------------------------------------------------------                       
DbcHandle    class,type,module('EnvHandle.clw'),link('EnvHandle.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! --------------------------------------------------------------
! a queue of hDbc handles.  typicall there will be only one 
! database connection handle.  however there are times when 
! multiple handles will be needed so store the handles in the 
! queue.   
! --------------------------------------------------------------
hDbcQue        &HandleQueue,private

! --------------------------------------------------
! allocates the env handle using the module level variable
! and sets the class properties to use that handle value and
! the module level counter
! --------------------------------------------------
construct       procedure()

! --------------------------------------------------
! frees the env handle when the reference count hits zero
! if the count is greater than than zero does not free
! --------------------------------------------------
destruct       procedure()

! --------------------------------------------------
! allocates the hDbc handle
! --------------------------------------------------
allocateHandle procedure(SQLHENV hEnv),sqlReturn,protected

! --------------------------------------------------
! set the version for the driver. 
! the ODBC version will be set to 3.8.  no reason 
! to use older versions.
! --------------------------------------------------
SetVersion procedure(),private,sqlReturn

! --------------------------------------------------
! enable connection pooling.   the default is to use 
! pooling.  
! --------------------------------------------------
EnablePooling procedure(),private,sqlReturn

! --------------------------------------------------
! returns the env handle 
! --------------------------------------------------
getHandle      procedure(),SQLHENV

! --------------------------------------------------
! returns the reference count for the objects
! --------------------------------------------------
getRefCount procedure(),long

             end 

_EndOfInclude_