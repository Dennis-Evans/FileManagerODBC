
!ABCIncludeFile

  OMIT('_EndOfInclude_',_EnvHandlePresent_)
_EnvHandlePresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  
! ------------------------------------------------------------------
! object to manage the env handle.  each file manager will have an instance 
! of this object.  there will be many instances of this but there will only 
! be one env. handle allocated and all file manager instances will use 
! the same environment handle
!
! typically there will be one environment handle.  while more could be created there
! is normally no reason for multiple environments.  
! 
! note, this object does not display any errors.  if the handle cannot be allocated,
! does not matter why the app will be shut down.  if the ODBC version cannot
! be set, again don't case why, the application will be shut down.
! ------------------------------------------------------------------                       
EnvHandle    class,type,module('EnvHandle.clw'),link('EnvHandle.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! ------------------------------------------
! the env handle for the file manager instances
! there will only be one environment for all the 
! file manager instances
! -----------------------------------------
handle            &SQLHENV,private

! ------------------------------------------
! reference counter to the enviornemnt handle for 
! the file manager instances.  this will be increased in the 
! constructor and decreased in the destructor.
! -----------------------------------------
refCount          &long

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
! allocates the env handle and sets the 
! ODBC version.  version will be set to 3.8
! --------------------------------------------------
allocateHandle procedure(),sqlReturn,private

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