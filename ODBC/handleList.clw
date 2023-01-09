  
  member()

  include('handleList.inc'),once 
  include('odbcTypes.inc'),once
  
  map 
  end

HandleList.Construct procedure()

  code

  self.handles &= new(HandleQueue)

  return 
! end destruct --------------------------------------------

HandleList.Destruct procedure()  !,virtual

  code

  free(self.Handles)
  dispose(self.handles)

  return 
! end destruct --------------------------------------------  

HandleList.FindHandle procedure(string label) !,long

retv   long,auto

   code

   self.handles.Label = label
   get(self.handles, self.handles.Label)
   if (errorcode() = 0) 
     retv = self.handles.handle
   else 
     retv = -1
   end

   return retv
! end FindHandle ----------------------------------------------------

HandleList.Addhandle    procedure(string label, SQLHANDLE h) !,long

retv   long,auto

  code

   self.handles.Label = label
   get(self.handles, self.handles.Label)
   if (errorcode() > 0)      
     retv = self.handles.handle
   else 
     retv = -1
   end
 

  return retv
! end AddHandle ----------------------------------------------------

HandleList.Removehandle procedure(string label)

   code

   return
! end RemoveHandle --------------------------------------------------



