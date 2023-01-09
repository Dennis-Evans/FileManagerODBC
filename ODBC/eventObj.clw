     member('OdbcDemo_Ms.clw')

     map
       module('win321')
         CreateEvent(long lpEventAttributes, BOOL bManualReset, BOOL bInitialState, *cstring lpName),long,pascal,raw,name('CreateEventA')
       end
       GetEventHandle(),long
     end

GetEventHandle procedure()

evname cstring('thisEvent')
retH   long,auto

  code

  retH = CreateEvent(0, false, false, evname)

  return retH
! ---------------------------------------------------------------------------