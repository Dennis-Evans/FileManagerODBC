   member('fmOdbcDemo')

   map
     module('os')
       GetHeapProcess(),long,pascal
       ReleaseFree(HANDLE hHeap, long  dwFlags, long lpMem),bool,pascal,name('HeapFree')
       HeapAlloc(HANDLE hHeap, long dwFlags, long dwBytes),long,pascal
     end
     module('clib')
       memcpy(long lpDest, long lpSource, long nCount),long,proc,name('_memcpy')
     end
   end

fmOdbcTVPUpdate procedure(fileMgrODBC fmOdbc)

hHeap       long,auto
tablePtr    long,auto
numberBytes long,auto
retv        long,auto

tableOffset long,auto
tableSrc    long,auto
tableDest   long(0)

  code

  ! handle theheap
  hHeap = GetHeapProcess()

  tablePtr = HeapAlloc(hHeap, 0, 20000)

  tableOffset = size(demoQUeue)
  tableSrc = address(demoQueue.SysId)

  loop 3 times
    tableDest = tablePtr + tableOffset
    tableSrc += tableOffset
    memcpy(tableDest, tableSrc, tableOffset)
 
    tableOffset += tableOffset
  end

  retv = ReleaseFree(hHeap, 0, tablePtr)
  
  return