   member('fmOdbcDemo')

   map
   end

! --------------------------------------------------
! fills the queue from a query using the page loading 
! or reading ability added to sql server in 2012
! 
! note Oracle, PostGre have like constructs, 
! the syntax varies some.  other verndors probably
! have something, check the docs for what the back end 
! is in use. 
!
! prior to 2012 there was the window functions that could 
! be used to acccomplish, these were added in 2005
!
! prior to 2005 there were other ways to accemplish page 
! loading and not use cursors.  the code was not trivial 
! but it was doable.
! --------------------------------------------------
pageLoad procedure(fileMgrODBC fmOdbc, long currentRow)

pageSize   long,auto
dynStr     &IDynStr
retv       byte,auto
openedHere byte,auto

  code

  fm.columns.AddColumn(demoQueue.sysId)
  fm.columns.AddColumn(demoQueue.Label)
  fm.columns.AddColumn(demoQueue.amount)

  fm.parameters.AddInParameter(currentRow)
  pageSize = pageLoadSize
  fm.parameters.AddInParameter(pageSize)

  dynStr &= newDynStr()
  ! note the additions to the order by clause of the offset and fetch next
  dynStr.cat('select ld.sysId, ld.Label, ld.amount ' & |
            'from dbo.labeldemo ld ' & |
            'order by ld.SysId ' & |
            'offset ? rows ' & |
            'fetch next ? rows only;')

  retv = fm.ExecuteQuery(dynStr, demoQueue)

  return 
! end fillSp ---------------------------------------------------