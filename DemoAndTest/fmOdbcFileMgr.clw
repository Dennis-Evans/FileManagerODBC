   member('fmOdbcDemo')

   map
   end

! --------------------------------------------------
! fill a queue using the typcial file manager access
! --------------------------------------------------
fileFill procedure(fileMgrODBC fmOdbc)

retv   byte,auto

  code
  
  fmOdbc.open()
  fmOdbc.useFile()
  !buffer(fmOdbc.file, 2000)
  set(labelDemo)

  loop
    if (fm.next() = level:Benign)
      demoQueue.sysId = labelDemo.Sysid
      demoQueue.Label = labelDemo.Label
      demoQueue.amount = labelDemo.amount
      add(demoQueue)
    else 

      break;
    end
  end

  fmOdbc.close()

  return
! end fileFill -----------------------------------------------

! --------------------------------------------------
! fill a queue using a prop:sql statement 
! --------------------------------------------------
propSqlFill procedure()

retv   byte,auto

  code

  open(labeldemo)
  buffer(labeldemo, 2000)
  labeldemo{prop:sql} = 'select ld.SysId, ld.Label, ld.amount from dbo.LabelDemo ld'
    
  loop
     next(labeldemo)
     if (errorcode() > 0)
      break
    end

    demoQueue.sysId = labelDemo.Sysid
    demoQueue.Label = labelDemo.Label
    demoQueue.amount = labelDemo.amount
    add(demoQueue)

  end

  close(labeldemo)

  return
! end propSqlFill -----------------------------------------------

! --------------------------------------------------
! fill a queue using a simple view definition
! --------------------------------------------------
ViewFill procedure(fileMgrODBC fmOdbc)

retv   byte,auto

  code
  
  fmOdbc.open()
  fmOdbc.useFile()
    
  open(demoView)
  set(demoView)

  buffer(demoView, 50000)

  loop
    next(demoView)
    if (errorcode() > 0)
      break
    end

    demoQueue.sysId = labelDemo.Sysid
    demoQueue.Label = labelDemo.Label
    demoQueue.amount = labelDemo.amount

    add(demoQueue)
  end

  close(demoView)
  fmOdbc.close()

  return
! end ViewFill -----------------------------------------------