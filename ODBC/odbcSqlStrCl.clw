
  member()
  
  include('odbcSqlStrCl.inc'),once 

  map 
    module('c6')
      memmove(long dest, long src, long num),long,proc,name('_memmove')
      strLen(long cstr),long,name('_strlen')
    end 
  end

! ---------------------------------------------------------------------------
  
sqlStrClType.init procedure()

  code
  
  self.sqlStr &= newdynstr()
  if (self.sqlStr &= null) 
    return level:notify
  end 
    
  return level:benign

sqlStrClType.init procedure(string sqlcode)

  code 
  
  if (self.sqlStr &= null) 
    self.init()
  end   
  
  self.sqlStr.kill()
  self.sqlStr.cat(clip(sqlcode))
  self.paramCount = 0
  self.paramMarker = 0
  
  return level:benign
    
sqlStrClType.init procedure(*IdynStr sqlcode)

  code 

  self.init(sqlcode.str())
  
  return level:benign
  
sqlStrClType.kill procedure()

  code 

  if (~self.sqlStr &= null)  
    self.sqlStr.kill()
    disposedynstr(self.sqlStr)
    self.sqlStr &= null
  end  
  
  return 
    
sqlStrClType.destruct procedure()  

  code 
  
  self.kill()
      
  return 
  
sqlStrClType.str procedure() !,string

  code
  return self.sqlStr.str()
  
sqlStrClType.cstr procedure() !,*cstring

  code
  return self.sqlStr.cstr()
  
sqlStrClType.strlen         procedure() !,long
  
  code
  return self.sqlStr.strLen()
  
sqlStrClType.cat procedure(string sqlcode)

  code
  
  self.sqlStr.cat(sqlCode)
  
  return

sqlStrClType.replaceStr procedure(*cstring workstr)  

  code 
  
  self.sqlstr.kill()
  self.sqlStr.cat(workstr)
  
  return 
  
sqlStrClType.replaceName procedure(*ParametersClass params, *cstring workStr) !,long  

retv   long

  code 
  
  self.paramMarker = instring(eParamIdChar, workStr, 1, self.paramMarker + 1)
  if (self.paramMarker > 0)
    self.paramCount += 1    
    if (self.paramCount > 0) 
      memmove(address(workstr) + self.paramMarker, address(workstr) + self.endParam - 1, strlen(address(workstr[self.endParam])) + 1)
      workstr[self.paramMarker] = eOdbcParamIdChar
    end  
  else 
    self.paramCount = 0  
  end 
  
  return self.paramCount
 
sqlStrClType.replaceFieldList procedure(*columnsClass cols)

workstr   &IDynStr 

  code 
      
  self.paramMarker = 0
    
  self.paramMarker = instring(eFieldListLabel, self.sqlStr.str(), 1, self.paramMarker + 1)
  if (self.paramMarker > 0)
    workStr &= newDynStr() 
    
    !workstr.cat(eSelectLabel & cols.fieldList.str())
    workstr.cat(eSpaceChar & sub(self.sqlStr.str(), self.paramMarker + size(eFieldListLabel), self.sqlStr.strlen() - 17))
    self.replaceStr(workstr.cstr())
    
    disposeDynstr(workStr)
  end 
           
  return 
! replaceFieldList
! ------------------------------------------------------------------------------------------------  

sqlStrClType.findEnd procedure(*cstring workStr)

retv   sqlreturn(sql_Success)

  code

  self.endParam = instring(eSpaceChar, workStr, 1, self.paramMarker + 1)  
  if (self.endParam = 0) 
    self.endParam = instring(eCloseParaen,  workStr, 1, self.paramMarker + 1)
    if (self.endParam = 0) 
      retv = sql_error
    end
  end

  loop while (workStr[self.endParam - 1] = eCloseParaen)
    self.endParam -= 1
  end ! loop
  
  if (self.endParam <= 1)
    retv = sql_error
  end    
  
  return retv
  
sqlStrClType.formatScalarCall procedure(string spName)
count   long
recCount long

  code 
  
  self.sqlStr.Kill()
  self.sqlStr.cat(eScalarCallLabel & spName & eOpenParen & ')}')

  return

sqlStrClType.formatScalarCall procedure(string spName, *ParametersClass params)
count   long
recCount long

  code 
  
  self.sqlStr.Kill()
  self.sqlStr.cat(eScalarCallLabel & spName & eOpenParen)
  if (params.FillPlaceHolders(self, 2) = 0) 
    self.sqlStr.cat(')}')
  end       

  return

sqlStrClType.formatSpCall procedure(string spName, *ParametersClass params) 
  
count   long
recCount long

  code 
  
  self.sqlStr.Kill()
  self.sqlStr.cat(eCallLabel & spName & eOpenParen)
  if (params.FillPlaceHolders(self) = 0) 
    self.sqlStr.cat(')}')
  end       

  return
  
sqlStrClType.formatSpCall procedure(string spName) 
  
count    long
recCount long

  code 
  
  self.sqlStr.Kill()
  self.sqlStr.cat(eCallLabel & spName & '}') ! & eOpenParen & ')}')

  return  
  
sqlStrClType.addOrderBy procedure(string fld)

  code 
  
  self.sqlStr.cat(' order by ' & clip(fld))
  
  return 
  
sqlStrClType.addOrderBy procedure(string fld, string dir)
  
  code 
  
  self.sqlStr.cat('order by ' & clip(fld) & ' ' & clip(dir))
  
  return 
  
sqlStrClType.addWhere procedure(string fld, string cond, string pName)  

  code 
  
  self.sqlStr.cat(' where ' & fld & ' ' & cond & ' ' & pName)
  
  return