
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCSqlStrClassPresent_)
_ODBCSqlStrPresent_   EQUATE(1)

  include('dynStr.inc'),once
  include('odbcConn.inc'),once
  include('odbcTypes.inc'),once
  include('odbcParamsCl.inc'),once 
  include('odbcColumnsCl.inc'),once 

               
sqlStrClType class,type,module('odbcSqlStrCl.clw'),link('odbcSqlStrCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
sqlStr         &IdynStr
orderByStr     &IDynStr
whereClause    &IDynStr
paramMarker    long
endParam       long
paramCount     long
init           procedure(),byte,proc
init           procedure(string sqlcode),byte,proc
init           procedure(*IdynStr sqlcode),byte,proc
kill           procedure()
destruct       procedure()
str            procedure(),string
cstr           procedure(),*cstring
cat            procedure(string sqlcode) 
strlen         procedure(),long
addWhere       procedure(string fld, string cond, string pName)
addOrderBy     procedure(string fld)
addOrderBy     procedure(string fld, string dir)
replaceStr     procedure(*cstring workStr)
replaceName    procedure(*ParametersClass params, *cstring workStr),long
replaceFieldList procedure(*columnsClass cols)
formatSpCall   procedure(string spName) 
formatSpCall   procedure(string spName, *ParametersClass params)
formatScalarCall procedure(string spName) 
formatScalarCall procedure(string spName, *ParametersClass params)
findEnd        procedure(*cstring workStr),sqlReturn,private
             end
             
  _EndOfInclude_             
