// cla_bcp.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "cla_bcp.h"
#include <atlstr.h>


DllExport HENV ClaBcpInit() {

  HENV hEnv = 0;
  SQLRETURN result; 

  result = SQLAllocHandle(SQL_HANDLE_ENV, NULL, &hEnv);
  //result = SQLAllocEnv(&hEnv);
  if ((result != SQL_SUCCESS) && (result != SQL_SUCCESS_WITH_INFO)) {
    hEnv = 0;
  }
  else {
    result = SQLSetEnvAttr(hEnv, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3_80, SQL_IS_INTEGER);
  }

  return hEnv;
} // end ClaBcpInit
// ------------------------------------------------------------------------

DllExport int ClaBcpKill(HENV hEnv, HDBC hDbc) {

  if (hDbc > 0) {
    SQLFreeHandle(SQL_HANDLE_DBC, hDbc);
  }
  if (hEnv > 0) {
    SQLFreeHandle(SQL_HANDLE_ENV, hEnv);
  }

  return 0;
} // end ClaBcpKill
// -----------------------------------------------------------------------

DllExport HDBC ClaBcpConnect(HENV hEnv, char* connStr) {

  SQLWCHAR		sqlState[16];
  SQLINTEGER	errPtr;
  SQLWCHAR		msgTxt[256];
  SQLSMALLINT txtLen;

  SQLRETURN result;
  HDBC hDbc = 0;
  wchar_t		  outconnStr[1024];
  SQLSMALLINT	connStrLen;
  CString cs = connStr;

  result = SQLAllocConnect(hEnv, &hDbc);
  if (result != SQL_SUCCESS && result != SQL_SUCCESS_WITH_INFO)  {
    return false;
  }

  result = SQLSetConnectAttr(hDbc, SQL_COPT_SS_BCP, (void *)SQL_BCP_ON, SQL_IS_INTEGER);

  if (result != SQL_SUCCESS && result != SQL_SUCCESS_WITH_INFO) {
    SQLFreeConnect(hDbc);
    hDbc = NULL;
    return false;
  }

  result = SQLDriverConnect(hDbc, NULL, (SQLWCHAR *)cs.GetBuffer(), cs.GetLength(), outconnStr, sizeof(outconnStr) - 1, &connStrLen, SQL_DRIVER_NOPROMPT);
  if (result != SQL_SUCCESS && result != SQL_SUCCESS_WITH_INFO) {
    SQLGetDiagRec(SQL_HANDLE_DBC, hDbc, 1, sqlState, &errPtr, msgTxt, 255, &txtLen);
    SQLFreeConnect(hDbc);
    hDbc = NULL;
    return false;
  }

  return hDbc;
} // end ClaBcpConnect
// ------------------------------------------------------------------------

// there are other functions in the API but this is all we need for inserts
DllExport bool init_Bcp(HDBC hDbc, char *tName) {

  CString tableName = tName;
  bool retv = true;

  if (bcp_init(hDbc, tableName.GetBuffer(), NULL, NULL, DB_IN) == FAIL) {
    retv = false;
  }

  return retv;
} // end init_bcp
// -----------------------------------------------------------

DllExport bool sendRow_Bcp(HDBC hDbc) {

  bool retv = true;

  if (bcp_sendrow(hDbc) == FAIL) {
    retv = false;
  }

  return retv;
} // end sendRow_bcp 
// -----------------------------------------------------------------

DllExport int batch_Bcp(HDBC hDbc) {

  int retv;

  retv = bcp_batch(hDbc);

  return retv;
} // end batch_bcp
// ----------------------------------------------------------------

DllExport int done_Bcp(HDBC hDbc) {

  int retv;

  retv = bcp_done(hDbc);

  return retv;
} // end done_bcp
// ------------------------------------------------------------------

// calls bcp_bind for the data type
DllExport bool bind_BcpBy(HDBC hDbc, byte *colv, long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, sizeof(byte), NULL, 0, SQLINT1, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -------------------------------------------------------------

DllExport bool bind_BcpSh(HDBC hDbc, short *colv, long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, sizeof(short), NULL, 0, SQLINT2, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -------------------------------------------------------------

// calls bcp_bind for the data type
DllExport bool bind_Bcpl(HDBC hDbc, long *colv, long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, sizeof(long), NULL, 0, SQLINT4, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -------------------------------------------------------------

DllExport bool bind_Bcpb(HDBC hDbc, bool *colv, long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, sizeof(bool), NULL, 0, SQLBIT, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -------------------------------------------------------------

DllExport bool bind_Bcpd(HDBC hDbc, DATE_STRUCT *colv, long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, sizeof(DATE_STRUCT), NULL, 0, SQLDATEN, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -------------------------------------------------------------

DllExport bool bind_BcpDt(HDBC hDbc, char colv[], long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, SQL_VARLEN_DATA, (LPCBYTE)"", 1, SQLCHARACTER, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -------------------------------------------------------------

DllExport bool bind_Bcpf(HDBC hDbc, double *colv, long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, sizeof(double), NULL, 0, SQLFLT8, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -------------------------------------------------------------
 
DllExport bool bind_Bcps(HDBC hDbc, char colv[], long colOrd, long slen) {

  bool retv = true;
  long w = 1;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, slen, NULL, 0, SQLCHARACTER, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -------------------------------------------------------------

DllExport bool bind_Bcpcs(HDBC hDbc, char colv[], long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, SQL_VARLEN_DATA, (LPCBYTE)"", 1, SQLCHARACTER, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -----------------------------------------------------------

DllExport bool bind_Bcpsf(HDBC hDbc, float *colv, long colOrd)  {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, sizeof(float), NULL, 0, SQLFLT4, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -----------------------------------------------------------

DllExport bool bind_bcpT(HDBC hDbc, char colv[], long colOrd) {

  bool retv = true;

  if (bcp_bind(hDbc, (LPCBYTE)colv, 0, SQL_VARLEN_DATA, (LPCBYTE)"", 1, SQLCHARACTER, colOrd) == FAIL) {
    retv = false;
  }

  return retv;
} // end bind_Bcp
// -----------------------------------------------------------
