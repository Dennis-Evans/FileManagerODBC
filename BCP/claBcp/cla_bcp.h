#pragma once

#include <sql.h>
#include <sqlext.h>

#include <sqltypes.h>

#include <atlstr.h>

//#define _SQLNCLI_ODBC_

// example path, adjust as needed for local system
#include "C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\170\SDK\Include\msodbcsql.h"			

#define DllExport   __declspec( dllexport )


extern "C" {
  DllExport HENV ClaBcpInit();

  DllExport int ClaBcpKill(HENV hEnv, HDBC hDbc);

  DllExport HDBC ClaBcpConnect(HENV hEnv, char *connStr);

  // there are other functions in the API but this is all we need for inserts
  DllExport bool init_Bcp(HDBC hDbc, char *tName);
  DllExport bool sendRow_Bcp(HDBC hDbc);
  DllExport int batch_Bcp(HDBC hDbc);
  DllExport int done_Bcp(HDBC hDbc);
  
  // calls bcp_bind for the data type
  DllExport bool bind_BcpBy(HDBC hDbc, byte *colv, long colOrd);
  DllExport bool bind_BcpSh(HDBC hDbc, short *colv, long colOrd);
  DllExport bool bind_Bcpl(HDBC hDbc, long *colv, long colOrd);
  DllExport bool bind_Bcpb(HDBC hDbc, bool *colv, long colOrd);
  DllExport bool bind_Bcpd(HDBC hDbc, DATE_STRUCT *colv, long colOrd);
  DllExport bool bind_Bcpf(HDBC hDbc, double *colv, long colOrd);
  DllExport bool bind_Bcps(HDBC hDbc, char colv[], long colOrd, long slen);
  DllExport bool bind_Bcpcs(HDBC hDbc, char colv[], long colOrd);
  DllExport bool bind_Bcpsf(HDBC hDbc, float *colv, long colOrd);
  DllExport bool bind_BcpDt(HDBC hDbc, char colv[], long colOrd);
  DllExport bool bind_bcpT(HDBC hDbc, char colv[], long colOrd);
  char *tableName;

  // holds the environment handle
  HENV    hEnv;

  // holds the connection handle from the calling instance database
  // set once and used until done
  HDBC		hDbc;
}