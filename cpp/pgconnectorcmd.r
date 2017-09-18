/*----------------------------------------------------------------------+
|                                                                       |
| $Copyright: (c) 2016 Azimut. All rights reserved. $
|                                                                       |
| Limited permission is hereby granted to reproduce and modify this     |
| copyrighted material provided that the resulting code is used only in |
| conjunction with Bentley Systems products under the terms of the      |
| license agreement provided therein, and that this notice is retained  |
| in its entirety in any such reproduction or modification.             |
|                                                                       |
+----------------------------------------------------------------------*/
/*----------------------------------------------------------------------+
|                                                                       |
|   $Logfile$
|   $Workfile$
|   $Revision$
|   $Date$
|                                                                       |
+----------------------------------------------------------------------*/
/*----------------------------------------------------------------------+
|                                                                       |
|   Function -                                                          |
|                                                                       |
|        Main PGCONNECTOR Commands                                              |
|                                                                       |
+----------------------------------------------------------------------*/
#include <dlogids.h>
#include <rscdefs.h>
#include <cmdclass.h>

/*-----------------------------------------------------------------------
 Setup for native code only MDL app
-----------------------------------------------------------------------*/
#define  DLLAPP_PRIMARY     1

DllMdlApp   DLLAPP_PRIMARY =
    {
    "PGCONNECTOR", "pgconnector"
    }

#define CT_NONE                          0
#define CT_MAIN                          1
#define CT_PGCONNECTOR                   2
#define CT_PGC_SYNC                      3
#define CT_PGC_CONNECTION                4

Table CT_MAIN =
{
    {  1, CT_PGCONNECTOR,       DATABASE,       REQ,        "PGC" },
};

/*------------------------------------------------ */
/*      PGCONNECTOR Subtable                       */
/*------------------------------------------------ */
Table CT_PGCONNECTOR =
{
    {  1, CT_NONE,              INHERIT,        DEF | TRY,  "ATTACH" },
    {  2, CT_NONE,              INHERIT,        TRY,        "CHECKOUT" },
    {  3, CT_NONE,              INHERIT,        TRY,        "CHECKIN"},
    {  4, CT_NONE,              INHERIT,        DEF,        "CONNECT"},
    {  5, CT_NONE,              INHERIT,        DEF,        "DISCONNECT"},
    {  6, CT_PGC_CONNECTION,    INHERIT,        REQ,        "CONNECTION"},
    {  7, CT_PGC_SYNC,          INHERIT,        REQ,        "SYNC"}
};

Table CT_PGC_CONNECTION =
{
    {  1, CT_NONE,              INHERIT,        DEF | TRY,  "ADD" },
    {  2, CT_NONE,              INHERIT,        TRY,        "DELETE" }
}
/*------------------------------------------------ */
/*      PGCONNECTOR SYNC Subtable                  */
/*------------------------------------------------ */
Table CT_PGC_SYNC =
{
    {  1, CT_NONE,              INHERIT,        DEF | TRY,  "ON" },
    {  2, CT_NONE,              INHERIT,        TRY,        "OFF" }
}
