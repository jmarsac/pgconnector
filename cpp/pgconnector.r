/*--------------------------------------------------------------------------------------+
|
|     $Source$
|
|  $Copyright: (c) 2016 Azimut. All rights reserved. $
|
+--------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------+
|									|
|    $Logfile$
|   $Workfile$
|   $Revision$
|   	$Date$
|									|
+----------------------------------------------------------------------*/
/*----------------------------------------------------------------------+
|									|
|   Function -								|
|									|
|	PGCONNECTOR Example dialog box resources				|
|									|
+----------------------------------------------------------------------*/
#include <dlogbox.h>
#include <dlogids.h>

#include "pgconnectorid.h"
#include "pgconnectortext.h"

/*----------------------------------------------------------------------+
|									|
|   Parcel Locate Dialog Box						|
|									|
+----------------------------------------------------------------------*/
#undef	    XC
#define	    XC		(DCOORD_RESOLUTION/2) * ASPECT_LOCATEPARCEL

#define DW	(35*XC)
#define DH	(5.5*YC)

#define X1	(15*XC)		/* text field */
#define X2	(10*XC)		/* Locate button */

#define Y1	(YC)		/* text field */
#define Y2	(2.75*YC)	/* Locate button */

DialogBoxRsc DIALOGID_Locate =
    {
    DIALOGATTR_DEFAULT,
    DW, DH,
    NOHELP, MHELP, NOHOOK, NOPARENTID,
    TXT_ParcelLocation,
{
{{ X1, Y1, 16*XC, 0},
	Text, TEXTID_ParcelID,  ON, 0, TXT_ParcelNumber, ""},
{{ X2, Y2, (16*XC), 0},
	PushButton, PUSHBUTTONID_Locate, ON, 0, "", ""},
}
};

/*----------------------------------------------------------------------+
|									|
|    PushButton Resources						|
|									|
+----------------------------------------------------------------------*/
DItem_PushButtonRsc PUSHBUTTONID_Locate =
    {
    NOT_DEFAULT_BUTTON, NOHELP, MHELP,
    HOOKITEMID_LocateButton, NOARG, NOCMD, MCMD, "",
    TXT_LocateParcel
    }

/*----------------------------------------------------------------------+
|									|
|    Text Resources							|
|									|
+----------------------------------------------------------------------*/
DItem_TextRsc TEXTID_ParcelID =
    {
    NOCMD, MCMD, NOSYNONYM, NOHELP, MHELP,
    NOHOOK, NOARG,
    16, "%s", "%s", "", "", NOMASK, CONCAT,
    "",
    ""
    };
