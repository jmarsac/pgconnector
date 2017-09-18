/*--------------------------------------------------------------------------------------+
|
|     $Source$
|
|  $Copyright: (c) 2016 Azimut. All rights reserved. $
|
+--------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------+
|									|
|     $Source$
|   $Workfile$
|   $Revision$
|   	$Date$
|									|
+----------------------------------------------------------------------*/
/*----------------------------------------------------------------------+
|									|
|   Function -								|
|									|
|	PGCONNECTOR Example application message resources			|
|									|
+----------------------------------------------------------------------*/
#include    <dlogbox.h>
#include    <dlogids.h>

#include    "pgconnectorid.h"

MessageList MESSAGELISTID_PGCONNECTORMessages =
    {
      {
      { 1,  "PGCONNECTOR commands loaded" },
      { 2,  "Enter a parcel (CLT) number" },
      { 3,  "ATTACH PARCEL Exited" },
      { 4,  "Locating parcel in database.." },
      { 5,  "Identify parcel centroid on map" },
      { 6,  "Parcel attributes attached" },
      { 7,  "Retrieving MAPID from MAPS table" },
      { 8,  "Updating MAPID in parcel database" },
      { 9,  "Parcel row updated with MAPID" },
      { 10, "Retrieving MAPID from PARCEL table" },
      { 11, "Retrieving PARCEL data from MSCATALOG" },
      { 12, "Retrieving MSLINK from PARCEL table" },
      { 13, "Retrieving MAPNAME from MAPS table" },
      { 14, "Scanning design file for parcel" },
      { 15, "Parcel not found" },
      { 16, "Parcel (%s) not located" },
      { 17, "Loading %s" },
      { 18, "Enter a real estate parcel number" },
      { 19, "PARCEL LOCATE Exited" },
      { 20, "Datapoint to place parcel pushpins" },
      { 21, "PUSHPIN PARCEL Exited" },
      { 22, "Can not find color file: %s" },
      { 23, "Can not open color file: %s" },
      { 24, "Loading displayable attributes..." },
      { 25, "Displayable attributes loaded" },
      }
    };
