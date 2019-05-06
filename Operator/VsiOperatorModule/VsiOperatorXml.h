/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorXml.h
**
**	Description:
**		Defines Configuration's XML text
**
********************************************************************************/

#pragma once

// Version
#define VSI_OPERATOR_VERSION				L"version"
	#define VSI_OPERATOR_VERSION_0_0			L""		// Attribute not set - Password not encrypted
	#define VSI_OPERATOR_VERSION_2_0			L"2.0"	// 2012-04-10: Password encrypted
	#define VSI_OPERATOR_VERSION_LATEST			VSI_OPERATOR_VERSION_2_0

// Operator
#define VSI_OPERATOR_ELM_OPERATORS			L"operators"
#define VSI_OPERATOR_ELM_OPERATOR			L"operator"

// Operator attributes
#define VSI_OPERATOR_ATT_UNIQUEID			L"id"
#define VSI_OPERATOR_ATT_NAME				L"name"
#define VSI_OPERATOR_ATT_TIME				L"time"
#define VSI_OPERATOR_ATT_PASSWORD			L"password"
#define VSI_OPERATOR_ATT_TYPE				L"rights"
#define VSI_OPERATOR_ATT_STATE				L"state"
#define VSI_OPERATOR_ATT_STUDY_ACCESS		L"studyAccess"
#define VSI_OPERATOR_ATT_GROUP				L"group"

// Service attribute
#define VSI_OPERATOR_ATT_SERVICE_RECORD		L"sr"




// Backup XML defines
#define VSI_OPERATOR_BACKUP_ELM_BACKUP					L"backup"
	#define VSI_OPERATOR_BACKUP_ATT_VERSION					L"version"
			#define VSI_OPERATOR_BACKUP_VALUE_VERSION			L"1.0"
	#define VSI_OPERATOR_BACKUP_ATT_TYPE					L"type"
		#define VSI_OPERATOR_BACKUP_VALUE_TYPE_OPERATOR			L"operator"
	#define VSI_OPERATOR_BACKUP_ATT_SOFTWARE_VERSION		L"versionSoftware"
	#define VSI_OPERATOR_BACKUP_ATT_DATE					L"date"
	#define VSI_OPERATOR_BACKUP_ATT_BY						L"by"

