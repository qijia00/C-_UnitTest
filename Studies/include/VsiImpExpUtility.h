/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiImpExpUtility.h
**
**	Description:
**		Implementation of CVsiImpExpUtility
**
********************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiGlobalDef.h>
#include <VsiStudyModule.h>


// Study Mover flags
#define	VSI_STUDY_MOVER_OVERWRITE	   0x00000001
#define VSI_STUDY_MOVER_AUTODELETE	   0x00000002


namespace VsiImpExpUtility
{
	BOOL IsValidChar(WCHAR c);
	BOOL IsValidClipboardChar(WCHAR c);
	LPWSTR FilterFileName(LPWSTR pszFileName, int length);
	BOOL BuildDefaultDirectoryName(IVsiStudy *pStudy, WCHAR *lpszBuffer, int iBufferLen);
};  // namespace VsiImpExpUtility
