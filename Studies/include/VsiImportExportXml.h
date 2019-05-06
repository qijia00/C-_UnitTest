/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiImportExportXml.h
**
**	Description:
**		Import/Export XML helpers
**
*******************************************************************************/

#pragma once

#include <VsiImportExportXmlTags.h>



interface IVsiStudy;
interface IVsiSeries;
interface IVsiImage;


HRESULT VsiCreateStudyElement(
	IXMLDOMDocument* pDoc,
	IXMLDOMElement* pParentElement,
	IVsiStudy *pStudy,
	IXMLDOMElement **ppStudyRecordElement,
	LPCWSTR pszNewOwner,
	BOOL bTempTarget,
	BOOL bCalculateSize);

HRESULT VsiCreateSeriesElement(
	IXMLDOMDocument* pDoc,
	IXMLDOMElement* pParentElement,
	IVsiSeries *pSeries,
	IXMLDOMElement **ppSeriesElement,
	INT64 *piTotalSize);
	
HRESULT VsiCreateImageElement(
	IXMLDOMDocument* pDoc,
	IXMLDOMElement* pParentElement,
	IVsiImage *pImage,
	IXMLDOMElement **ppImageElement,
	INT64 *piTotalSize);