/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyXml.h
**
**	Description:
**		Study XML support
**
*******************************************************************************/

#pragma once


#define VSI_STUDY_XML_ELM_STUDY						L"study"

#define VSI_STUDY_XML_ATT_ID						L"id"
#define VSI_STUDY_XML_ATT_ID_CREATED				L"idCreated"
#define VSI_STUDY_XML_ATT_ID_COPIED					L"idCopied"
#define VSI_STUDY_XML_ATT_NAME						L"name"
#define VSI_STUDY_XML_ATT_CREATED_DATE				L"createdDate"
#define VSI_STUDY_XML_ATT_LOCK						L"locked"
#define VSI_STUDY_XML_ATT_OWNER						L"owner"
#define VSI_STUDY_XML_ATT_EXPORTED					L"exported"
#define VSI_STUDY_XML_ATT_NOTES						L"notes"
#define VSI_STUDY_XML_ATT_GRANTING_INSTITUTION		L"grantingInstitution"
#define VSI_STUDY_XML_ATT_VERSION_REQUIRED			L"versionRequired"
#define VSI_STUDY_XML_ATT_VERSION_CREATED			L"versionCreated"
#define VSI_STUDY_XML_ATT_VERSION_MODIFIED			L"versionModified"
#define VSI_STUDY_XML_ATT_ACCESS					L"access"
#define VSI_STUDY_XML_ATT_LAST_MODIFIED_BY			L"lastModifiedBy"
#define VSI_STUDY_XML_ATT_CREATED_LAST_MODIFIED_DATE	L"lastModifiedDate"

#define VSI_SERIES_XML_ELM_SERIES					L"series"
#define VSI_SERIES_XML_ATT_ID						L"id"
#define VSI_SERIES_XML_ATT_ID_STUDY					L"idStudy"
#define VSI_SERIES_XML_ATT_NAME						L"name"
#define VSI_SERIES_XML_ATT_CREATED_DATE				L"createdDate"
#define VSI_SERIES_XML_ATT_ACQUIRED_BY				L"acquiredBy"
#define VSI_SERIES_XML_ATT_EXPORTED					L"exported"
#define VSI_SERIES_XML_ATT_VERSION_REQUIRED			L"versionRequired"
#define VSI_SERIES_XML_ATT_VERSION_CREATED			L"versionCreated"
#define VSI_SERIES_XML_ATT_VERSION_MODIFIED			L"versionModified"
#define VSI_SERIES_XML_ATT_NOTES					L"notes"
#define VSI_SERIES_XML_ATT_APPLICATION				L"application"
#define VSI_SERIES_XML_ATT_MSMNT_PACKAGE			L"msmntPackage"
#define VSI_SERIES_XML_ATT_ANIMAL_ID				L"animalId"
#define VSI_SERIES_XML_ATT_ANIMAL_COLOR				L"animalColor"
#define VSI_SERIES_XML_ATT_STRAIN					L"strain"
#define VSI_SERIES_XML_ATT_SOURCE					L"source"
#define VSI_SERIES_XML_ATT_WEIGHT					L"weight"
#define VSI_SERIES_XML_ATT_TYPE						L"type"
#define VSI_SERIES_XML_ATT_DATE_OF_BIRTH			L"dateOfBirth"
#define VSI_SERIES_XML_ATT_HEART_RATE				L"heartRate"
#define VSI_SERIES_XML_ATT_BODY_TEMP				L"bodyTemp"
#define VSI_SERIES_XML_ATT_SEX						L"sex"
#define VSI_SERIES_XML_VAL_MALE						L"male"
#define VSI_SERIES_XML_VAL_FEMALE					L"female"
#define VSI_SERIES_XML_ATT_PREGNANT					L"pregnant"
#define VSI_SERIES_XML_ATT_DATE_MATED				L"dateMated"
#define VSI_SERIES_XML_ATT_DATE_PLUGGED				L"datePlugged"
#define VSI_SERIES_XML_ATT_CUSTOM1					L"c1"
#define VSI_SERIES_XML_ATT_CUSTOM2					L"c2"
#define VSI_SERIES_XML_ATT_CUSTOM3					L"c3"
#define VSI_SERIES_XML_ATT_CUSTOM4					L"c4"
#define VSI_SERIES_XML_ATT_CUSTOM5					L"c5"
#define VSI_SERIES_XML_ATT_CUSTOM6					L"c6"
#define VSI_SERIES_XML_ATT_CUSTOM7					L"c7"
#define VSI_SERIES_XML_ATT_CUSTOM8					L"c8"
#define VSI_SERIES_XML_ATT_CUSTOM9					L"c9"
#define VSI_SERIES_XML_ATT_CUSTOM10					L"c10"
#define VSI_SERIES_XML_ATT_CUSTOM11					L"c11"
#define VSI_SERIES_XML_ATT_CUSTOM12					L"c12"
#define VSI_SERIES_XML_ATT_CUSTOM13					L"c13"
#define VSI_SERIES_XML_ATT_CUSTOM14					L"c14"
#define VSI_SERIES_XML_ATT_COLOR_INDEX				L"colorIndex"
#define VSI_SERIES_XML_ATT_LAST_MODIFIED_BY			L"lastModifiedBy"
#define VSI_SERIES_XML_ATT_CREATED_LAST_MODIFIED_DATE		L"lastModifiedDate"

#define VSI_SERIES_XML_ELM_FIELDS					L"fields"
#define VSI_SERIES_XML_ELM_FIELD					L"field"
#define VSI_SERIES_XML_FIELD_ATT_NAME				L"name"
#define VSI_SERIES_XML_FIELD_ATT_LABEL				L"label"

#define VSI_IMAGE_XML_ELM_IMAGE						L"image"
#define VSI_IMAGE_XML_ATT_ID						L"id"
#define VSI_IMAGE_XML_ATT_ID_SERIES					L"idSeries"
#define VSI_IMAGE_XML_ATT_ID_STUDY					L"idStudy"
#define VSI_IMAGE_XML_ATT_NAME						L"name"
#define VSI_IMAGE_XML_ATT_ACQUIRED_DATE				L"acquiredDate"
#define VSI_IMAGE_XML_ATT_CREATED_DATE				L"createdDate"
#define VSI_IMAGE_XML_ATT_MODIFIED_DATE				L"modifiedDate"
#define VSI_IMAGE_XML_ATT_MODE						L"mode"
#define VSI_IMAGE_XML_ATT_MODE_DISPLAY				L"modeDisplay"
#define VSI_IMAGE_XML_ATT_LENGTH					L"length"
#define VSI_IMAGE_XML_ATT_ACQID						L"acqId"
#define VSI_IMAGE_XML_ATT_KEY						L"key"
#define VSI_IMAGE_XML_ATT_VERSION_REQUIRED			L"versionRequired"
#define VSI_IMAGE_XML_ATT_VERSION_CREATED			L"versionCreated"
#define VSI_IMAGE_XML_ATT_VERSION_MODIFIED			L"versionModified"
#define VSI_IMAGE_XML_ATT_FRAME_TYPE				L"frameType"
#define VSI_IMAGE_XML_ATT_VERSION					L"version"
#define VSI_IMAGE_XML_ATT_THUMBNAIL_FRAME			L"thumbnail"
#define VSI_IMAGE_XML_ATT_RF_DATA					L"rfData"

#define VSI_IMAGE_XML_ATT_VEVOSTRAIN_UPDATED		L"VevoStrainUpdated"
#define VSI_IMAGE_XML_ATT_VEVOSTRAIN_SERIES_NAME	L"VevoStrainSeriesName"
#define VSI_IMAGE_XML_ATT_VEVOSTRAIN_STUDY_NAME		L"VevoStrainStudyName"
#define VSI_IMAGE_XML_ATT_VEVOSTRAIN_NAME			L"VevoStrainName"
#define VSI_IMAGE_XML_ATT_VEVOSTRAIN_ANIMAL_ID		L"VevoStrainAnimalId"

#define VSI_IMAGE_XML_ATT_SONOVEVO_UPDATED			L"SonoVevoUpdated"
#define VSI_IMAGE_XML_ATT_SONOVEVO_SERIES_NAME		L"SonoVevoSeriesName"
#define VSI_IMAGE_XML_ATT_SONOVEVO_STUDY_NAME		L"SonoVevoStudyName"
#define VSI_IMAGE_XML_ATT_SONOVEVO_NAME				L"SonoVevoName"
#define VSI_IMAGE_XML_ATT_SONOVEVO_ANIMAL_ID		L"SonoVevoAnimalId"

#define VSI_IMAGE_XML_ATT_LAST_MODIFIED_BY			L"lastModifiedBy"
#define VSI_IMAGE_XML_ATT_CREATED_LAST_MODIFIED_DATE	L"lastModifiedDate"

#define VSI_IMAGE_XML_ELM_MODE_DATA					L"modeData"

#define VSI_IMAGE_XML_ELM_MODE_VIEW					L"modeView"

#define VSI_IMAGE_XML_ELM_SETTINGS					L"settings"

// Update
#define VSI_UPDATE_XML_ELM_UPDATE					L"update"
	#define VSI_UPDATE_XML_ATT_RESET_SELECTION		L"resetSelection"

#define VSI_UPDATE_XML_ELM_ADD						L"add"
#define VSI_UPDATE_XML_ELM_REMOVE					L"remove"
	#define VSI_REMOVE_XML_ATT_NUMBER_STUDIES		L"numStudies"
	#define VSI_REMOVE_XML_ATT_NUMBER_SERIES		L"numSeries"
	#define VSI_REMOVE_XML_ATT_NUMBER_IMAGES		L"numImages"
#define VSI_UPDATE_XML_ELM_REFRESH					L"refresh"
#define VSI_UPDATE_XML_ELM_ITEM						L"item"
	#define VSI_UPDATE_XML_ATT_NS					L"ns"
	#define VSI_UPDATE_XML_ATT_PATH					L"path"
	#define VSI_REMOVE_XML_ATT_ERROR				L"error"

#define VSI_UPDATE_XML_SELECT						L"select"
#define VSI_UPDATE_XML_COLLAPSE						L"collapse"

#define VSI_IMAGE_XML_ATT_PA_ACQ_TYPE_ID			L"PaAcqType"

#define VSI_IMAGE_XML_ATT_EKV_ACQ_TYPE_ID			L"EKV"

// Update XML example
// <update>
//   <add>
//		<item ns="1234" path="c:/study/1234567"/>
//   </add>
//   <remove>
//      <item ns="3456/4567/5678"/>
//   </remove>
//   <refresh>
//      <item ns="6789"/>
//   </refresh>
// </update>

