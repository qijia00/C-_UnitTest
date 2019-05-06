/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  dllmain.h
**
**	Description:
**		Declaration of module class.
**
********************************************************************************/

class CVsiImportExportModule : public CAtlDllModuleT< CVsiImportExportModule >
{
public :
	DECLARE_LIBID(LIBID_VsiImportExportModuleLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_VSIIMPORTEXPORTMODULE, "{1AE7F21A-6E99-46CC-9B0D-BF7448A9CC59}")
};

extern class CVsiImportExportModule _AtlModule;
