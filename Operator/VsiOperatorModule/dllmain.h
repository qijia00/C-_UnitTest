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

class CVsiOperatorModule : public CAtlDllModuleT< CVsiOperatorModule >
{
public:
	DECLARE_LIBID(LIBID_VsiOperatorModuleLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_VSIOPERATORMODULE, "{811E78C7-E02C-48C5-945B-53C16327E1CB}")
};

extern class CVsiOperatorModule _AtlModule;
