/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportTable.h
**
**	Description:
**		Declaration of the CVsiExportTable
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>


class CVsiExportTable
{
private:
	HWND m_hwndTable;

	int m_iNumColumns;
	int m_piColumnWidth[32];

public:
	CVsiExportTable();
	~CVsiExportTable();

	void SetTable(HWND hwndTable);
	HRESULT DoExport(LPCWSTR pszPath);

private:
	int FindMaxString(int iCol, int iMaxSoFar);
	HRESULT WriteHeader(HANDLE hFile);
	HRESULT WriteTable(HANDLE hFile);
};

