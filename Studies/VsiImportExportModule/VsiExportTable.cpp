/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportTable.cpp
**
**	Description:
**		Implementation of CVsiExportTable
**
*******************************************************************************/

#include "stdafx.h"
#include "VsiExportTable.h"


CVsiExportTable::CVsiExportTable() :
	m_hwndTable(NULL),
	m_iNumColumns(0)
{
	ZeroMemory(m_piColumnWidth, sizeof(m_piColumnWidth));
}

CVsiExportTable::~CVsiExportTable()
{
}

void CVsiExportTable::SetTable(HWND hwndTable)
{
	m_hwndTable = hwndTable;
}

HRESULT CVsiExportTable::DoExport(LPCWSTR pszPath)
{
	_ASSERT(IsWindow(m_hwndTable));
	HRESULT hr;
	HANDLE hFile(NULL);

	try
	{
		WCHAR szBuffer[512];

		LVCOLUMN lvCol;
		memset(&lvCol, 0x00, sizeof(lvCol));
		lvCol.mask = LVCF_TEXT;
		lvCol.pszText = szBuffer;
		lvCol.cchTextMax = _countof(szBuffer);

		// Spin through all of the columns on the table and find the longest
		// string in each column
		BOOL bMoreColumns(FALSE);

		do
		{
			lvCol.iSubItem = m_iNumColumns;
			bMoreColumns = ListView_GetColumn(m_hwndTable, m_iNumColumns, &lvCol);
			if (bMoreColumns)
			{
				m_piColumnWidth[m_iNumColumns] = FindMaxString(m_iNumColumns, lstrlen(szBuffer));
				++m_iNumColumns;
			}
		} while (bMoreColumns);

		hFile = CreateFile(
			pszPath,
			GENERIC_WRITE,
			FILE_SHARE_READ,
			NULL,
			OPEN_ALWAYS,
			FILE_ATTRIBUTE_NORMAL,
			NULL);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			// Not good
			VSI_WIN32_FAIL(
				GetLastError(),
				VSI_LOG_WARNING,
				L"Unable to create table file!");
		}
		else
		{
			hr = WriteHeader(hFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Unable to write table file");

			hr = WriteTable(hFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Unable to write table file");

			BOOL bRet = SetEndOfFile(hFile);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Unable to write table file");

			CloseHandle(hFile);
			hFile = NULL;
		}
	}
	VSI_CATCH(hr);

	if (NULL != hFile)
		CloseHandle(hFile);

	return hr;
}

int CVsiExportTable::FindMaxString(int iCol, int iMaxSoFar)
{
	WCHAR szBuffer[500];
	LVITEM lvi = { LVIF_INDENT | LVIF_TEXT };
	lvi.pszText = szBuffer;
	lvi.cchTextMax = _countof(szBuffer);
	lvi.iItem = 0;
	lvi.iSubItem = iCol;

	int iNumItems = ListView_GetItemCount(m_hwndTable);
	for (lvi.iItem = 0; lvi.iItem < iNumItems; ++lvi.iItem)
	{
		ListView_GetItem(m_hwndTable, &lvi);

		int iLen = (int)lstrlen(szBuffer);
		if (0 == iCol && 0 < lvi.iIndent)
		{
			iLen += lvi.iIndent * 2;
		}

		iMaxSoFar = max(iMaxSoFar, iLen);
	}

	// Add a bit of white space between the columns.
	return iMaxSoFar + 2;
}

HRESULT CVsiExportTable::WriteHeader(HANDLE hFile)
{
	HRESULT hr(S_OK);

	try
	{
		WCHAR szBuffer[500];

		LVCOLUMN lvCol = { LVCF_TEXT };
		lvCol.pszText = szBuffer;
		lvCol.cchTextMax = _countof(szBuffer);

		BOOL bRet;
		int iTableWidth(0);
		DWORD dwWritten;

		for (int col = 0; col < m_iNumColumns; ++col)
		{
			ListView_GetColumn(m_hwndTable, col, &lvCol);

			int iCount = m_piColumnWidth[col] - lstrlen(szBuffer);
			for (int i = 0; i < iCount; ++i)
			{
				lstrcat(szBuffer, L" ");
			}

			DWORD dwDataSize = lstrlen(szBuffer);
			bRet = WriteFile(hFile, CW2A(szBuffer), dwDataSize, &dwWritten, NULL);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Failed to write data buffer");

			iTableWidth += dwDataSize;
		}

		bRet = WriteFile(hFile, "\r\n", 2, &dwWritten, NULL);
		VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Failed to write data buffer");

		// Write a separator between the headings and the table contents.
		for (int i = 0; i < iTableWidth; ++i)
		{
			bRet = WriteFile(hFile, "-", 1, &dwWritten, NULL);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Failed to write data buffer");
		}

		bRet = WriteFile(hFile, "\r\n", 2, &dwWritten, NULL);
		VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Failed to write data buffer");
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiExportTable::WriteTable(HANDLE hFile)
{
	HRESULT hr(S_OK);

	try
	{
		WCHAR szBuffer[500];
		LVITEM lvi = { LVIF_INDENT | LVIF_TEXT };
		lvi.pszText = szBuffer;
		lvi.cchTextMax = _countof(szBuffer);

		DWORD dwDataSize;
		DWORD dwWritten;
		BOOL bRet;

		int iNumItems = ListView_GetItemCount(m_hwndTable);
		for (lvi.iItem = 0; lvi.iItem < iNumItems; ++lvi.iItem)
		{
			for (lvi.iSubItem = 0; lvi.iSubItem < m_iNumColumns; ++lvi.iSubItem)
			{
				ListView_GetItem(m_hwndTable, &lvi);

				int iCount = m_piColumnWidth[lvi.iSubItem] - lstrlen(szBuffer);

				if (0 == lvi.iSubItem && 0 < lvi.iIndent)
				{
					CStringA strIndent;
					for (int i = 0; i < lvi.iIndent; ++i)
					{
						strIndent += "  ";
						iCount -= 2;
					}

					dwDataSize = strIndent.GetLength();
					bRet = WriteFile(hFile, strIndent.GetString(), dwDataSize, &dwWritten, NULL);
					VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Failed to write data buffer");
				}

				for (int i = 0; i < iCount; ++i)
				{
					lstrcat(szBuffer, L" ");
				}

				dwDataSize = lstrlen(szBuffer);
				bRet = WriteFile(hFile, CW2A(szBuffer), dwDataSize, &dwWritten, NULL);
				VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Failed to write data buffer");
			}

			bRet = WriteFile(hFile, "\r\n", 2, &dwWritten, NULL);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Failed to write data buffer");
		}
	}
	VSI_CATCH(hr);

	return hr;
}

