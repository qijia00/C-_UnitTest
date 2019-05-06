/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiDeleteConfirmation.h
**
**	Description:
**		Declaration of the CVsiDeleteConfirmation
**
*******************************************************************************/

#pragma once


#include <VsiRes.h>
#include <VsiStudyModule.h>


class CVsiDeleteConfirmation :
	public CDialogImpl<CVsiDeleteConfirmation>
{
private:
	IUnknown **m_ppStudy;
	IUnknown **m_ppSeries;
	IUnknown **m_ppImage;
	int m_iStudyCount;
	int m_iSeriesCount;
	int m_iImageCount;

public:
	enum { IDD = IDD_DELETE_STUDY_ITEM_CONFIRM };

	CVsiDeleteConfirmation(
		IUnknown **ppStudy,
		int iStudy,
		IUnknown **ppSeries,
		int iSeries,
		IUnknown **ppImage,
		int iImage);
	~CVsiDeleteConfirmation();

//	BOOL GetIsDeleteUnexportedItems() { return m_bDeleteUnexportedItems; }

protected:
	BEGIN_MSG_MAP(CVsiDeleteConfirmation)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		COMMAND_ID_HANDLER(IDYES, OnCmdYes)
		COMMAND_ID_HANDLER(IDNO, OnCmdNo)
		COMMAND_ID_HANDLER(IDCANCEL, OnCmdNo)
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnClose(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnCmdYes(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdNo(WORD, WORD, HWND, BOOL&);
	
private:
	int CheckForMeasurementsOnImage();
};

