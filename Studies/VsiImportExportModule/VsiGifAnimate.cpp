/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiGifAnimate.h
**
**	Description:
**		Code to enable animation of gifEncode files
**
********************************************************************************/

#include "stdafx.h"
#pragma comment(lib, "Vfw32.lib")
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include "VsiGifAnimate.h"


CVsiGifAnimate::CVsiGifAnimate() :
	m_hwndProgress(NULL),
	m_bError(FALSE),
	m_dwXRes(0),
	m_dwYRes(0),
	m_dwNumSourceFrames(0),
	m_bFirstDataSet(TRUE)
{
	ZeroMemory(m_ppbtTransparentMask, sizeof(m_ppbtTransparentMask));
	ZeroMemory(m_ppbtImageDataSrc, sizeof(m_ppbtImageDataSrc));
	ZeroMemory(m_ppbtImageDataDst, sizeof(m_ppbtImageDataDst));

	// Number of threads to use
	m_dwNumThreads = omp_get_num_procs();
	if (m_dwNumThreads > 4)
		m_dwNumThreads = 4;
}

CVsiGifAnimate::~CVsiGifAnimate()
{
	for (int i=0; i<_countof(m_ppbtTransparentMask); ++i)
		_ASSERT(NULL == m_ppbtTransparentMask[i]);
	for (int i=0; i<_countof(m_ppbtImageDataSrc); ++i)
		_ASSERT(NULL == m_ppbtImageDataSrc[i]);
	for (int i=0; i<_countof(m_ppbtImageDataDst); ++i)
		_ASSERT(NULL == m_ppbtImageDataDst[i]);
	for (int i=0; i<_countof(m_pColorReduce); ++i)
		_ASSERT(NULL == m_pColorReduce[i]);		
}

STDMETHODIMP CVsiGifAnimate::Start(
	DWORD dwXRes,
	DWORD dwYRes,
	DWORD dwRate,
	DWORD dwScale,
	LPCOLESTR pszOutputFileName)
{
	HRESULT hr = S_OK;

	UNREFERENCED_PARAMETER(dwRate);
	UNREFERENCED_PARAMETER(dwScale);
	
	try
	{
		m_dwXRes = dwXRes;
		m_dwYRes = dwYRes;
		m_dwNumSourceFrames = 0;
		m_bFirstDataSet = TRUE;

		for(int i=0; i<_countof(m_pColorReduce);++i)
		{
			hr = m_pColorReduce[i].CoCreateInstance(__uuidof(VsiColorQuantize));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Allocate memory for masks
		for(int i=0; i<_countof(m_ppbtTransparentMask);++i)
		{
			m_ppbtTransparentMask[i] = (BYTE *)_aligned_malloc(sizeof(BYTE) * m_dwXRes * m_dwYRes, 32);
			VSI_CHECK_MEM(m_ppbtTransparentMask[i], VSI_LOG_ERROR, NULL);
		}

		// Allocate memory for images
		for(int i=0; i<_countof(m_ppbtImageDataSrc);++i)
		{
			m_ppbtImageDataSrc[i] = (BYTE *)_aligned_malloc(3 * m_dwXRes * m_dwYRes, 32);
			VSI_CHECK_MEM(m_ppbtImageDataSrc[i], VSI_LOG_ERROR, NULL);
		}
	
		DWORD dwImageDestSize = m_dwXRes * m_dwYRes * 3;
		for(int i=0; i<_countof(m_ppbtImageDataDst);++i)
		{
			m_ppbtImageDataDst[i] = (BYTE *)_aligned_malloc(dwImageDestSize, 32);
			VSI_CHECK_MEM(m_ppbtImageDataDst[i], VSI_LOG_ERROR, NULL);
		}

		// Calculate frame rate from avi file		
		USHORT usDelay = 0;

		hr = m_gifEncode.Initialize(pszOutputFileName, (USHORT)m_dwXRes, (USHORT)m_dwYRes, usDelay);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Create color reduce objects (one per thread)
		for (int i=0; i<_countof(m_pColorReduce);++i)
		{
			hr = m_pColorReduce[i]->Initialize(
				m_dwXRes,
				m_dwYRes,
				m_dwXRes*3,
				256,
				TRUE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiGifAnimate::AddFrame(BYTE *pbtData, DWORD dwSize)
{
	HRESULT hr = S_OK;

	_ASSERT(_countof(m_ppbtImageDataSrc) < 64);
	DWORD pdwGifDataSize[64];
	ZeroMemory(pdwGifDataSize, sizeof(pdwGifDataSize) );
	
	try
	{
		VSI_VERIFY(dwSize == m_dwXRes * m_dwYRes * 3, VSI_LOG_ERROR, NULL);

		CopyMemory(m_ppbtImageDataSrc[m_dwNumSourceFrames++], pbtData, dwSize);

		// Check to see if we are ready to compress
		if (m_dwNumSourceFrames == _countof(m_ppbtImageDataSrc))
		{		
			// We only get here if we have a complete set of frames
			// Create transparency masks for these frames - this can be parallel region		
			#pragma omp parallel for num_threads(m_dwNumThreads)
				for (int j=0; j<(int)m_dwNumSourceFrames-1; ++j)
				{
					CreateCompareMask(
						m_ppbtImageDataSrc[j],
						m_ppbtImageDataSrc[j+1],
						m_ppbtTransparentMask[j+1],
						m_dwXRes * m_dwYRes);
					
					pdwGifDataSize[j] = m_dwXRes * m_dwYRes * 3;
				}

			#pragma omp parallel for num_threads(m_dwNumThreads)
				for (int j=0; j<(int)m_dwNumSourceFrames-1; ++j)
				{
					int iThreadNum = omp_get_thread_num();

					// Both next and previous images exist, we have a transperent region
					// Reduce color depth to 8 bits (minus 1 for transparency)
					if (m_bFirstDataSet && j==0)
					{
						hr = m_pColorReduce[iThreadNum]->Convert24To8Bit(
							m_ppbtImageDataSrc[j], 
							m_ppbtImageDataDst[j],
							NULL,					// no transparency for very first frame
							&pdwGifDataSize[j]);
					}
					else
					{
						hr = m_pColorReduce[iThreadNum]->Convert24To8Bit(
							m_ppbtImageDataSrc[j], 
							m_ppbtImageDataDst[j],
							m_ppbtTransparentMask[j],
							&pdwGifDataSize[j]);
					}
					if (FAILED(hr))
					{
						VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure in Convert24To8Bit");
					}
				}

			// Add frames to output - must be serial
			// If first data set then export all frames, otherwise do not export first frame as its hold over for transparency			
			for(int j=0; j<(int)m_dwNumSourceFrames - 1; ++j)
			{
				hr = m_gifEncode.AddImage(m_ppbtImageDataDst[j], pdwGifDataSize[j]);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Writing GIF file failed");
			}

			// Prep next sequence
			CopyMemory(m_ppbtImageDataSrc[0], m_ppbtImageDataSrc[_countof(m_ppbtImageDataSrc) - 1], dwSize);
			CopyMemory(m_ppbtTransparentMask[0], m_ppbtTransparentMask[_countof(m_ppbtImageDataSrc) - 1], m_dwXRes * m_dwYRes);
			m_dwNumSourceFrames = 1;

			m_bFirstDataSet = FALSE;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiGifAnimate::Stop()
{
	HRESULT hr = S_OK;
	
	try
	{
		// Check to see if there are lingering frames that need to be pushed		
		if (m_dwNumSourceFrames > 0)
		{
			// We have frames to complete
			DWORD pdwGifDataSize[64];
			ZeroMemory(pdwGifDataSize, sizeof(pdwGifDataSize) );
		
			// We only get here if we have a complete set of frames
			// Create transparency masks for these frames - this can be parallel region		
			#pragma omp parallel for num_threads(m_dwNumThreads)
				for (int j=0; j<(int)m_dwNumSourceFrames; ++j)
				{
					if (j < (int)m_dwNumSourceFrames-1)
					{
						CreateCompareMask(
							m_ppbtImageDataSrc[j],
							m_ppbtImageDataSrc[j+1],
							m_ppbtTransparentMask[j+1],
							m_dwXRes * m_dwYRes);
					}
					
					pdwGifDataSize[j] = m_dwXRes * m_dwYRes * 3;
				}

			#pragma omp parallel for num_threads(m_dwNumThreads)
				for (int j=0; j<(int)m_dwNumSourceFrames; ++j)
				{
					int iThreadNum = omp_get_thread_num();

					// Both next and previous images exist, we have a transperent region
					// Reduce color depth to 8 bits (minus 1 for transparency)
					if ((m_bFirstDataSet && j==0) ||
						j == (int)m_dwNumSourceFrames-1)
					{
						hr = m_pColorReduce[iThreadNum]->Convert24To8Bit(
							m_ppbtImageDataSrc[j], 
							m_ppbtImageDataDst[j],
							NULL,					// no transparency for very first frame
							&pdwGifDataSize[j]);
					}
					else
					{
						hr = m_pColorReduce[iThreadNum]->Convert24To8Bit(
							m_ppbtImageDataSrc[j], 
							m_ppbtImageDataDst[j],
							m_ppbtTransparentMask[j],
							&pdwGifDataSize[j]);
					}
					if (FAILED(hr))
					{
						VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure in Convert24To8Bit");
					}
				}
			
			// Add frames to output - must be serial
			// If first data set then export all frames, otherwise do not export first frame as its hold over for transparency			
			for(int j=0; j<(int)m_dwNumSourceFrames - 1; ++j)
			{
				hr = m_gifEncode.AddImage(m_ppbtImageDataDst[j], pdwGifDataSize[j]);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Writing GIF file failed");
			}
		}		
	}
	VSI_CATCH(hr);

	// Complete
	m_gifEncode.Uninitialize();

	for(int i=0; i<_countof(m_pColorReduce);++i)
	{
		if (NULL != m_pColorReduce[i])
		{
			m_pColorReduce[i]->Uninitialize();
			m_pColorReduce[i].Release();
		}
	}

	// Free memory
	for(int i=0; i<_countof(m_ppbtTransparentMask);++i)
	{
		if(NULL != m_ppbtTransparentMask[i])
		{
			_aligned_free(m_ppbtTransparentMask[i]);
			m_ppbtTransparentMask[i] = NULL;
		}
	}

	// Allocate memory for images
	for(int i=0; i<_countof(m_ppbtImageDataSrc);++i)
	{
		if(NULL != m_ppbtImageDataSrc[i])
		{
			_aligned_free(m_ppbtImageDataSrc[i]);
			m_ppbtImageDataSrc[i] = NULL;
		}
	}	
	
	for(int i=0; i<_countof(m_ppbtImageDataDst);++i)
	{
		if(NULL != m_ppbtImageDataDst[i])
		{
			_aligned_free(m_ppbtImageDataDst[i]);
			m_ppbtImageDataDst[i] = NULL;
		}
	}

	return hr;
}

STDMETHODIMP CVsiGifAnimate::ConvertAviToGif(
	HWND hwndProgress, 
	LPCOLESTR pszInputFileName, 
	LPCOLESTR pszOutputFileName)
{
	HRESULT hr = S_OK;
	
	try
	{
		VSI_VERIFY(S_FALSE == IsThreadRunning(), VSI_LOG_ERROR, L"Processing thread already active");

		m_hwndProgress = hwndProgress;
		m_strInputFileName = pszInputFileName;
		m_strOutputFileName = pszOutputFileName;		
		m_bError = FALSE;

		// Start thread so we can exit back to UI
		hr = CVsiThread<CVsiGifAnimate>::RunThread(L"VsiGifAnimateThread", THREAD_PRIORITY_NORMAL, FALSE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure starting GIF Animate thread");
		
		HANDLE phThread[] = { m_hThread };
		for (;;)
		{
			DWORD dwWait = MsgWaitForMultipleObjects(
				1,
				phThread,
				FALSE,
				INFINITE,
				QS_ALLINPUT);

			if (WAIT_FAILED == dwWait)
			{
				VSI_LOG_MSG(VSI_LOG_ERROR,
					L"MsgWaitForMultipleObjects received an error condition");
				
				// Continue and hopefully it is going to recover
				continue;
			}

			// Checks shutdown
			if (WAIT_OBJECT_0 == dwWait)
				break;

			MSG msg;
			while (PeekMessage(&msg, NULL, 0, 0, TRUE))
			{
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
	}
	VSI_CATCH(hr);

	if (m_bError)
		hr = E_FAIL;

	return hr;
}

STDMETHODIMP CVsiGifAnimate::ConvertAviToGifDirect(
	HWND hwndProgress, 
	LPCOLESTR pszInputFileName, 
	LPCOLESTR pszOutputFileName)
{
	HRESULT hr = S_OK;
	
	CVsiGifEncode gifEncode;
	CComPtr<IVsiColorQuantize> pColorReduce[4];

	BYTE *ppbtTransparentMask[4] = { NULL };
	BYTE *ppbtImageDataSrc[5] = { NULL };
	BYTE *ppbtImageDataDst[4] = { NULL };
	BYTE *pbtImageDataNextFirst(NULL);

	PAVIFILE pAviFile(NULL);
	PAVISTREAM pAviStream(NULL);
	PGETFRAME pGetFrame(NULL);

	AVIFileInit(); // Intitialize AVI library
	
	try
	{
		for(int i=0; i<_countof(pColorReduce);++i)
		{
			hr = pColorReduce[i].CoCreateInstance(__uuidof(VsiColorQuantize));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Open the file		
		hr = AVIFileOpen(&pAviFile, pszInputFileName, OF_READ, 0);
		VSI_VERIFY(AVIERR_OK == hr, VSI_LOG_WARNING, L"Failed to open AVI file");

		// Get the video stream		
		hr = AVIFileGetStream(
			pAviFile,
			&pAviStream,
			streamtypeVIDEO, 0);
		VSI_VERIFY(AVIERR_OK == hr, VSI_LOG_ERROR, L"Failure in AVIFileGetStream");

		// Get additional information regarding the format
		int iInfoWidth(-1);
		int iInfoHeight(-1);
		{
			LONG lSize; // in bytes
			hr = AVIStreamReadFormat(pAviStream, AVIStreamStart(pAviStream), NULL, &lSize);
			VSI_VERIFY(AVIERR_OK == hr, VSI_LOG_ERROR, L"Failure in AVIStreamReadFormat");
			 
			CHeapPtr<BITMAPINFO> pInfo;
			bool bRet = pInfo.AllocateBytes(lSize);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, NULL);

			hr = AVIStreamReadFormat(pAviStream, AVIStreamStart(pAviStream), pInfo.operator LPBITMAPINFO(), &lSize);
			VSI_VERIFY(AVIERR_OK == hr, VSI_LOG_ERROR, L"Failure in AVIStreamReadFormat");

			iInfoWidth = pInfo->bmiHeader.biWidth;
			iInfoHeight = pInfo->bmiHeader.biHeight;
		}
		
		// get start position and count of frames
		int iFirstFrame = AVIStreamStart(pAviStream);
		int iCountFrames = AVIStreamLength(pAviStream);
		int iLastFrame = iFirstFrame + iCountFrames;

		// get header information            
		AVISTREAMINFO si;
		hr = AVIStreamInfo(pAviStream, &si, sizeof(si));
		VSI_VERIFY(AVIERR_OK == hr, VSI_LOG_ERROR, L"Failure in AVIStreamInfo");

		// Calculate frame rate 
		//double dbFrameRate = (double)si.dwRate / (double)si.dwScale;
		
		{
			// construct the expected bitmap header - always 24 bit please
			BITMAPINFOHEADER bih;
			ZeroMemory(&bih, sizeof(BITMAPINFOHEADER));
			bih.biBitCount = 24;
			bih.biCompression = BI_RGB;

			bih.biHeight = iInfoHeight;
			bih.biWidth = iInfoWidth;
			bih.biPlanes = 1;
			bih.biSize = sizeof(bih);

			// Allocate memory for masks
			for(int i=0; i<_countof(ppbtTransparentMask);++i)
			{
				ppbtTransparentMask[i] = (BYTE *)_aligned_malloc(sizeof(BYTE) * bih.biHeight * bih.biWidth, 32);
				VSI_CHECK_MEM(ppbtTransparentMask[i], VSI_LOG_ERROR, NULL);
			}

			// Allocate memory for images
			for(int i=0; i<_countof(ppbtImageDataSrc);++i)
			{
				ppbtImageDataSrc[i] = (BYTE *)_aligned_malloc(3 * bih.biHeight * bih.biWidth, 32);
				VSI_CHECK_MEM(ppbtImageDataSrc[i], VSI_LOG_ERROR, NULL);
			}

			BOOL bHasNextFirst(FALSE);
			pbtImageDataNextFirst = (BYTE *)_aligned_malloc(3 * bih.biHeight * bih.biWidth, 32);

			DWORD dwImageDestSize = bih.biHeight * bih.biWidth * 3;
			for(int i=0; i<_countof(ppbtImageDataDst);++i)
			{
				ppbtImageDataDst[i] = (BYTE *)_aligned_malloc(dwImageDestSize, 32);
				VSI_CHECK_MEM(ppbtImageDataDst[i], VSI_LOG_ERROR, NULL);
			}

			// Calculate frame rate from avi file
			//double dbDelaySec = 1.0 / dbFrameRate;
			USHORT usDelay = 0;//(USHORT)(0.5 + dbDelaySec * 100.0);

			hr = gifEncode.Initialize(pszOutputFileName, (USHORT)bih.biWidth, (USHORT)bih.biHeight, usDelay);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Create color reduce objects (one per thread)
			for (int i=0; i<_countof(pColorReduce);++i)
			{
				hr = pColorReduce[i]->Initialize(
					bih.biWidth,
					bih.biHeight,
					bih.biWidth*3,
					256,
					TRUE);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		
			if (NULL != hwndProgress)
			{
				::SendMessage(hwndProgress,
					WM_VSI_PROGRESS_SETMAX, iLastFrame - iFirstFrame + 1, 0);

				::SendMessage(hwndProgress,
					WM_VSI_PROGRESS_SETPOS, 0, 0);
			}

		
			// prepare to decompress DIBs (device independent bitmaps)			
			pGetFrame = AVIStreamGetFrameOpen(pAviStream, &bih);			
			if (NULL != pGetFrame)
			{
				BOOL bCancel(FALSE);

				// Process 4 frames at a time
				for(int i=iFirstFrame; i<iLastFrame && !bCancel; i+=4)
				{
					// Update progress
					//szText.Format(L"Processing %d/%d\n", i-iFirstFrame, iCountFrames);
					//cWnd.SetWindowText(szText);											

					// Read/Extract data for next 5 frames so we can do parallel processing
					// We may not get a full 5 frame set (between 1 and 5)
					BOOL pbValidFrames[5] = {FALSE, FALSE, FALSE, FALSE, FALSE};
					DWORD pdwGifDataSize[4] = {0, 0, 0, 0};
					DWORD dwNumValidFrames(0);

					// Data is only valid until next call to AVIStreamGetFrameOpen - must be serial
					// I don't think we can request reverse frames (i.e. a frame we have acquired in the past)
					for(int j=0; j<4; ++j)
					{
						if( (i+j) >= iFirstFrame && (i+j) < iLastFrame)
						{
							BYTE *pBitmapData = (BYTE *)AVIStreamGetFrame(pGetFrame, i+j);
							VSI_VERIFY(pBitmapData != NULL, VSI_LOG_ERROR, VsiFormatMsg(L"Failure getting frame %d", i+j));

							CopyMemory(ppbtImageDataSrc[j+1], pBitmapData + sizeof(BITMAPINFOHEADER), 3 * bih.biHeight * bih.biWidth);							

							pbValidFrames[j+1] = TRUE;
							dwNumValidFrames++;
						}
					}

					if ((BOOL)::SendMessage(hwndProgress, WM_VSI_PROGRESS_GETCANCEL, 0, 0))
					{
						VSI_THROW_HR(S_OK, VSI_LOG_INFO, L"Writing GIF file cancelled");
					}

					// Get previous frame
					if(bHasNextFirst)
					{
						pbValidFrames[0] = TRUE;
						CopyMemory(ppbtImageDataSrc[0], pbtImageDataNextFirst, 3 * bih.biHeight * bih.biWidth);
					}

					// Get next first frame
					bHasNextFirst = TRUE;
					CopyMemory(pbtImageDataNextFirst, ppbtImageDataSrc[4], 3 * bih.biHeight * bih.biWidth);					

					// Create transparency masks for these frames - this can be parallel region				
					#pragma omp parallel for num_threads(m_dwNumThreads)
						for(int j=0;j<4;++j)
						{
							// Both frame pairs must be valid - go in reverse direction so we don't overlap
							if(pbValidFrames[4-j] && pbValidFrames[3-j])
							{
								// Create mask
								CreateCompareMask(
									ppbtImageDataSrc[3-j], 
									ppbtImageDataSrc[4-j], 
									ppbtTransparentMask[3-j], 
									bih.biHeight * bih.biWidth);
							}

							pdwGifDataSize[3-j] = dwImageDestSize;
						}

					#pragma omp parallel for num_threads(m_dwNumThreads)
						for(int j=0;j<4;++j)
						{
							int iThreadNum = omp_get_thread_num();

							// Both next and previous images exist, we have a transperent region
							if(pbValidFrames[4-j] && pbValidFrames[3-j])
							{
								// Reduce color depth to 8 bits (minus 1 for transparency)
								hr = pColorReduce[iThreadNum]->Convert24To8Bit(
									ppbtImageDataSrc[4-j], 
									ppbtImageDataDst[3-j],
									ppbtTransparentMask[3-j],
									&pdwGifDataSize[3-j]);
							}
							// No previous image, no transparency
							else if(pbValidFrames[4-j])
							{
								// Reduce color depth to 8 bits (minus 1 for transparency)
								hr = pColorReduce[iThreadNum]->Convert24To8Bit(
									ppbtImageDataSrc[4-j], 
									ppbtImageDataDst[3-j],
									NULL,
									&pdwGifDataSize[3-j]);
							}
							if (FAILED(hr))
							{
								VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure in Convert24To8Bit");
							}
						}

					// Add frames to output - must be serial
					for(int j=0;j<4;++j)
					{
						if(pbValidFrames[j+1])
						{							
							if (S_OK != gifEncode.AddImage(ppbtImageDataDst[j], pdwGifDataSize[j]))
								VSI_THROW_HR(S_FALSE, VSI_LOG_WARNING, L"Writing GIF file failed");
						}
					}

					if (NULL != hwndProgress)
					{
						::SendMessage(hwndProgress, WM_VSI_PROGRESS_SETPOS, i - iFirstFrame + 4, 0);

						// Get cancel state
						bCancel = (BOOL)::SendMessage(hwndProgress, WM_VSI_PROGRESS_GETCANCEL, 0, 0);

						if (bCancel)
							hr = S_FALSE;
					}
				}				
				
				gifEncode.Uninitialize();
			}
			else
			{
				VSI_FAIL(VSI_LOG_ERROR, L"Failure decompressing frame");
			}
		}		
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
		{
			VSI_ERROR_LOG(err);
		}

		m_bError = TRUE;
	}

	if (pGetFrame != NULL)
	{
		AVIStreamGetFrameClose(pGetFrame);
		pGetFrame = NULL;
	}
	
	if (pAviStream != NULL)
	{
		AVIStreamRelease(pAviStream);
		pAviStream = NULL;
	}

	// Free AVI functions
	if(NULL != pAviFile)
	{
		AVIFileRelease(pAviFile);
		pAviFile = NULL;
	}

	AVIFileExit();

	// Make sure processors are uninitialized
	gifEncode.Uninitialize();

	for(int i=0; i<_countof(pColorReduce);++i)
	{
		if (NULL != pColorReduce[i])
		{
			pColorReduce[i]->Uninitialize();
			pColorReduce[i].Release();
		}
	}

	// Free memory
	for(int i=0; i<_countof(ppbtTransparentMask);++i)
	{
		if(NULL != ppbtTransparentMask[i])
		{
			_aligned_free(ppbtTransparentMask[i]);
			ppbtTransparentMask[i] = NULL;
		}
	}

	// Allocate memory for images
	for(int i=0; i<_countof(ppbtImageDataSrc);++i)
	{
		if(NULL != ppbtImageDataSrc[i])
		{
			_aligned_free(ppbtImageDataSrc[i]);
			ppbtImageDataSrc[i] = NULL;
		}
	}

	if(NULL != pbtImageDataNextFirst)
	{
		_aligned_free(pbtImageDataNextFirst);
		pbtImageDataNextFirst = NULL;
	}
	
	for(int i=0; i<_countof(ppbtImageDataDst);++i)
	{
		if(NULL != ppbtImageDataDst[i])
		{
			_aligned_free(ppbtImageDataDst[i]);
			ppbtImageDataDst[i] = NULL;
		}
	}

	return hr;
}

DWORD CVsiGifAnimate::ThreadProc()
{
	HRESULT hr = CoInitializeEx(0, COINIT_MULTITHREADED);

	try
	{
		hr = ConvertAviToGifDirect(m_hwndProgress, m_strInputFileName, m_strOutputFileName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure converting AVI to GIF");
	}
	VSI_CATCH(hr);

	CoUninitialize();

	return 0;
}

// Compare two 24 bits images and set a mask with 1 where values are the same
void CVsiGifAnimate::CreateCompareMask(
	const BYTE *pbtSource0, 
	const BYTE *pbtSource1, 
	BYTE *pbtMask, 
	DWORD dwLength)
{		
	memset(pbtMask, 1, dwLength);

	// Mark same values with 1
	for(DWORD i=0; i<dwLength; ++i)
	{
		if(pbtSource0[0] == pbtSource1[0] &&
			pbtSource0[1] == pbtSource1[1] &&
			pbtSource0[2] == pbtSource1[2])
		{
			pbtMask[i] = 0;
		}
		
		pbtSource0 += 3;
		pbtSource1 += 3;
	}
}

// Set pixels in an image to zero if they are transparent
//void CVsiGifAnimate::SetFramemask(
//		BYTE *pbtSource,
//		BYTE *pbtMask, 
//		DWORD dwLength)
//{
//	for(DWORD i=0; i<dwLength; ++i)
//	{
//		pbtSource[0] *= pbtMask[i];
//		pbtSource[1] *= pbtMask[i];
//		pbtSource[2] *= pbtMask[i];		
//		
//		pbtSource += 3;		
//	}
//}