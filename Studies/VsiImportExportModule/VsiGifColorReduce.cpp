/*******************************************************************************
**
**  Copyright (c) 1999-2011 VisualSonics Inc. All Rights Reserved.
**
**  VsiGifColorReduce.cpp
**
**	Description:
**		Implementation of CVsiGifColorReduce
**
********************************************************************************/

#include "stdafx.h"
#include "VsiGifColorReduce.h"
#include <GdiplusImaging.h>

HRESULT CVsiGifColorReduce::Initialize(int iXRes, int iYRes, int iPitch)
{
	HRESULT hr = S_OK;

	try
	{
		Uninitialize();

		m_iXRes = iXRes;
		m_iYRes = iYRes;
		m_iPitchSrc = iPitch;

		m_ImgDst.Create(m_iXRes, m_iYRes, 8);

		// Get source data
		m_iPitchDst = m_ImgDst.GetPitch();
		m_pbtDataDst = (BYTE *)m_ImgDst.GetBits();	
		if(m_iPitchDst<0)
		{
			m_pbtDataDst = m_pbtDataDst + (m_iYRes-1) * m_iPitchDst;
		}
		m_iPitchDst = abs(m_iPitchDst);

		// Make sure we initialize for 255 colours (not 256)
		hr = m_quantize.Initialize(m_iXRes, m_iYRes, m_iPitchSrc, 24, TRUE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiGifColorReduce::Uninitialize()
{
	m_quantize.Uninitialize();
	m_ImgDst.Destroy();

	m_pbtDataDst = NULL;
	m_pbtDataSrc = NULL;

	return S_OK;
}	

// Base time 
//
// Quant 53 ms
// BPP 363 ms
// Total 415 ms

HRESULT CVsiGifColorReduce::Convert24To8Bit(BYTE *btDataSrc, BYTE *btDataDst, BYTE *pdwMask, DWORD *pdwDstSize)
{
	HRESULT hr = S_OK;
	
	// Reset palette match
	m_dwFilledPreviousPal = 0;
	m_dwWritePtr = 0;

	try
	{		
		// Quantize colour spectrum		
		m_quantize.ProcessImage(btDataSrc);			
		m_quantize.GetColorTable(m_prgbPal);
		m_pColorMap = m_quantize.GetColorMap();
		//printf("Num unique colours = (%d)\n", m_pColorMap->size());		
		
		// Set palette - index 255 is transparent
		m_prgbPal[255].rgbBlue = 0;
		m_prgbPal[255].rgbGreen = 0;
		m_prgbPal[255].rgbRed = 0;
		m_prgbPal[255].rgbReserved = 0;
		m_ImgDst.SetColorTable(0, 256, m_prgbPal);

		m_pbtDataSrc = btDataSrc;		

		// Decrease color depth
		DecreaseBpp();

		// Set mask if this has transparency
		if(NULL != pdwMask)
		{
			for(int y=0, iCount=0; y<m_iYRes; ++y)
			{
				for(int x=0; x<m_iXRes; ++x, ++iCount)
				{
					if(pdwMask[iCount] == 0)
					{
						// Set to transparent color index if 0
						m_pbtDataDst[x + y*m_iPitchDst] = 255;
					}
				}
			}
		}

		// Get image
		hr = GetGifBuffer(btDataDst, pdwDstSize);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);	
	}
	VSI_CATCH(hr);

	m_pbtDataSrc = NULL;
	m_pColorMap = NULL;
		
	return hr;
}

HRESULT CVsiGifColorReduce::GetGifBuffer(BYTE *pbtGifData, DWORD *pdwFileSize)
{
	HRESULT hr = S_OK;
	CComPtr<IStream> pIStream;

	try
	{		
		// Create stream with 0 size
		hr = CreateStreamOnHGlobal(NULL, TRUE, (LPSTREAM*)&pIStream);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
		// DEFINE_GUID(ImageFormatGIF, 0xb96b3cb0,0x0728,0x11d3,0x9d,0x7b,0x00,0x00,0xf8,0x1e,0xf3,0x2e);
		
		GUID ImageFormatGIF;
		ImageFormatGIF.Data1 = 0xb96b3cb0;
		ImageFormatGIF.Data2 = 0x0728;
		ImageFormatGIF.Data3 = 0x11d3;
		ImageFormatGIF.Data4[0] = 0x9d;
		ImageFormatGIF.Data4[1] = 0x7b;
		ImageFormatGIF.Data4[2] = 0x00;
		ImageFormatGIF.Data4[3] = 0x00;
		ImageFormatGIF.Data4[4] = 0xf8;
		ImageFormatGIF.Data4[5] = 0x1e;
		ImageFormatGIF.Data4[6] = 0xf3;
		ImageFormatGIF.Data4[7] = 0x2e;

		hr = m_ImgDst.Save(pIStream, ImageFormatGIF);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// get the size of the stream
		ULARGE_INTEGER ulnSize;
		LARGE_INTEGER lnOffset;
		lnOffset.QuadPart = 0;
		hr = pIStream->Seek(lnOffset, STREAM_SEEK_END, &ulnSize);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Ensure that the file size is going to be large enough
		VSI_VERIFY(*pdwFileSize >= ulnSize.QuadPart, VSI_LOG_ERROR, L"Insuficient memory allocated");

		// now move the pointer to the beginning of the file
		hr = pIStream->Seek(lnOffset, STREAM_SEEK_SET, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		ULONG ulBytesRead;
        hr = pIStream->Read(pbtGifData, (ULONG)(ulnSize.QuadPart), &ulBytesRead);
        VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
		// Return new size
		*pdwFileSize = (DWORD)(ulnSize.QuadPart);
	}
	VSI_CATCH(hr);

	return hr;
}

void CVsiGifColorReduce::DecreaseBpp()
{	
	int er, eg, eb;	
	RGBQUAD q;
	RGBQUAD q2;
	BYTE btIndex;	
	int coeff(0);

	ZeroMemory(m_pbtDataDst, m_iYRes * m_iPitchDst);

	q.rgbReserved = 0;
	q2.rgbReserved = 0;

	for (int y=1; y<m_iYRes-1; y++)
	{		
		for (int x=1; x<m_iXRes-1; x++)
		{
			GetColor24Bit(x, y, &q);
			
			btIndex = SetColor8Bit(x, y, q);
			
			er = (int)q.rgbRed - (int)m_prgbPal[btIndex].rgbRed;
			eg = (int)q.rgbGreen - (int)m_prgbPal[btIndex].rgbGreen;
			eb = (int)q.rgbBlue - (int)m_prgbPal[btIndex].rgbBlue;
					
			GetColor24Bit(x+1, y, &q);
			q2.rgbRed = (BYTE)min(255L, max(0L, (int)q.rgbRed + ((er*7)/16)));
			q2.rgbGreen = (BYTE)min(255L, max(0L, (int)q.rgbGreen + ((eg*7)/16)));
			q2.rgbBlue = (BYTE)min(255L, max(0L, (int)q.rgbBlue + ((eb*7)/16)));
															
			SetColor24Bit(x+1, y, q2);
						
			for(int i=-1; i<2; i++)
			{
				switch(i)
				{
					case -1:
						coeff=2; 
						break;
					case 0:
						coeff=4; 
						break;
					case 1:
						coeff=1; 
						break;
				}
				
				GetColor24Bit(x+i, y+1, &q);
				q2.rgbRed = (BYTE)min(255L, max(0L, (int)q.rgbRed + ((er * coeff)/16)));
				q2.rgbGreen = (BYTE)min(255L, max(0L, (int)q.rgbGreen + ((eg * coeff)/16)));
				q2.rgbBlue = (BYTE)min(255L, max(0L, (int)q.rgbBlue + ((eb * coeff)/16)));
							
				SetColor24Bit(x+i,y+1, q2);
			}			
		}
	}
}