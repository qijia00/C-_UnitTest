/*******************************************************************************
**
**  Copyright (c) 1999-2011 VisualSonics Inc. All Rights Reserved.
**
**  VsiGifColorReduce.h
**
**	Description:
**		Implementation of CVsiGifColorReduce
**
********************************************************************************/

#include <atlimage.h>
#include "VsiGifColorQuantizer.h"

#define VSI_MAX_PREVIOUS_PALETTES	1024

class CVsiGifColorReduce
{
private:
	int m_iXRes;
	int m_iYRes;

	int m_iPitchSrc;
	int m_iPitchDst;

	BYTE *m_pbtDataSrc;
	BYTE *m_pbtDataDst;

	RGBQUAD *m_prgbPal;	
	DWORD *m_pdwDiffs;
	
	CImage m_ImgDst;

	// Previously computed palettes
	DWORD *m_pdwPreviousPalettes;
	BYTE *m_pbtPreviousIndexs;
	DWORD m_dwFilledPreviousPal;
	DWORD m_dwWritePtr;
		
	CVsiGifColorQuantizer m_quantize;
	CVsiQuantizeMap *m_pColorMap;

public:
	CVsiGifColorReduce() :
		m_iXRes(0),
		m_iYRes(0),
		m_iPitchSrc(0),
		m_iPitchDst(0),
		m_pbtDataSrc(NULL),
		m_pbtDataDst(NULL)
	{		
		m_prgbPal = (RGBQUAD *)_aligned_malloc(sizeof(RGBQUAD) * 256, 32);
		m_pdwDiffs = (DWORD *)_aligned_malloc(sizeof(DWORD) * 256, 32);
		
		m_pdwPreviousPalettes = (DWORD *)_aligned_malloc(sizeof(DWORD) * VSI_MAX_PREVIOUS_PALETTES, 32);
		m_pbtPreviousIndexs = (BYTE *)_aligned_malloc(sizeof(BYTE) * VSI_MAX_PREVIOUS_PALETTES, 32);
		m_dwFilledPreviousPal = 0;
		m_dwWritePtr = 0;
	}

	~CVsiGifColorReduce()
	{
		if(NULL != m_prgbPal)
		{
			_aligned_free(m_prgbPal);
			m_prgbPal = NULL;
		}

		if(NULL != m_pdwDiffs)
		{
			_aligned_free(m_pdwDiffs);
			m_pdwDiffs = NULL;
		}

		if(NULL != m_pdwPreviousPalettes)
		{
			_aligned_free(m_pdwPreviousPalettes);
			m_pdwPreviousPalettes = NULL;
		}

		if(NULL != m_pbtPreviousIndexs)
		{
			_aligned_free(m_pbtPreviousIndexs);
			m_pbtPreviousIndexs = NULL;
		}
	}

	HRESULT Initialize(int iXRes, int iYRes, int iPitch);
	HRESULT Convert24To8Bit(BYTE *btDataSrc, BYTE *btDataDst, BYTE *pdwMask, DWORD *pdwDstSize);
	HRESULT Uninitialize();
	
private:
	HRESULT GetGifBuffer(BYTE *pbtGifData, DWORD *pdwFileSize);

	inline BYTE GetNearestIndex(RGBQUAD q)
	{
		// Index zero is always 0, 0, 0 - most common colour most likely
		if(q.rgbBlue == 0 && q.rgbGreen == 0 && q.rgbRed == 0)
			return 0;

		// Check to see if the colour is in the map
		DWORD *pdwTempQ = (DWORD *)&q;
		CVsiQuantizeMapIter iter = m_pColorMap->find(*pdwTempQ);

		// Colour is not in the map - do a specific search
		if(iter == m_pColorMap->end())
		{
			BYTE btIndex = GetNearestIndexSearch(q);

			// Add to list for subsequent matches
			m_pColorMap->insert(CVsiQuantizeMap::value_type(*pdwTempQ, btIndex));

			return btIndex;
		}
		else
		{
			return iter->second;
		}
	}
	
	inline BYTE GetNearestIndexSearch(RGBQUAD q)
	{	
		DWORD dwErrorBest(UINT_MAX);		
		DWORD dwError;
		BYTE btIndexBest(0);
		
		// Index 255 is a transparent index
		for(int i=0;i<255;i++)
		{
			// Not sure what one is more efficient

			//dwError = (m_prgbPal[i].rgbBlue-q.rgbBlue)*(m_prgbPal[i].rgbBlue-q.rgbBlue)+
			//	(m_prgbPal[i].rgbGreen-q.rgbGreen)*(m_prgbPal[i].rgbGreen-q.rgbGreen)+
			//	(m_prgbPal[i].rgbRed-q.rgbRed)*(m_prgbPal[i].rgbRed-q.rgbRed);

			dwError = abs(m_prgbPal[i].rgbBlue-q.rgbBlue)+
				abs(m_prgbPal[i].rgbGreen-q.rgbGreen)+
				abs(m_prgbPal[i].rgbRed-q.rgbRed);

			if (dwError == 0)
			{
				return (BYTE)i;
			}
			if (dwError<dwErrorBest)
			{
				dwErrorBest = dwError;
				btIndexBest = (BYTE)i;
			}
		}

		return btIndexBest;
	}
	
	inline void GetColor24Bit(int x, int y, RGBQUAD *pq)
	{
		BYTE *pbtData = &m_pbtDataSrc[x * 3 + y * m_iPitchSrc];
		pq->rgbBlue = *pbtData++;
		pq->rgbGreen = *pbtData++;
		pq->rgbRed = *pbtData;
	}

	// Sets the best colour available from the palette into the 8 bit array
	inline BYTE SetColor8Bit(int x, int y, RGBQUAD &q)
	{	
		BYTE btIndex = GetNearestIndex(q);

		m_pbtDataDst[x + y * m_iPitchDst] = btIndex;

		return btIndex;
	}

	inline void SetColor24Bit(int x, int y, const RGBQUAD &q)
	{		
		m_pbtDataSrc[x * 3 + y * m_iPitchSrc] = q.rgbBlue;
		m_pbtDataSrc[x * 3 + y * m_iPitchSrc + 1] = q.rgbGreen;
		m_pbtDataSrc[x * 3 + y * m_iPitchSrc + 2] = q.rgbRed;
	}

	void DecreaseBpp();
};