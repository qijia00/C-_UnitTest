/*******************************************************************************
**
**  Copyright (c) 1999-2011 VisualSonics Inc. All Rights Reserved.
**
**  VsiGifColorQuantizer.h
**
**	Description:
**		Implementation of CVsiGifColorQuantizer
**		Based on CVsiGifColorQuantizer from Jeff Prosise, Davide Pizzolato - www.xdp.it
**
********************************************************************************/

#include "stdafx.h"
#include "VsiGifColorQuantizer.h"

CVsiGifColorQuantizer::CVsiGifColorQuantizer() :
	m_dwColorBits(0),
	m_dwMaxColors(0),	
	m_pTree(NULL),
	m_dwLeafCount(0),
	m_iWidth(0),
	m_iHeight(0),
	m_iPitch(0)

{
	for	(DWORD i=0; i<_countof(m_pReducibleNodes); i++)
		m_pReducibleNodes[i] = NULL;

	ZeroMemory(m_prgbPal, sizeof(m_prgbPal));
}

CVsiGifColorQuantizer::~CVsiGifColorQuantizer()
{
	if (m_pTree	!= NULL)
	{
		DeleteTree (&m_pTree);
	}
}

HRESULT CVsiGifColorQuantizer::Initialize(
	int iWidth, 
	int iHeight, 
	int iPitch,
	DWORD dwBPP,
	BOOL bAllowTransparency)
{
	Uninitialize();

	// Check for 24 BPP (3 bytes)
	if(dwBPP != 24)
	{
		VSI_LOG_MSG(VSI_LOG_ERROR, L"Image must be 24 BPP");
		return E_FAIL;
	}

	if(bAllowTransparency)
	{
		m_dwColorBits = 8;
		m_dwMaxColors = 255;	// index 255 is transparent
	}
	else
	{
		m_dwColorBits = 8;
		m_dwMaxColors = 256;
	}

	m_iWidth = iWidth;
	m_iHeight = iHeight;
	m_iPitch = iPitch;

	return S_OK;
}

HRESULT CVsiGifColorQuantizer::Uninitialize()
{	
	m_iWidth = 0;
	m_iHeight = 0;
	m_iPitch = 0;

	return S_OK;
}


HRESULT CVsiGifColorQuantizer::ProcessImage(BYTE *pbBits)
{	
	int	i, j;
	
	// Correct for pitch adjustments
	int	iPad = m_iPitch - m_iWidth * 3;

	RGBQUAD q;
	ZeroMemory(&q, sizeof(RGBQUAD));

	m_colorMap.clear();

	DWORD *pdwQ = (DWORD *)&q;	
	
	for	(i=0; i<m_iHeight; i++)
	{
		for	(j=0; j<m_iWidth; j++)
		{
			q.rgbBlue =	*pbBits++;
			q.rgbGreen = *pbBits++;
			q.rgbRed =	*pbBits++;
			
			// Add this to unique color list as long as the value is not zero (black)
			if(*pdwQ != 0)
			{				
				m_colorMap.insert(CVsiQuantizeMap::value_type(*pdwQ, 0));
			}

			AddColor (
				&m_pTree,
				q.rgbRed, q.rgbGreen, q.rgbBlue, 0, 
				m_dwColorBits,
				0, 
				&m_dwLeafCount,
				m_pReducibleNodes);

			while (m_dwLeafCount > m_dwMaxColors)
				ReduceTree(m_dwColorBits, &m_dwLeafCount, m_pReducibleNodes);
		}
		pbBits += iPad;
	}

	//DWORD dwNumUniqueColors = (DWORD)m_colorMap.size();
	//VSI_LOG_MSG(VSI_LOG_INFO, L"Num unique colours = %d", dwNumUniqueColors);	

	// Create color table
	CalculateColorTable(m_prgbPal);

	// Do color table mapping
	{
		DWORD dwColor;
		RGBQUAD *prgbColor = (RGBQUAD *)&dwColor;
		BYTE btIndex;

		CVsiQuantizeMapIter iter;
		for (iter = m_colorMap.begin(); iter != m_colorMap.end(); ++iter)
		{
			dwColor = iter->first;
			btIndex = GetNearestIndex(*prgbColor);
			iter->second = btIndex;
		}		
	}
	
	return S_OK;
}

void CVsiGifColorQuantizer::AddColor(
	NODE** ppNode, 
	BYTE r, 
	BYTE g, 
	BYTE b, 
	BYTE a,
	DWORD dwColorBits, 
	DWORD dwLevel, 
	DWORD *pLeafCount, 
	NODE **pReducibleNodes)
{
	static BYTE	mask[8]	= {	0x80, 0x40,	0x20, 0x10,	0x08, 0x04,	0x02, 0x01 };

	// If the node doesn't exist, create it.
	if (*ppNode	== NULL)
		*ppNode	= (NODE*) CreateNode(dwLevel, dwColorBits, pLeafCount, pReducibleNodes);

	// Update color	information	if it's	a leaf node.
	if ((*ppNode)->bIsLeaf)	
	{
		(*ppNode)->dwPixelCount++;
		(*ppNode)->dwRedSum += r;
		(*ppNode)->dwGreenSum += g;
		(*ppNode)->dwBlueSum += b;
		(*ppNode)->dwAlphaSum += a;
	} 
	else
	{	
		// Recurse a level deeper if the node is not a leaf.
		int	iShift = 7 - dwLevel;

		int	iIndex =(((r & mask[dwLevel]) >> iShift) << 2) |
					(((g & mask[dwLevel]) >> iShift) << 1) |
					(( b & mask[dwLevel]) >> iShift);

		AddColor (&((*ppNode)->pChild[iIndex]),	r, g, b, a, dwColorBits,
					dwLevel + 1,	pLeafCount,	pReducibleNodes);
	}
}

void *CVsiGifColorQuantizer::CreateNode(
	DWORD dwLevel,
	DWORD dwColorBits,
	DWORD *pLeafCount,
	NODE **pReducibleNodes)
{
	NODE* pNode = (NODE*)calloc(1,sizeof(NODE));	

	if (pNode== NULL) 
		return NULL;

	pNode->bIsLeaf = (dwLevel == dwColorBits) ? TRUE : FALSE;
	if (pNode->bIsLeaf)
	{
		(*pLeafCount)++;
	}
	else 
	{
		pNode->pNext = pReducibleNodes[dwLevel];
		pReducibleNodes[dwLevel]	= pNode;
	}

	return pNode;
}

void CVsiGifColorQuantizer::ReduceTree(
	DWORD dwColorBits, 
	DWORD *pLeafCount,
	NODE **pReducibleNodes)
{
	// Find	the	deepest	level containing at	least one reducible	node.
	int i;
	for	(i=dwColorBits - 1; (i>0) && (pReducibleNodes[i] == NULL); i--);

	// Reduce the node most	recently added to the list at level	i.
	NODE* pNode	= pReducibleNodes[i];
	pReducibleNodes[i] = pNode->pNext;

	DWORD dwRedSum = 0;
	DWORD dwGreenSum = 0;
	DWORD dwBlueSum = 0;
	DWORD dwAlphaSum = 0;
	DWORD nChildren = 0;

	for	(i=0; i<8; i++)	
	{
		if (pNode->pChild[i] !=	NULL) 
		{
			dwRedSum += pNode->pChild[i]->dwRedSum;
			dwGreenSum += pNode->pChild[i]->dwGreenSum;
			dwBlueSum += pNode->pChild[i]->dwBlueSum;
			dwAlphaSum += pNode->pChild[i]->dwAlphaSum;

			pNode->dwPixelCount += pNode->pChild[i]->dwPixelCount;
			free(pNode->pChild[i]);
			pNode->pChild[i] = NULL;
			nChildren++;
		}
	}

	pNode->bIsLeaf = TRUE;
	pNode->dwRedSum = dwRedSum;
	pNode->dwGreenSum = dwGreenSum;
	pNode->dwBlueSum = dwBlueSum;
	pNode->dwAlphaSum = dwAlphaSum;

	*pLeafCount	-= (nChildren -	1);
}

void CVsiGifColorQuantizer::DeleteTree(NODE**	ppNode)
{
	for	(int i=0; i<8; i++)	
	{
		if ((*ppNode)->pChild[i] !=	NULL) 
			DeleteTree (&((*ppNode)->pChild[i]));
	}

	free(*ppNode);
	*ppNode	= NULL;
}

void CVsiGifColorQuantizer::GetPaletteColors(
	NODE* pTree,	
	RGBQUAD *prgb, 
	DWORD *pIndex, DWORD *pSum)
{
	if (pTree)
	{
		if (pTree->bIsLeaf)
		{
			prgb[*pIndex].rgbRed = (BYTE)((pTree->dwRedSum)/(pTree->dwPixelCount));
			prgb[*pIndex].rgbGreen = (BYTE)((pTree->dwGreenSum)/(pTree->dwPixelCount));
			prgb[*pIndex].rgbBlue = (BYTE)((pTree->dwBlueSum)/(pTree->dwPixelCount));
			prgb[*pIndex].rgbReserved =	(BYTE)((pTree->dwAlphaSum)/(pTree->dwPixelCount));

			if (pSum) pSum[*pIndex] = pTree->dwPixelCount;
			(*pIndex)++;
		}
		else
		{
			for	(int i=0; i<8; i++)
			{
				if (pTree->pChild[i] !=	NULL)
				{
					GetPaletteColors(pTree->pChild[i], prgb, pIndex, pSum);
				}
			}
		}
	}
}

CVsiQuantizeMap *CVsiGifColorQuantizer::GetColorMap()
{
	return &m_colorMap;
}

DWORD CVsiGifColorQuantizer::GetColorCount()
{
	return m_dwLeafCount;
}

void CVsiGifColorQuantizer::GetColorTable(RGBQUAD* prgb)
{
	CopyMemory(prgb, m_prgbPal, sizeof(m_prgbPal));
}

void CVsiGifColorQuantizer::CalculateColorTable(RGBQUAD* prgb)
{
	DWORD iIndex = 0;

	GetPaletteColors(m_pTree, prgb, &iIndex, 0);

	// Sort palette based on relative intensity
	DWORD pdwIntensity[256];
	ZeroMemory(pdwIntensity, sizeof(pdwIntensity));

	for(DWORD i=0; i<m_dwMaxColors; ++i)
	{
		pdwIntensity[i] = prgb[i].rgbBlue + prgb[i].rgbGreen + prgb[i].rgbRed;
	}

	// Ensure that the reserved values all all zero
	for(DWORD i=0; i<m_dwMaxColors; ++i)
	{
		prgb[i].rgbReserved = 0;
	}

	RGBQUAD rgbtemp;
	DWORD dwTemp;
	BOOL bComplete(FALSE);
	while(!bComplete)
	{
		bComplete = TRUE;
		for(DWORD i=0; i<(m_dwMaxColors - 1); ++i)
		{
			if(pdwIntensity[i] > pdwIntensity[i+1])
			{
				rgbtemp = prgb[i];
				prgb[i] = prgb[i+1];
				prgb[i+1] = rgbtemp;

				dwTemp = pdwIntensity[i];
				pdwIntensity[i] = pdwIntensity[i+1];
				pdwIntensity[i+1] = dwTemp;
				bComplete = FALSE;
			}
		}
	}

	// Ensure that position 0 is 0,0,0
	prgb[0].rgbBlue = 0;
	prgb[0].rgbGreen = 0;
	prgb[0].rgbRed = 0;
}