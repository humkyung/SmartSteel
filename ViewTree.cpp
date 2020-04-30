// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface 
// (the "Fluent UI") and is provided only as referential material to supplement the 
// Microsoft Foundation Classes Reference and related electronic documentation 
// included with the MFC C++ library software.  
// License terms to copy, use or distribute the Fluent UI are available separately.  
// To learn more about our Fluent UI licensing program, please visit 
// http://msdn.microsoft.com/officeui.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

#include "stdafx.h"
#include "ViewTree.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CViewTree

CViewTree::CViewTree()
{
}

CViewTree::~CViewTree()
{
}

BEGIN_MESSAGE_MAP(CViewTree, CTreeCtrl)
	ON_WM_LBUTTONDOWN()
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CViewTree message handlers

BOOL CViewTree::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	BOOL bRes = CTreeCtrl::OnNotify(wParam, lParam, pResult);

	NMHDR* pNMHDR = (NMHDR*)lParam;
	ASSERT(pNMHDR != NULL);

	if (pNMHDR && pNMHDR->code == TTN_SHOW && GetToolTips() != NULL)
	{
		GetToolTips()->SetWindowPos(&wndTop, -1, -1, -1, -1, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOSIZE);
	}

	return bRes;
}

void CViewTree::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CTreeCtrl::OnLButtonDown(nFlags , point);

	UINT nHitFlags = 0;
	HTREEITEM hItem = HitTest(point , &nHitFlags);
	if (hItem != NULL)
	{
		if (nHitFlags & TVHT_ONITEMSTATEICON || nHitFlags & TVHT_ONITEMLABEL ) 
		{
			this->SelectItem(hItem);
			
			CSmartSteelDoc* pDoc = GetSDIActiveDocument();
			if(pDoc)
			{
				int nImage = -1 , nSelectedImage = -1;
				this->GetItemImage(hItem , nImage , nSelectedImage);
				if((GUSSET_PLATE_ICON == nImage) || (END_PLATE_ICON == nImage))
				{
					CSteelPlate* pPlate = (CSteelPlate*)(this->GetItemData(hItem));
					pDoc->ZoomOCCEntity(pPlate);
				}
				else if((HBRACE_ICON == nImage) || (BEAM_ICON == nImage))
				{
					CSDNFLinearMember* pMember = (CSDNFLinearMember*)(this->GetItemData(hItem));
					pDoc->ZoomOCCEntity(pMember);
				}
			}
		}
	}
}