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
#include <assert.h>
#include "mainfrm.h"
#include "FileView.h"
#include "Resource.h"
#include "SmartSteel.h"

#include "Command/DeleteCommand.h"
#include "Command/AppCommandManager.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CFileView

CFileView::CFileView()
{
	m_hRoot = m_hColumn = m_hBeam = m_hHBrace = m_hVBrace = NULL;

	m_oElmCountMap.insert(make_pair(CSDNFLinearMember::COLUMN , 0));
	m_oElmCountMap.insert(make_pair(CSDNFLinearMember::BEAM , 0));
	m_oElmCountMap.insert(make_pair(CSDNFLinearMember::HBRACE , 0));
	m_oElmCountMap.insert(make_pair(CSDNFLinearMember::VBRACE , 0));

	m_hGussetPlate = m_hEndPlate = NULL;
}

CFileView::~CFileView()
{
}

BEGIN_MESSAGE_MAP(CFileView, CDockablePane)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_CONTEXTMENU()
	ON_COMMAND(ID_PROPERTIES, OnProperties)
	ON_COMMAND(ID_OPEN, OnFileOpen)
	ON_COMMAND(ID_OPEN_WITH, OnFileOpenWith)
	ON_COMMAND(ID_DUMMY_COMPILE, OnDummyCompile)
	ON_COMMAND(ID_EDIT_CUT, OnEditCut)
	ON_COMMAND(ID_EDIT_COPY, OnEditCopy)
	ON_COMMAND(ID_EDIT_DELETE, OnEditDelete)
	ON_WM_PAINT()
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CWorkspaceBar message handlers

int CFileView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDockablePane::OnCreate(lpCreateStruct) == -1)
		return -1;

	CRect rectDummy;
	rectDummy.SetRectEmpty();

	// Create view:
	const DWORD dwViewStyle = WS_CHILD | WS_VISIBLE | TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS | TVS_SHOWSELALWAYS;

	if (!m_wndFileView.Create(dwViewStyle, rectDummy, this, 4))
	{
		TRACE0("Failed to create file view\n");
		return -1;      // fail to create
	}

	/// Load view images:
	m_FileViewImages.Create(IDB_FILE_VIEW, 16, 0, RGB(255, 0, 255));
	m_wndFileView.SetImageList(&m_FileViewImages, TVSIL_NORMAL);

	m_wndToolBar.Create(this, AFX_DEFAULT_TOOLBAR_STYLE, IDR_EXPLORER);
	m_wndToolBar.LoadToolBar(IDR_EXPLORER, 0, 0, TRUE /* Is locked */);

	OnChangeVisualStyle();

	m_wndToolBar.SetPaneStyle(m_wndToolBar.GetPaneStyle() | CBRS_TOOLTIPS | CBRS_FLYBY);

	m_wndToolBar.SetPaneStyle(m_wndToolBar.GetPaneStyle() & ~(CBRS_GRIPPER | CBRS_SIZE_DYNAMIC | CBRS_BORDER_TOP | CBRS_BORDER_BOTTOM | CBRS_BORDER_LEFT | CBRS_BORDER_RIGHT));

	m_wndToolBar.SetOwner(this);

	/// All commands will be routed via this control , not via the parent frame:
	m_wndToolBar.SetRouteCommandsViaFrame(FALSE);

	/// Fill in some static tree view data (dummy code, nothing magic here)
	FillFileView();
	AdjustLayout();

	return 0;
}

void CFileView::OnSize(UINT nType, int cx, int cy)
{
	CDockablePane::OnSize(nType, cx, cy);
	AdjustLayout();
}

/**
	@brief	
	@author	humkyung
*/
void CFileView::FillFileView()
{
	m_hRoot = m_wndFileView.InsertItem(_T("SDNF elements"), 0, 0);
	m_wndFileView.SetItemState(m_hRoot, TVIS_BOLD, TVIS_BOLD);

	m_hColumn = m_wndFileView.InsertItem(_T("Columns(0)"), 0, 0, m_hRoot);
	m_hBeam = m_wndFileView.InsertItem(_T("Beams(0)"), 0, 0, m_hRoot);
	m_hHBrace = m_wndFileView.InsertItem(_T("HBraces(0)"), 0, 0, m_hRoot);
	m_hVBrace = m_wndFileView.InsertItem(_T("VBraces(0)"), 0, 0, m_hRoot);

	m_wndFileView.Expand(m_hRoot , TVE_EXPAND);

	m_hGussetPlate = m_wndFileView.InsertItem(_T("Gusset Plates(0)"), 0, 0);
	m_wndFileView.SetItemData(m_hGussetPlate , 0);
	m_hEndPlate = m_wndFileView.InsertItem(_T("End Plates(0)"), 0, 0);
	m_wndFileView.SetItemData(m_hEndPlate , 0);
}

void CFileView::OnContextMenu(CWnd* pWnd, CPoint point)
{
	CTreeCtrl* pWndTree = (CTreeCtrl*) &m_wndFileView;
	ASSERT_VALID(pWndTree);

	if (pWnd != pWndTree)
	{
		CDockablePane::OnContextMenu(pWnd, point);
		return;
	}

	if (point != CPoint(-1, -1))
	{
		// Select clicked item:
		CPoint ptTree = point;
		pWndTree->ScreenToClient(&ptTree);

		UINT flags = 0;
		HTREEITEM hTreeItem = pWndTree->HitTest(ptTree, &flags);
		if ((hTreeItem != NULL) && (flags & TVHT_ONITEMSTATEICON || flags & TVHT_ONITEMLABEL ))
		{
			pWndTree->SelectItem(hTreeItem);

			int nImage = -1 , nSelectedImage = -1;
			pWndTree->GetItemImage(hTreeItem , nImage , nSelectedImage);
			pWndTree->SetFocus();
			if((6 == nImage) || (7 == nImage))
			{
				theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EXPLORER, point.x, point.y, this, TRUE);
			}
		}
	}
}

void CFileView::AdjustLayout()
{
	if (GetSafeHwnd() == NULL)
	{
		return;
	}

	CRect rectClient;
	GetClientRect(rectClient);

	int cyTlb = m_wndToolBar.CalcFixedLayout(FALSE, TRUE).cy;

	m_wndToolBar.SetWindowPos(NULL, rectClient.left, rectClient.top, rectClient.Width(), cyTlb, SWP_NOACTIVATE | SWP_NOZORDER);
	m_wndFileView.SetWindowPos(NULL, rectClient.left + 1, rectClient.top + cyTlb + 1, rectClient.Width() - 2, rectClient.Height() - cyTlb - 2, SWP_NOACTIVATE | SWP_NOZORDER);
}

void CFileView::OnProperties()
{
	AfxMessageBox(_T("Properties...."));

}

void CFileView::OnFileOpen()
{
	// TODO: Add your command handler code here
}

void CFileView::OnFileOpenWith()
{
	// TODO: Add your command handler code here
}

void CFileView::OnDummyCompile()
{
	// TODO: Add your command handler code here
}

void CFileView::OnEditCut()
{
	// TODO: Add your command handler code here
}

void CFileView::OnEditCopy()
{
	// TODO: Add your command handler code here
}

void CFileView::OnEditDelete()
{
	CTreeCtrl* pWndTree = (CTreeCtrl*) &m_wndFileView;
	ASSERT_VALID(pWndTree);

	HTREEITEM hItem = pWndTree->GetSelectedItem();
	if(NULL != hItem)
	{
		CSteelPlate* pPlate = (CSteelPlate*)(pWndTree->GetItemData(hItem));
		if(NULL != pPlate)
		{
			vector<CSteelPlate*> oPlateSet;
			oPlateSet.push_back( pPlate );

			Command::CAppCommandManager& inst = Command::CAppCommandManager::GetInstance();
			inst.Add(new Command::CDeleteCommand(oPlateSet));
		}
	}
}

void CFileView::OnPaint()
{
	CPaintDC dc(this); // device context for painting

	CRect rectTree;
	m_wndFileView.GetWindowRect(rectTree);
	ScreenToClient(rectTree);

	rectTree.InflateRect(1, 1);
	dc.Draw3dRect(rectTree, ::GetSysColor(COLOR_3DSHADOW), ::GetSysColor(COLOR_3DSHADOW));
}

void CFileView::OnSetFocus(CWnd* pOldWnd)
{
	CDockablePane::OnSetFocus(pOldWnd);

	m_wndFileView.SetFocus();
}

void CFileView::OnChangeVisualStyle()
{
	m_wndToolBar.CleanUpLockedImages();
	m_wndToolBar.LoadBitmap(theApp.m_bHiColorIcons ? IDB_EXPLORER_24 : IDR_EXPLORER, 0, 0, TRUE /* Locked */);

	m_FileViewImages.DeleteImageList();

	UINT uiBmpId = theApp.m_bHiColorIcons ? IDB_FILE_VIEW_24 : IDB_FILE_VIEW;

	CBitmap bmp;
	if (!bmp.LoadBitmap(uiBmpId))
	{
		TRACE(_T("Can't load bitmap: %x\n"), uiBmpId);
		ASSERT(FALSE);
		return;
	}

	BITMAP bmpObj;
	bmp.GetBitmap(&bmpObj);

	UINT nFlags = ILC_MASK;

	nFlags |= (theApp.m_bHiColorIcons) ? ILC_COLOR24 : ILC_COLOR4;

	m_FileViewImages.Create(16, bmpObj.bmHeight, nFlags, 0, 0);
	m_FileViewImages.Add(&bmp, RGB(255, 0, 255));

	m_wndFileView.SetImageList(&m_FileViewImages, TVSIL_NORMAL);
}

/**
	@brief	add sdnfelement to tree
	@author	humkyung
	@date	2013.05.29
*/
int CFileView::Add(CSDNFLinearMember* pElm)
{
	assert(pElm && "pElm is NULL");

	if(pElm)
	{
		CString str;

		CSDNFLinearMember::ElmType type = pElm->Type();
		OSTRINGSTREAM_T oss;
		oss << ++(m_oElmCountMap[type]);
		switch(type)
		{
			case CSDNFLinearMember::COLUMN:
			{
				HTREEITEM hItem = m_wndFileView.InsertItem(pElm->MemberID().c_str() , 3, 3, m_hColumn);
				if(NULL != hItem)
				{
					m_wndFileView.SetItemData(hItem , DWORD_PTR(pElm));
					str.Format(_T("Columns(%d)") , m_oElmCountMap[type]);
					m_wndFileView.SetItemText(m_hColumn , str);
				}
			}
				break;
			case CSDNFLinearMember::BEAM:
			{
				HTREEITEM hItem = m_wndFileView.InsertItem(pElm->MemberID().c_str() , 4, 4, m_hBeam);
				if(NULL != hItem)
				{
					m_wndFileView.SetItemData(hItem , DWORD_PTR(pElm));
					str.Format(_T("Beams(%d)") , m_oElmCountMap[type]);
					m_wndFileView.SetItemText(m_hBeam , str);
				}
			}
				break;
			case CSDNFLinearMember::HBRACE:
			{
				HTREEITEM hItem = m_wndFileView.InsertItem(pElm->MemberID().c_str() , 5, 5, m_hHBrace);
				if(NULL != hItem)
				{
					m_wndFileView.SetItemData(hItem , DWORD_PTR(pElm));
					str.Format(_T("HBraces(%d)") , m_oElmCountMap[type]);
					m_wndFileView.SetItemText(m_hHBrace , str);
				}
			}
				break;
			case CSDNFLinearMember::VBRACE:
			{
				HTREEITEM hItem = m_wndFileView.InsertItem(pElm->MemberID().c_str() , 5, 5, m_hVBrace);
				if(NULL != hItem)
				{
					m_wndFileView.SetItemData(hItem , DWORD_PTR(pElm));
					str.Format(_T("VBraces(%d)") , m_oElmCountMap[type]);
					m_wndFileView.SetItemText(m_hVBrace , str);
				}
			}
				break;
			default:
				break;
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	add gusset plate to tree
	@author	humkyung
	@date	2013.07.09
*/
int CFileView::Add(CGussetPlate* pGussetPlate)
{
	assert(pGussetPlate && "pGussetPlate is NULL");

	if(pGussetPlate)
	{
		int iChildCount = int(m_wndFileView.GetItemData(m_hGussetPlate));
		OSTRINGSTREAM_T oss;
		oss << _T("Gusset Plate - ") << ++iChildCount;
		HTREEITEM hItem = m_wndFileView.InsertItem(oss.str().c_str() , 6 , 10 , m_hGussetPlate);
		if(NULL != hItem)
		{
			m_wndFileView.SetItemData(hItem , DWORD_PTR(pGussetPlate));
			pGussetPlate->id() = iChildCount;
			if(722 == pGussetPlate->id())
			{
				int d = 0;
			}
		}

		CString str;
		str.Format(_T("Gusset Plates(%d)") , iChildCount);
		m_wndFileView.SetItemText(m_hGussetPlate , str);
		m_wndFileView.SetItemData(m_hGussetPlate , DWORD_PTR(iChildCount));

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	add end plate to tree

	@author	humkyung

	@date	2013.07.09
*/
int CFileView::Add(CEndPlate* pEndPlate)
{
	assert(pEndPlate && "pEndPlate is NULL");

	if(pEndPlate)
	{
		int iChildCount = int(m_wndFileView.GetItemData(m_hEndPlate));
		OSTRINGSTREAM_T oss;
		oss << _T("End Plate - ") << ++iChildCount;
		HTREEITEM hItem = m_wndFileView.InsertItem(oss.str().c_str() , 7 , 11 , m_hEndPlate);
		if(NULL != hItem)
		{
			m_wndFileView.SetItemData(hItem , DWORD_PTR(pEndPlate));
			pEndPlate->id() = iChildCount;
		}

		CString str;
		str.Format(_T("End Plates(%d)") , iChildCount);
		m_wndFileView.SetItemText(m_hEndPlate , str);
		m_wndFileView.SetItemData(m_hEndPlate , DWORD_PTR(iChildCount));

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	return child count of tree item

	@author	humkyung

	@date	2013.07.09
*/
int CFileView::GetChildCountOf(HTREEITEM hParent)
{
	int res = 0;

	HTREEITEM hChild = m_wndFileView.GetChildItem(hParent);
	while(hChild)
	{
		hChild = m_wndFileView.GetNextSiblingItem(hChild);
		++res;
	}

	return res;
}

/**
	@brief	find tree item which has given plate as item data
	@author	humkyung
*/
HTREEITEM CFileView::FindItemWith(CSteelPlate* pPlate)
{
	assert(pPlate && "pPlate is NULL");
	if(pPlate)
	{
		if(NULL != m_hGussetPlate)
		{
			HTREEITEM hChild = m_wndFileView.GetChildItem(m_hGussetPlate);
			while(hChild)
			{
				if(pPlate == (CSteelPlate*)(m_wndFileView.GetItemData(hChild))) return hChild;
				hChild =  m_wndFileView.GetNextSiblingItem(hChild);
			}
		}

		if(NULL != m_hEndPlate)
		{
			HTREEITEM hChild = m_wndFileView.GetChildItem(m_hEndPlate);
			while(hChild)
			{
				if(pPlate == (CSteelPlate*)(m_wndFileView.GetItemData(hChild))) return hChild;
				hChild =  m_wndFileView.GetNextSiblingItem(hChild);
			}
		}
	}

	return NULL;
}

/**
	@brief	select tree item which has given plate
	@author	humkyung
	@date	2013.07.30
*/
int CFileView::SelectItemWith(CSteelPlate* pPlate)
{
	assert(pPlate && "pPlate is NULL");
	if(pPlate)
	{
		if(NULL != m_hGussetPlate)
		{
			HTREEITEM hChild = m_wndFileView.GetChildItem(m_hGussetPlate);
			while(hChild)
			{
				if(pPlate == (CSteelPlate*)(m_wndFileView.GetItemData(hChild)))
				{
					m_wndFileView.SelectItem(hChild);
					return ERROR_SUCCESS;
				}
				hChild =  m_wndFileView.GetNextSiblingItem(hChild);
			}
		}

		if(NULL != m_hEndPlate)
		{
			HTREEITEM hChild = m_wndFileView.GetChildItem(m_hEndPlate);
			while(hChild)
			{
				if(pPlate == (CSteelPlate*)(m_wndFileView.GetItemData(hChild)))
				{
					m_wndFileView.SelectItem(hChild);
					return ERROR_SUCCESS;
				}
				hChild =  m_wndFileView.GetNextSiblingItem(hChild);
			}
		}
	}

	return ERROR_SUCCESS;
}

/**
	@brief	reset contents
	@author	humkyung
	@date	2013.08.02
*/
int CFileView::ResetContents()
{
	if(NULL != m_hRoot) m_wndFileView.DeleteItem(m_hRoot);
	if(NULL != m_hGussetPlate) m_wndFileView.DeleteItem(m_hGussetPlate);
	if(NULL != m_hGussetPlate) m_wndFileView.DeleteItem(m_hEndPlate);
	if(!m_oElmCountMap.empty())
	{
		for(map<CSDNFLinearMember::ElmType , int>::iterator itr = m_oElmCountMap.begin();itr != m_oElmCountMap.end();++itr)
		{
			itr->second = 0;
		}
	}
	FillFileView();

	return ERROR_SUCCESS;
}

/**
	@brief	reset plate contents
	@author	humkyung
	@date	2013.08.02
*/
int CFileView::ResetPlateContents()
{
	if(NULL != m_hGussetPlate) m_wndFileView.DeleteItem(m_hGussetPlate);
	if(NULL != m_hGussetPlate) m_wndFileView.DeleteItem(m_hEndPlate);
	{
		m_hGussetPlate = m_wndFileView.InsertItem(_T("Gusset Plates(0)"), 0, 0);
		m_wndFileView.SetItemData(m_hGussetPlate , 0);
		m_hEndPlate = m_wndFileView.InsertItem(_T("End Plates(0)"), 0, 0);
		m_wndFileView.SetItemData(m_hEndPlate , 0);
	}

	return ERROR_SUCCESS;
}