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

#pragma once

#include <SDNFFile.h>
#include <SDNFLinearMember.h>
#include "ViewTree.h"
#include "GussetPlate.h"
#include "EndPlate.h"

#include <map>
using namespace std;

class CFileViewToolBar : public CMFCToolBar
{
	virtual void OnUpdateCmdUI(CFrameWnd* /*pTarget*/, BOOL bDisableIfNoHndler)
	{
		CMFCToolBar::OnUpdateCmdUI((CFrameWnd*) GetOwner(), bDisableIfNoHndler);
	}

	virtual BOOL AllowShowOnList() const { return FALSE; }
};

class CFileView : public CDockablePane
{
// Construction
public:
	CFileView();

	void AdjustLayout();
	void OnChangeVisualStyle();

	/// Reset contents
	int ResetContents();
	int ResetPlateContents();

	int SelectItemWith(CSteelPlate* pPlate);
	HTREEITEM FindItemWith(CSteelPlate* pPlate);
	CViewTree& GetViewTree(){ return m_wndFileView;}
// Attributes
protected:
	CViewTree m_wndFileView;
	CImageList m_FileViewImages;
	CFileViewToolBar m_wndToolBar;

	HTREEITEM m_hRoot , m_hColumn , m_hBeam , m_hHBrace , m_hVBrace;
	HTREEITEM m_hGussetPlate , m_hEndPlate;
	map<CSDNFLinearMember::ElmType , int> m_oElmCountMap;
protected:
	///@brief	return child count of given item
	int GetChildCountOf(HTREEITEM hParent);
	void FillFileView();
	
// Implementation
public:
	virtual ~CFileView();
	int Add(CSDNFLinearMember* pElm);
	int Add(CGussetPlate* pGussetPlate);
	int Add(CEndPlate* pEndPlate);
protected:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	afx_msg void OnProperties();
	afx_msg void OnFileOpen();
	afx_msg void OnFileOpenWith();
	afx_msg void OnDummyCompile();
	afx_msg void OnEditCut();
	afx_msg void OnEditCopy();
	afx_msg void OnEditDelete();
	afx_msg void OnPaint();
	afx_msg void OnSetFocus(CWnd* pOldWnd);

	DECLARE_MESSAGE_MAP()
};

