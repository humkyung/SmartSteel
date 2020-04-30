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

// SmartSteelView.h : interface of the CSmartSteelView class
//

#pragma once
#include <OCC_3dView.h>
#include "SmartSteelDoc.h"
#include "SelectionSet.h"
#include <memory>

class CSmartSteelView : public OCC_3dView
{
protected: // create from serialization only
	CSmartSteelView();
	DECLARE_DYNCREATE(CSmartSteelView)

// Attributes
public:
	CSmartSteelDoc* GetDocument() const;

// Operations
public:

// Overrides
public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// Implementation
public:
	virtual ~CSmartSteelView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:
#if _MSC_VER >= 1700
	std::shared_ptr<CSelectionSet> m_hSelectionSet;
#else
	tr1::shared_ptr<CSelectionSet> m_hSelectionSet;
#endif
	CToolTipCtrl m_ctrlToolTip;	/// 2014.09.05 added by humkyung
// Generated message map functions
protected:
	afx_msg void OnFilePrintPreview();
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	afx_msg void OnViewInformation();
	afx_msg void OnUpdateViewInformation(CCmdUI* pCmdUI);
	afx_msg void OnViewRotate();
	afx_msg void OnUpdateViewRotate(CCmdUI* pCmdUI);
#ifdef SMART_STEEL
	afx_msg void OnViewRotateAboutAxis();
	afx_msg void OnUpdateViewRotateAboutAxis(CCmdUI* pCmdUI);
#endif
	afx_msg void OnViewZoom();
	afx_msg void OnUpdateViewZoom(CCmdUI* pCmdUI);
	afx_msg void OnViewFit();
	afx_msg void OnOCCViewTop();
	afx_msg void OnOCCViewBottom();
	afx_msg void OnOCCViewLeft();
	afx_msg void OnOCCViewRight();
	afx_msg void OnOCCViewFront();
	afx_msg void OnOCCViewBack();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	virtual void OnInitialUpdate();
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnEditDelete();
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnOCCViewISO();
};

#ifndef _DEBUG  // debug version in SmartSteelView.cpp
inline CSmartSteelDoc* CSmartSteelView::GetDocument() const
   { return reinterpret_cast<CSmartSteelDoc*>(m_pDocument); }
#endif

