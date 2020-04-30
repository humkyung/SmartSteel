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

// occlibTestView.cpp : implementation of the CocclibTestView class
//

#include "stdafx.h"
#include "occlibTest.h"

#include "occlibTestDoc.h"
#include "occlibTestView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CocclibTestView

IMPLEMENT_DYNCREATE(CocclibTestView, OCC_3dView)

BEGIN_MESSAGE_MAP(CocclibTestView, OCC_3dView)
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, &OCC_3dView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &OCC_3dView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CocclibTestView::OnFilePrintPreview)
	ON_COMMAND(ID_OCC_ROT  , &CocclibTestView::OnOCCRot)
	ON_COMMAND(ID_OCC_ZOOM , &CocclibTestView::OnOCCZoom)
	ON_COMMAND(ID_OCC_FIT  , &CocclibTestView::OnOCCFit)
END_MESSAGE_MAP()

// CocclibTestView construction/destruction

CocclibTestView::CocclibTestView()
{
	// TODO: add construction code here
}

CocclibTestView::~CocclibTestView()
{
}

BOOL CocclibTestView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return OCC_3dView::PreCreateWindow(cs);
}

// CocclibTestView drawing

void CocclibTestView::OnDraw(CDC* pDC)
{
	CocclibTestDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)
		return;

	OCC_3dView::OnDraw(pDC);
}


// CocclibTestView printing


void CocclibTestView::OnFilePrintPreview()
{
	AFXPrintPreview(this);
}

void CocclibTestView::OnOCCRot()
{
	OCC_3dView::OnBUTTONRot();
}

void CocclibTestView::OnOCCZoom()
{
	OCC_3dView::OnBUTTONZoomWin();
}

void CocclibTestView::OnOCCFit()
{
	OCC_3dView::OnBUTTONZoomAll();
}

BOOL CocclibTestView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);
}

void CocclibTestView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing
}

void CocclibTestView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}

void CocclibTestView::OnRButtonUp(UINT nFlags, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void CocclibTestView::OnContextMenu(CWnd* pWnd, CPoint point)
{
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
}


// CocclibTestView diagnostics

#ifdef _DEBUG
void CocclibTestView::AssertValid() const
{
	OCC_3dView::AssertValid();
}

void CocclibTestView::Dump(CDumpContext& dc) const
{
	OCC_3dView::Dump(dc);
}

CocclibTestDoc* CocclibTestView::GetDocument() const // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CocclibTestDoc)));
	return (CocclibTestDoc*)m_pDocument;
}
#endif //_DEBUG


// CocclibTestView message handlers
