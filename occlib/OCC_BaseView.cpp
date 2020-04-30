// OCC_BaseView.cpp: implementation of the OCC_BaseView class.
//
//////////////////////////////////////////////////////////////////////

#include "Stdafx.h"
#include "OCC_BaseView.h"

IMPLEMENT_DYNCREATE(OCC_BaseView, CView)
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

OCC_BaseView::OCC_BaseView()
{
	m_pDC  = NULL;
}

OCC_BaseView::~OCC_BaseView()
{

}

void OCC_BaseView::OnDraw(CDC* pDC)
{
	CView::OnDraw(pDC);
}