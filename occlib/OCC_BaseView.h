// OCC_BaseView.h: interface for the OCC_BaseView class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_OCC_BASEVIEW_H__2E048CCA_38F9_11D7_8611_0060B0EE281E__INCLUDED_)
#define AFX_OCC_BASEVIEW_H__2E048CCA_38F9_11D7_8611_0060B0EE281E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

//#include "Stdafx.h"

class AFX_EXT_CLASS OCC_BaseView  : public CView
{
	DECLARE_DYNCREATE(OCC_BaseView)
public:
	OCC_BaseView();
	virtual ~OCC_BaseView();
	
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
protected:
	CDC*  m_pDC;

	Standard_Real        m_dWidth; 
	Standard_Real        m_dHeight; 
	Standard_Integer     m_Atx  , m_Aty  , m_Atz  ;//
	Standard_Integer     m_Eyex , m_Eyey , m_Eyez ;//
};

#endif // !defined(AFX_OCC_BASEVIEW_H__2E048CCA_38F9_11D7_8611_0060B0EE281E__INCLUDED_)
