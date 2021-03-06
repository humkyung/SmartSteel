// OCC_3dView.h: interface for the OCC_3dView class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_OCC_3DVIEW_H__1F4065AE_39C4_11D7_8611_0060B0EE281E__INCLUDED_)
#define AFX_OCC_3DVIEW_H__1F4065AE_39C4_11D7_8611_0060B0EE281E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

using namespace std;

#include <Standard_Macro.hxx>
#include <Visual3d_Layer.hxx>
///#include "ColorBar.h"
#include "OCC_BaseView.h"
#include "OCC_3dBaseDoc.h"
#include "jxglTracker.h"

enum CurAction3d 
{ 
	CurAction3d_Nothing,
	CurAction3d_DynamicZooming,
	CurAction3d_WindowZooming,
	CurAction3d_DynamicPanning,
	CurAction3d_GlobalPanning,
	CurAction3d_DynamicRotation,
	CurAction3d_RotationAboutAxis
};

class AFX_EXT_CLASS OCC_3dView : public OCC_BaseView  
{
	DECLARE_DYNCREATE(OCC_3dView)
public:
	OCC_3dView();
	virtual ~OCC_3dView();

	static BOOL IsAeroEnabled();	/// 2014.08.24 added by humkyung
	OCC_3dBaseDoc* GetDocument();
	
	Standard_Boolean Convert2dPntTo3dPnt(const Standard_Integer aX2d, const Standard_Integer aY2d, gp_Pnt& a3dPoint);
	int SetBackgroundColor(const unsigned char& r , const unsigned char& g , const unsigned char& b);	/// 2011.12.09 added by humkyung
	
	void ZoomWin(const Bnd_Box& oBndBox);	/// 2011.12.04 added by humkyung
	void ZoomWin(){m_CurrentMode = CurAction3d_WindowZooming;};
	void FitAll() {   m_hView->FitAll();  m_hView->ZFitAll();  };
	void Redraw() {   m_hView->Redraw(); };

	void SetZoom ( const Quantity_Factor& Coef  ) {   m_hView->SetZoom ( Coef  );  };

	// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(OCC_3dView)
public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual void OnInitialUpdate();
	//}}AFX_VIRTUAL

	// Generated message map functions
protected:
	//{{AFX_MSG(OCC_3dView)
	afx_msg void OnBUTTONAxo();
	afx_msg void OnBUTTONBack();
	afx_msg void OnBUTTONBottom();
	afx_msg void OnBUTTONFront();
	afx_msg void OnBUTTONHlrOff();
	afx_msg void OnBUTTONHlrOn();
	afx_msg void OnBUTTONLeft();
	afx_msg void OnBUTTONPan();
	afx_msg void OnBUTTONPanGlo();
	afx_msg void OnBUTTONReset();
	afx_msg void OnBUTTONRight();
	afx_msg void OnBUTTONRot();
	afx_msg void OnBUTTONTop();
	afx_msg void OnBUTTONZoomAll();
	afx_msg void OnFileExportImage();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnBUTTONZoomProg();
	afx_msg void OnBUTTONZoomWin();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnUpdateBUTTONHlrOff(CCmdUI* pCmdUI);
	afx_msg void OnUpdateBUTTONHlrOn(CCmdUI* pCmdUI);
	afx_msg void OnUpdateBUTTONPanGlo(CCmdUI* pCmdUI);
	afx_msg void OnUpdateBUTTONPan(CCmdUI* pCmdUI);
	afx_msg void OnUpdateBUTTONZoomProg(CCmdUI* pCmdUI);
	afx_msg void OnUpdateBUTTONZoomWin(CCmdUI* pCmdUI);
	afx_msg void OnUpdateBUTTONRot(CCmdUI* pCmdUI);
	afx_msg void OnModifyChangeBackground();
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:
	bool m_bInitialized;
	Handle_V3d_View m_hView;
#ifdef WNT
	Handle_WNT_Window m_hWindow; 
#else
	Handle_Xw_Window  m_hWindow;
#endif
	CurAction3d			m_CurrentMode , m_PrevMode;
	Standard_Integer	m_iPrevPosX;
	Standard_Integer	m_iPrevPosY;  
	Standard_Integer	m_iCurrentPosX;
	Standard_Integer	m_iCurrentPosY;
	Quantity_Factor		m_dCurZoom;
	Standard_Boolean	myDegenerateModeIsOn;
	Standard_Integer	myWidth;
	Standard_Integer   	myHeight;

	Handle(Visual3d_Layer) m_hGradientBkgndLayer;	/// layer for gradient background - 2011.09.16	added by humkyung
	Handle(Visual3d_Layer) m_hColorBarLayer;		/// layer for ColorBar - 2011.09.16	added by humkyung
	///OCC::CColorBar*	m_oColorBar;	/// 2011.09.16	added by humkyung

	jxglTracker m_tracker;
public:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);	
private:
	bool m_bDrawBackground;

	int DrawGradientBackground(void);
	int DrawColorBar(void);
};

#ifndef _DEBUG  // debug version in OCC_3dView.cpp
inline OCC_3dBaseDoc* OCC_3dView::GetDocument()
{ return (OCC_3dBaseDoc*)m_pDocument; }
#endif

#endif // !defined(AFX_OCC_3DVIEW_H__1F4065AE_39C4_11D7_8611_0060B0EE281E__INCLUDED_)
