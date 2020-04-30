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

// SmartSteelView.cpp : implementation of the CSmartSteelView class
//

#include "stdafx.h"
#include "SmartSteel.h"

#include "MainFrm.h"
#include "AppDocData.h"
#include "SmartSteelDoc.h"
#include "SmartSteelView.h"
#include "SDNFAttribute.h"

#include "Command/DeleteCommand.h"
#include "Command/AppCommandManager.h"

#include <TPrsStd_AISPresentation.hxx>
#include <TDF_Label.hxx>
#include <DNaming.hxx>
#include <TFunction_Function.hxx>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CSmartSteelView

IMPLEMENT_DYNCREATE(CSmartSteelView, OCC_3dView)

BEGIN_MESSAGE_MAP(CSmartSteelView, OCC_3dView)
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, &OCC_3dView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &OCC_3dView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CSmartSteelView::OnFilePrintPreview)
	ON_COMMAND(ID_VIEW_INFO  , &CSmartSteelView::OnViewInformation)
	ON_UPDATE_COMMAND_UI(ID_VIEW_INFO, &CSmartSteelView::OnUpdateViewInformation)
	ON_COMMAND(ID_VIEW_ROTATE  , &CSmartSteelView::OnViewRotate)
	ON_UPDATE_COMMAND_UI(ID_VIEW_ROTATE, &CSmartSteelView::OnUpdateViewRotate)
#ifdef SMART_STEEL
	ON_COMMAND(ID_ROTATE_ABOUT_AXIS  , &CSmartSteelView::OnViewRotateAboutAxis)
	ON_UPDATE_COMMAND_UI(ID_ROTATE_ABOUT_AXIS , &CSmartSteelView::OnUpdateViewRotateAboutAxis)
#endif
	ON_COMMAND(ID_VIEW_ZOOM , &CSmartSteelView::OnViewZoom)
	ON_UPDATE_COMMAND_UI(ID_VIEW_ZOOM, &CSmartSteelView::OnUpdateViewZoom)
	ON_COMMAND(ID_VIEW_FIT  , &CSmartSteelView::OnViewFit)
	ON_COMMAND(ID_OCC_VIEW_TOP , &CSmartSteelView::OnOCCViewTop)
	ON_COMMAND(ID_OCC_VIEW_BOTTOM , &CSmartSteelView::OnOCCViewBottom)
	ON_COMMAND(ID_OCC_VIEW_LEFT , &CSmartSteelView::OnOCCViewLeft)
	ON_COMMAND(ID_OCC_VIEW_RIGHT , &CSmartSteelView::OnOCCViewRight)
	ON_COMMAND(ID_OCC_VIEW_FRONT , &CSmartSteelView::OnOCCViewFront)
	ON_COMMAND(ID_OCC_VIEW_BACK , &CSmartSteelView::OnOCCViewBack)
	ON_COMMAND(ID_OCC_VIEW_ISO  , &CSmartSteelView::OnOCCViewISO)
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_RBUTTONUP()

	ON_COMMAND(ID_EDIT_DELETE, &CSmartSteelView::OnEditDelete)
	ON_WM_KEYDOWN()
END_MESSAGE_MAP()

// CSmartSteelView construction/destruction

CSmartSteelView::CSmartSteelView()/* : m_hSelectionSet(NULL)*/
{
}

CSmartSteelView::~CSmartSteelView()
{
	try
	{
		//SAFE_DELETE(m_hSelectionSet);
	}
	catch(...)
	{
	}
}

BOOL CSmartSteelView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return OCC_3dView::PreCreateWindow(cs);
}

BOOL CSmartSteelView::PreTranslateMessage(MSG* pMsg) 
{
    m_ctrlToolTip.RelayEvent(pMsg);
    return OCC_3dView::PreTranslateMessage(pMsg);
}

// CSmartSteelView drawing
void CSmartSteelView::OnDraw(CDC* pDC)
{
	CSmartSteelDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)	return;

	OCC_3dView::OnDraw(pDC);
}

// CSmartSteelView printing
void CSmartSteelView::OnFilePrintPreview()
{
	AFXPrintPreview(this);
}

BOOL CSmartSteelView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);
}

void CSmartSteelView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing
}

void CSmartSteelView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}

/**
	@brief	show tooltip when gusset or end plate is highlighted
	@author	humkyung
	@date	2014.09.05
*/
void CSmartSteelView::OnMouseMove(UINT nFlags, CPoint point)
{
	OCC_3dView::OnMouseMove(nFlags , point);

	if(GetDocument()->GetAISContext()->HasDetected())
	{
		Handle(AIS_InteractiveObject) ais = GetDocument()->GetAISContext()->DetectedInteractive();
		if(!ais.IsNull())
		{
			Handle_AIS_Shape aisShape = Handle_AIS_Shape::DownCast(ais);
			if(!aisShape.IsNull())
			{
				TopoDS_Shape aShape = aisShape->Shape();
				if(!aShape.IsNull())
				{
					CSteelPlate* pPlate = GetDocument()->GetPlateFrom(aShape);
					if(NULL != pPlate)
					{
						m_ctrlToolTip.Activate(TRUE);
						STRINGSTREAM_T oss;
						oss << _T("Type=") << pPlate->GetTypeString() << _T("\nID=") << pPlate->id();
						m_ctrlToolTip.UpdateTipText(oss.str().c_str() , this);
						return;
					}
				}
			}
		}
	}
	m_ctrlToolTip.Activate(FALSE);
}

void CSmartSteelView::OnRButtonUp(UINT nFlags, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

/**
	@brief	show context menu
	@author	humkyung
	@date	
*/
void CSmartSteelView::OnContextMenu(CWnd* pWnd, CPoint point)
{
	if(CurAction3d_Nothing == m_CurrentMode)
	{
		vector<TopoDS_Shape> shapeList;
		if(m_hSelectionSet->GetSelected(shapeList))
		{
			theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EXPLORER, point.x, point.y, this, TRUE);
		}
	}
}

// CSmartSteelView diagnostics

#ifdef _DEBUG
void CSmartSteelView::AssertValid() const
{
	OCC_3dView::AssertValid();
}

void CSmartSteelView::Dump(CDumpContext& dc) const
{
	OCC_3dView::Dump(dc);
}

CSmartSteelDoc* CSmartSteelView::GetDocument() const // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CSmartSteelDoc)));
	return (CSmartSteelDoc*)m_pDocument;
}
#endif //_DEBUG


// CSmartSteelView message handlers

/**
	@brief	view information
	@author	humkyung
	@date	2013.08.18
*/
void CSmartSteelView::OnViewInformation()
{
	m_CurrentMode = CurAction3d_Nothing;
}

/**
	@brief	set check if current mode is CurAction3d_Nothing
	@author	humkyung
	@date	2013.08.18
*/
void CSmartSteelView::OnUpdateViewInformation(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(CurAction3d_Nothing == m_CurrentMode);
}

void CSmartSteelView::OnViewRotate()
{
	OCC_3dView::OnBUTTONRot();
}

void CSmartSteelView::OnUpdateViewRotate(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(CurAction3d_DynamicRotation == m_CurrentMode);
}

#ifdef SMART_STEEL
void CSmartSteelView::OnViewRotateAboutAxis()
{
	m_CurrentMode = CurAction3d_RotationAboutAxis;
}

void CSmartSteelView::OnUpdateViewRotateAboutAxis(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(CurAction3d_RotationAboutAxis == m_CurrentMode);
}
#endif

void CSmartSteelView::OnViewZoom()
{
	OCC_3dView::OnBUTTONZoomWin();
}

void CSmartSteelView::OnUpdateViewZoom(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(CurAction3d_WindowZooming == m_CurrentMode);
}

void CSmartSteelView::OnViewFit()
{
	OCC_3dView::OnBUTTONZoomAll();
}

void CSmartSteelView::OnOCCViewTop()
{
	OCC_3dView::OnBUTTONTop();
}

void CSmartSteelView::OnOCCViewBottom()
{
	OCC_3dView::OnBUTTONBottom();
}

void CSmartSteelView::OnOCCViewLeft()
{
	OCC_3dView::OnBUTTONLeft();
}

void CSmartSteelView::OnOCCViewRight()
{
	OCC_3dView::OnBUTTONRight();
}

void CSmartSteelView::OnOCCViewFront()
{
	OCC_3dView::OnBUTTONFront();
}

void CSmartSteelView::OnOCCViewBack()
{
	OCC_3dView::OnBUTTONBack();
}

void CSmartSteelView::OnOCCViewISO()
{
	OCC_3dView::OnBUTTONAxo();
}

/**
	@brief
	@author	
*/
void CSmartSteelView::OnLButtonUp(UINT nFlags, CPoint point)
{
	CurAction3d SavedMode = m_CurrentMode;

	OCC_3dView::OnLButtonUp(nFlags, point);
	if(CurAction3d_Nothing == m_CurrentMode)
	{
		if(GetDocument()->GetAISContext()->HasDetected())
		{
			Handle(AIS_InteractiveObject) ais = GetDocument()->GetAISContext()->DetectedInteractive();
			if(!ais.IsNull())
			{
				Handle_AIS_Shape aisShape = Handle_AIS_Shape::DownCast(ais);
				if(!aisShape.IsNull())
				{
					CSmartSteelDoc* pDoc = GetDocument();

					TopoDS_Shape aShape = aisShape->Shape();
					if(!aShape.IsNull())
					{
						OCC::COCCEntity* pEnt = pDoc->GetOCCEntityFrom(aShape);
						if(NULL != pEnt)
						{
							CSDNFAttribute* pAttr = static_cast<CSDNFAttribute*>(pEnt->GetAttributeAt(0));
							if(NULL != pAttr)
							{
								CMainFrame* pFrame = (CMainFrame*)(AfxGetApp()->GetMainWnd());
								if(NULL != pFrame) pFrame->FillPropertiesOf(pAttr);
							}
						}
					}
					else
					{
						CSteelPlate* pPlate = pDoc->GetPlateFrom(aShape);
						if(NULL != pPlate)
						{
							CMainFrame* pFrame = CMainFrame::GetInstance();
							CFileView& oFileView = pFrame->GetFileView();
							oFileView.SelectItemWith(pPlate);
							if(NULL != pFrame) pFrame->FillPropertiesOf(pPlate);
						}
					}
				}
			}
		}
		/*
		vector<TopoDS_Shape> shapeList;
		if(m_hSelectionSet->GetSelected(shapeList))
		{
			CSmartSteelDoc* pDoc = GetDocument();
			OCC::COCCEntity* pEnt = pDoc->GetOCCEntityFrom(shapeList.back());
			if(NULL != pEnt)
			{
				CSDNFAttribute* pAttr = static_cast<CSDNFAttribute*>(pEnt->GetAttributeAt(0));
				if(NULL != pAttr)
				{
					CMainFrame* pFrame = (CMainFrame*)(AfxGetApp()->GetMainWnd());
					if(NULL != pFrame) pFrame->FillPropertiesOf(pAttr);

				}
			}
			else
			{
				CSteelPlate* pPlate = pDoc->GetPlateFrom(shapeList.back());
				if(NULL != pPlate)
				{
					CMainFrame* pFrame = CMainFrame::GetInstance();
					CFileView& oFileView = pFrame->GetFileView();
					oFileView.SelectItemWith(pPlate);
					if(NULL != pFrame) pFrame->FillPropertiesOf(pPlate);
				}
			}
		}
		*/
	}
	
	if((CurAction3d_DynamicRotation == SavedMode) || (CurAction3d_WindowZooming == SavedMode)) m_CurrentMode = SavedMode;
}

void CSmartSteelView::OnInitialUpdate()
{
	OCC_3dView::OnInitialUpdate();

	m_hSelectionSet = tr1::shared_ptr<CSelectionSet>(new CSelectionSet(GetDocument()->GetAISContext() , m_hView));
#ifdef	SMART_STEEL
	CAppDocData& docData = CAppDocData::GetInstance();
	if(_T("DEMO") != docData.GetProjectName())
	{
		m_hView->SetBgGradientColors(Quantity_NOC_BLACK , Quantity_NOC_MATRAGRAY , Aspect_GFM_DIAG1 , Standard_False);
	}
	else
	{
		m_hView->SetBackgroundImage(GetExecPath() + _T("\\demo.bmp") , Aspect_FM_CENTERED  , Standard_False);
	}
#else
	m_hView->SetBgGradientColors(Quantity_NOC_BLACK , Quantity_NOC_MATRAGRAY , Aspect_GFM_DIAG1 , Standard_False);
#endif

	EnableToolTips(TRUE);
	if(TRUE == m_ctrlToolTip.Create(this))
	{
		m_ctrlToolTip.Activate(TRUE);
		m_ctrlToolTip.AddTool(this,  (LPCTSTR)_T(""));
		m_ctrlToolTip.SetMaxTipWidth(SHRT_MAX);
		m_ctrlToolTip.SendMessage(TTM_SETDELAYTIME, TTDT_RESHOW, 10000);
	}
}

/**
	@brief	delete selected steel plate
	@author	humkyung
	@date	2013.08.03
*/
void CSmartSteelView::OnEditDelete()
{
	vector<TopoDS_Shape> shapeList;
	if(m_hSelectionSet->GetSelected(shapeList))
	{
		vector<CSteelPlate*> oPlateSet;
		for(vector<TopoDS_Shape>::iterator itr = shapeList.begin();itr != shapeList.end();++itr)
		{
			CSteelPlate* pPlate = GetDocument()->GetPlateFrom(*itr);
			if(NULL != pPlate) oPlateSet.push_back( pPlate );
		}
		
		Command::CAppCommandManager& inst = Command::CAppCommandManager::GetInstance();
		inst.Add(new Command::CDeleteCommand(oPlateSet));
	}
}

/**
	@brief	
	@author	humkyung
	@date	2013.08.06
*/
void CSmartSteelView::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if(VK_ESCAPE == nChar)
	{
		m_CurrentMode = CurAction3d_Nothing;
	}

	OCC_3dView::OnKeyDown(nChar, nRepCnt, nFlags);
}