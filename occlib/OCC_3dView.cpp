// OCC_3dView.cpp: implementation of the OCC_3dView class.
//

#include "stdafx.h"
#include <gl/gl.h>
#include <gl\glu.h>			// Header File For The GLu32 Library
//#include <gl\glaux.h>		// Header File For The Glaux Library
#include <IsTools.h>
#include <Resource.h>
#include "OCC_3dView.h"

#include "OCC_3dApp.h"
#include "OCC_3dBaseDoc.h"

#include <Visual3d_View.hxx>
#include <Graphic3d_ExportFormat.hxx>
#include <BRepPrimAPI_MakeSphere.hxx>
#include <AIS_Trihedron.hxx>
#include <Geom_Axis2Placement.hxx>
#include <Visual3d_View.hxx>

#ifndef  _ElSLib_HeaderFile
#include <ElSLib.hxx>
#endif
#ifndef  _ProjLib_HeaderFile
#include <ProjLib.hxx>
#endif
//#include <Graphic3d_NameOfFont.hxx>

#define ValZWMin 1

IMPLEMENT_DYNCREATE(OCC_3dView, OCC_BaseView)

BEGIN_MESSAGE_MAP(OCC_3dView, OCC_BaseView)
	//{{AFX_MSG_MAP(OCC_3dView)
	/*ON_COMMAND(ID_BUTTONAxo, OnBUTTONAxo)
	ON_COMMAND(ID_BUTTONBack, OnBUTTONBack)
	ON_COMMAND(ID_BUTTONBottom, OnBUTTONBottom)
	ON_COMMAND(ID_BUTTONFront, OnBUTTONFront)
	ON_COMMAND(ID_BUTTONHlrOff, OnBUTTONHlrOff)
	ON_COMMAND(ID_BUTTONHlrOn, OnBUTTONHlrOn)
	ON_COMMAND(ID_BUTTONLeft, OnBUTTONLeft)
	ON_COMMAND(ID_BUTTONPan, OnBUTTONPan)
	ON_COMMAND(ID_BUTTONPanGlo, OnBUTTONPanGlo)
	ON_COMMAND(ID_BUTTONReset, OnBUTTONReset)
	ON_COMMAND(ID_BUTTONRight, OnBUTTONRight)
	ON_COMMAND(ID_BUTTONRot, OnBUTTONRot)
	ON_COMMAND(ID_BUTTONTop, OnBUTTONTop)
	ON_COMMAND(ID_BUTTONZoomAll, OnBUTTONZoomAll)*/
	ON_WM_SIZE()
	///ON_COMMAND(ID_FILE_EXPORT_IMAGE, OnFileExportImage)
	/*ON_COMMAND(ID_BUTTONZoomProg, OnBUTTONZoomProg)
	ON_COMMAND(ID_BUTTONZoomWin, OnBUTTONZoomWin)*/
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MBUTTONDOWN()
	ON_WM_MBUTTONUP()
	ON_WM_MOUSEMOVE()
	ON_WM_RBUTTONDOWN()
	ON_WM_RBUTTONUP()
	/*ON_UPDATE_COMMAND_UI(ID_BUTTONHlrOff, OnUpdateBUTTONHlrOff)
	ON_UPDATE_COMMAND_UI(ID_BUTTONHlrOn, OnUpdateBUTTONHlrOn)
	ON_UPDATE_COMMAND_UI(ID_BUTTONPanGlo, OnUpdateBUTTONPanGlo)
	ON_UPDATE_COMMAND_UI(ID_BUTTONPan, OnUpdateBUTTONPan)
	ON_UPDATE_COMMAND_UI(ID_BUTTONZoomProg, OnUpdateBUTTONZoomProg)
	ON_UPDATE_COMMAND_UI(ID_BUTTONZoomWin, OnUpdateBUTTONZoomWin)
	ON_UPDATE_COMMAND_UI(ID_BUTTONRot, OnUpdateBUTTONRot)
	ON_COMMAND(ID_Modify_ChangeBackground     , OnModifyChangeBackground)*/
	//}}AFX_MSG_MAP
	ON_WM_CREATE()
	ON_WM_MOUSEWHEEL()
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// OCC_3dView construction/destruction

OCC_3dView::OCC_3dView() : m_bInitialized(false) , m_bDrawBackground(false) , m_hView(NULL)/// , m_oColorBar(NULL)
{
	// TODO: add construction code here
	m_iPrevPosX=0;
	m_iPrevPosY=0;  
	m_iCurrentPosX=0;
	m_iCurrentPosY=0;
	m_dCurZoom=0;
	myWidth=0;
	myHeight=0;
	/// will be set in OnInitial update, but, for more security :
	m_CurrentMode = m_PrevMode = CurAction3d_Nothing;
	myDegenerateModeIsOn=Standard_True;
	//m_Pen = NULL;
}

OCC_3dView::~OCC_3dView()
{
	try
	{
		m_hView->Remove();
		/*if (m_Pen != NULL)
		{
			delete m_Pen;
			m_Pen = NULL;
		}*/

		if(NULL != m_pDC)
		{
			delete m_pDC;
			m_pDC = NULL;
		}
	}
	catch(...)
	{
	}

	//if (m_Pen) delete m_Pen;
}

BOOL OCC_3dView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return OCC_BaseView::PreCreateWindow(cs);
}

//Like this...
//
//void C+ProjectName+Doc::DisplayGrid(gp_Pln aPlane, Quantity_Length XOrigin, Quantity_Length YOrigin, Quantity_Length XStep, Quantity_Length YStep, int Type, Quantity_Length Rotation, Quantity_Length XSize, Quantity_Length YSize, Quantity_Length Offset)
//{
//     gp_Ax3 Ax3(aPlane.Location(), aPlane.Axis().Direction());
//     myViewer->SetPrivilegedPlane(Ax3);
//     myViewer->SetRectangularGridValues(XOrigin, YOrigin, XStep, YStep, Rotation);
//     myViewer->SetRectangularGridGraphicValues(XSize, YSize, Offset);
//     myViewer->ActivateGrid(Aspect_GT_Rectangular, Aspect_GDM_Lines);
//}

/////////////////////////////////////////////////////////////////////////////
// OCC_3dView drawing
void OCC_3dView::OnInitialUpdate() 
{
	OCC_BaseView::OnInitialUpdate();

	if(false == m_bInitialized)
	{
		m_pDC = new CClientDC(this);
		HGLRC hRC = ::wglCreateContext(m_pDC->GetSafeHdc());
		::wglMakeCurrent(m_pDC->GetSafeHdc(), hRC);
		///m_nDefaultBitmapFontID = GenerateBitmapFont(NULL);
		::wglMakeCurrent(NULL, NULL);

		m_hView = GetDocument()->GetViewer()->CreateView();

		// set the default mode in wireframe ( not hidden line ! )
		//m_hView->SetDegenerateModeOn();	: removed at OCC ver 6.6.0
		// store for restore state after rotation (witch is in Degenerated mode)
		myDegenerateModeIsOn = Standard_True;
		// Creating new Window and attaching it to View
		m_hView->SetProj(V3d_Zpos);
#ifdef WNT
		m_hWindow = new WNT_Window(/*theGraphicDevice , */GetSafeHwnd());
#else
		m_hWindow = new Xw_Window(pApp->GetGraphicDevice() , winId());
#endif
		m_hView->SetWindow(m_hWindow);
		if (!m_hWindow->IsMapped()) m_hWindow->Map();

		// store the mode ( nothing , dynamic zooming, dynamic ... )
		m_CurrentMode = CurAction3d_DynamicRotation;
		OnBUTTONTop();

		///GetDocument()->GetViewer()->ActivateGrid(Aspect_GT_Rectangular , Aspect_GDM_Lines);
		///GetDocument()->GetViewer()->SetGridColor(Quantity_Color(Quantity_NOC_WHITE), Quantity_Color(Quantity_NOC_WHITE));
		
		/// 2011.09.16 added by humkyung
		///m_oColorBar = new OCC::CColorBar(m_hView);
		/// up to here

		m_hView->ZBufferTriedronSetup();///Quantity_NOC_RED,Quantity_NOC_GREEN,Quantity_NOC_BLUE1,0.5,0.025,6);
		m_hView->TriedronDisplay(Aspect_TOTP_LEFT_LOWER,Quantity_NOC_WHITE,0.05,V3d_ZBUFFER);
		m_hView->Redraw();

		m_bInitialized = true;
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-08-23
    @class      OCC_3dView
    @function   OnDraw
    @return     void
    @param      CDC*    pDC
    @brief
******************************************************************************/
void OCC_3dView::OnDraw(CDC* pDC)
{
	CRect aRect;
	GetWindowRect(aRect);
	if(m_dWidth != aRect.Width() || m_dHeight != aRect.Height()) 
	{
		m_dWidth = aRect.Width();
		m_dHeight = aRect.Height();
		///::PostMessage ( GetSafeHwnd () , WM_SIZE , SW_SHOW , m_dWidth + m_dHeight*65536 );
	}
	
	if(true == m_bInitialized)
	{
		if(true == m_bDrawBackground)
		{
			DrawGradientBackground();
		}

		m_hView->Redraw();
	}
}

/////////////////////////////////////////////////////////////////////////////
// OCC_3dView diagnostics

#ifdef _DEBUG
void OCC_3dView::AssertValid() const
{
	OCC_BaseView::AssertValid();
}

void OCC_3dView::Dump(CDumpContext& dc) const
{
	OCC_BaseView::Dump(dc);
}

OCC_3dBaseDoc* OCC_3dView::GetDocument() // non-debug version is inline
{
	//	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(OCC_3dBaseDoc)));
	return (OCC_3dBaseDoc*)m_pDocument;
}

#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// OCC_3dView message handlers
void OCC_3dView::OnFileExportImage()
{
	LPCTSTR filter;
#ifdef WNT
	filter = _T("BMP Files (*.BMP)|*.bmp|GIF Files (*.GIF)|*.gif|XWD Files (*.XWD)|*.xwd|PS Files (*.PS)|*.ps|EMF Files (*.EMF)|*.emf||");
#else
	filter = _T("BMP Files (*.BMP)|*.bmp|GIF Files (*.GIF)|*.gif|XWD Files (*.XWD)|*.xwd|PS Files (*.PS)|*.ps||");
#endif //WNT
	CFileDialog dlg(FALSE,_T("*.BMP"),NULL,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		filter, 
		NULL );

	if (IDOK == dlg.DoModal()) 
	{
		SetCursor(AfxGetApp()->LoadStandardCursor(IDC_WAIT));
		CString filename = dlg.GetPathName();
		TCHAR* theFile = new TCHAR[filename.GetLength()+1];
		_tcscpy(theFile,filename);
		CString ext = dlg.GetFileExt();
		if (ext == _T("ps") || ext == _T("emf"))
		{
			Graphic3d_ExportFormat exFormat;
			if (ext == _T("ps")) exFormat = Graphic3d_EF_PostScript;
			else             exFormat = Graphic3d_EF_EMF;
			m_hView->View()->Export( CStringA(theFile) , exFormat );
			return;
		}
		/*Handle(Aspect_Window) anAspectWindow = m_hView->Window();
		Handle(WNT_Window) aWNTWindow = Handle(WNT_Window)::DownCast(anAspectWindow);
		if (ext == _T("bmp"))     aWNTWindow->SetOutputFormat ( WNT_TOI_BMP );
		if (ext == _T("gif"))     aWNTWindow->SetOutputFormat ( WNT_TOI_GIF );
		if (ext == _T("xwd"))     aWNTWindow->SetOutputFormat ( WNT_TOI_XWD );
		aWNTWindow->Dump ((Standard_CString)(LPCTSTR)filename);
		SetCursor(AfxGetApp()->LoadStandardCursor(IDC_ARROW));*/
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-08-23
    @class      OCC_3dView
    @function   OnSize
    @return     void
    @param      UINT    nType
    @param      int     cx
    @param      int     cy
    @brief
******************************************************************************/
void OCC_3dView::OnSize(UINT nType, int cx, int cy) 
{
	OCC_BaseView::OnSize(nType , cx , cy);

	m_dWidth = (double)cx;
	m_dHeight= (double)cy;
	if (!m_hView.IsNull())
	{
		m_hView->MustBeResized();
		m_hView->Size(m_dWidth, m_dHeight);

		/// 왜 여기서 호출해야 제대로 될까?
		///DrawColorBar();
	}
}

void OCC_3dView::OnBUTTONBack() 
{ m_hView->SetProj(V3d_Xneg); } // See the back View
void OCC_3dView::OnBUTTONFront() 
{ m_hView->SetProj(V3d_Xpos); } // See the front View

void OCC_3dView::OnBUTTONBottom() 
{ m_hView->SetProj(V3d_Zneg); } // See the bottom View
void OCC_3dView::OnBUTTONTop() 
{ m_hView->SetProj(V3d_Zpos); } // See the top View	

void OCC_3dView::OnBUTTONLeft() 
{ m_hView->SetProj(V3d_Ypos); } // See the left View	
void OCC_3dView::OnBUTTONRight() 
{ m_hView->SetProj(V3d_Yneg); } // See the right View

void OCC_3dView::OnBUTTONAxo() 
{ m_hView->SetProj(V3d_XposYnegZpos); } // See the axonometric View

void OCC_3dView::OnBUTTONHlrOff() 
{
	//m_hView->SetDegenerateModeOn();
	myDegenerateModeIsOn = Standard_True;
}

void OCC_3dView::OnBUTTONHlrOn() 
{
	SetCursor(AfxGetApp()->LoadStandardCursor(IDC_WAIT));
	//m_hView->SetDegenerateModeOff();
	myDegenerateModeIsOn = Standard_False;
	SetCursor(AfxGetApp()->LoadStandardCursor(IDC_ARROW));
}

void OCC_3dView::OnBUTTONPan() 
{  m_CurrentMode = CurAction3d_DynamicPanning; }

void OCC_3dView::OnBUTTONPanGlo() 
{
	// save the current zoom value 
	m_dCurZoom = m_hView->Scale();
	// Do a Global Zoom 
	//m_hView->FitAll();
	// Set the mode 
	m_CurrentMode = CurAction3d_GlobalPanning;
}

void OCC_3dView::OnBUTTONReset() 
{   m_hView->Reset(); }

void OCC_3dView::OnBUTTONRot() 
{   m_CurrentMode = CurAction3d_DynamicRotation; }


void OCC_3dView::OnBUTTONZoomAll() 
{
	m_hView->FitAll();
	m_hView->ZFitAll();
}

void OCC_3dView::OnBUTTONZoomProg() 
{  m_CurrentMode = CurAction3d_DynamicZooming; }

void OCC_3dView::OnBUTTONZoomWin() 
{  m_CurrentMode = CurAction3d_WindowZooming; }

HRESULT (WINAPI *pfDwmIsCompositionEnabled)(BOOL *pfEnabled);
/******************************************************************************
    @brief		return true if aero is enabled
	@author     humkyung
    @date       2014-08-24
    @class      OCC_3dView
    @function   IsAeroEnabled
    @return     BOOL
******************************************************************************/
BOOL OCC_3dView::IsAeroEnabled()
{
	HMODULE hDwmDLL = LoadLibrary(_T("dwmapi.dll"));	/// Loads the DWM DLL
	if(!hDwmDLL) return FALSE;

	// Everything is fine upto here, we can get function address
	*(FARPROC *)&pfDwmIsCompositionEnabled = GetProcAddress(hDwmDLL,_T("DwmIsCompositionEnabled"));
	BOOL bDwmEnabled = false;
	pfDwmIsCompositionEnabled(&bDwmEnabled);
	return bDwmEnabled;
}

void OCC_3dView::OnLButtonDown(UINT nFlags, CPoint point) 
{
	///  save the current mouse coordinate in min 
	m_iPrevPosX=point.x;  m_iPrevPosY=point.y;
	m_iCurrentPosX=point.x;  m_iCurrentPosY=point.y;

	const BOOL bDwmEnabled = IsAeroEnabled();
	if ( nFlags & MK_CONTROL ) 
	{
		// Button MB1 down Control :start zomming 
		// SetCursor(AfxGetApp()->LoadStandardCursor());
	}
	else // if ( Ctrl )
	{
		switch (m_CurrentMode)
		{
		case CurAction3d_Nothing : /// start a drag
		{
			if (nFlags & MK_SHIFT)
				GetDocument()->ShiftDragEvent(m_iCurrentPosX,m_iCurrentPosY,-1,m_hView);
			else
			{
				GetDocument()->DragEvent(m_iCurrentPosX,m_iCurrentPosY,-1,m_hView);
			}

			if(TRUE == bDwmEnabled)
			{
				CClientDC dc(this); /// device context for painting
				m_tracker.SetContext(dc.m_hDC ,::wglGetCurrentContext() );
				m_tracker.TrackRubberBand(this , point ,TRUE);
			}
		}
		break;
		case CurAction3d_DynamicZooming :	/// noting
			break;
		case CurAction3d_WindowZooming :
		{
			if(TRUE == bDwmEnabled)
			{
				CClientDC dc(this); /// device context for painting
				m_tracker.SetContext(dc.m_hDC ,::wglGetCurrentContext() );
				m_tracker.TrackRubberBand(this , point ,TRUE);
			}
		}
		break;
		case CurAction3d_DynamicPanning :	/// noting
			break;
		case CurAction3d_GlobalPanning :	/// noting
			break;
		case  CurAction3d_DynamicRotation :
		{
			m_hView->StartRotation(point.x , point.y);  
		}
		break;
		case  CurAction3d_RotationAboutAxis:
		{
			if(GetDocument()->GetAISContext()->HasDetected())
			{
				Handle(AIS_InteractiveObject) ais = GetDocument()->GetAISContext()->DetectedInteractive();
				if(!ais.IsNull())
				{
					Handle_AIS_Shape aisShape = Handle_AIS_Shape::DownCast(ais);
					if(!aisShape.IsNull())
					{
						TopoDS_Shape aShape = aisShape->Shape();
						for( TopExp_Explorer oEdgeExplorer( aShape , TopAbs_EDGE ) ; oEdgeExplorer.More() ; oEdgeExplorer.Next() )
						{
							TopoDS_Edge anEdge = TopoDS::Edge( oEdgeExplorer.Current() );
							Standard_Real s0, s1;
							Handle(Geom_Curve) aCurve = BRep_Tool::Curve(anEdge, s0, s1);
							if(!aCurve.IsNull())
							{
								gp_Pnt aStart = aCurve->Value(s0);
								gp_Pnt aEnd = aCurve->Value(s1);
								gp_Vec aDir(aStart , aEnd);
								aDir.Normalize();

								m_hView->SetAxis(aStart.X() , aStart.Y() , aStart.Z() , aDir.X() , aDir.Y() , aDir.Z());
								break;
							}
						}
					}
				}
			}
			else
			{
				m_CurrentMode = CurAction3d_DynamicRotation;
				m_hView->StartRotation(point.x , point.y);
			}
		}
		break;
		default :
			Standard_Failure::Raise(" incompatible Current Mode ");
			break;
		}
	}

	OCC_BaseView::OnLButtonDown(nFlags, point);
}

void OCC_3dView::OnLButtonUp(UINT nFlags, CPoint point) 
{
	BOOL bDwmEnabled = OCC_3dView::IsAeroEnabled();
	if ( nFlags & MK_CONTROL ) 
	{
		return;
	}
	else // if ( Ctrl )
	{
		switch (m_CurrentMode)
		{
		case CurAction3d_Nothing :
			if (point.x == m_iPrevPosX && point.y == m_iPrevPosY)
			{ // no offset between down and up --> selectEvent
				m_iCurrentPosX=point.x;  
				m_iCurrentPosY=point.y;
				if (nFlags & MK_SHIFT )
					GetDocument()->ShiftInputEvent(point.x,point.y,m_hView);
				else
					GetDocument()->InputEvent     (point.x,point.y,m_hView);
			} 
			else
			{
				m_iCurrentPosX=point.x;    m_iCurrentPosY=point.y;
				if(FALSE == bDwmEnabled)
				{
					CClientDC dc(this);
					m_tracker.DrawRectangle(dc , m_iPrevPosX,m_iPrevPosY,m_iCurrentPosX,m_iCurrentPosY,Standard_False);
				}
				if (nFlags & MK_SHIFT)
					GetDocument()->ShiftDragEvent(point.x,point.y,0,m_hView);
				else
					GetDocument()->DragEvent(point.x,point.y,0,m_hView);
			}
			break;
		case CurAction3d_DynamicZooming :
			// SetCursor(AfxGetApp()->LoadStandardCursor());         
			m_CurrentMode = CurAction3d_Nothing;
			break;
		case CurAction3d_WindowZooming :
			m_iCurrentPosX=point.x;        m_iCurrentPosY=point.y;
			if(FALSE == bDwmEnabled)
			{
				CClientDC dc(this);
				m_tracker.DrawRectangle(dc , m_iPrevPosX,m_iPrevPosY,m_iCurrentPosX,m_iCurrentPosY,Standard_False);
			}
			if ((abs(m_iPrevPosX-m_iCurrentPosX)>ValZWMin) || (abs(m_iPrevPosY-m_iCurrentPosY)>ValZWMin))
				// Test if the zoom window is greater than a minimale window.
			{
				// Do the zoom window between Pmin and Pmax
				m_hView->WindowFitAll(m_iPrevPosX,m_iPrevPosY,m_iCurrentPosX,m_iCurrentPosY);  
			}  
			m_CurrentMode = CurAction3d_Nothing;
			break;
		case CurAction3d_DynamicPanning :
			m_CurrentMode = CurAction3d_Nothing;
			break;
		case CurAction3d_GlobalPanning :
			m_hView->Place(point.x,point.y,m_dCurZoom); 
			m_CurrentMode = CurAction3d_Nothing;
			break;
		case  CurAction3d_DynamicRotation :
			m_CurrentMode = CurAction3d_Nothing;
			if (!myDegenerateModeIsOn)
			{  
				CWaitCursor aWaitCursor;
				//m_hView->SetDegenerateModeOff();
				myDegenerateModeIsOn = Standard_False;
			}
			else
			{
				//m_hView->SetDegenerateModeOn();
				myDegenerateModeIsOn = Standard_True;
			}
			break;
		case  CurAction3d_RotationAboutAxis:
		{
		}
		break;
		default :
			Standard_Failure::Raise(" incompatible Current Mode ");
			break;
		} //switch (m_CurrentMode)
	} //	else // if ( Ctrl )
}

/******************************************************************************
    @author     humkyung
    @date       2011-08-25
    @class      OCC_3dView
    @function   OnMButtonDown
    @return     void
    @param      UINT    nFlags
    @param      CPoint  point
    @brief
******************************************************************************/
void OCC_3dView::OnMButtonDown(UINT nFlags, CPoint point)
{
	if ( nFlags & MK_CONTROL ) 
	{
		// Button MB2 down Control : panning init  
		// SetCursor(AfxGetApp()->LoadStandardCursor());   
	}
	
	///  save the current mouse coordinate in min 
	m_iPrevPosX=point.x;  m_iPrevPosY=point.y;
	m_iCurrentPosX=point.x;  m_iCurrentPosY=point.y;

	/// save current mode
	m_PrevMode = m_CurrentMode;

	m_CurrentMode = CurAction3d_DynamicPanning;

	OCC_BaseView::OnMButtonDown(nFlags, point);
}

/******************************************************************************
    @author     humkyung
    @date       2011-08-25
    @class      OCC_3dView
    @function   OnMButtonUp
    @return     void
    @param      UINT    nFlags
    @param      CPoint  point
    @brief
******************************************************************************/
void OCC_3dView::OnMButtonUp(UINT nFlags, CPoint point) 
{
	if ( nFlags & MK_CONTROL ) 
	{
		// Button MB2 down Control : panning init  
		// SetCursor(AfxGetApp()->LoadStandardCursor());   
	}

	//  save the current mouse coordinate in min 
	m_iPrevPosX=point.x;  m_iPrevPosY=point.y;
	m_iCurrentPosX=point.x;  m_iCurrentPosY=point.y;

	m_CurrentMode = m_PrevMode;

	OCC_BaseView::OnMButtonDown(nFlags, point);
}

void OCC_3dView::OnRButtonDown(UINT nFlags, CPoint point) 
{
	if ( nFlags & MK_CONTROL ) 
	{
		// SetCursor(AfxGetApp()->LoadStandardCursor());   
		/*if (!myDegenerateModeIsOn)
			m_hView->SetDegenerateModeOn();*/
		m_hView->StartRotation(point.x,point.y);  
	}
	else // if ( Ctrl )
	{
		GetDocument()->Popup(point.x,point.y,m_hView);
	}	
}

void OCC_3dView::OnRButtonUp(UINT nFlags, CPoint point) 
{
	SetCursor(AfxGetApp()->LoadStandardCursor(IDC_WAIT));
	if (!myDegenerateModeIsOn)
	{  
		//m_hView->SetDegenerateModeOff();
		myDegenerateModeIsOn = Standard_False;
	} else
	{
		//m_hView->SetDegenerateModeOn();
		myDegenerateModeIsOn = Standard_True;
	}
	SetCursor(AfxGetApp()->LoadStandardCursor(IDC_ARROW));
}

void OCC_3dView::OnMouseMove(UINT nFlags, CPoint point) 
{
	BOOL bDwmEnabled = OCC_3dView::IsAeroEnabled();

	//   ============================  LEFT BUTTON =======================
	if ( nFlags & MK_LBUTTON)
	{
		if ( nFlags & MK_CONTROL ) 
		{
			// move with MB1 and Control : on the dynamic zooming  
			// Do the zoom in function of mouse's coordinates  
			m_hView->Zoom(m_iCurrentPosX,m_iCurrentPosY,point.x,point.y); 
			// save the current mouse coordinate in min 
			m_iCurrentPosX = point.x; 
			m_iCurrentPosY = point.y;	
		}
		else // if ( Ctrl )
		{
			switch (m_CurrentMode)
			{
			case CurAction3d_Nothing :
			{
				CClientDC dc(this);
				m_iCurrentPosX = point.x;  m_iCurrentPosY = point.y;
				if(FALSE == bDwmEnabled)
				{
					m_tracker.DrawRectangle(dc , m_iPrevPosX,m_iPrevPosY,m_iCurrentPosX,m_iCurrentPosY,Standard_False);
				}
				/*if (nFlags & MK_SHIFT)		
					GetDocument()->ShiftDragEvent(m_iCurrentPosX,m_iCurrentPosY,0,m_hView);
				else
					GetDocument()->DragEvent(m_iCurrentPosX,m_iCurrentPosY,0,m_hView);*/
				if(FALSE == bDwmEnabled)
				{
					m_tracker.DrawRectangle(dc , m_iPrevPosX,m_iPrevPosY,m_iCurrentPosX,m_iCurrentPosY,Standard_True);
				}
			}
				break;
			case CurAction3d_DynamicZooming :
				m_hView->Zoom(m_iCurrentPosX,m_iCurrentPosY,point.x,point.y); 
				// save the current mouse coordinate in min \n";
				m_iCurrentPosX=point.x;  m_iCurrentPosY=point.y;
				break;
			case CurAction3d_WindowZooming :
				m_iCurrentPosX = point.x; m_iCurrentPosY = point.y;	
				if(FALSE == bDwmEnabled)
				{
					CClientDC dc(this);
					m_tracker.DrawRectangle(dc , m_iPrevPosX,m_iPrevPosY,m_iCurrentPosX,m_iCurrentPosY,Standard_False);
					m_tracker.DrawRectangle(dc , m_iPrevPosX,m_iPrevPosY,m_iCurrentPosX,m_iCurrentPosY,Standard_True);
				}
				break;
			case CurAction3d_DynamicPanning :
				if ( nFlags & MK_MBUTTON)
				{
					m_hView->Pan(point.x - m_iCurrentPosX , m_iCurrentPosY - point.y); // Realize the panning
					m_iCurrentPosX = point.x; m_iCurrentPosY = point.y;
				}
				break;
				///m_hView->Pan(point.x-m_iCurrentPosX,m_iCurrentPosY-point.y); // Realize the panning
				///m_iCurrentPosX = point.x; m_iCurrentPosY = point.y;	
				break;
			case CurAction3d_GlobalPanning : // nothing           
				break;
			case  CurAction3d_DynamicRotation :
			{
				m_hView->Rotation(point.x,point.y);
			}
			break;
			case  CurAction3d_RotationAboutAxis:
			{
				CRect rect;
				GetWindowRect(&rect);
				const double dR = max(rect.Width() , rect.Height());

				const int dx = (m_iCurrentPosX - point.x);
				const int dy = (m_iCurrentPosX - point.y);
				const double dLen = ::sqrt(double(dx*dx) + double(dy*dy));
				const double dRad = asin(dLen/2 / dR);

				m_hView->Rotate((dx >= 0) ? dRad : -dRad ,TRUE);

				m_iCurrentPosX = point.x;
				m_iCurrentPosY = point.y;
			}
			break;
			default :
				Standard_Failure::Raise(" incompatible Current Mode ");
				break;
			}//  switch (m_CurrentMode)
		}// if ( nFlags & MK_CONTROL )  else 
	} 
	else //   if ( nFlags & MK_LBUTTON) 
	{
		//   ============================  MIDDLE BUTTON =======================
		if ( nFlags & MK_MBUTTON)
		{
			if ( nFlags & MK_CONTROL ) 
			{
				m_hView->Pan(point.x-m_iCurrentPosX,m_iCurrentPosY-point.y); // Realize the panning
				m_iCurrentPosX = point.x; m_iCurrentPosY = point.y;	
			}
			else
			{
				m_hView->Pan(point.x-m_iCurrentPosX,m_iCurrentPosY-point.y); // Realize the panning
				m_iCurrentPosX = point.x; m_iCurrentPosY = point.y;	
			}
		} else //  if ( nFlags & MK_MBUTTON)
			//   ============================  RIGHT BUTTON =======================
			if ( nFlags & MK_RBUTTON)
			{
				if ( nFlags & MK_CONTROL ) 
				{
					m_hView->Rotation(point.x,point.y);
				}
			}
			else
			///   ============================  NO BUTTON =======================
			{
				m_iCurrentPosX = point.x; m_iCurrentPosY = point.y;	
				if (nFlags & MK_SHIFT)
					GetDocument()->ShiftMoveEvent(point.x,point.y,m_hView);
				else
					GetDocument()->MoveEvent(point.x,point.y,m_hView);
			}
	}
}

void OCC_3dView::OnUpdateBUTTONHlrOff(CCmdUI* pCmdUI) 
{
	pCmdUI->SetCheck (myDegenerateModeIsOn);
	pCmdUI->Enable   (!myDegenerateModeIsOn);	
}

void OCC_3dView::OnUpdateBUTTONHlrOn(CCmdUI* pCmdUI) 
{
	pCmdUI->SetCheck (!myDegenerateModeIsOn);
	pCmdUI->Enable   (myDegenerateModeIsOn);	
}

void OCC_3dView::OnUpdateBUTTONPanGlo(CCmdUI* pCmdUI) 
{
	pCmdUI->SetCheck (m_CurrentMode == CurAction3d_GlobalPanning);
	pCmdUI->Enable   (m_CurrentMode != CurAction3d_GlobalPanning);	

}

void OCC_3dView::OnUpdateBUTTONPan(CCmdUI* pCmdUI) 
{
	pCmdUI->SetCheck (m_CurrentMode == CurAction3d_DynamicPanning);
	pCmdUI->Enable   (m_CurrentMode != CurAction3d_DynamicPanning );	
}

void OCC_3dView::OnUpdateBUTTONZoomProg(CCmdUI* pCmdUI) 
{
	pCmdUI->SetCheck (m_CurrentMode == CurAction3d_DynamicZooming );
	pCmdUI->Enable   (m_CurrentMode != CurAction3d_DynamicZooming);	
}

void OCC_3dView::OnUpdateBUTTONZoomWin(CCmdUI* pCmdUI) 
{
	pCmdUI->SetCheck (m_CurrentMode == CurAction3d_WindowZooming);
	pCmdUI->Enable   (m_CurrentMode != CurAction3d_WindowZooming);	
}

void OCC_3dView::OnUpdateBUTTONRot(CCmdUI* pCmdUI) 
{
	pCmdUI->SetCheck (m_CurrentMode == CurAction3d_DynamicRotation);
	pCmdUI->Enable   (m_CurrentMode != CurAction3d_DynamicRotation);	
}

void OCC_3dView::OnModifyChangeBackground() 
{
	Standard_Real R1;
	Standard_Real G1;
	Standard_Real B1;
	m_hView->BackgroundColor(Quantity_TOC_RGB,R1,G1,B1);
	COLORREF m_clr ;
	m_clr = RGB(R1*255,G1*255,B1*255);

	CColorDialog dlgColor(m_clr);
	if (dlgColor.DoModal() == IDOK)
	{
		m_clr = dlgColor.GetColor();
		R1 = GetRValue(m_clr)/255.;
		G1 = GetGValue(m_clr)/255.;
		B1 = GetBValue(m_clr)/255.;
		m_hView->SetBackgroundColor(Quantity_TOC_RGB,R1,G1,B1);
	}
	m_hView->Redraw();
}

int OCC_3dView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (OCC_BaseView::OnCreate(lpCreateStruct) == -1)
		return -1;

	return 0;
}

static Standard_Integer cZoomStep = 20;
BOOL OCC_3dView::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	// move with MB1 and Control : on the dynamic zooming  
	// Do the zoom in function of mouse's coordinates  

	m_dCurZoom = m_hView->Scale();

	Standard_Real cXmax = pt.x + zDelta/cZoomStep; 
	Standard_Real cYmax = pt.y + zDelta/cZoomStep;

	if (cYmax < 0)    cYmax = 0;
	if (cXmax < 0)    cXmax = 0;

	m_hView->Zoom(cXmax,cYmax,pt.x,pt.y);

	return TRUE;

	//     return CWnd::OnMouseWheel(nFlags, zDelta, pt);
}

/**
	@brief	do nothing
	@author	HumKyung
	@date	????.??.??
*/
BOOL OCC_3dView::OnEraseBkgnd(CDC* pDC) 
{
	return TRUE;
}

int OCC_3dView::DrawGradientBackground(void)
{
	if (!m_hView.IsNull())
	{
		if (m_hGradientBkgndLayer.IsNull())
		{
			Standard_Boolean aSizeDependant = Standard_True; //each window to have particular mapped layer?
			m_hGradientBkgndLayer = new Visual3d_Layer(m_hView->Viewer()->Viewer() , Aspect_TOL_UNDERLAY, aSizeDependant); 
		}
		m_hGradientBkgndLayer->Clear(); //! make sure we draw on a clean layer
		m_hGradientBkgndLayer->Begin();

		Standard_Integer w, h;
		Handle(Aspect_Window) hWin = m_hView->Window();
		hWin->Size(w, h);

		m_hGradientBkgndLayer->SetViewport (w ? w : 1 , h ? h : 1);
		m_hGradientBkgndLayer->SetOrtho(0 , 10 , 0 , 10 , Aspect_TOC_TOP_LEFT);

		double left = 0,bottom = 0;
		double right=10,top = 10 , delta = 0.0;
		const double dAspectRatio = float(w) / float(h);
		if (dAspectRatio >= 1.0) 
		{ /* fenetre horizontale */
			delta = (float )((top - bottom)/2.0);
			/* Aspect_TOC_TOP_LEFT */
			bottom = top - 2*delta/dAspectRatio;			
		}
		else
		{ /* fenetre verticale */
			delta = (float )((right - left)/2.0);
			/* Aspect_TOC_TOP_LEFT */
			right = left + 2*delta*dAspectRatio;	
		}

		m_hGradientBkgndLayer->BeginPolygon();
		m_hGradientBkgndLayer->SetColor(Quantity_NOC_BLACK);
		m_hGradientBkgndLayer->AddVertex (left,top);
		m_hGradientBkgndLayer->AddVertex (right,top);
		m_hGradientBkgndLayer->SetColor (Quantity_NOC_MATRAGRAY);
		m_hGradientBkgndLayer->AddVertex (right,bottom);
		m_hGradientBkgndLayer->AddVertex (left,bottom);
		m_hGradientBkgndLayer->ClosePrimitive();

		m_hGradientBkgndLayer->End();
	}

	return ERROR_SUCCESS;
}

int OCC_3dView::DrawColorBar(void)
{	
	if (m_hView.IsNull()) return ERROR_SUCCESS;
		
	//!
	CRect rect; 
	GetClientRect(&rect); 

	if (m_hColorBarLayer.IsNull())
	{
		Standard_Boolean aSizeDependant = Standard_True; /// each window to have particular mapped layer?
		m_hColorBarLayer = new Visual3d_Layer(m_hView->Viewer()->Viewer(), Aspect_TOL_OVERLAY, aSizeDependant); 
	}
	
	Standard_Integer w, h;
	Handle(Aspect_Window) hWin = m_hView->Window();
	hWin->Size(w, h);
	if (w < h)
	{
		rect.right  = LONG(h ? 10. * w / h : 10.);
		rect.bottom = 0;
	}
	else
	{
		rect.right  = 10;
		rect.bottom = LONG(w ? 10. * (1. - static_cast<double>(h) / w) : 0.);
	}

	m_hColorBarLayer->SetViewport(w ? w : 1 , h ? h : 1);
	m_hColorBarLayer->SetOrtho(0 , 10 , 0 , 10 , Aspect_TOC_TOP_LEFT);
	
	double left = 0,bottom = 0;
	double right=10,top = 10 , delta = 0.0;
	//! aspect ratio 적용해야 한다.
	const double dAspectRatio = float(w) / float(h);

	m_hColorBarLayer->Clear(); //make sure we draw on a clean layer
	m_hColorBarLayer->Begin();
	if (dAspectRatio >= 1.0) 
	{ /* fenetre horizontale */
		delta = (float )((top - bottom)/2.0);
		/* Aspect_TOC_TOP_LEFT */
		bottom = top - 2*delta/dAspectRatio;			
	}
	else
	{ /* fenetre verticale */
		delta = (float )((right - left)/2.0);
		/* Aspect_TOC_TOP_LEFT */
		right = left + 2*delta*dAspectRatio;	
	}

	glPushAttrib (GL_LIST_BIT);
	///glListBase (m_nDefaultBitmapFontID);
	TCHAR buf[20] = {'\0' ,};
	Quantity_Color color(1.0 , 1.0 , 0.0 , Quantity_TOC_RGB);

	//! pixel 단위의 colorbar의 min,max를 구한다.
	double minx = right - (right - left)*70/w;
	double maxx = right - (right - left)*50/w;
	double maxy = top - (top - bottom)*20/h;
	double miny = top - (top - bottom)*300/h;

	const double dStep = (maxy - miny) / 7.0;
	for(int i = 0;i < 7;++i)
	{
		glBegin( GL_QUADS );
		glNormal3f( 0.f , 0.f , 1.f );
		glColor3f(
			(float)((255 - i*255.0/7.0) / 255.f),
			(float)((255 - i*255.0/7.0) / 255.f),
			(float)((255 - i*255.0/7.0) / 255.f)
			);


		glVertex3d (minx , maxy - i*dStep , 0.f);
		glVertex3d (maxx , maxy - i*dStep , 0.f);
		glVertex3d (maxx , maxy - (i+1)*dStep , 0.f);
		glVertex3d (minx , maxy - (i+1)*dStep , 0.f);
		glEnd();

		//! print text
		m_hColorBarLayer->SetColor(color);
		m_hColorBarLayer->SetTextAttributes(/*Graphic3d_NOF_ASCII_MONO*/"Courier", Aspect_TODT_NORMAL, color);                              
		SPRINTF_T (buf, _T("%.2lf") , i * 10.0);
		m_hColorBarLayer->DrawText(CStringA(buf) , maxx , maxy - i*dStep, 14);
	}
	glPopAttrib ();

	m_hColorBarLayer->End();

	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-12-04
    @class      OCC_3dView
    @function   ZoomWin
    @return     void
    @param      const   Bnd_Box&
    @param      oBndBox
    @brief		zoom around bouding box
******************************************************************************/
void OCC_3dView::ZoomWin(const Bnd_Box& oBndBox)
{
	double aXmin = 0.0 , aYmin = 0.0 , aZmin = 0.0 , aXmax = 0.0 , aYmax = 0.0 , aZmax = 0.0;
	oBndBox.Get(aXmin , aYmin , aZmin , aXmax , aYmax , aZmax);
	if(!m_hView.IsNull())
	{
		int iXmin = 0 , iYmin = 0 , iXmax = 0 , iYmax = 0;
		m_hView->Convert(aXmin , aYmin , aZmin , iXmin , iYmin);
		m_hView->Convert(aXmax , aYmax , aZmax , iXmax , iYmax);
		
		if(iXmin > iXmax) swap(iXmin , iXmax);
		if(iYmin > iYmax) swap(iYmin , iYmax);
		const int width = (iXmax - iXmin);
		const int height= (iYmax - iYmin);
		m_hView->WindowFitAll(iXmin, iYmin, iXmax, iYmax);
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-12-09
    @class      OCC_3dView
    @function   SetBackgroundColor
    @return     int
    @param      const   unsigned
    @param      char&   r
    @param      const   unsigned
    @param      char&   g
    @param      const   unsigned
    @param      char&   b
    @brief
******************************************************************************/
int OCC_3dView::SetBackgroundColor(const unsigned char& r , const unsigned char& g , const unsigned char& b)
{
	const Quantity_Color aColor = Quantity_Color (double(r) / 255. , double(g) /255. , double(b) /255. , Quantity_TOC_RGB);
	m_hView->SetBackgroundColor(aColor);

	return ERROR_SUCCESS;
}

//==================================================================================
// Function name	: OCC_3dView::Convert2dPntTo3dPnt
//==================================================================================
// Written by	    : Stephane Routelous - 2001-11-08 19:02:49
// Description	    : 
// Return type		: Standard_Boolean 
//==================================================================================
// Argument         : const Standard_Integer aX2d
// Argument         : const Standard_Integer aY2d
// Argument         : gp_Pnt& a3dPoint
Standard_Boolean OCC_3dView::Convert2dPntTo3dPnt(const Standard_Integer aX2d, const Standard_Integer aY2d, gp_Pnt& a3dPoint)
{
	if (m_hView.IsNull())
		return Standard_False;
	
	// get the eye and the target points
	V3d_Coordinate theXEye, theYEye, theZEye, theXAt, theYAt, theZAt;
	m_hView->Eye(theXEye, theYEye, theZEye);
	m_hView->At(theXAt, theYAt, theZAt);
	gp_Pnt theEyePoint(theXEye, theYEye, theZEye);
	gp_Pnt theAtPoint(theXAt, theYAt, theZAt);
	
	// create the direction
	gp_Vec theEyeVector(theEyePoint, theAtPoint);
	gp_Dir theEyeDir(theEyeVector);
	
	// make a plane perpendicular to this direction
	gp_Pln thePlaneOfTheView = gp_Pln(theAtPoint, theEyeDir);
	
	// convert the 2d point into 3d
	Standard_Real theX, theY, theZ;
	m_hView->Convert(aX2d, aY2d, theX, theY, theZ);
	gp_Pnt theConvertedPoint(theX, theY, theZ);
	
	// project the converted point to the plane
	gp_Pnt2d theConvertedPointOnPlane = ProjLib::Project(thePlaneOfTheView, theConvertedPoint);
	
	// get the 3d point of this 2d point
	gp_Pnt theResultPoint = ElSLib::Value(theConvertedPointOnPlane.X(),	theConvertedPointOnPlane.Y(), thePlaneOfTheView);
	a3dPoint = theResultPoint;

	return Standard_True;
}